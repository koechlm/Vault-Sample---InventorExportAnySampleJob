/*=====================================================================
  
  This file is an educational sample working with Autodesk Vault API.

  Copyright (C) Autodesk Inc.  All rights reserved.

THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
PARTICULAR PURPOSE.
=====================================================================*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.IO;
using Autodesk.Connectivity.Extensibility.Framework;
using ACET = Autodesk.Connectivity.Explorer.ExtensibilityTools;
using VDF = Autodesk.DataManagement.Client.Framework;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections;
using Autodesk.DataManagement.Client.Framework.Vault.Currency.Properties;
using Autodesk.DataManagement.Client.Framework.Vault.Settings;
using Autodesk.DataManagement.Client.Framework.Currency;

using Autodesk.Connectivity.JobProcessor.Extensibility;
using ACW = Autodesk.Connectivity.WebServices;
using Inventor;

[assembly: ApiVersion("15.0")]
[assembly: ExtensionId("385b4efe-36be-485d-a533-1e84369d1bea")]


namespace Autodesk.VltInvSrv.ExportSampleJob
{
    public class JobExtension : IJobHandler
    {
        private static string JOB_TYPE = "Autodesk.VltInvSrv.ExportSampleJob";
        private static Settings mSettings = Settings.Load();
        private static string mLogDir = JobExtension.mSettings.LogFileLocation;
        private static string mLogFile = JOB_TYPE + ".log";
        private TextWriterTraceListener mTrace = new TextWriterTraceListener(System.IO.Path.Combine(
            mLogDir, mLogFile), "mJobTrace");

        #region IJobHandler Implementation
        public bool CanProcess(string jobType)
        {
            return jobType == JOB_TYPE;
        }

        public JobOutcome Execute(IJobProcessorServices context, IJob job)
        {
            try
            {
                FileInfo mLogFileInfo = new FileInfo(System.IO.Path.Combine(
                    mLogDir, mLogFile));
                if (mLogFileInfo.Exists) mLogFileInfo.Delete();
                mTrace.WriteLine("Starting Job...");

                //start step export
                mCreateExport(context, job);

                mTrace.IndentLevel = 0;
                mTrace.WriteLine("... successfully ending Job.");
                mTrace.Flush();
                mTrace.Close();

                return JobOutcome.Success;
            }
            catch (Exception ex)
            {
                context.Log(ex, "Autodesk.VltInvSrv.ExportSampleJob failed: " + ex.ToString() + " ");

                mTrace.IndentLevel = 0;
                mTrace.WriteLine("... ending Job with failures.");
                mTrace.Flush();
                mTrace.Close();

                return JobOutcome.Failure;
            }

        }

        public void OnJobProcessorShutdown(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }

        public void OnJobProcessorSleep(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }

        public void OnJobProcessorStartup(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }

        public void OnJobProcessorWake(IJobProcessorServices context)
        {
            //throw new NotImplementedException();
        }
        #endregion IJobHandler Implementation

        #region Job Execution
        private void mCreateExport(IJobProcessorServices context, IJob job)
        {
            List<string> mExpFrmts = new List<string>();
            List<string> mUploadFiles = new List<string>();

            // read target export formats from settings file
            Settings settings = Settings.Load();

            #region validate execution rules

            mTrace.IndentLevel += 1;
            mTrace.WriteLine("Translator Job started...");

            //pick up this job's context
            Connection connection = context.Connection;
            Autodesk.Connectivity.WebServicesTools.WebServiceManager mWsMgr = connection.WebServiceManager;
            long mEntId = Convert.ToInt64(job.Params["EntityId"]);
            string mEntClsId = job.Params["EntityClassId"];

            // only run the job for files
            if (mEntClsId != "FILE")
                return;

            // only run the job for source file types, supported by exports (as of today)
            List<string> mFileExtensions = new List<string> { ".ipt", ".iam", ".idw"}; //ipn is not supported by InventorServer
            ACW.File mFile = mWsMgr.DocumentService.GetFileById(mEntId);
            if (!mFileExtensions.Any(n => mFile.Name.ToLower().EndsWith(n)))
            {
                mTrace.WriteLine("Translator job exits: file extension is not supported.");
                return;
            }

            // apply execution filters, e.g., exclude files of classification "substitute" etc.
            List<string> mFileClassific = new List<string> { "ConfigurationFactory", "DesignSubstitute" }; //add "DesignDocumentation" for 3D Exporters only
            if (mFileClassific.Any(n => mFile.FileClass.ToString().Contains(n)))
            {
                mTrace.WriteLine("Translator job exits: file classification 'ConfigurationFactory' or 'DesignSubstitute' are not supported.");
                return;
            }

            // you may add addtional execution filters, e.g., category name == "Sheet Metal Part"

            //only run the job for implemented/available combinations of source file and export file formats
            List<string> mIptExpFrmts = new List<string> { "STP", "JT", "SMDXF", "3DDWG", "IMAGE"};
            List<string> mIamExpFrmts = new List<string> { "STP", "JT", "3DDWG", "IMAGE"};
            List<string> mIdwExpFrmts = new List<string> { "2DDWG", "IMAGE" };

            if (settings.ExportFormats == null)
                throw new Exception("Settings expect to list at least one export format!");
            if (settings.ExportFormats.Contains(","))
            {
                mExpFrmts = settings.ExportFormats.Replace(" ", "").Split(',').ToList();
            }
            else
            {
                mExpFrmts.Add(settings.ExportFormats);
            }

            //filter export formats according source file type
            if (mFile.Name.ToLower().EndsWith(".ipt"))
            {
                if (!mExpFrmts.Any(n => mIptExpFrmts.Contains(n)))
                {
                    mTrace.WriteLine("Translator job exits: no export file format matches the source file type IPT.");
                    return;
                }

                //remove SM formats, if source isn't sheet metal
                if (mFile.Cat.CatName != settings.SmCatDispName)
                {
                    if (mExpFrmts.Contains("SMDXF")) mExpFrmts.Remove("SMDXF");
                    //if (mExpFrmts.Contains("SMSAT")) mExpFrmts.Remove("SMSAT"); sat is not implemented as of today
                }
                //remove 2D formats
                if (mExpFrmts.Contains("2DDWG")) mExpFrmts.Remove("2DDWG");
            }
            if (mFile.Name.ToLower().EndsWith(".iam"))
            {
                if (!mExpFrmts.Any(n => mIamExpFrmts.Contains(n)))
                {
                    mTrace.WriteLine("Translator job exits: no export file format matches the source file type IAM.");
                    return;
                }
                //remove sheet metal and 2D formats
                if (mExpFrmts.Contains("SMDXF")) mExpFrmts.Remove("SMDXF");
                if (mExpFrmts.Contains("2DDWG")) mExpFrmts.Remove("2DDWG");
            }
            if (mFile.Name.ToLower().EndsWith(".idw") || mFile.Name.ToLower().EndsWith(".dwg"))
            {
                if (!mExpFrmts.Any(n => mIdwExpFrmts.Contains(n)))
                {
                    mTrace.WriteLine("Translator job exits: no export file format matches the source file type IDW/DWG.");
                    return;
                }
                //remove 3D formats, if source is 2D
                if (mExpFrmts.Contains("STP")) mExpFrmts.Remove("STP");
                if (mExpFrmts.Contains("JT")) mExpFrmts.Remove("JT");
                if (mExpFrmts.Contains("SMDXF")) mExpFrmts.Remove("SMDXF");
                if (mExpFrmts.Contains("3DDWG")) mExpFrmts.Remove("3DDWG");
            }

            //if (mFile.Name.ToLower().EndsWith(".ipn")) //note - requires Inventor application instead of InventorServer
            //{
            //    string[] mTmpList = new string[mExpFrmts.Count];
            //    mTmpList = mExpFrmts.ToArray();
            //    foreach (string mFrmt in mTmpList.ToList())
            //    {
            //        if (mFrmt != "IMAGE") mExpFrmts.Remove(mFrmt);
            //    }
            //}

            //validate that at least one export format is in the list
            if (mExpFrmts.Count < 1)
            {
                mTrace.WriteLine("Translator job exits: no matching source file type/export type found.");
                return;
            }

            mTrace.WriteLine("Job execution rules validated.");

            #endregion validate execution rules

            #region VaultInventorServer IPJ activation
            //establish InventorServer environment including translator addins; differentiate her in case full Inventor.exe is used
            Inventor.InventorServer mInv = context.InventorObject as InventorServer;
            ApplicationAddIns mInvSrvAddIns = mInv.ApplicationAddIns;

            //override InventorServer default project settings by your Vault specific ones
            Inventor.DesignProjectManager projectManager;
            Inventor.DesignProject mSaveProject = null, mProject = null;
            String mIpjPath = "";
            String mWfPath = "";
            String mIpjLocalPath = "";
            ACW.File mProjFile;
            VDF.Vault.Currency.Entities.FileIteration mIpjFileIter = null;

            //download and activate the Inventor Project file in VaultInventorServer
            mTrace.IndentLevel += 1;
            mTrace.WriteLine("Job tries activating Inventor project file as enforced in Vault behavior configurations.");
            try
            {
                //Download enforced ipj file
                if (mWsMgr.DocumentService.GetEnforceWorkingFolder() && mWsMgr.DocumentService.GetEnforceInventorProjectFile())
                {
                    mIpjPath = mWsMgr.DocumentService.GetInventorProjectFileLocation();
                    mWfPath = mWsMgr.DocumentService.GetRequiredWorkingFolderLocation();
                }
                else
                {
                    throw new Exception("Job requires both settings enabled: 'Enforce Workingfolder' and 'Enforce Inventor Project'.");
                }

                String[] mIpjFullFileName = mIpjPath.Split(new string[] { "/" }, StringSplitOptions.None);
                String mIpjFileName = mIpjFullFileName.LastOrDefault();

                //get the projects file object for download
                ACW.PropDef[] filePropDefs = mWsMgr.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                ACW.PropDef mNamePropDef = filePropDefs.Single(n => n.SysName == "ClientFileName");
                ACW.SrchCond mSrchCond = new ACW.SrchCond()
                {
                    PropDefId = mNamePropDef.Id,
                    PropTyp = ACW.PropertySearchType.SingleProperty,
                    SrchOper = 3, // is equal
                    SrchRule = ACW.SearchRuleType.Must,
                    SrchTxt = mIpjFileName
                };
                string bookmark = string.Empty;
                ACW.SrchStatus status = null;
                List<ACW.File> totalResults = new List<ACW.File>();
                while (status == null || totalResults.Count < status.TotalHits)
                {
                    ACW.File[] results = mWsMgr.DocumentService.FindFilesBySearchConditions(new ACW.SrchCond[] { mSrchCond },
                        null, null, false, true, ref bookmark, out status);
                    if (results != null)
                        totalResults.AddRange(results);
                    else
                        break;
                }
                if (totalResults.Count == 1)
                {
                    mProjFile = totalResults[0];
                }
                else
                {
                    throw new Exception("Job execution stopped due to ambigous project file definitions; single project file per Vault expected");
                }

                //define download settings for the project file
                VDF.Vault.Settings.AcquireFilesSettings mDownloadSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection);
                mDownloadSettings.LocalPath = new VDF.Currency.FolderPathAbsolute(mWfPath);
                mIpjFileIter = new VDF.Vault.Currency.Entities.FileIteration(connection, mProjFile);
                mDownloadSettings.AddFileToAcquire(mIpjFileIter, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);

                //download project file and get local path
                VDF.Vault.Results.AcquireFilesResults mDownLoadResult;
                VDF.Vault.Results.FileAcquisitionResult fileAcquisitionResult;
                mDownLoadResult = connection.FileManager.AcquireFiles(mDownloadSettings);
                fileAcquisitionResult = mDownLoadResult.FileResults.FirstOrDefault();
                mIpjLocalPath = fileAcquisitionResult.LocalPath.FullPath;

                //activate the given project file for this job only
                projectManager = mInv.DesignProjectManager;
                //VaultInventorServer might fail with unhandled exeption on fresh installed machines, if no IPJ had been used before
                try
                {
                    if (projectManager.ActiveDesignProject != null && projectManager.ActiveDesignProject.FullFileName != mIpjLocalPath)
                    {
                        mSaveProject = projectManager.ActiveDesignProject;
                    }
                }
                catch (Exception)
                { }
                mProject = projectManager.DesignProjects.AddExisting(mIpjLocalPath);
                mProject.Activate();

                //[Optionally:] get Inventor Design Data settings and download all related files ---------

                mTrace.WriteLine("Job successfully activated Inventor IPJ.");
            }
            catch (Exception ex)
            {
                throw new Exception("Job was not able to activate Inventor project file. - Note: The ipj must not be checked out by another user.", ex.InnerException);
            }
            #endregion VaultInventorServer IPJ activation

            #region download source file(s)
            mTrace.IndentLevel += 1;
            mTrace.WriteLine("Job downloads source file(s) for translation.");
            //download the source file iteration, enforcing overwrite if local files exist
            VDF.Vault.Settings.AcquireFilesSettings mDownloadSettings2 = new VDF.Vault.Settings.AcquireFilesSettings(connection);
            VDF.Vault.Currency.Entities.FileIteration mFileIteration = new VDF.Vault.Currency.Entities.FileIteration(connection, mFile);
            mDownloadSettings2.AddFileToAcquire(mFileIteration, VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Download);
            mDownloadSettings2.OrganizeFilesRelativeToCommonVaultRoot = true;
            mDownloadSettings2.OptionsRelationshipGathering.FileRelationshipSettings.IncludeChildren = true;
            mDownloadSettings2.OptionsRelationshipGathering.FileRelationshipSettings.RecurseChildren = true;
            mDownloadSettings2.OptionsRelationshipGathering.FileRelationshipSettings.IncludeLibraryContents = true;
            mDownloadSettings2.OptionsRelationshipGathering.FileRelationshipSettings.ReleaseBiased = true;
            VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions mResOpt = new VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions();
            mResOpt.OverwriteOption = VDF.Vault.Settings.AcquireFilesSettings.AcquireFileResolutionOptions.OverwriteOptions.ForceOverwriteAll;
            mResOpt.SyncWithRemoteSiteSetting = VDF.Vault.Settings.AcquireFilesSettings.SyncWithRemoteSite.Always;

            //execute download
            VDF.Vault.Results.AcquireFilesResults mDownLoadResult2 = connection.FileManager.AcquireFiles(mDownloadSettings2);
            //pickup result details
            VDF.Vault.Results.FileAcquisitionResult fileAcquisitionResult2 = mDownLoadResult2.FileResults.Where(n => n.File.EntityName == mFileIteration.EntityName).FirstOrDefault();

            if (fileAcquisitionResult2 == null)
            {
                mResetIpj(mSaveProject);
                throw new Exception("Job stopped execution as the source file to translate did not download");
            }
            string mDocPath = fileAcquisitionResult2.LocalPath.FullPath;
            string mExt = System.IO.Path.GetExtension(mDocPath); //mDocPath.Split('.').Last();
            mTrace.WriteLine("Job successfully downloaded source file(s) for translation.");
            #endregion download source file(s)

            #region VaultInventorServer CAD Export

            mTrace.WriteLine("Job opens source file.");
            Document mDoc = null;
            mDoc = mInv.Documents.Open(mDocPath);

            //use the matching export addin and export options
            foreach (string item in mExpFrmts)
            {
                if (item == "STP")
                {
                    //activate STEP Translator environment,
                    try
                    {
                        TranslatorAddIn mStepTrans = mInvSrvAddIns.ItemById["{90AF7F40-0C01-11D5-8E83-0010B541CD80}"] as TranslatorAddIn;
                        if (mStepTrans == null)
                        {
                            //switch temporarily used project file back to original one
                            mResetIpj(mSaveProject);
                            throw new Exception("Job stopped execution, because indicated translator addin is not available.");
                        }
                        TranslationContext mTransContext = mInv.TransientObjects.CreateTranslationContext();
                        NameValueMap mTransOptions = mInv.TransientObjects.CreateNameValueMap();
                        if (mStepTrans.HasSaveCopyAsOptions[mDoc, mTransContext, mTransOptions] == true)
                        {
                            //open, and translate the source file
                            mTrace.IndentLevel += 1;
                            mTrace.WriteLine("Job opens source file.");
                            mTransOptions.Value["ApplicationProtocolType"] = 3; //AP 2014, Automotive Design
                            mTransOptions.Value["Description"] = "Sample-Job Step Translator using VaultInventorServer";
                            mTransContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
                            //delete local file if exists, as the export wouldn't overwrite
                            string mExpFileName = mDocPath + ".stp";
                            if (System.IO.File.Exists(mExpFileName))
                            {
                                System.IO.File.SetAttributes(mExpFileName, System.IO.FileAttributes.Normal);
                                System.IO.File.Delete(mExpFileName);
                            };
                            DataMedium mData = mInv.TransientObjects.CreateDataMedium();
                            mData.FileName = mExpFileName;
                            mStepTrans.SaveCopyAs(mDoc, mTransContext, mTransOptions, mData);
                            //collect all export files for later upload
                            mUploadFiles.Add(mExpFileName);
                            System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(mUploadFiles.LastOrDefault());
                            if (mExportFileInfo.Exists)
                            {
                                mTrace.WriteLine("STEP Translator created file: " + mUploadFiles.LastOrDefault());
                                mTrace.IndentLevel -= 1;
                            }
                            else
                            {
                                mResetIpj(mSaveProject);
                                throw new Exception("Validating the export file before upload failed.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        mResetIpj(mSaveProject);
                        mTrace.WriteLine("STEP Export Failed: " + ex.Message);
                    }
                }

                if (item == "JT")
                {
                    //activate JT Translator environment,
                    try
                    {
                        TranslatorAddIn mJtTrans = mInvSrvAddIns.ItemById["{16625A0E-F58C-4488-A969-E7EC4F99CACD}"] as TranslatorAddIn;
                        if (mJtTrans == null)
                        {
                            //switch temporarily used project file back to original one
                            mTrace.WriteLine("JT Translator not found.");
                            break;
                        }
                        TranslationContext mTransContext = mInv.TransientObjects.CreateTranslationContext();
                        NameValueMap mTransOptions = mInv.TransientObjects.CreateNameValueMap();
                        if (mJtTrans.HasSaveCopyAsOptions[mDoc, mTransContext, mTransOptions] == true)
                        {
                            //open, and translate the source file
                            mTrace.IndentLevel += 1;

                            mTransOptions.Value["Version"] = 102; //default
                            mTransContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;
                            //delete local file if exists, as the export wouldn't overwrite
                            string mExpFileName = mDocPath + ".jt";
                            if (System.IO.File.Exists(mExpFileName))
                            {
                                System.IO.File.SetAttributes(mExpFileName, System.IO.FileAttributes.Normal);
                                System.IO.File.Delete(mExpFileName);
                            };
                            DataMedium mData = mInv.TransientObjects.CreateDataMedium();
                            mData.FileName = mExpFileName;
                            mJtTrans.SaveCopyAs(mDoc, mTransContext, mTransOptions, mData);
                            //collect all export files for later upload
                            mUploadFiles.Add(mExpFileName);
                            System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(mUploadFiles.LastOrDefault());
                            if (mExportFileInfo.Exists)
                            {
                                mTrace.WriteLine("JT Translator created file: " + mUploadFiles.LastOrDefault());
                                mTrace.IndentLevel -= 1;
                            }
                            else
                            {
                                mResetIpj(mSaveProject);
                                throw new Exception("Validating the export file before upload failed.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        mResetIpj(mSaveProject);
                        mTrace.WriteLine("JT Export Failed: " + ex.Message);
                    }
                }

                if (item == "SMDXF")
                {
                    //activate DXF translator
                    try
                    {
                        TranslatorAddIn mDXFTrans = mInvSrvAddIns.ItemById["{C24E3AC4-122E-11D5-8E91-0010B541CD80}"] as TranslatorAddIn;
                        mDXFTrans.Activate();
                        if (mDXFTrans == null)
                        {
                            mTrace.WriteLine("DXF Translator not found.");
                            break;
                        }

                        string mExpFileName = mDocPath + ".dxf";
                        if (System.IO.File.Exists(mExpFileName))
                        {
                            System.IO.FileInfo fileInfo = new FileInfo(mExpFileName);
                            fileInfo.IsReadOnly = false;
                            fileInfo.Delete();
                        }

                        PartDocument mPartDoc = (PartDocument)mDoc;
                        DataIO mDataIO = mPartDoc.ComponentDefinition.DataIO;
                        String mOut = "FLAT PATTERN DXF?AcadVersion=R12&OuterProfileLayer=Outer";
                        mDataIO.WriteDataToFile(mOut, mExpFileName);
                        //collect all export files for later upload
                        mUploadFiles.Add(mExpFileName);
                        System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(mUploadFiles.LastOrDefault());
                        if (mExportFileInfo.Exists)
                        {
                            mTrace.WriteLine("SheetMetal DXF Translator created file: " + mUploadFiles.LastOrDefault());
                            mTrace.IndentLevel -= 1;
                        }
                        else
                        {
                            mResetIpj(mSaveProject);
                            throw new Exception("Validating the export file before upload failed.");
                        }
                    }
                    catch (Exception ex)
                    {
                        mResetIpj(mSaveProject);
                        mTrace.WriteLine("SMDXF Export Failed: " + ex.Message);
                    }
                }

                if (item == "3DDWG" || item == "2DDWG")
                {
                    try
                    {
                        //activate DXF/DWG translator
                        TranslatorAddIn mDwgTrans = mInvSrvAddIns.ItemById["{C24E3AC2-122E-11D5-8E91-0010B541CD80}"] as TranslatorAddIn;
                        mDwgTrans.Activate();
                        if (mDwgTrans == null)
                        {
                            mTrace.WriteLine("Dwg Translator not found.");
                            break;
                        }

                        //delete existing export file; note the resulting file name is e.g. "Drawing.idw.dwg
                        string mExpFileName = mDocPath + ".dwg";
                        if (System.IO.File.Exists(mExpFileName))
                        {
                            System.IO.FileInfo fileInfo = new FileInfo(mExpFileName);
                            fileInfo.IsReadOnly = false;
                            fileInfo.Delete();
                        }

                        mTrace.IndentLevel += 1;
                        mTrace.WriteLine("DWG Export starts...");

                        //create the TranslationContext
                        Inventor.TranslationContext mTranslationContext = mInv.TransientObjects.CreateTranslationContext();
                        mTranslationContext.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                        //create NameValueMap object
                        NameValueMap mOptions = mInv.TransientObjects.CreateNameValueMap();

                        //Create a DataMedium object
                        DataMedium mDataMedium = mInv.TransientObjects.CreateDataMedium();

                        //Check whether the translator has 'SaveCopyAs' options and add the export configuration *.ini file
                        if (mDwgTrans.HasSaveCopyAsOptions[mDoc, mTranslationContext, mOptions] == true)
                        {
                            //validate that the ini file exist at the given path
                            if (item == "2DDWG")
                            {
                                System.IO.FileInfo mIniFileInfo = new System.IO.FileInfo(mSettings.DwgIniFile2D);
                                if (mIniFileInfo.Exists)
                                {
                                    mOptions.Value["Export_Acad_IniFile"] = mSettings.DwgIniFile2D;
                                }
                                else
                                {
                                    mResetIpj(mSaveProject);
                                    throw new Exception("Job could not find the DWG export configuration file " + mSettings.DwgIniFile2D);
                                }
                            }
                            if (item == "3DDWG")
                            {
                                mOptions.Value["DwgVersion"] = 32;
                                mOptions.Value["Solid"] = true;
                                mOptions.Value["Surface"] = false;
                                mOptions.Value["Sketch"] = false;
                            }

                        }

                        //create the export file
                        mDataMedium.FileName = mExpFileName;
                        try
                        {
                            mDwgTrans.SaveCopyAs(mDoc, mTranslationContext, mOptions, mDataMedium);
                        }
                        catch (Exception ex)
                        {
                            mResetIpj(mSaveProject);
                            throw new Exception("DWG Translator Add-In failed to export DWG file (TransAddIn.SaveCopyAs()): " + ex.Message);
                        }
                        //collect all export files for later upload
                        mUploadFiles.Add(mExpFileName);
                        System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(mUploadFiles.LastOrDefault());
                        if (mExportFileInfo.Exists)
                        {
                            mTrace.WriteLine("DWG Translator created file: " + mUploadFiles.LastOrDefault());
                            mTrace.IndentLevel -= 1;
                        }
                        else
                        {
                            mResetIpj(mSaveProject);
                            throw new Exception("Validating the export file before upload failed.");
                        }

                    }
                    catch (Exception ex)
                    {
                        mResetIpj(mSaveProject);
                        throw new Exception("Failed to activate DWG Translator Add-in or prepairing the export options: " + ex.Message);
                    }
                }

                if (item == "IMAGE")
                {
                    //delete existing export file; note the resulting file name is e.g. "Drawing.idw.dwg
                    string mExpFileName = mDocPath + "." + mSettings.ImgFileType.ToLower();
                    if (System.IO.File.Exists(mExpFileName))
                    {
                        System.IO.FileInfo fileInfo = new FileInfo(mExpFileName);
                        fileInfo.IsReadOnly = false;
                        fileInfo.Delete();
                    }

                    mTrace.IndentLevel += 1;
                    mTrace.WriteLine("Image Export starts...");

                    //create camera object; note InventorServer does not provide document views (=saved camera)
                    Camera mCamera = mInv.TransientObjects.CreateCamera();
                    PartDocument mPartDoc = null;
                    AssemblyDocument mAssyDoc = null;
                    DrawingDocument mDrwDoc = null;
                    //PresentationDocument mIpnDoc = null; //note - IPN require Inventor application instead of InventorServer

                    //assign the scene object according the doc type: ComponentDefinition for IPT/IAM, Sheet for IDW/DWG, PresentationScene for IPN
                    if (mDoc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                    {
                        mPartDoc = (PartDocument)mDoc;
                        mCamera.SceneObject = mPartDoc.ComponentDefinition;
                        //orient the camera
                        mCamera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
                    }
                    if (mDoc.DocumentType == DocumentTypeEnum.kAssemblyDocumentObject)
                    {
                        mAssyDoc = (AssemblyDocument)mDoc;
                        mCamera.SceneObject = mAssyDoc.ComponentDefinition;
                        //orient the camera
                        mCamera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
                    }
                    if (mDoc.DocumentType == DocumentTypeEnum.kDrawingDocumentObject)
                    {
                        mDrwDoc = (DrawingDocument)mDoc;
                        mCamera.SceneObject = mDrwDoc.ActiveSheet;
                    }
                    //if (mDoc.DocumentType == DocumentTypeEnum.kPresentationDocumentObject) //note - requires Inventor application instead of InventorServer
                    //{
                    //    mIpnDoc = (PresentationDocument)mDoc;
                    //    mCamera.SceneObject = mIpnDoc.ActiveScene;
                    //    mCamera.ViewOrientationType = ViewOrientationTypeEnum.kCurrentViewOrientation;
                    //}

                    //zoom all 
                    mCamera.Fit();
                    mCamera.ApplyWithoutTransition();

                    //set the background color; set different top and bottom color to get a gradient
                    Color mTopClr = mInv.TransientObjects.CreateColor(255, 255, 255, 1); //white
                    Color mBtmClr = mInv.TransientObjects.CreateColor(211, 211, 211, 1); //light grey

                    mCamera.SaveAsBitmap(mExpFileName, 1280, 768, mTopClr, mBtmClr);

                    //collect all export files for later upload
                    mUploadFiles.Add(mExpFileName);
                    System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(mUploadFiles.LastOrDefault());
                    if (mExportFileInfo.Exists)
                    {
                        mTrace.WriteLine("Inventor Image Export created the file: " + mUploadFiles.LastOrDefault());
                        mTrace.IndentLevel -= 1;
                    }
                    else
                    {
                        mResetIpj(mSaveProject);
                        throw new Exception("Validating the export file " + mExpFileName + " before upload failed.");
                    }
                }
            }

            mDoc.Close(true);
            mTrace.WriteLine("Source file closed");

            //switch temporarily used project file back to original one
            mResetIpj(mSaveProject);

            mTrace.WriteLine("Job exported file(s); continues uploading.");
            mTrace.IndentLevel -= 1;

            #endregion VaultInventorServer CAD Export

            #region Vault File Management

            foreach (string file in mUploadFiles)
            {
                ACW.File mExpFile = null;
                System.IO.FileInfo mExportFileInfo = new System.IO.FileInfo(file);
                if (mExportFileInfo.Exists)
                {
                    //copy file to output location
                    if (settings.OutPutPath != "")
                    {
                        System.IO.FileInfo fileInfo = new FileInfo(settings.OutPutPath + "\\" + mExportFileInfo.Name);
                        if (fileInfo.Exists)
                        {
                            fileInfo.IsReadOnly = false;
                            fileInfo.Delete();
                        }
                        System.IO.File.Copy(mExportFileInfo.FullName, settings.OutPutPath + "\\" + mExportFileInfo.Name, true);
                    }

                    //add resulting export file to Vault if it doesn't exist, otherwise update the existing one

                    ACW.Folder mFolder = mWsMgr.DocumentService.FindFoldersByIds(new long[] { mFile.FolderId }).FirstOrDefault();
                    string vaultFilePath = System.IO.Path.Combine(mFolder.FullName, mExportFileInfo.Name).Replace("\\", "/");

                    ACW.File wsFile = mWsMgr.DocumentService.FindLatestFilesByPaths(new string[] { vaultFilePath }).First();
                    VDF.Currency.FilePathAbsolute vdfPath = new VDF.Currency.FilePathAbsolute(mExportFileInfo.FullName);
                    VDF.Vault.Currency.Entities.FileIteration vdfFile = null;
                    VDF.Vault.Currency.Entities.FileIteration addedFile = null;
                    VDF.Vault.Currency.Entities.FileIteration mUploadedFile = null;
                    if (wsFile == null || wsFile.Id < 0)
                    {
                        // add new file to Vault
                        mTrace.WriteLine("Job adds " + mExportFileInfo.Name + " as new file.");

                        if (mFolder == null || mFolder.Id == -1)
                            throw new Exception("Vault folder " + mFolder.FullName + " not found");

                        var folderEntity = new Autodesk.DataManagement.Client.Framework.Vault.Currency.Entities.Folder(connection, mFolder);
                        try
                        {
                            addedFile = connection.FileManager.AddFile(folderEntity, "Created by ExportSampleJob", null, null, ACW.FileClassification.DesignRepresentation, false, vdfPath);
                            mExpFile = addedFile;
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Job could not add export file " + vdfPath + "Exception: ", ex);
                        }

                    }
                    else
                    {
                        // checkin new file version
                        mTrace.WriteLine("Job uploads " + mExportFileInfo.Name + " as new file version.");

                        VDF.Vault.Settings.AcquireFilesSettings aqSettings = new VDF.Vault.Settings.AcquireFilesSettings(connection)
                        {
                            DefaultAcquisitionOption = VDF.Vault.Settings.AcquireFilesSettings.AcquisitionOption.Checkout
                        };
                        vdfFile = new VDF.Vault.Currency.Entities.FileIteration(connection, wsFile);
                        aqSettings.AddEntityToAcquire(vdfFile);
                        var results = connection.FileManager.AcquireFiles(aqSettings);
                        try
                        {
                            mUploadedFile = connection.FileManager.CheckinFile(results.FileResults.First().File, "Created by ExportSampleJob", false, null, null, false, null, ACW.FileClassification.DesignRepresentation, false, vdfPath);
                            mExpFile = mUploadedFile;
                        }
                        catch (Exception ex)
                        {
                            throw new Exception("Job could not update existing export file " + vdfFile + "Exception: ", ex);
                        }
                    }
                }
                else
                {
                    throw new Exception("Job could not find the export result file: " + mDocPath.Replace(mExt, ".stp"));
                }

                mTrace.IndentLevel += 1;

                //update the new file's revision
                try
                {
                    mTrace.WriteLine("Job tries synchronizing " + mExpFile.Name + "'s revision in Vault.");
                    mWsMgr.DocumentServiceExtensions.UpdateFileRevisionNumbers(new long[] { mExpFile.Id }, new string[] { mFile.FileRev.Label }, "Rev Index synchronized by ExportSampleJob");
                    mExpFile = (mWsMgr.DocumentService.GetLatestFileByMasterId(mExpFile.MasterId));
                }
                catch (Exception)
                {
                    //the job will not stop execution in this sample, if revision labels don't synchronize
                }

                //synchronize source file properties to export file properties for UDPs assigned to both
                try
                {
                    mTrace.WriteLine(mExpFile.Name + ": Job tries synchronizing properties in Vault.");
                    //get the design rep category's user properties
                    ACET.IExplorerUtil mExplUtil = Autodesk.Connectivity.Explorer.ExtensibilityTools.ExplorerLoader.LoadExplorerUtil(
                                connection.Server, connection.Vault, connection.UserID, connection.Ticket);
                    Dictionary<ACW.PropDef, object> mPropDictonary = new Dictionary<ACW.PropDef, object>();

                    //get property definitions filtered to UDPs
                    VDF.Vault.Currency.Properties.PropertyDefinitionDictionary mPropDefDic = connection.PropertyManager.GetPropertyDefinitions(
                        VDF.Vault.Currency.Entities.EntityClassIds.Files, null, VDF.Vault.Currency.Properties.PropertyDefinitionFilter.IncludeUserDefined);

                    VDF.Vault.Currency.Properties.PropertyDefinition mPropDef = new PropertyDefinition();
                    ACW.PropInst[] mSourcePropInsts = mWsMgr.PropertyService.GetProperties("FILE", new long[] { mFile.Id }, new long[] { mPropDef.Id });

                    //get property definitions assigned to Design Representation category
                    ACW.CatCfg catCfg1 = mWsMgr.CategoryService.GetCategoryConfigurationById(mExpFile.Cat.CatId, new string[] { "UserDefinedProperty" });
                    List<long> mFilePropDefs = new List<long>();
                    foreach (ACW.Bhv bhv in catCfg1.BhvCfgArray.First().BhvArray)
                    {
                        mFilePropDefs.Add(bhv.Id);
                    }

                    //get properties assigned to source file and add definition/value pair to dictionary
                    mSourcePropInsts = mWsMgr.PropertyService.GetProperties("FILE", new long[] { mFile.Id }, mFilePropDefs.ToArray());
                    ACW.PropDef[] propDefs = connection.WebServiceManager.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE");
                    foreach (ACW.PropInst item in mSourcePropInsts)
                    {
                        mPropDef = connection.PropertyManager.GetPropertyDefinitionById(item.PropDefId);
                        ACW.PropDef propDef = propDefs.SingleOrDefault(n => n.Id == item.PropDefId);
                        mPropDictonary.Add(propDef, item.Val);
                    }

                    //update export file using the property dictionary; note this the IExplorerUtil method bumps file iteration and requires no check out
                    mExplUtil.UpdateFileProperties(mExpFile, mPropDictonary);
                    mExpFile = (mWsMgr.DocumentService.GetLatestFileByMasterId(mExpFile.MasterId));
                }
                catch (Exception ex)
                {
                    mTrace.WriteLine("Job failed copying properties from source file " + mFile.Name + " to export file: " + mExpFile.Name + " . Exception details: " + ex);
                    //you may uncomment the action below if the job should abort executing due to failures copying property values
                    //throw new Exception("Job failed copying properties from source file " + mFile.Name + " to export file: " + mExpFile.Name + " . Exception details: " + ex.ToString() + " ");
                }

                //align lifecycle states of export to source file's state name
                try
                {
                    mTrace.WriteLine(mExpFile.Name + ": Job tries synchronizing lifecycle state in Vault.");
                    Dictionary<string, long> mTargetStateNames = new Dictionary<string, long>();
                    ACW.LfCycDef mTargetLfcDef = (mWsMgr.LifeCycleService.GetLifeCycleDefinitionsByIds(new long[] { mExpFile.FileLfCyc.LfCycDefId })).FirstOrDefault();
                    foreach (var item in mTargetLfcDef.StateArray)
                    {
                        mTargetStateNames.Add(item.DispName, item.Id);
                    }
                    mTargetStateNames.TryGetValue(mFile.FileLfCyc.LfCycStateName, out long mTargetLfcStateId);
                    mWsMgr.DocumentServiceExtensions.UpdateFileLifeCycleStates(new long[] { mExpFile.MasterId }, new long[] { mTargetLfcStateId }, "Lifecycle state synchronized ExportSampleJob");
                }
                catch (Exception ex)
                {
                    mTrace.WriteLine("Job failed aligning lifecycle states of source file " + mFile.Name + " and export file: " + mExpFile.Name + " . Exception details: " + ex);
                }

                //attach export file to source file leveraging design representation attachment type
                try
                {
                    mTrace.WriteLine(mExpFile.Name + ": Job tries to attach to its source in Vault.");
                    ACW.FileAssocParam mAssocParam = new ACW.FileAssocParam();
                    mAssocParam.CldFileId = (mWsMgr.DocumentService.GetLatestFileByMasterId(mExpFile.MasterId)).Id;
                    mAssocParam.ExpectedVaultPath = mWsMgr.DocumentService.FindFoldersByIds(new long[] { mFile.FolderId }).First().FullName;
                    mAssocParam.RefId = null;
                    mAssocParam.Source = null;
                    mAssocParam.Typ = ACW.AssociationType.Attachment;
                    //refresh the parent file to the latest version id; default jobs like sync props or update rev.table might have updated the parent already
                    mFile = (mWsMgr.DocumentService.GetLatestFileByMasterId(mFile.MasterId));
                    mWsMgr.DocumentService.AddDesignRepresentationFileAttachment(mFile.Id, mAssocParam);
                }
                catch (Exception ex)
                {
                    mTrace.WriteLine("Job failed attaching the exported file " + mExpFile.Name + " to the source file: " + mFile.Name + " . Exception details: " + ex);
                }

                mTrace.IndentLevel -= 1;

            }

            //evaluate PDF creation as latest step
            //more links: https://forums.autodesk.com/t5/vault-customization/in-what-order-job-types-are-processed-default-job-types-and-user/td-p/9492378
            //more links: https://support.coolorange.com/support/solutions/articles/22000201245-how-to-manually-queue-autodesk-vault-jobs
            //if (mFile.Name.EndsWith(".idw"))
            //{
            //    ACW.JobParam mJobParamFileVerId = new ACW.JobParam();
            //    ACW.JobParam mJobParamUptViewOpt = new ACW.JobParam();
            //    List<ACW.JobParam> mJobParamsList = new List<ACW.JobParam>();

            //    mJobParamFileVerId.Name = "FileVersionId";
            //    mJobParamFileVerId.Val = (mWsMgr.DocumentService.GetLatestFileByMasterId(mFile.MasterId)).Id.ToString();
            //    mJobParamsList.Add(mJobParamFileVerId);

            //    mJobParamUptViewOpt.Name = "UpdateViewOption";
            //    mJobParamUptViewOpt.Val = "False";
            //    mJobParamsList.Add(mJobParamUptViewOpt);

            //    mWsMgr.JobService.AddJob("Autodesk.Vault.PDF.Create.idw", "Create PDF: " + mFile.Name + " (submitted by ExportSampleJob)", mJobParamsList.ToArray(), 1);
            //}

            #endregion Vault File Management

            mTrace.IndentLevel = 1;
            mTrace.WriteLine("Job finished all steps.");
            mTrace.Flush();
            mTrace.Close();
        }

        private static void mResetIpj(DesignProject mSaveProject)
        {
            if (mSaveProject != null)
            {
                mSaveProject.Activate();
            }
        }
        #endregion Job Execution
    }
}
