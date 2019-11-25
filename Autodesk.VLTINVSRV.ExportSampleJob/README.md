# Vault-Inventor-Server EXPORT JOB SAMPLE

This custom job creates Inventor export files using VaultInventorServer.

INTRODUCTION:
---------------------------------
This sample walks you through the basic concept handling Inventor part or assembly files submitted to the job queue by VaultInventorServer.
To open 3D part files as well as full assemblies maintaining all file references, the job leverages Vault project file settings initializing the job execution.
File download and running the Inventor translator add-in for CAD-Export formats follow as major execution step.
The final steps adds export files to Vault, copying user properties, trying to synchronize revision index and lifecycle state, and finally attaches the result to their source respectively.

REQUIREMENTS:
---------------------------------
Vault Workgroup, Vault Professional 2020 or newer. This job leverages the Vault Inventor Server component and does not require Inventor installation or Inventor license.
The job is valid for any Vault configuration fulfilling these requirements:
- Enforce Workingfolder = Enabled
- Enforce Inventor Project File = Enabled
- Single project file in Vault.
- Category Design Representation (as default configuration "Manufacturing")
- Category Design Representation Assignement Rule based on File classification "Design Representation" (as default configuration "Manufacturing")
- Matching Revision scheme for source file category and design representation category
- Matching Lifecycle State names for lifecycles applied to source and to export files
- Job processor user need to have transition access to life cycle of design representation category

TO CONFIGURE:
---------------------------------
1) Copy the folder Autodesk.VltInvSrv.ExportSampleJob to %ProgramData%\Autodesk\Vault 2020\Extensions\.
2) Edit the settings.xml to configure the logfile directory; the default is "C:\Temp\".
3) Edit the settings.xml to configure the export format(s) per Job; multiple formats are allowed e.g., to create a STEP for Sheet Metal folded model 
	and a DXF for the Sheet Metal Flat Pattern add the values STP, SMDXF to <ExportFomats>STP,SMDXF</ExportFomats>. To enable dual export for Sheet Metal components
	a category for these part files is expected; configure the name of the category using the element SmCadDispName e.g., <SmCatDispName>Sheet Metal Part</SmCatDispName>
4) Add the job to the Job Queue activating job transitions. To achieve this, integrate this job into a custom lifecycle transition by adding the Job-Type name
"Autodesk.VltInvSrv.ExportSampleJob" to the transition's 'Custom Job Types' tab.

DISCLAIMER:
---------------------------------
In any case, all binaries, configuration code, templates and snippets of this solution are of "work in progress" character. This also applies to GitHub "Release" versions.
Neither Markus Koechl, nor Autodesk represents that these samples are reliable, accurate, complete, or otherwise valid. 
Accordingly, those configuration samples are provided “as is” with no warranty of any kind and you use the applications at your own risk.


NOTES/KNOWN ISSUES:
---------------------------------
The job expects that all library definition files configured in the Inventor project file are available in the file system. Otherwise, the job might fail with unhandled exception due to missing style library or other settings files.

VERSION HISTORY / RELEASE NOTES:
---------------------------------
2020.25.1.0 - Added Export-Format Configuration, Output Folder option, named and published as Autodesk.VltInvSrv.ExportSampleJob, 
				implemented formats are STEP, JT and Sheet Metal Flat Pattern DXF
2020.25.0.3 - Added Logging, named and published as Autodesk.STEP.ExportSampleJob
2020.25.0.0 - Initial Version, named and published as Autodesk.STEP.ExportSampleJob
 
---------------------------------

Thank you for your interest in Autodesk Vault solutions and API samples.

Sincerely,

Markus Koechl, Autodesk
