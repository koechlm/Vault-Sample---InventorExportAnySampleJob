using System;
using Autodesk.Navisworks.Api.Automation;

namespace NavisworksExportExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize the Navisworks Automation object
            NavisworksApplication mNavisworksAutomation = new NavisworksApplication();

            // Path to the document to open
            string mDocPath = @"C:\path\to\your\file.nwd";
            // Path to save the new Navisworks file
            string mNWDName = @"C:\path\to\output\file.nwd";
            // Path to save the export file
            string mExportFilePath = @"C:\path\to\output\file";

            // Condition to check for exporting format
            string exportFormat = "DWF"; // Change this to the desired format

            try
            {
                // Open the file with Navisworks; opening other file formats creates a new Navisworks file appending the import file
                mNavisworksAutomation.OpenFile(mDocPath);

                // Save the new Navisworks file
                mNavisworksAutomation.SaveFile(mNWDName);

                // Export to the desired format using the appropriate plugin
                switch (exportFormat)
                {
                    case "DWF":
                        mNavisworksAutomation.ExecuteAddInPlugin("NativeExportPluginAdaptor_LcDwfExporterPlugin_Export.Navisworks", mExportFilePath + ".dwf");
                        Console.WriteLine("Export to DWF successful!");
                        break;
                    case "FBX":
                        mNavisworksAutomation.ExecuteAddInPlugin("NativeExportPluginAdaptor_LcFbxExporterPlugin_Export.Navisworks", mExportFilePath + ".fbx");
                        Console.WriteLine("Export to FBX successful!");
                        break;
                    case "RVT":
                        mNavisworksAutomation.ExecuteAddInPlugin("NativeExportPluginAdaptor_LcRvtExporterPlugin_Export.Navisworks", mExportFilePath + ".rvt");
                        Console.WriteLine("Export to RVT successful!");
                        break;
                    // Add more cases for other formats as needed
                    default:
                        Console.WriteLine("Unsupported export format.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
            finally
            {
                // Clean up the Navisworks Automation object
                mNavisworksAutomation.Dispose();
            }
        }
    }
}
