using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Printing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web;
using UDC;

namespace UDC.Models
{
    public static class TiffConversion
    {
        private static readonly string _conversionFilePath = new AppSettingsReader().GetValue("ConversionFilePath", typeof(string)).ToString();
        [DllImport("Winspool.drv")]
        private static extern bool SetDefaultPrinter(string printerName);

        /// <summary>
        /// Convert afile to tiff
        /// The application that can print the file must exist on the server 
        /// and be able to print in non-attended mode
        /// </summary>
        /// <param name="filename">the file to be converted.  The file must reside in the C:\Windows\Temp\UDC folder.</param>
        /// <returns></returns>
        public static string ConvertToTiff(string filename)
        {
            try
            {
                //Create a UDC object and get its interfaces
                var objUdc = new APIWrapper();
                var printer = objUdc.get_Printers("Universal Document Converter");
                var profile = printer.Profile;
                //Use Universal Document Converter API to change settings of converted document
                profile.PageSetup.ResolutionX = 600;
                profile.PageSetup.ResolutionY = 600;
                profile.FileFormat.ActualFormat = FormatID.FMT_TIFF;
                profile.FileFormat.TIFF.ColorSpace = ColorSpaceID.CS_TRUECOLOR;
                profile.FileFormat.TIFF.Multipage = MultipageModeID.MM_MULTI;
                profile.OutputLocation.Mode = LocationModeID.LM_PREDEFINED;
                profile.OutputLocation.FolderPath = _conversionFilePath;
                profile.OutputLocation.FileName = @"&[DocName(0)].&[ImageType]";
                profile.OutputLocation.OverwriteExistingFile = false;
                profile.PostProcessing.Mode = PostProcessingModeID.PP_NONE;
                profile.Advanced.ShowNotifications = false;
                profile.Advanced.ShowProgressWnd = false;
                // Change the default printer for the user running the VowApi.
                SetDefaultPrinter("Universal Document Converter");
                // Use default file association of existing file to print to UDC. 
                var info = new ProcessStartInfo(string.Format(@"{0}\{1}", _conversionFilePath, filename));
                info.Verb = "Print";
                info.CreateNoWindow = true;
                info.WindowStyle = ProcessWindowStyle.Hidden;
                using (var process = new Process())
                {
                    process.StartInfo = info;
                    process.Start();
                    process.WaitForExit();
                }
                WaitForPrinterJobToComplete(filename);
                filename = filename.Substring(0, filename.LastIndexOf(".")) + ".tif";
                return filename;

            }
            catch (Exception ex)
            {
                var message = string.Format(@"Exception Message: [{0}]   ...   Stack Trace: [{1}]",
                    ex.Message, ex.StackTrace);
                throw;
            }
        }

        public static PrintSystemJobInfo GetPrinterJob()
        {
            var jobInfos = new List<PrintSystemJobInfo>();
            var server = new LocalPrintServer();
            var printQueue = server.GetPrintQueue("Universal Document Converter");
            jobInfos = printQueue.GetPrintJobInfoCollection().ToList();
            return jobInfos[0];
        }

        public static void WaitForPrinterJobToComplete(string filename)
        {
            var jobInfos = new List<PrintSystemJobInfo>();
            var server = new LocalPrintServer();
            var printQueue = server.GetPrintQueue("Universal Document Converter");
            jobInfos = printQueue.GetPrintJobInfoCollection().ToList();
            foreach (var jobInfo in jobInfos)
            {
                if (jobInfo.Name.ToLower().IndexOf(filename.ToLower()) >= 0)
                {
                    var isJobComplete = false;
                    while (!isJobComplete)
                    {
                        try
                        {
                            var update = printQueue.GetJob(jobInfo.JobIdentifier);
                            Thread.Sleep(250);
                        }
                        catch (Exception)
                        {
                            isJobComplete = true;
                        }
                    }
                    return;
                }
            }
        }
    }
}