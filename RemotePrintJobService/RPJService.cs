using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
//using System.Threading;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net;
using System.Timers;
using System.Configuration;
using PdfiumViewer;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.Reflection;

namespace RemotePrintJobService
{
    public partial class RPJService : ServiceBase
    {
        private System.Timers.Timer queryRemoteJobsTimer;
        private HttpClient CLIENT;
        private EventLog LOG;
        private string PRINTER_NAME;
        private short PRINTER_COPIES;
        private string PAPER_SIZE;

        public RPJService()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            LOG = new EventLog("Application");
            LOG.Source = "RemotePrintJobService";
            Configuration config;
            try
            {
                Assembly executingAssembly = Assembly.GetAssembly(typeof(ProjectInstaller));
                string targetDir = executingAssembly.Location;
                config = ConfigurationManager.OpenExeConfiguration(targetDir);
                LOG.WriteEntry("Target dir: " + targetDir);
            }
            catch (Exception ex)
            {
                LOG.WriteEntry("Error: Could not read from config. Exception: " + ex.Message);
                this.Stop();
                return;
            }

            try
            {
                string token = config.AppSettings.Settings["Token"].Value;
                string baseUri = config.AppSettings.Settings["BaseURI"].Value;
                CLIENT = new HttpClient();
                CLIENT.DefaultRequestHeaders.Add("Token", token);
                CLIENT.BaseAddress = new Uri(baseUri);
            }
            catch (Exception ex)
            {
                LOG.WriteEntry("Error: Could not initialize HttpClient. Exception: " + ex.Message);
                this.Stop();
                return;
            }

            try
            {
                PRINTER_NAME = config.AppSettings.Settings["Printer"].Value;
                PRINTER_COPIES = Int16.Parse(config.AppSettings.Settings["PrintCopies"].Value);
                PAPER_SIZE = config.AppSettings.Settings["PaperSize"].Value;
            }
            catch (Exception ex)
            {
                LOG.WriteEntry("Error: Could not find the printer name and print copies in the config. Exception: " + ex.Message);
                this.Stop();
                return;
            }

            try
            {
                string qrjTimerString = config.AppSettings.Settings["QueryRemoteJobsTimer"].Value;
                double qrjTimerValue = Double.Parse(qrjTimerString);
                this.queryRemoteJobsTimer = new System.Timers.Timer(qrjTimerValue) { AutoReset = true };
                queryRemoteJobsTimer.Elapsed += QueryRemoteJobs;
                queryRemoteJobsTimer.Start();
            }
            catch (Exception ex)
            {
                LOG.WriteEntry("Error: Could not create a timer to query remote jobs becasue of exception: " + ex.Message);
                this.Stop();
                return;
            }
        }

        protected override void OnStop()
        {
        }

        // Query for remote jobs from the server
        private async void QueryRemoteJobs(object sender, ElapsedEventArgs e)
        {
            HttpResponseMessage jobList;
            try
            { 
                jobList = await CLIENT.GetAsync("list");
            }
            catch (Exception ex)
            {
                LOG.WriteEntry("Warning: Could not reach server when trying to GET `list`. Exception: " + ex.Message);
                return;
            }
            if (!jobList.IsSuccessStatusCode)
            {
                LOG.WriteEntry("Warning: Server returned a failure response code. Code: " + jobList.StatusCode + " | Body: " + jobList.Content);
                return;
            }
            string jobListContent = await jobList.Content.ReadAsStringAsync();
            string[] jobs = jobListContent.Split(new char[] { '\n' });
            foreach (string filename in jobs)
            {
                // Assume that all files returned by the server are valid
                HttpResponseMessage file;
                using (var request = new HttpRequestMessage(HttpMethod.Post, "file"))
                {
                    request.Headers.Add("Filename", filename);
                    try
                    {
                        file = await CLIENT.SendAsync(request);
                    }
                    catch (Exception ex)
                    {
                        LOG.WriteEntry("Warning: Could not reach server when trying to GET `file` with filename: " + filename + " | Exception: " + ex.Message);
                        continue;
                    }
                }
                if (!file.IsSuccessStatusCode)
                {
                    LOG.WriteEntry("Server returned a failure response code when trying to retrieve file: " + filename + " | Code: " + file.StatusCode + " | Body: " + file.Content);
                    continue;
                }

                // Print file. This is really banking on the file being a PDF.
                Stream pdfStream = await file.Content.ReadAsStreamAsync();
                try
                {
                    PrintPDF(pdfStream);
                }
                catch (Exception ex)
                {
                    LOG.WriteEntry("Error: Could not print PDF. Exception: " + ex.Message);
                    continue;
                }
            }
        }

        // Print the PDF on the printer (according to the config)
        public void PrintPDF(Stream pdfStream) 
        {
            PrinterSettings printerSettings;
            // Create printer settings
            printerSettings = new PrinterSettings
            {
                PrinterName = PRINTER_NAME,
                Copies = PRINTER_COPIES,
            };


            // Create our page settings for the paper size selected
            var pageSettings = new PageSettings(printerSettings)
            {
                Margins = new Margins(0, 0, 0, 0),
            };
                
            foreach (PaperSize paperSize in printerSettings.PaperSizes)
            {
                if (paperSize.PaperName == PAPER_SIZE)
                {
                    pageSettings.PaperSize = paperSize;
                    break;
                }
            }

            // Now print the PDF document
            using (var document = PdfDocument.Load(pdfStream))
            {
                using (var printDocument = document.CreatePrintDocument())
                {
                    printDocument.PrinterSettings = printerSettings;
                    printDocument.DefaultPageSettings = pageSettings;
                    printDocument.PrintController = new StandardPrintController();
                    printDocument.Print();
                }
            }
        }
    }
}
