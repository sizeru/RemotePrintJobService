using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

namespace RemotePrintJobService
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            // Create an event log source for the program. I don't think
            // this line of code should ever be hit. Since this just loads
            // the service into the service manager, a restart is not required
            // to create a new EventLog source, unlike with most regular
            // programs
            //if (!EventLog.SourceExists("RemotePrintJobService"))
            //{
            //    string source = "RemotePrintJobService";
            //    string logName = "Application";
            //    EventLog.CreateEventSource(source, logName);
            //}
            // Load the service into the Services Control Manager
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new RPJService()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
