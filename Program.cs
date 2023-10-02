using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CronService_Processor
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            CronService_Processor ServicesToRun = new CronService_Processor();
            if (!Debugger.IsAttached)
            {
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                ServicesToRun.BeginService();
            }
        }
    }
}
