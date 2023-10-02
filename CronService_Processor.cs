using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace CronService_Processor
{
    public partial class CronService_Processor : ServiceBase
    {
        ITaskScheduler scheduler;
        public CronService_Processor()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            scheduler = new TaskScheduler();
            scheduler.Run();
        }

        protected override void OnStop()
        {
            if (scheduler != null)
            {
                scheduler.Stop();
            }
        }

        public void BeginService()
        {
            scheduler = new TaskScheduler();
            scheduler.Run();
        }
    }
}
