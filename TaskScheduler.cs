using Quartz.Impl;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Quartz.Core;
using System.Configuration;

namespace CronService_Processor
{
    public class TaskScheduler : ITaskScheduler
    {
        private IScheduler _scheduler;
        public string Name
        {
            get { return GetType().Name; }
        }

        public void Run()
        {
            ISchedulerFactory schedulerFactory = new StdSchedulerFactory();
            _scheduler = schedulerFactory.GetScheduler();

            string taskGroup = "Test Group";
            string taskName = "Test Name";

            IJobDetail testJob = JobBuilder.Create<ScheduledJobs>()
                    .WithIdentity(taskName, taskGroup)
                    .Build();

            ITrigger testTrigger = TriggerBuilder.Create()
                    .WithIdentity(taskName, taskGroup)
                    .StartNow()
                    .WithCronSchedule(ConfigurationManager.AppSettings["Interval"])
                    .Build();

            var dictionary = new Dictionary<IJobDetail, Quartz.Collection.ISet<ITrigger>>();

            dictionary.Add(testJob, new Quartz.Collection.HashSet<ITrigger>()
                                {
                                    testTrigger
                                });


            _scheduler.ScheduleJobs(dictionary, false);
            _scheduler.Start();
        }

        public void Stop()
        {
            _scheduler.Shutdown();
        }
    }
}
