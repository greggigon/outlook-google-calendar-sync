using System;
using System.Threading;
using System.Configuration;
using Quartz;
using Quartz.Impl;


namespace OutlookCalendarExporter
{
    class Program
    {
        static ManualResetEvent _quitEvent = new ManualResetEvent(false);

        static void Main(string[] args)
        {
            var calendarId = ConfigurationManager.AppSettings["googleCalendarId"];
            var apiCredentialsFileName = ConfigurationManager.AppSettings["apiCredentialsFileName"];
            var syncIntervalInMinutes = int.Parse(ConfigurationManager.AppSettings["syncIntervalInMinutes"]);
            var numberOfDaysToSync = int.Parse(ConfigurationManager.AppSettings["numberOfDaysToSync"]);

            try {
                var scheduler = StdSchedulerFactory.GetDefaultScheduler();
                scheduler.Start();

                var calendarSyncJob = JobBuilder.Create<SyncJob>()
                    .WithIdentity("sync-outlook-into-google", "calendar-sync")
                    .UsingJobData("calendarId", calendarId)
                    .UsingJobData("calendarApiCredentialsFileName", apiCredentialsFileName)
                    .UsingJobData("numberOfDaysToSync", numberOfDaysToSync)
                    .Build();

                var calendarSyncTrigger = TriggerBuilder.Create()
                    .WithIdentity("sync-outlook-into-google", "calendar-sync")
                    .StartNow()
                    .WithSimpleSchedule(x => x
                        .WithIntervalInMinutes(syncIntervalInMinutes)
                        .RepeatForever())
                    .Build();

                scheduler.ScheduleJob(calendarSyncJob, calendarSyncTrigger);

                Console.CancelKeyPress += (sender, eArgs) => {
                    _quitEvent.Set();
                    eArgs.Cancel = true;
                };

                Console.WriteLine("*** Ctrl+C will exit this application ****");
                _quitEvent.WaitOne();

                scheduler.Shutdown();

            }catch(SchedulerException ex)
            {
                Console.WriteLine(ex);
            }

        }

       
    }
}
