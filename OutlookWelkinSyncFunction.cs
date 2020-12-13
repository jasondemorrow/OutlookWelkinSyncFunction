namespace OutlookWelkinSyncFunction
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Azure.WebJobs;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;
    using Ninject.Parameters;
    using Sync = OutlookWelkinSync;

    public static class OutlookWelkinSyncFunction
    {
        [FunctionName("OutlookWelkinSyncFunction")]
        public static void Run([TimerTrigger("%TimerSchedule%")]TimerInfo timerInfo, ILogger log)
        {
            log.LogInformation($"Starting Welkin/Outlook events sync at: {DateTime.Now}");
            
            Sync.NinjectModules.CurrentLogger = log;
            IKernel ninject = new StandardKernel(Sync.NinjectModules.CurrentModule);
            Sync.WelkinClient welkinClient = ninject.Get<Sync.WelkinClient>();
            Sync.OutlookClient outlookClient = ninject.Get<Sync.OutlookClient>();
            Sync.OutlookEventRetrieval outlookEventRetrieval = ninject.Get<Sync.OutlookEventRetrieval>();

            List<Sync.WelkinSyncTask> welkinSyncTasks = new List<Sync.WelkinSyncTask>();
            List<Sync.OutlookSyncTask> outlookSyncTasks = new List<Sync.OutlookSyncTask>();

            DateTime lastRun = timerInfo?.ScheduleStatus?.Last ?? DateTime.UtcNow.AddHours(-2);
            TimeSpan historySpan = DateTime.UtcNow - lastRun.AddMinutes(-1);

            // 1. Get all recently updated Welkin events (sync is Welkin-driven since this set of users will be smaller)
            IEnumerable<Sync.WelkinEvent> welkinEvents = welkinClient.RetrieveEventsUpdatedSince(historySpan);
            foreach (Sync.WelkinEvent welkinEvent in welkinEvents)
            {
                log.LogInformation($"Found a new Welkin event, ID {welkinEvent.Id}.");
                ConstructorArgument argument = new ConstructorArgument("welkinEvent", welkinEvent);
                Sync.WelkinSyncTask welkinSyncTask = ninject.Get<Sync.WelkinSyncTask>(argument);
                welkinSyncTasks.Add(welkinSyncTask);
            }

            // 2. Run Outlook event retrieval, which checks all Welkin workers' Outlook calendars or a shared calendar
            IEnumerable<Event> outlookEvents = outlookEventRetrieval.RetrieveAllUpdatedSince(historySpan);
            foreach (Event outlookEvent in outlookEvents)
            {
                log.LogInformation($"Found a new Outlook event, ID {outlookEvent.ICalUId}.");
                ConstructorArgument argument = new ConstructorArgument("outlookEvent", outlookEvent);
                Sync.OutlookSyncTask outlookSyncTask = ninject.Get<Sync.OutlookSyncTask>(argument);
                outlookSyncTasks.Add(outlookSyncTask);
            }

            // 3. Run all Welkin sync tasks created for newly updated events, creating corresponding placeholder events in Outlook
            foreach (Sync.OutlookSyncTask outlookSyncTask in outlookSyncTasks)
            {
                try
                {
                    outlookSyncTask.Sync();
                }
                catch (Exception ex)
                {
                    log.LogError($"Exception while running {outlookSyncTask.ToString()}: {ex.Message} {ex.StackTrace}");
                }
            }

            // 4. Run all Outlook sync tasks created for newly updated events, creating corresponding placeholder events in Welkin
            foreach (Sync.WelkinSyncTask welkinSyncTask in welkinSyncTasks)
            {
                try
                {
                    welkinSyncTask.Sync();
                }
                catch (Exception ex)
                {
                    log.LogError($"Exception while running {welkinSyncTask.ToString()}: {ex.Message} {ex.StackTrace}");
                }
            }

            log.LogInformation("Done!");
        }
    }
}