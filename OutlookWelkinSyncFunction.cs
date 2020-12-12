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

            // 2. Create Outlook event retrieval, which checks all Welkin workers' Outlook calendars or a shared calendar
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
            
            /*
            OutlookClientOld outlookClient = new OutlookClientOld(new OutlookConfig(), log);
            WelkinClientOld welkinClient = new WelkinClientOld(new WelkinConfig(), log);
            OutlookEventSync outlookEventSync = new OutlookEventSync(outlookClient, welkinClient, log);
            WelkinEventSync welkinEventSync = new WelkinEventSync(outlookClient, welkinClient, log);
            DateTime lastRun = timerInfo?.ScheduleStatus?.Last ?? DateTime.UtcNow.AddHours(-24);
            TimeSpan historySpan = DateTime.UtcNow - lastRun;
            IEnumerable<User> outlookUsers = new List<User>();
            IEnumerable<WelkinPractitioner> welkinUsers = new List<WelkinPractitioner>();

            try
            {
                // TODO: Pagination and re-factoring
                outlookUsers = outlookClient.GetAllUsers();
                welkinUsers = welkinClient.GetAllPractitioners();
            }
            catch (Exception e)
            {
                string trace = Exceptions.ToStringRecursively(e);
                log.LogError($"While retrieving users: {trace}");
            }

            IDictionary<string, string> welkinUserNamesByCalendarId = new Dictionary<string, string>();
            IDictionary<string, string> welkinCalendarIdsByUserName = new Dictionary<string, string>();
            IDictionary<string, WelkinPractitioner> welkinPractitionerByUserName = new Dictionary<string, WelkinPractitioner>();

            foreach (WelkinPractitioner welkinUser in welkinUsers) // Build mappings of Welkin users to their calendars in Welkin
            {
                string userName = UserNameFrom(welkinUser.Email);
                //log.LogInformation($"Found Welkin user {welkinUser}");
                if (string.IsNullOrEmpty(userName))
                {
                    continue;
                }

                try
                {
                    WelkinCalendar calendar = welkinClient.GetCalendarForPractitioner(welkinUser);
                    if (calendar == null)
                    {
                        log.LogWarning($"Welkin calendar not found for user {userName}");
                        continue;
                    }
                    welkinUserNamesByCalendarId[calendar.Id] = userName;
                    welkinCalendarIdsByUserName[userName] = calendar.Id;
                    welkinPractitionerByUserName[userName] = welkinUser;
                }
                catch (Exception e)
                {
                    string trace = Exceptions.ToStringRecursively(e);
                    log.LogError($"While retrieving Welkin calendar for {userName}: {trace}");
                }
            }

            IEnumerable<WelkinEvent> welkinEvents = welkinClient.GetEventsUpdatedSince(historySpan);
            Dictionary<string, Dictionary<string, WelkinEvent>> welkinEventsByUserNameThenEventId = new Dictionary<string, Dictionary<string, WelkinEvent>>();

            // Cache all recently updated Welkin events, keyed by user name first then by event ID.
            foreach (WelkinEvent welkinEvent in welkinEvents)
            {
                if (!welkinUserNamesByCalendarId.ContainsKey(welkinEvent.CalendarId))
                {
                    continue;
                }

                string userName = welkinUserNamesByCalendarId[welkinEvent.CalendarId];

                if (string.IsNullOrEmpty(userName))
                {
                    log.LogWarning($"Welkin event {welkinEvent.Id} has no known user.");
                    continue;
                }

                if (!welkinEventsByUserNameThenEventId.ContainsKey(userName))
                {
                    welkinEventsByUserNameThenEventId[userName] = new Dictionary<string, WelkinEvent>();
                }

                welkinEventsByUserNameThenEventId[userName][welkinEvent.Id] = welkinEvent;
            }

            // Find common users
            Dictionary<string, string> welkinIdToOutlookPrincipal = new Dictionary<string, string>();
            foreach (User outlookUser in outlookUsers)
            {
                // For users with accounts in both Welkin and Outlook, sync their recently updated events
                string userName = UserNameFrom(outlookUser.UserPrincipalName);
                if (string.IsNullOrEmpty(userName) || !welkinCalendarIdsByUserName.ContainsKey(userName) || !welkinPractitionerByUserName.ContainsKey(userName))
                {
                    log.LogWarning($"Unknown user ({userName}) or missing calendar or practitioner for user.");
                    continue;
                }

                IEnumerable<Event> recentlyUpdatedOutlookEvents = new List<Event>();
                try
                {
                    recentlyUpdatedOutlookEvents = outlookClient.GetEventsForUserUpdatedSince(outlookUser, historySpan, Constants.OutlookEventExtensionsNamespace);
                }
                catch (Exception e)
                {
                    string trace = Exceptions.ToStringRecursively(e);
                    log.LogError(e, $"While retrieving Outlook events for user {userName}: {trace}");
                }

                // First, sync newly updated Outlook events for user
                foreach (Event outlookEvent in recentlyUpdatedOutlookEvents)
                {
                    log.LogInformation($"Found newly updated Outlook event '{outlookEvent.ICalUId}' for user {userName}.");
                    try
                    {
                        string syncId = outlookEventSync.Sync(
                            outlookEvent, 
                            outlookUser, 
                            welkinPractitionerByUserName[userName],
                            welkinCalendarIdsByUserName[userName]);
                
                        if (syncId != null && welkinEventsByUserNameThenEventId.ContainsKey(userName) && 
                            welkinEventsByUserNameThenEventId[userName].ContainsKey(syncId))
                        {
                            // If the existing Welkin event has also been recently updated, we can skip it later
                            welkinEventsByUserNameThenEventId[userName].Remove(syncId);
                            log.LogInformation($@"Welkin event with ID {syncId} has recently been updated, " +
                                                "but will be skipped since its corresponding Outlook event " +
                                                "with ID {evt.ICalUId} has also been recently updated and " +
                                                "therefore sync'ed.");
                        }
                    }
                    catch (Exception e)
                    {
                        string trace = Exceptions.ToStringRecursively(e);
                        log.LogError($"Sync failed with Outlook event {outlookEvent.ICalUId} for user {userName}: {trace}");
                    }
                }

                // Second, sync newly updated Welkin events for user, if any
                if (welkinEventsByUserNameThenEventId.ContainsKey(userName))
                {
                    foreach (WelkinEvent welkinEvent in welkinEventsByUserNameThenEventId[userName].Values)
                    {
                        log.LogInformation($"Found newly updated Welkin event '{welkinEvent}' for user {userName}.");
                        try
                        {
                            string syncId = welkinEventSync.Sync(
                                welkinEvent, 
                                outlookUser, 
                                welkinPractitionerByUserName[userName], 
                                welkinCalendarIdsByUserName[userName]);
                        }
                        catch (Exception e)
                        {
                            string trace = Exceptions.ToStringRecursively(e);
                            log.LogError($"Sync failed for Welkin event {welkinEvent}: {trace}");
                        }
                    }
                }
            }
            */
            log.LogInformation("Done!");
        }

        private static string UserNameFrom(string email)
        {
            if(string.IsNullOrEmpty(email)) return null;
            int idx = email.IndexOf('+');
            if(idx == -1) idx = email.IndexOf('@');
            if(idx == -1) return null;
            return email.Substring(0, idx);
        }
    }
}