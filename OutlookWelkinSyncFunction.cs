using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OutlookWelkinSyncFunction
{
    public static class OutlookWelkinSyncFunction
    {
        [FunctionName("OutlookWelkinSyncFunction")]
        public static void Run([TimerTrigger("%TimerSchedule%")]TimerInfo timerInfo, ILogger log)
        {
            log.LogInformation($"Starting Welkin/Outlook events sync at: {DateTime.Now}");
            OutlookClient outlookClient = new OutlookClient(new OutlookConfig(), log);
            WelkinClient welkinClient = new WelkinClient(new WelkinConfig(), log);
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
            foreach (User user in outlookUsers)
            {
                string userName = UserNameFrom(user.UserPrincipalName);
                if (string.IsNullOrEmpty(userName) || !welkinCalendarIdsByUserName.ContainsKey(userName) || !welkinPractitionerByUserName.ContainsKey(userName))
                {
                    log.LogWarning($"Unknown user ({userName}) or missing calendar or practitioner for user.");
                    continue;
                }

                // For users with accounts in both Welkin and Outlook, sync their recently updated events
                try
                {
                    // First, sync newly updated Outlook events for user
                    IEnumerable<Event> recentlyUpdatedOutlookEvents = 
                        outlookClient.GetEventsForUserUpdatedSince(user, historySpan, Constants.OutlookEventExtensionsNamespace);
                    foreach (Event evt in recentlyUpdatedOutlookEvents)
                    {
                        outlookEventSync.Sync(
                            evt, 
                            user, 
                            welkinPractitionerByUserName[userName], 
                            welkinEventsByUserNameThenEventId, 
                            welkinCalendarIdsByUserName[userName], 
                            userName);
                    }

                    // Second, sync newly updated Welkin events for user, if any
                    if (welkinEventsByUserNameThenEventId.ContainsKey(userName))
                    {
                        foreach (WelkinEvent evt in welkinEventsByUserNameThenEventId[userName].Values)
                        {
                            welkinEventSync.Sync(
                                evt, 
                                user, 
                                welkinPractitionerByUserName[userName], 
                                welkinEventsByUserNameThenEventId, 
                                welkinCalendarIdsByUserName[userName], 
                                userName);
                        }
                    }
                }
                catch (Exception e)
                {
                    string trace = Exceptions.ToStringRecursively(e);
                    log.LogError(e, $"While sync'ing events for user {userName}: {trace}");
                }
            }

            log.LogInformation("Done!");
        }

        private static string UserNameFrom(string email)
        {
            if(string.IsNullOrEmpty(email)) return null;
            int idx = email.IndexOf('@');
            if(idx == -1) return null;
            return email.Substring(0, idx);
        }
    }
}