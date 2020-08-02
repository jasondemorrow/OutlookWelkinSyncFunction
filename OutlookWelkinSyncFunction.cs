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
            DateTime lastRun = timerInfo.ScheduleStatus.Last;
            TimeSpan historySpan = DateTime.UtcNow - lastRun;

            IEnumerable<User> outlookUsers = outlookClient.GetAllUsers();
            IEnumerable<WelkinPractitioner> welkinUsers = welkinClient.GetAllPractitioners();
            ISet<string> welkinUserNames = new HashSet<string>();
            IDictionary<string, string> welkinUserNamesByCalendarId = new Dictionary<string, string>();

            foreach (WelkinPractitioner welkinUser in welkinUsers)
            {
                string userName = UserNameFrom(welkinUser.Email);
                WelkinCalendar calendar = welkinClient.GetCalendarForPractitioner(welkinUser);

                if (string.IsNullOrEmpty(userName) || calendar == null)
                {
                    continue;
                }

                welkinUserNames.Add(userName);
                welkinUserNamesByCalendarId[calendar.Id] = userName;
            }

            IEnumerable<WelkinEvent> welkinEvents = welkinClient.GetEventsUpdatedSince(historySpan);
            Dictionary<string, Dictionary<string, WelkinEvent>> welkinEventsByUserNameThenEventId = new Dictionary<string, Dictionary<string, WelkinEvent>>();

            // Cache all recently updated Welkin events, keyed by user name first then by event ID.
            foreach (WelkinEvent welkinEvent in welkinEvents)
            {
                string userName = welkinUserNamesByCalendarId[welkinEvent.CalendarId];

                if (string.IsNullOrEmpty(userName))
                {
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
                if (string.IsNullOrEmpty(userName) || !welkinUserNames.Contains(userName))
                {
                    continue;
                }

                // For users in both Welkin and Outlook, sync their recently updated events
                try
                {
                    IEnumerable<Event> recentlyUpdatedOutlookEvents = outlookClient.GetEventsForUserUpdatedSince(user, TimeSpan.FromDays(7), Constants.OutlookExtensionsNamespace);
                    foreach (Event evt in recentlyUpdatedOutlookEvents)
                    {
                        log.LogInformation($"Found newly updated Outlook event '{evt.Subject}' for user {userName}.");
                        if (evt.Extensions != null) // This Outlook event is already sync'ed with Welkin and needs to be updated there
                        {
                            if (evt.Extensions.AdditionalData != null && evt.Extensions.AdditionalData.ContainsKey(Constants.LinkedWelkinEventIdKey))
                            {
                                WelkinEvent linkedWelkinEvent;
                            }
                            else
                            {
                                log.LogError($"Outlook event {evt.Id} missing expected extension data.")
                            }
                            foreach(Extension ext in evt.Extensions)
                            {
                                log.LogInformation($"Found extension {ext.ToString()}");
                            }
                        }
                    }
                }
                catch (Exception e)
                {
                    log.LogError(e, $"While retrieving Outlook events for user {userName}.");
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
