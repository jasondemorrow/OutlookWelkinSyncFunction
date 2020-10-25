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
            DateTime lastRun = timerInfo?.ScheduleStatus?.Last ?? DateTime.UtcNow.AddHours(-24);
            TimeSpan historySpan = DateTime.UtcNow - lastRun;
            IEnumerable<User> outlookUsers = new List<User>();
            IEnumerable<WelkinPractitioner> welkinUsers = new List<WelkinPractitioner>();

            try
            {
                // TODO: Obviously the following doesn't scale well. Factoring this out and adding pagination will help.
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
                    EventLink eventLink = new EventLink(null, null, outlookClient, welkinClient, user, welkinPractitionerByUserName[userName], log);
                    // First, sync newly updated Outlook events for user.
                    IEnumerable<Event> recentlyUpdatedOutlookEvents = 
                        outlookClient.GetEventsForUserUpdatedSince(user, historySpan, Constants.OutlookEventExtensionsNamespace);
                    foreach (Event evt in recentlyUpdatedOutlookEvents)
                    {
                        log.LogInformation($"Found newly updated Outlook event '{evt.ICalUId}' for user {userName}.");
                        if (OutlookClient.IsPlaceHolderEvent(evt))
                        {
                            log.LogInformation("This is a placeholder event created for a Welkin event. Skipping...");
                            continue;
                        }

                        DateTime? lastSync = OutlookClient.GetLastSyncDateTime(evt);
                        if (lastSync != null && evt.LastModifiedDateTime != null && lastSync.Value >= evt.LastModifiedDateTime.Value.UtcDateTime)
                        {
                            log.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                            continue;
                        }
                        
                        eventLink.Clear();

                        try
                        {
                            eventLink.TargetOutlookEvent = evt;
                            bool createdPlaceholderWelkinEvent = false;
                            if (!eventLink.Exists(EventLink.Direction.OutlookToWelkin))
                            {
                                try
                                {
                                    WelkinEvent placeholderEvent = WelkinEvent.CreateDefaultForCalendar(welkinCalendarIdsByUserName[userName]);
                                    placeholderEvent.SyncWith(evt);
                                    eventLink.TargetWelkinEvent = welkinClient.CreateOrUpdateEvent(placeholderEvent);
                                    createdPlaceholderWelkinEvent = (eventLink.TargetWelkinEvent != null);
                                    eventLink.Ensure(EventLink.Direction.OutlookToWelkin);
                                }
                                catch (Exception e)
                                {
                                    string trace = Exceptions.ToStringRecursively(e);
                                    log.LogError($"While ensuring Outlook to Welkin link for Outlook event {evt.ICalUId}: {trace}");
                                    if (createdPlaceholderWelkinEvent)
                                    {
                                        welkinClient.DeleteEvent(eventLink.TargetWelkinEvent);
                                        eventLink.TargetWelkinEvent = null;
                                    }
                                }
                            }

                            bool welkinEventNeedsUpdate = !createdPlaceholderWelkinEvent && eventLink.LinkedWelkinEvent.SyncWith(evt);
                            if (welkinEventNeedsUpdate)
                            {
                                welkinClient.CreateOrUpdateEvent(eventLink.LinkedWelkinEvent, eventLink.LinkedWelkinEvent.Id);
                            }
                            else if (!createdPlaceholderWelkinEvent) // Outlook event needs update
                            {
                                Event updatedEvent = outlookClient.Update(user, evt);
                            }
                            
                            if (welkinEventsByUserNameThenEventId.ContainsKey(userName) && 
                                welkinEventsByUserNameThenEventId[userName].ContainsKey(eventLink.LinkedWelkinEvent.Id))
                            {
                                // If the existing Welkin event has also been recently updated, we can skip it later
                                welkinEventsByUserNameThenEventId[userName].Remove(eventLink.LinkedWelkinEvent.Id);
                                log.LogInformation($@"Welkin event with ID {eventLink.LinkedWelkinEvent.Id} has recently been updated, " +
                                                    "but will be skipped since its corresponding Outlook event with ID {evt.ICalUId} has " + 
                                                    "also been recently updated and therefore sync'ed.");
                            }

                            outlookClient.SetLastSyncDateTime(user, evt);
                            log.LogInformation($"Successfully sync'ed Outlook event {evt.ICalUId} with Welkin event {eventLink.LinkedWelkinEvent.Id}.");
                        }
                        catch (Exception e)
                        {
                            string trace = Exceptions.ToStringRecursively(e);
                            log.LogError($"While sync'ing Outlook event {evt.ICalUId} for user {userName}: {trace}");
                        }
                    }

                    // Second, sync newly updated Welkin events for user, if any
                    if (welkinEventsByUserNameThenEventId.ContainsKey(userName))
                    {
                        foreach (WelkinEvent evt in welkinEventsByUserNameThenEventId[userName].Values)
                        {
                            log.LogInformation($"Found newly updated Welkin event '{evt.Id}' for user {userName}.");
                            if (welkinClient.IsPlaceHolderEvent(evt))
                            {
                                log.LogInformation("This is a placeholder event created for an Outlook event. Skipping...");
                                continue;
                            }

                            DateTime? lastSync = welkinClient.FindLastSyncDateTimeFor(evt);
                            if (lastSync != null && evt.Updated != null && lastSync.Value >= evt.Updated.Value)
                            {
                                log.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                                continue;
                            }

                            eventLink.Clear();

                            eventLink.TargetWelkinEvent = evt;
                            bool createdPlaceholderOutlookEvent = false;
                            if (!eventLink.Exists(EventLink.Direction.WelkinToOutlook))
                            {
                                eventLink.TargetOutlookEvent = 
                                    outlookClient.CreateOutlookEventFromWelkinEvent(user, evt, welkinPractitionerByUserName[userName]);
                                createdPlaceholderOutlookEvent = true;
                                eventLink.Ensure(EventLink.Direction.WelkinToOutlook);
                            }

                            log.LogInformation($"Outlook event with ID {eventLink.LinkedOutlookEvent.ICalUId} associated with Welkin event {evt.Id}.");

                            if (evt.SyncWith(eventLink.LinkedOutlookEvent))
                            {
                                welkinClient.CreateOrUpdateEvent(evt, evt.Id);
                                welkinClient.SetLastSyncDateTimeFor(evt);
                            }
                            else if (!createdPlaceholderOutlookEvent)
                            {
                                outlookClient.Update(user, eventLink.LinkedOutlookEvent);
                            }

                            log.LogInformation($"Successfully sync'ed Welkin event {evt.Id} with Outlook event {eventLink.LinkedOutlookEvent.ICalUId}.");
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
