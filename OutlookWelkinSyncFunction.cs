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

                // For users in both Welkin and Outlook, sync their recently updated events
                try
                {
                    // First, sync newly update Outlook events for user.
                    IEnumerable<Event> recentlyUpdatedOutlookEvents = 
                        outlookClient.GetEventsForUserUpdatedSince(user, historySpan, Constants.OutlookEventExtensionsNamespace);
                    foreach (Event evt in recentlyUpdatedOutlookEvents)
                    {
                        try
                        {
                            WelkinEvent linkedWelkinEvent = WelkinEvent.CreateDefaultForCalendar(welkinCalendarIdsByUserName[userName]);
                            bool isNew = true;
                            log.LogInformation($"Found newly updated Outlook event '{evt.Id}' for user {userName}.");

                            if (evt.Extensions != null) // This Outlook event is already sync'ed with Welkin and needs to be updated there
                            {
                                if (evt.Extensions.AdditionalData != null && evt.Extensions.AdditionalData.ContainsKey(Constants.LinkedWelkinEventIdKey))
                                {
                                    string linkedEventId = evt.Extensions.AdditionalData[Constants.LinkedWelkinEventIdKey].ToString();
                                    linkedWelkinEvent = welkinClient.GetEvent(linkedEventId);
                                    isNew = false;
                                }
                                else
                                {
                                    log.LogError($"Outlook event {evt.Id} for user {userName} is missing expected extension data.");
                                }
                            }

                            linkedWelkinEvent.SyncWith(evt); // Will set start and end date-times of both events to match whichever event was more recently updated
                            linkedWelkinEvent = welkinClient.CreateOrUpdateEvent(linkedWelkinEvent, isNew);

                            if (isNew) // Associate new Welkin event with current Outlook event
                            {
                                // First by an extension property on the Outlook event
                                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                                keyValuePairs[Constants.LinkedWelkinEventIdKey] = linkedWelkinEvent.Id;
                                outlookClient.SetOpenExtensionPropertiesOnEvent(user, evt, keyValuePairs, Constants.OutlookEventExtensionsNamespace);
                                // Then by an external ID in Welkin
                                WelkinExternalId welkinExternalId = new WelkinExternalId
                                {
                                    Resource = Constants.CalendarEventResourceName,
                                    ExternalId = evt.Id,
                                    InternalId = linkedWelkinEvent.Id,
                                    Namespace = Constants.WelkinEventExtensionNamespace
                                };
                                welkinExternalId = welkinClient.CreateOrUpdateExternalId(welkinExternalId, true);
                            }
                            else if (welkinEventsByUserNameThenEventId[userName].ContainsKey(linkedWelkinEvent.Id))
                            {
                                // If the existing Welkin event has also been recently updated, we can skip it later
                                welkinEventsByUserNameThenEventId[userName].Remove(linkedWelkinEvent.Id);
                                log.LogInformation($@"Welkin event with ID {linkedWelkinEvent.Id} has recently been updated, 
                                                        but will be skipped since its corresponding Outlook event with ID 
                                                        {evt.Id} has also been recently updated and therefore sync'ed.");
                            }

                            string newOrExisting = isNew ? "new" : "existing";
                            log.LogInformation($"Successfully sync'ed Outlook event {evt.Id} with {newOrExisting} Welkin event {linkedWelkinEvent.Id}.");
                        }
                        catch (Exception e)
                        {
                            log.LogError(e, $"While sync'ing Outlook event {evt.Id} for user {userName}.");
                        }
                    }

                    // Second, sync newly updated Welkin events for user
                    foreach (WelkinEvent evt in welkinEventsByUserNameThenEventId[userName].Values)
                    {
                        WelkinExternalId externalId = welkinClient.FindExternalMappingFor(evt);
                        Event linkedOutlookEvent = null;

                        // This Welkin event is already associated with an Outlook event, let's retrieve it
                        if (externalId != null)
                        {
                            linkedOutlookEvent = outlookClient.GetEventForUserWithId(user, externalId.ExternalId);
                        }

                        string verb = "Retrieved";

                        // There is no associated Outlook event, let's create and link it
                        if (linkedOutlookEvent == null)
                        {
                            linkedOutlookEvent = outlookClient.CreateOutlookEventFromWelkinEvent(user, evt, welkinPractitionerByUserName[userName]);
                            verb = "Created";
                        }

                        log.LogInformation($"{verb} Outlook event with ID {linkedOutlookEvent.Id} associated with Welkin event (ID {evt.Id}).");

                        evt.SyncWith(linkedOutlookEvent);
                        welkinClient.CreateOrUpdateEvent(evt, false);
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
