using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OutlookWelkinSyncFunction
{
    public class OutlookEventSync
    {
        private readonly ILogger log;
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;

        public OutlookEventSync(OutlookClient outlookClient, WelkinClient welkinClient, ILogger log)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.log = log;
        }
        
        public void Sync(Event outlookEvent,
                         User outlookUser,
                         WelkinPractitioner practitioner,
                         Dictionary<string, Dictionary<string, WelkinEvent>> welkinEventsByUserNameThenEventId,
                         string welkinCalendarId,
                         string commonUserName)
        {
            WelkinEvent placeholderEvent = null;
            try
            {
                log.LogInformation($"Found newly updated Outlook event '{outlookEvent.ICalUId}' for user {commonUserName}.");
                if (OutlookClient.IsPlaceHolderEvent(outlookEvent))
                {
                    log.LogInformation("This is a placeholder event created for a Welkin event. Skipping...");
                    return;
                }

                DateTime? lastSync = OutlookClient.GetLastSyncDateTime(outlookEvent);
                if (lastSync != null && 
                    outlookEvent.LastModifiedDateTime != null && 
                    lastSync.Value >= outlookEvent.LastModifiedDateTime.Value.UtcDateTime)
                {
                    log.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                    return;
                }
                        
                EventLink eventLink = 
                    new EventLink(outlookEvent, null, outlookClient, welkinClient, outlookUser, practitioner, log);
                bool createdPlaceholderWelkinEvent = false;
                if (!eventLink.FetchAndPopulateIfExists(EventLink.Direction.OutlookToWelkin))
                {
                    try
                    {
                        placeholderEvent = welkinClient.GeneratePlaceholderEventForCalendar(welkinCalendarId);
                        placeholderEvent.SyncWith(outlookEvent);
                        eventLink.TargetWelkinEvent = 
                            welkinClient.CreateOrUpdateEvent(placeholderEvent, placeholderEvent.Id);
                        createdPlaceholderWelkinEvent = (eventLink.TargetWelkinEvent != null);
                        eventLink.Ensure(EventLink.Direction.OutlookToWelkin);
                    }
                    catch (Exception e)
                    {
                        string trace = Exceptions.ToStringRecursively(e);
                        log.LogError(
                            $"While ensuring Outlook to Welkin link for Outlook event {outlookEvent.ICalUId}: {trace}");
                        if (createdPlaceholderWelkinEvent)
                        {
                            welkinClient.DeleteEvent(eventLink.TargetWelkinEvent);
                            eventLink.TargetWelkinEvent = null;
                        }
                    }
                }

                bool welkinEventNeedsUpdate = !createdPlaceholderWelkinEvent && 
                                              eventLink.LinkedWelkinEvent.SyncWith(outlookEvent);
                if (welkinEventNeedsUpdate)
                {
                    welkinClient.CreateOrUpdateEvent(eventLink.LinkedWelkinEvent, eventLink.LinkedWelkinEvent.Id);
                }
                else if (!createdPlaceholderWelkinEvent) // Outlook event needs update
                {
                    Event updatedEvent = outlookClient.Update(outlookUser, outlookEvent);
                }
                
                if (welkinEventsByUserNameThenEventId.ContainsKey(commonUserName) && 
                    welkinEventsByUserNameThenEventId[commonUserName].ContainsKey(eventLink.LinkedWelkinEvent.Id))
                {
                    // If the existing Welkin event has also been recently updated, we can skip it later
                    welkinEventsByUserNameThenEventId[commonUserName].Remove(eventLink.LinkedWelkinEvent.Id);
                    log.LogInformation($@"Welkin event with ID {eventLink.LinkedWelkinEvent.Id} has recently been updated, " +
                                        "but will be skipped since its corresponding Outlook event with ID {evt.ICalUId} has " + 
                                        "also been recently updated and therefore sync'ed.");
                }

                outlookClient.SetLastSyncDateTime(outlookUser, outlookEvent);
                log.LogInformation(
                    $"Successfully sync'ed Outlook event {outlookEvent.ICalUId} with Welkin event {eventLink.LinkedWelkinEvent.Id}.");
            }
            catch (Exception e)
            {
                string trace = Exceptions.ToStringRecursively(e);
                log.LogError($"Sync failed with Outlook event {outlookEvent.ICalUId} for user {commonUserName}: {trace}");
                if (placeholderEvent != null)
                {
                    log.LogInformation("Deleting created placeholder event in Welkin...");
                    this.welkinClient.DeleteEvent(placeholderEvent);
                }
            }
        }
    }
}