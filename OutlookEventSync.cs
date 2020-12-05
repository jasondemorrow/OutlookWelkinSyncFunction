using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OutlookWelkinSyncFunction
{
    public class OutlookEventSync
    {
        private readonly ILogger log;
        private readonly OutlookClientOld outlookClient;
        private readonly WelkinClientOld welkinClient;

        public OutlookEventSync(OutlookClientOld outlookClient, WelkinClientOld welkinClient, ILogger log)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.log = log;
        }
        
        public string Sync(Event outlookEvent,
                         User outlookUser,
                         WelkinPractitioner practitioner,
                         string welkinCalendarId,
                         string calendarName = null)
        {
            string syncId = null;
            bool updateSyncTime = true;
            bool placeholderEventCreated = false;
            DateTimeOffset presumptiveLastSyncTime = DateTimeOffset.UtcNow; // must be before any update
            EventLink eventLink = new EventLink(outlookEvent, null, outlookClient, welkinClient, outlookUser, practitioner, log);

            try
            {
                if (OutlookClientOld.IsPlaceHolderEvent(outlookEvent))
                {
                    log.LogInformation("This is a placeholder event created for a Welkin event. Skipping...");
                    return syncId;
                }

                DateTime? lastSync = OutlookClientOld.GetLastSyncDateTime(outlookEvent);
                if (lastSync != null && 
                    outlookEvent.LastModifiedDateTime != null && 
                    lastSync.Value >= outlookEvent.LastModifiedDateTime.Value.UtcDateTime)
                {
                    log.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                    updateSyncTime = false;
                    return syncId;
                }

                if (!eventLink.FetchAndPopulateIfExists(EventLink.Direction.OutlookToWelkin))
                {
                    WelkinEvent placeholderEvent = welkinClient.GeneratePlaceholderEventForCalendar(welkinCalendarId);
                    placeholderEvent.SyncWith(outlookEvent);
                    eventLink.TargetWelkinEvent = 
                        welkinClient.CreateOrUpdateEvent(placeholderEvent, placeholderEvent.Id);
                    placeholderEventCreated = (eventLink.TargetWelkinEvent != null);
                    eventLink.Ensure(EventLink.Direction.OutlookToWelkin); // if successful, LinkedWelkinEvent will be non-null
                }

                bool welkinEventNeedsUpdate = !placeholderEventCreated && 
                                              eventLink.LinkedWelkinEvent != null &&
                                              eventLink.LinkedWelkinEvent.SyncWith(outlookEvent);
                if (welkinEventNeedsUpdate)
                {
                    welkinClient.CreateOrUpdateEvent(eventLink.LinkedWelkinEvent, eventLink.LinkedWelkinEvent.Id);
                }
                else if (!placeholderEventCreated) // Outlook event needs update
                {
                    Event updatedEvent = outlookClient.Update(outlookUser, outlookEvent, calendarName);
                }

                log.LogInformation(
                    $"Successfully sync'ed Outlook event {outlookEvent.ICalUId} with Welkin event {eventLink.LinkedWelkinEvent.Id}.");
                syncId = eventLink.LinkedWelkinEvent.Id;
            }
            catch (Exception e)
            {
                if (placeholderEventCreated)
                {
                    log.LogInformation("Deleting created placeholder event in Welkin...");
                    this.welkinClient.DeleteEvent(eventLink.TargetWelkinEvent);
                }
                throw e;
            }
            finally
            {
                if (updateSyncTime)
                {
                    outlookClient.SetLastSyncDateTime(outlookUser, outlookEvent, presumptiveLastSyncTime);
                }
            }

            return syncId;
        }
    }
}