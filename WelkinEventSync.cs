using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OutlookWelkinSyncFunction
{
    public class WelkinEventSync
    {
        private readonly ILogger log;
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;

        public WelkinEventSync(OutlookClient outlookClient, WelkinClient welkinClient, ILogger log)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.log = log;
        }
        
        public void Sync(WelkinEvent welkinEvent,
                         User outlookUser,
                         WelkinPractitioner practitioner,
                         Dictionary<string, Dictionary<string, WelkinEvent>> welkinEventsByUserNameThenEventId,
                         string welkinCalendarId,
                         string commonUserName)
        {
            log.LogInformation($"Found newly updated Welkin event '{welkinEvent}' for user {commonUserName}.");
            WelkinLastSyncEntry lastSync = null;
            Event createdOutlookEvent = null;
            EventLink eventLink = null;

            // We ignore unavailable times and working hours when updating event. The reason for this is that we want 
            // the user to see when they've scheduled a conflict in Outlook. If we don't sync, they might not see it.
            bool originalIgnoreTimes = welkinEvent.IgnoreUnavailableTimes;
            bool originalIgnoreHours = welkinEvent.IgnoreWorkingHours;
            welkinEvent.IgnoreUnavailableTimes = true;
            welkinEvent.IgnoreWorkingHours = true;

            try
            {
                if (welkinClient.IsPlaceHolderEvent(welkinEvent))
                {
                    log.LogInformation("This is a placeholder event created for an Outlook event. Skipping...");
                    return;
                }

                lastSync = welkinClient.FindLastSyncEntryFor(welkinEvent);
                if (lastSync != null && lastSync.IsValid() && welkinEvent.Updated != null && lastSync.Time >= welkinEvent.Updated.Value)
                {
                    log.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                    return;
                }

                eventLink = new EventLink(null, welkinEvent, outlookClient, welkinClient, outlookUser, practitioner, log);
                eventLink.TargetWelkinEvent = welkinEvent;
                bool createdPlaceholderOutlookEvent = false;
                if (!eventLink.FetchAndPopulateIfExists(EventLink.Direction.WelkinToOutlook))
                {
                    eventLink.TargetOutlookEvent = 
                        outlookClient.CreateOutlookEventFromWelkinEvent(outlookUser, welkinEvent, practitioner);
                    createdPlaceholderOutlookEvent = true;
                    createdOutlookEvent = eventLink.TargetOutlookEvent;
                    eventLink.Ensure(EventLink.Direction.WelkinToOutlook);
                }

                log.LogInformation($"Outlook event with ID {eventLink.LinkedOutlookEvent.ICalUId} associated with Welkin event {welkinEvent}.");

                if (welkinEvent.SyncWith(eventLink.LinkedOutlookEvent))
                {
                    welkinClient.CreateOrUpdateEvent(welkinEvent, welkinEvent.Id);
                }
                else if (!createdPlaceholderOutlookEvent)
                {
                    outlookClient.Update(outlookUser, eventLink.LinkedOutlookEvent);
                }

                log.LogInformation($"Successfully sync'ed Welkin event {welkinEvent} with Outlook event {eventLink.LinkedOutlookEvent.ICalUId}.");
            }
            catch (Exception e)
            {
                string trace = Exceptions.ToStringRecursively(e);
                log.LogError($"Sync failed for Welkin event {welkinEvent}: {trace}");
                if (createdOutlookEvent != null)
                {
                    log.LogInformation("Deleting created placeholder event in Outlook...");
                    this.outlookClient.Delete(outlookUser, createdOutlookEvent);
                }
            }
            finally
            {
                welkinEvent.IgnoreUnavailableTimes = originalIgnoreTimes;
                welkinEvent.IgnoreWorkingHours = originalIgnoreHours;
                string lastSyncEntryId = (lastSync != null && lastSync.IsValid()) ? lastSync.ExternalId.Id : null;
                welkinClient.SetLastSyncDateTimeFor(welkinEvent, lastSyncEntryId);
            }
        }
    }
}