namespace OutlookWelkinSync
{
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    /// <summary>
    /// For the outlook event given, look for a linked welkin event and sync if it exists. 
    /// If not, get user that created the outlook event. If they have a welkin user with 
    /// the same user name, create a new, corresponding event in that welkin user's 
    /// schedule and link it with the outlook event.
    /// </summary>
    public class NameBasedOutlookSyncTask : OutlookSyncTask
    {
        public NameBasedOutlookSyncTask(Event outlookEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger) 
        : base(outlookEvent, outlookClient, welkinClient, logger)
        {
        }

        public override WelkinEvent Sync()
        {
            WelkinEvent syncedTo = null;
            if (!this.ShouldSync())
            {
                return syncedTo;
            }

            string linkedWelkinEventId = this.outlookClient.LinkedWelkinEventIdFrom(this.outlookEvent);
            if (!string.IsNullOrEmpty(linkedWelkinEventId))
            {
                syncedTo = this.welkinClient.RetrieveEvent(linkedWelkinEventId);
                if (syncedTo.SyncWith(this.outlookEvent)) // Welkin needs to be updated
                {
                    syncedTo = this.welkinClient.CreateOrUpdateEvent(syncedTo, syncedTo.Id);
                }
                else // Outlook needs to be updated
                {
                    this.outlookClient.UpdateEvent(this.outlookEvent);
                }
            }
            else // Welkin needs to be created
            {
                // Find the Welkin user and calendar for the Outlook event owner
                string eventOwnerEmail = this.outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString();
                WelkinWorker worker = this.welkinClient.FindWorker(eventOwnerEmail);
                WelkinCalendar calendar = this.welkinClient.RetrieveCalendarFor(worker);
                Throw.IfAnyAreNull(eventOwnerEmail, worker, calendar);

                // Generate and save a placeholder event in Welkin with a dummy patient TODO: Need to check first not placeholder
                WelkinEvent placeholderEvent = this.welkinClient.GeneratePlaceholderEventForCalendar(calendar);
                placeholderEvent.SyncWith(this.outlookEvent);
                placeholderEvent = this.welkinClient.CreateOrUpdateEvent(placeholderEvent, placeholderEvent.Id);

                // Link the Outlook and Welkin events using external metadata fields
                OutlookToWelkinLink outlookToWelkinLink = new OutlookToWelkinLink(
                    this.outlookClient, this.welkinClient, this.outlookEvent, placeholderEvent, this.logger);

                if (outlookToWelkinLink.CreateIfMissing())
                {
                    // Link did not previously exist and needs to be created from Welkin to Outlook as well
                    WelkinToOutlookLink welkinToOutlookLink = new WelkinToOutlookLink(
                        this.outlookClient, this.welkinClient, placeholderEvent, this.outlookEvent, this.logger);
                    
                    if (!welkinToOutlookLink.CreateIfMissing())
                    {
                        outlookToWelkinLink.Rollback();
                        throw new LinkException(
                            $"Failed to create link from Welkin event {placeholderEvent.Id} " +
                            $"to Outlook event {this.outlookEvent.ICalUId}.");
                    }

                    syncedTo = placeholderEvent;
                }
            }

            return syncedTo;
        }
    }
}