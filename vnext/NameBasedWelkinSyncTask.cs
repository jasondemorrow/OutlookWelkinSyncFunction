namespace OutlookWelkinSync
{
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    /// <summary>
    /// For the welkin event given, look for a linked outlook event and sync if it exists. 
    /// If not, get user that created the welkin event. If they have an outlook user with 
    /// the same user name, create a new, corresponding event in that outlook user's 
    /// calendar and link it with the welkin event.
    /// </summary>
    public class NameBasedWelkinSyncTask : WelkinSyncTask
    {
        public NameBasedWelkinSyncTask(WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger) 
        : base(welkinEvent, outlookClient, welkinClient, logger)
        {
        }

        public override Event Sync()
        {
            Event syncedTo = null;

            if (!this.ShouldSync())
            {
                this.logger.LogInformation($"Not going to sync Welkin event {this.welkinEvent}.");
                return syncedTo;
            }

            WelkinExternalId externalId = this.welkinClient.FindExternalMappingFor(this.welkinEvent);
            WelkinCalendar calendar = this.welkinClient.RetrieveCalendar(this.welkinEvent.CalendarId);
            WelkinWorker worker = this.welkinClient.RetrieveWorker(calendar.WorkerId);

            // If there's already an Outlook event linked to this Welkin event
            if (externalId != null && !string.IsNullOrEmpty(externalId.Namespace))
            {
                string outlookICalId = externalId.Namespace.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);
                this.logger.LogInformation($"Found Outklook event {outlookICalId} associated with Welkin event {welkinEvent.Id}.");
                // With name-based sync, we require that Welkin user principals are the same as Outlook
                syncedTo = this.outlookClient.RetrieveEventWithICalId(worker.Email, outlookICalId);
                if (this.welkinEvent.SyncWith(syncedTo)) // Welkin needs to be updated
                {
                    this.welkinEvent = this.welkinClient.CreateOrUpdateEvent(this.welkinEvent, this.welkinEvent.Id);
                }
                else // Outlook needs to be updated
                {
                    this.outlookClient.UpdateEvent(syncedTo);
                }
            }
            else // An Outlook event needs to be created and linked
            {
                // This will also create and persist the Outlook->Welkin link
                syncedTo = this.outlookClient.CreateOutlookEventFromWelkinEvent(this.welkinEvent, worker);
                WelkinToOutlookLink welkinToOutlookLink = new WelkinToOutlookLink(
                    this.outlookClient, this.welkinClient, this.welkinEvent, syncedTo, this.logger);

                if (!welkinToOutlookLink.CreateIfMissing())
                {
                    // Failed for some reason, need to roll back
                    this.outlookClient.DeleteEvent(syncedTo);
                    throw new LinkException(
                        $"Failed to create link from Welkin event {this.welkinEvent.Id} " +
                        $"to Outlook event {syncedTo.ICalUId}.");
                }
            }

            return syncedTo;
        }
    }
}