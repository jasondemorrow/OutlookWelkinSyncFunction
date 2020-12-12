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
            if (!this.ShouldSync())
            {
                return null;
            }

            WelkinExternalId externalId = this.welkinClient.FindExternalMappingFor(this.welkinEvent);
            WelkinCalendar calendar = this.welkinClient.RetrieveCalendar(this.welkinEvent.CalendarId);
            WelkinWorker worker = this.welkinClient.RetrieveWorker(calendar.WorkerId);
            Event syncedTo = null;

            // If there's already an Outlook event linked to this Welkin event
            if (externalId != null && !string.IsNullOrEmpty(externalId.Namespace))
            {
                string outlookICalId = externalId.Namespace.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);
                this.logger.LogInformation($"Found Outlook event {outlookICalId} associated with Welkin event {welkinEvent.Id}.");
                User outlookUser = this.outlookClient.FindUserCorrespondingTo(worker);
                syncedTo = this.outlookClient.RetrieveEventWithICalId(outlookUser.UserPrincipalName, outlookICalId);
                syncedTo.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser; // TODO: put this part in the client
                // TODO: Sync can mess up start/end time
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

            this.welkinClient.UpdateLastSyncFor(this.welkinEvent, externalId?.Id);
            return syncedTo;
        }
    }
}