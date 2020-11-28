namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    /// <summary>
    /// For the outlook event given, look for a linked welkin event and sync if it exists. 
    /// If not, get user that created the outlook event. If they have a welkin user with 
    /// the same user name, create a new, corresponding event in that welkin user's 
    /// schedule and link it with the outlook event.
    /// </summary>
    public class NameBasedOutlookSyncTask : OutlookSyncTask
    {
        public NameBasedOutlookSyncTask(Event outlookEvent, OutlookClient outlookClient, WelkinClient welkinClient) 
        : base(outlookEvent, outlookClient, welkinClient)
        {
        }

        public override WelkinEvent Sync()
        {
            Throw.IfAnyAreNull(this.outlookClient, this.welkinClient, this.outlookEvent);
            WelkinEvent syncedTo = null;
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
                    this.outlookClient.Update(this.outlookEvent);
                }
            }
            else // Welkin needs to be created
            {
                string eventOwnerEmail = this.outlookEvent.Calendar.Owner.Address;
                WelkinWorker worker = this.welkinClient.FindWorker(eventOwnerEmail);
                WelkinCalendar calendar = this.welkinClient.RetrieveCalendarFor(worker);
                Throw.IfAnyAreNull(eventOwnerEmail, worker, calendar);
                WelkinEvent placeholderEvent = this.welkinClient.GeneratePlaceholderEventForCalendar(calendar);
                placeholderEvent.SyncWith(this.outlookEvent);
            }

            return syncedTo;
        }
    }
}