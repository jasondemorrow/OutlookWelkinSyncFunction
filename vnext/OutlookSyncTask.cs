namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    public abstract class OutlookSyncTask
    {
        protected readonly Event outlookEvent;
        protected readonly OutlookClient outlookClient;
        protected readonly WelkinClient welkinClient;

        protected OutlookSyncTask(Event outlookEvent, OutlookClient outlookClient, WelkinClient welkinClient)
        {
            this.outlookEvent = outlookEvent;
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
        }

        /// <summary>
        /// Perform sync of the given Outlook event to the corresponding Welkin schedule.
        /// </summary>
        /// <returns>The Welkin event created or updated as the result of the sync.</returns>
        public abstract WelkinEvent Sync();
    }
}