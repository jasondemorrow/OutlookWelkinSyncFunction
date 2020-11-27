namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    public abstract class WelkinSyncTask
    {
        protected readonly WelkinEvent welkinEvent;
        protected readonly OutlookClient outlookClient;
        protected readonly WelkinClient welkinClient;

        protected WelkinSyncTask(WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient)
        {
            this.welkinEvent = welkinEvent;
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
        }

        /// <summary>
        /// Perform sync of the given Welkin event to the corresponding Outlook calendar.
        /// </summary>
        /// <returns>The Outlook event created or updated as the result of the sync.</returns>
        public abstract Event Sync();
    }
}