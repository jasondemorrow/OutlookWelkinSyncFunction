namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    public abstract class WelkinSyncTask
    {
        protected readonly WelkinEvent welkinEvent;

        protected WelkinSyncTask(WelkinEvent welkinEvent)
        {
            this.welkinEvent = welkinEvent;
        }

        /// <summary>
        /// Perform sync of the given Welkin event to the corresponding Outlook calendar.
        /// </summary>
        /// <returns>The Outlook event created or updated as the result of the sync.</returns>
        public abstract Event Sync();
    }
}