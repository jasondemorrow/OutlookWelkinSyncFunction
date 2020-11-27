namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    public abstract class OutlookSyncTask
    {
        protected readonly Event outlookEvent;

        protected OutlookSyncTask(Event outlookEvent)
        {
            this.outlookEvent = outlookEvent;
        }

        /// <summary>
        /// Perform sync of the given Outlook event to the corresponding Welkin schedule.
        /// </summary>
        /// <returns>The Welkin event created or updated as the result of the sync.</returns>
        public abstract WelkinEvent Sync();
    }
}