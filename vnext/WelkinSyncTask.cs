namespace OutlookWelkinSync
{
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public abstract class WelkinSyncTask
    {
        protected readonly WelkinEvent welkinEvent;
        protected readonly OutlookClient outlookClient;
        protected readonly WelkinClient welkinClient;
        protected readonly ILogger logger;

        protected WelkinSyncTask(WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        {
            this.welkinEvent = welkinEvent;
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.logger = logger;
        }

        /// <summary>
        /// Perform sync of the given Welkin event to the corresponding Outlook calendar.
        /// </summary>
        /// <returns>The Outlook event created or updated as the result of the sync, or null if no sync need be performed.</returns>
        public abstract Event Sync();
    }
}