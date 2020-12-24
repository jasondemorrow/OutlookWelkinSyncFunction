namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class WelkinSyncTask
    {
        protected WelkinEvent welkinEvent;
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
        public virtual Event Sync()
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// Perform standard pre-sync checks.
        /// </summary>
        /// <returns>Whether sync should continue.</returns>
        protected bool ShouldSync()
        {
            Throw.IfAnyAreNull(this.outlookClient, this.welkinClient, this.welkinEvent);

            if (this.welkinClient.IsPlaceHolderEvent(this.welkinEvent))
            {
                this.logger.LogInformation($"Welkin event {this.welkinEvent.Id} is a placeholder event created in Welkin for an Outlook event. Skipping...");
                return false;
            }

            WelkinLastSyncEntry lastSync = welkinClient.RetrieveLastSyncFor(welkinEvent);
            if (lastSync != null && lastSync.IsValid() && this.welkinEvent.Updated != null && 
                lastSync.Time >= this.welkinEvent.Updated.Value.ToUniversalTime())
            {
                this.logger.LogInformation($"Welkin event {this.welkinEvent.Id} hasn't been updated since its last sync. Skipping...");
                return false;
            }

            return true;
        }
        
        public override string ToString()
        {
            return $"{this.GetType().FullName} for {this.welkinEvent.Id}";
        }
    }
}