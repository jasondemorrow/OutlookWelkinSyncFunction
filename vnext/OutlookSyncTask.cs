namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class OutlookSyncTask
    {
        protected Event outlookEvent;
        protected readonly OutlookClient outlookClient;
        protected readonly WelkinClient welkinClient;
        protected readonly ILogger logger;

        protected OutlookSyncTask(Event outlookEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        {
            this.outlookEvent = outlookEvent;
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.logger = logger;
        }

        /// <summary>
        /// Perform sync of the given Outlook event to the corresponding Welkin schedule.
        /// </summary>
        /// <returns>The Welkin event created or updated as the result of the sync, or null if no sync need be performed.</returns>
        public virtual WelkinEvent Sync()
        {
            throw new System.NotImplementedException();
        }

        /// <summary>
        /// Perform standard pre-sync checks.
        /// </summary>
        /// <returns>Whether sync should continue.</returns>
        protected bool ShouldSync()
        {
            Throw.IfAnyAreNull(this.outlookClient, this.welkinClient, this.outlookEvent);

            // If this is a placeholder event created during Welkin sync, we don't sync it.
            if (OutlookClient.IsPlaceHolderEvent(outlookEvent))
            {
                this.logger.LogInformation("This is a placeholder event created for a Welkin event. Skipping...");
                return false;
            }

            DateTime? lastSync = OutlookClient.GetLastSyncDateTime(outlookEvent);
            if (lastSync != null && 
                outlookEvent.LastModifiedDateTime != null && 
                lastSync.Value >= outlookEvent.LastModifiedDateTime.Value.UtcDateTime)
            {
                this.logger.LogInformation("This event hasn't been updated since its last sync. Skipping...");
                return false;
            }

            return true;
        }

        public override string ToString()
        {
            return $"{this.GetType().FullName} for {this.outlookEvent.ICalUId}";
        }
    }
}