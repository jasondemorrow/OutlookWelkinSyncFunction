namespace OutlookWelkinSync
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class OutlookToWelkinLink
    {
        private readonly OutlookClient outlookClient;
        private readonly IWelkinClient welkinClient;
        private readonly Event sourceOutlookEvent;
        private readonly WelkinEvent targetWelkinEvent;
        protected readonly ILogger logger;

        public OutlookToWelkinLink(OutlookClient outlookClient, IWelkinClient welkinClient, Event sourceOutlookEvent, WelkinEvent targetWelkinEvent, ILogger logger)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.sourceOutlookEvent = sourceOutlookEvent;
            this.targetWelkinEvent = targetWelkinEvent;
            this.logger = logger;
        }

        private string GetCurrentlyLinkedWelkinEventId()
        {
            Extension extensionForWelkin = this.sourceOutlookEvent.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData == null || !extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLinkedWelkinEventIdKey))
            {
                this.logger.LogInformation($"No linked Welkin event for Outlook event {this.sourceOutlookEvent.ICalUId}");
                return null;
            }

            string linkedEventId = extensionForWelkin.AdditionalData[Constants.OutlookLinkedWelkinEventIdKey]?.ToString();
            if (string.IsNullOrEmpty(linkedEventId))
            {
                this.logger.LogInformation($"Null or empty linked Welkin event ID for Outlook event {this.sourceOutlookEvent.ICalUId}");
                return null;
            }

            return linkedEventId;
        }

        /// <summary>
        /// Create a link from the source Outlook event to the destination Welkin event if not already linked.
        /// </summary>
        /// <returns>True if a new link was created, otherwise false.</returns>
        public bool CreateIfMissing()
        {
            string linkedWelkinId = this.GetCurrentlyLinkedWelkinEventId();

            if (string.IsNullOrEmpty(linkedWelkinId))
            {
                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = this.targetWelkinEvent.Id;
                this.outlookClient.MergeOpenExtensionPropertiesOnEvent(
                    this.sourceOutlookEvent, 
                    keyValuePairs, 
                    Constants.OutlookEventExtensionsNamespace);
                string msg = $"Successfully created link from Outlook event {this.sourceOutlookEvent.ICalUId} to Welkin event {linkedWelkinId}";
                this.logger.LogInformation(msg);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Roll back any changes to the target Outlook event's Welkin link.
        /// </summary>
        public void Rollback()
        {
            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            // TODO: This isn't so much a rollback as an erasure. Would need to keep state to make this more accurate.
            keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = ""; // Empty string => not set (unset)
            this.outlookClient.MergeOpenExtensionPropertiesOnEvent(
                this.sourceOutlookEvent, 
                keyValuePairs, 
                Constants.OutlookEventExtensionsNamespace);
            string msg = $"Successfully removed link from Outlook event {this.sourceOutlookEvent.ICalUId} to any Welkin event";
            this.logger.LogInformation(msg);
        }
    }
}