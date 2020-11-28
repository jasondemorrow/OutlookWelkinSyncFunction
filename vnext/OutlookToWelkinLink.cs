namespace OutlookWelkinSync
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class OutlookToWelkinLink
    {
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;
        private readonly Event sourceOutlookEvent;
        private readonly WelkinEvent targetWelkinEvent;
        protected readonly ILogger logger;

        public OutlookToWelkinLink(OutlookClient outlookClient, WelkinClient welkinClient, Event sourceOutlookEvent, WelkinEvent targetWelkinEvent, ILogger logger)
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
                keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = linkedWelkinId;
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
    }
}