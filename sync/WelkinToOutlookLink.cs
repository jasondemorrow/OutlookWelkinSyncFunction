namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class WelkinToOutlookLink
    {
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;
        private readonly WelkinEvent sourceWelkinEvent;
        private readonly Event targetOutlookEvent;
        protected readonly ILogger logger;

        public WelkinToOutlookLink(OutlookClient outlookClient, WelkinClient welkinClient, WelkinEvent sourceWelkinEvent, Event targetOutlookEvent, ILogger logger)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.sourceWelkinEvent = sourceWelkinEvent;
            this.targetOutlookEvent = targetOutlookEvent;
            this.logger = logger;
        }

        /// <summary>
        /// Create a link from the source Welkin event to the destination Outlook event if not already linked.
        /// </summary>
        /// <returns>True if a new link was created, otherwise false.</returns>
        public bool CreateIfMissing()
        {
            WelkinExternalId externalId = welkinClient.FindExternalMappingFor(this.sourceWelkinEvent, this.targetOutlookEvent);

            if (externalId == null || string.IsNullOrEmpty(externalId.ExternalId))
            {
                this.logger.LogInformation($"Linking Welkin event {this.sourceWelkinEvent.Id} to Outlook event {this.targetOutlookEvent.ICalUId}.");
                
                /**
                * Outlook's UUID for calendar events, ICalUId, is not a properly formed GUID, which the Welkin External ID API expects.
                * We use a hashing method to generate a consistent, synthetic GUID for ICalUId, appending its original value to the 
                * namespace string stored in the External ID. This makes the API happy while allowing us to derive the GUID
                * from ICalUId when necessary and vice versa, in a way that is adequately collision-resistent.
                */
                string derivedGuid = Guids.FromText(this.targetOutlookEvent.ICalUId).ToString();
                WelkinExternalId welkinExternalId = new WelkinExternalId
                {
                    Resource = Constants.CalendarEventResourceName,
                    ExternalId = derivedGuid,
                    InternalId = this.sourceWelkinEvent.Id,
                    Namespace = Constants.WelkinEventExtensionNamespacePrefix + this.targetOutlookEvent.ICalUId
                };
                welkinExternalId = welkinClient.CreateOrUpdateExternalId(welkinExternalId); // TODO: Catch exception here and roll back
                string outlookICalId = welkinExternalId?.Namespace?.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);

                if (outlookICalId != null && outlookICalId.Equals(this.targetOutlookEvent.ICalUId))
                {
                    this.logger.LogInformation($"Created link from Welkin event {this.sourceWelkinEvent.Id} to Outlook event {this.targetOutlookEvent.ICalUId}");
                    return true;
                }
            }

            return false;
        }
    }
}