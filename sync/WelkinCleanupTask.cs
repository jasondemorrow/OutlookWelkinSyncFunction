namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;

    public class WelkinCleanupTask
    {
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;
        private readonly ILogger logger;

        public WelkinCleanupTask(OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.logger = logger;
        }

        public void FindAndDeleteOrphanedPlaceholderEventsScheduledBetween(DateTimeOffset start, DateTimeOffset end)
        {
            IEnumerable<WelkinExternalId> mappings = this.welkinClient.FindExternalEventMappingsUpdatedBetween(start, end);
            foreach(WelkinExternalId mapping in mappings)
            {
                WelkinEvent welkinEvent = null;
                try
                {
                    welkinEvent = this.welkinClient.RetrieveEvent(mapping.InternalId);
                }
                catch (Exception e)
                {
                    this.logger.LogError(e, $"Exception while retrieving Welkin event {mapping.InternalId} from mapping.");
                }

                if (welkinEvent != null && this.welkinClient.IsPlaceHolderEvent(welkinEvent))
                {
                    string outlookICalId = mapping.Namespace.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);
                    //User outlookUser = this.outlookClient.FindUserCorrespondingTo(worker);
                    //syncedTo = this.outlookClient.RetrieveEventWithICalId(outlookUser, outlookICalId);
                }
            }
        }
    }
}