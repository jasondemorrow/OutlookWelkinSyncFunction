namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    public abstract class WelkinWorkerOutlookEventRetrieval : OutlookEventRetrieval
    {
        public WelkinWorkerOutlookEventRetrieval(OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        : base(outlookClient, welkinClient, logger)
        {
        }

        public override IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago) => null;
    }
}