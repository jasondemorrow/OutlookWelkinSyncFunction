namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    public abstract class OutlookEventRetrieval
    {
        protected readonly OutlookClient outlookClient;
        protected readonly WelkinClient welkinClient;
        protected readonly ILogger logger;

        protected OutlookEventRetrieval(OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        {
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.logger = logger;
        }

        public abstract IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago);
    }
}