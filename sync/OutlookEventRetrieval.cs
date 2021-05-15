namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    public class OutlookEventRetrieval
    {
        protected readonly OutlookClient outlookClient;
        protected readonly IWelkinClient welkinClient;
        protected readonly ILogger logger;

        protected OutlookEventRetrieval(OutlookClient outlookClient, IWelkinClient welkinClient, ILogger logger)
        {
            Throw.IfAnyAreNull(outlookClient, welkinClient, logger);
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.logger = logger;
        }

        public virtual IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago)
        {
            throw new System.NotImplementedException();
        }

        public virtual IEnumerable<Event> RetrieveAllOrphanedBetween(DateTimeOffset start, DateTimeOffset end)
        {
            throw new System.NotImplementedException();
        }

        public override string ToString()
        {
            return $"{this.GetType().FullName}";
        }
    }
}