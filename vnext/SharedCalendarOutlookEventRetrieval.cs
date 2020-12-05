namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    public abstract class SharedCalendarOutlookEventRetrieval : OutlookEventRetrieval
    {
        private readonly string sharedCalendarUser;
        private readonly string sharedCalendarName;

        public SharedCalendarOutlookEventRetrieval(
            OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.SharedCalUserEnvVarName)] string sharedCalendarUser,
            [Named(Constants.SharedCalNameEnvVarName)] string sharedCalendarName)
        : base(outlookClient, welkinClient, logger)
        {
            this.sharedCalendarUser = sharedCalendarUser;
            this.sharedCalendarName = sharedCalendarName;
        }

        public override IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago) => null;
    }
}