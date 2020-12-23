namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    public class SharedCalendarOutlookEventRetrieval : OutlookEventRetrieval
    {
        private readonly string sharedCalendarUser;
        private readonly string sharedCalendarName;
        private readonly User sharedCalendarOutlookUser;
        private readonly Calendar sharedOutlookCalendar;

        public SharedCalendarOutlookEventRetrieval(
            OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.SharedCalUserEnvVarName)] string sharedCalendarUser,
            [Named(Constants.SharedCalNameEnvVarName)] string sharedCalendarName)
        : base(outlookClient, welkinClient, logger)
        {
            this.sharedCalendarUser = sharedCalendarUser;
            this.sharedCalendarName = sharedCalendarName;
            this.sharedCalendarOutlookUser = this.outlookClient.RetrieveUser(this.sharedCalendarUser);
            this.sharedOutlookCalendar = this.outlookClient.RetrieveCalendar(this.sharedCalendarUser, this.sharedCalendarName);
        }

        public override IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago)
        {
            /*DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            return this.outlookClient.RetrieveEventsForUserScheduledBetween(
                this.sharedCalendarUser, 
                start, 
                end, 
                null, 
                this.sharedCalendarName);*/
            return new List<Event>(); // Outlook event sync from shared calendar not yet supported
        }

        public override IEnumerable<Event> RetrieveAllOrphanedBetween(DateTimeOffset start, DateTimeOffset end)
        {
            List<Event> orphaned = new List<Event>();
            IEnumerable<Event> events = this.outlookClient.RetrieveEventsForUserScheduledBetween(
                this.sharedCalendarOutlookUser, 
                start, 
                end, 
                Constants.OutlookEventExtensionsNamespace, 
                this.sharedOutlookCalendar.Id);

            foreach (Event outlookEvent in events)
            {
                if (OutlookClient.IsPlaceHolderEvent(outlookEvent))
                {
                    string linkedWelkinEventId = this.outlookClient.LinkedWelkinEventIdFrom(outlookEvent);
                    WelkinEvent syncedTo = null;
                    if (!string.IsNullOrEmpty(linkedWelkinEventId))
                    {
                        try
                        {
                            syncedTo = this.welkinClient.RetrieveEvent(linkedWelkinEventId);
                        }
                        catch (Exception e)
                        {
                            this.logger.LogError(e, $"Failed to retrieve Welkin event for placeholder Outlook event {outlookEvent.ICalUId}.");
                        }

                        if (syncedTo == null || syncedTo.IsCancelled)
                        {
                            orphaned.Add(outlookEvent);
                        }
                    }
                }
            }

            return orphaned;
        }
    }
}