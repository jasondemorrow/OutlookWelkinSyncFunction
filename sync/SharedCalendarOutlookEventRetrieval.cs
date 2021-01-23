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
            this.logger.LogInformation($"Running Outlook event retrieval on shared calendar {this.sharedOutlookCalendar?.Name} for user {this?.sharedCalendarOutlookUser?.Mail}");
            List<Event> events = new List<Event>();
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            IEnumerable<Event> retrieved = this.outlookClient.RetrieveEventsForUserScheduledBetween(
                this.sharedCalendarOutlookUser, 
                start, 
                end, 
                Constants.OutlookEventExtensionsNamespace, 
                this.sharedOutlookCalendar.Id);
            this.logger.LogInformation($"Outlook event retrieval successfully queried events.");

            // Save the Welkin worker email and owning user on each event for later sync
            foreach (Event outlookEvent in retrieved)
            {
                this.logger.LogInformation($"Outlook event retrieval found event {outlookEvent.ICalUId}.");
                try
                {
                    // Unlike name-based sync, we don't have a Welkin user at this point. We need to
                    // try and get the Welkin user from the organizer's email and save if successful.
                    // This user will later be used to create a placeholder event in Welkin if needed.
                    string userEmail = outlookEvent.Organizer?.EmailAddress?.Address;
                    if (string.IsNullOrEmpty(userEmail))
                    {
                        continue;
                    }

                    WelkinWorker worker = this.welkinClient.FindWorker(userEmail);
                    if (worker == null)
                    {
                        continue;
                    }

                    User outlookUser = this.outlookClient.FindUserCorrespondingTo(worker);
                    if (outlookUser == null)
                    {
                        continue;
                    }

                    this.logger.LogInformation($"Found new Outlook event {outlookEvent.ICalUId} " + 
                                               $"for user {userEmail} in shared calendar {this.sharedCalendarName}.");
                    outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey] = userEmail;
                    outlookEvent.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser;
                    events.Add(outlookEvent);
                }
                catch (Exception ex)
                {
                    this.logger.LogWarning($"Exception while running {this.ToString()}: {ex.Message} {ex.StackTrace}");
                }
            }

            return events;
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