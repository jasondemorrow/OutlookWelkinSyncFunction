namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    public class WelkinWorkerOutlookEventRetrieval : OutlookEventRetrieval
    {
        public WelkinWorkerOutlookEventRetrieval(OutlookClient outlookClient, IWelkinClient welkinClient, ILogger logger)
        : base(outlookClient, welkinClient, logger)
        {
        }

        public override IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago)
        {
            List<Event> events = new List<Event>();
            IEnumerable<WelkinWorker> workers = this.welkinClient.RetrieveAllWorkers();
            ISet<string> domains = this.outlookClient.RetrieveAllDomainsInCompany();
            ISet<string> successfulOutlookUsers = new HashSet<string>();
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;

            foreach (WelkinWorker worker in workers)
            {
                try
                {
                    User outlookUser = this.outlookClient.FindUserCorrespondingTo(worker);
                    if (outlookUser == null || successfulOutlookUsers.Contains(outlookUser.UserPrincipalName))
                    {
                        continue;
                    }

                    string userName = outlookUser.UserPrincipalName;
                    IEnumerable<Event> workerEvents = this.outlookClient.RetrieveEventsForUserUpdatedSince(
                        userName, ago, Constants.OutlookEventExtensionsNamespace);
                    successfulOutlookUsers.Add(userName);
                    this.logger.LogInformation($"Successfully retrieved events for {userName}.");

                    // Save the Welkin worker email and owning user on each event for later sync
                    foreach (Event workerEvent in workerEvents)
                    {
                        workerEvent.AdditionalData[Constants.WelkinWorkerEmailKey] = worker.Email;
                        workerEvent.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser;
                    }

                    events.AddRange(workerEvents);
                }
                catch (Exception ex)
                {
                    this.logger.LogInformation(
                        $"Exception while trying to retrieve events for {worker.Email}: {ex.Message}");
                }
            }

            return events;
        }

        public override IEnumerable<Event> RetrieveAllOrphanedBetween(DateTimeOffset start, DateTimeOffset end)
        {
            List<Event> orphaned = new List<Event>();
            IEnumerable<WelkinWorker> workers = this.welkinClient.RetrieveAllWorkers();
            ISet<string> successfulOutlookUsers = new HashSet<string>();

            foreach (WelkinWorker worker in workers)
            {
                try
                {
                    User outlookUser = this.outlookClient.FindUserCorrespondingTo(worker);
                    if (outlookUser == null || successfulOutlookUsers.Contains(outlookUser.UserPrincipalName))
                    {
                        continue;
                    }

                    IEnumerable<Event> events = this.outlookClient.RetrieveEventsForUserScheduledBetween(
                        outlookUser, 
                        start, 
                        end, 
                        Constants.OutlookEventExtensionsNamespace);

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
                }
                catch (Exception ex)
                {
                    this.logger.LogInformation(
                        $"Exception while trying to retrieve events for {worker.Email}: {ex.Message}");
                }
            }

            return orphaned;
        }
    }
}