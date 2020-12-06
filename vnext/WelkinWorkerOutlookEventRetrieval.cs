namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    
    public class WelkinWorkerOutlookEventRetrieval : OutlookEventRetrieval
    {
        public WelkinWorkerOutlookEventRetrieval(OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger)
        : base(outlookClient, welkinClient, logger)
        {
        }

        public override IEnumerable<Event> RetrieveAllUpdatedSince(TimeSpan ago)
        {
            List<Event> events = new List<Event>();
            IEnumerable<WelkinWorker> workers = this.welkinClient.RetrieveAllWorkers();
            ISet<string> domains = this.outlookClient.RetrieveAllDomainsInCompany();
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;

            foreach (WelkinWorker worker in workers)
            {
                ISet<string> candidateEmails = this.ProducePrincipalCandidates(worker, domains);
                foreach (string email in candidateEmails)
                {
                    try
                    {
                        IEnumerable<Event> workerEvents = this.outlookClient.RetrieveEventsForUserScheduledBetween(email, start, end);
                        this.logger.LogInformation($"Successfully retrieved events for {email}.");
                        events.AddRange(workerEvents);
                        break; // Stop once we find a working candidate
                    }
                    catch (Exception ex)
                    {
                        this.logger.LogError(
                            ex, 
                            $"Exception while retrieving Outlook events for {email}.");
                    }
                }
            }

            return events;
        }

        private ISet<string> ProducePrincipalCandidates(WelkinWorker worker, ISet<string> domains)
        {
            HashSet<string> candidates = new HashSet<string>();
            int idxIdAt = worker.Id.IndexOf("@");
            string idAt = (idxIdAt > -1) ? worker.Id.Substring(0, idxIdAt) : null;
            int idxIdPlus = worker.Id.IndexOf("+");
            string idPlus = (idxIdPlus > -1) ? worker.Id.Substring(0, idxIdPlus) : null;
            int idxEmailAt = worker.Email.IndexOf("@");
            string emailAt = (idxEmailAt > -1) ? worker.Email.Substring(0, idxEmailAt) : null;
            int idxEmailPlus = worker.Email.IndexOf("+");
            string emailPlus = (idxEmailPlus > -1) ? worker.Email.Substring(0, idxEmailPlus) : null;

            foreach (string domain in domains)
            {
                if (!string.IsNullOrEmpty(idAt))
                {
                    candidates.Add($"{idAt}@{domain}");
                }
                if (!string.IsNullOrEmpty(idPlus))
                {
                    candidates.Add($"{idPlus}@{domain}");
                }
                if (!string.IsNullOrEmpty(emailAt))
                {
                    candidates.Add($"{emailAt}@{domain}");
                }
                if (!string.IsNullOrEmpty(emailPlus))
                {
                    candidates.Add($"{emailPlus}@{domain}");
                }
            }

            return candidates;
        }
    }
}