namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;

    public class OutlookClient
    {
        private MemoryCache internalCache = new MemoryCache(new MemoryCacheOptions()
        {
            SizeLimit = 1024
        });
        private readonly MemoryCacheEntryOptions cacheEntryOptions = 
            new MemoryCacheEntryOptions()
                .SetAbsoluteExpiration(TimeSpan.FromSeconds(180))
                .SetSize(1);
        private readonly OutlookConfig config;
        private readonly string token;
        private readonly GraphServiceClient graphClient;
        private readonly ILogger logger;

        public OutlookClient(OutlookConfig config, ILogger logger)
        {
            this.config = config;
            this.logger = logger;
            IConfidentialClientApplication app = 
                        ConfidentialClientApplicationBuilder
                            .Create(config.ClientId)
                            .WithClientSecret(config.ClientSecret)
                            .WithAuthority(new Uri(config.Authority))
                            .Build();
                                                    
            string[] scopes = new string[] { $"{config.ApiUrl}.default" }; 
            
            AuthenticationResult result = app.AcquireTokenForClient(scopes).ExecuteAsync().GetAwaiter().GetResult();
            this.token = result.AccessToken;

            if (string.IsNullOrEmpty(this.token))
            {
                throw new ArgumentException($"Unable to retrieve a valid token using the credentials in env");
            }
            
            this.graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) => {
                requestMessage
                    .Headers
                    .Authorization = new AuthenticationHeaderValue("Bearer", this.token);

                return Task.FromResult(0);
            }));
        }

        public static bool IsPlaceHolderEvent(Event outlookEvent)
        {
            Extension extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData != null && extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookPlaceHolderEventKey))
            {
                return true;
            }

            return false;
        }

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(Event outlookEvent, string userPrincipal, string calendarId = null)
        {
            if (userPrincipal == null)
            {
                User outlookUser = outlookEvent.AdditionalData[Constants.OutlookUserObjectKey] as User;
                userPrincipal = outlookUser?.UserPrincipalName;
            }
            
            return CalendarRequestBuilderFrom(userPrincipal, calendarId);
        }

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(string userPrincipal, string calendarId = null)
        {
            IUserRequestBuilder userBuilder = this.graphClient.Users[userPrincipal];
            
            if (calendarId != null)
            {
                return userBuilder.Calendars[calendarId];
            }
            else
            {
                return userBuilder.Calendar;  // Use default calendar
            }
        }

        public Event RetrieveEventWithICalId(
            string userPrincipal, 
            string guid, 
            string extensionsNamespace = null, 
            string calendarId = null)
        {
            Event found;
            if (this.internalCache.TryGetValue(guid, out found))
            {
                return found;
            }

            string filter = $"iCalUId eq '{guid}'";

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarId)
                            .Events
                            .Request()
                            .Filter(filter);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }
            
            found = request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult()
                    .FirstOrDefault();

            this.internalCache.Set(guid, found, this.cacheEntryOptions);
            return found;
        }

        public IEnumerable<Event> RetrieveEventsForUserUpdatedSince(string userPrincipal, TimeSpan ago, string extensionsNamespace = null, string calendarId = null)
        {
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            string filter = $"lastModifiedDateTime lt {end.ToString("o")} and lastModifiedDateTime gt {start.ToString("o")}";

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarId)
                            .Events
                            .Request()
                            .Filter(filter);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }
            
            return request
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public IEnumerable<Event> RetrieveEventsForUserScheduledBetween(string userPrincipal, DateTime start, DateTime end, string extensionsNamespace = null, string calendarId = null)
        {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startdatetime", start.ToString("o")),
                new QueryOption("enddatetime", end.ToString("o"))
            };

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarId)
                            .Events
                            .Request(queryOptions);

            if (extensionsNamespace != null)
            {
                request = request.Expand($"extensions($filter=id eq '{extensionsNamespace}')");
            }

            IEnumerable<Event> events = request
                                        .GetAsync()
                                        .GetAwaiter()
                                        .GetResult();

            // Cache for later individual retrieval by ICalUId
            foreach (Event outlookEvent in events)
            {
                this.internalCache.Set(outlookEvent.ICalUId, outlookEvent, this.cacheEntryOptions);
            }

            return events;
        }

        public ISet<string> RetrieveAllDomainsInCompany()
        {
            HashSet<string> domains;
            string key = "domains";

            if (this.internalCache.TryGetValue(key, out domains))
            {
                return domains;
            }

            var page = this.graphClient.Domains.Request().GetAsync().GetAwaiter().GetResult();
            domains = page.Select(r => r.Id).ToHashSet();

            this.internalCache.Set(key, domains, this.cacheEntryOptions);
            return domains;
        }
        
        public Event UpdateEvent(Event outlookEvent, string userName = null, string calendarId = null)
        {
            return CalendarRequestBuilderFrom(outlookEvent, userName, calendarId)
                .Events[outlookEvent.Id]
                .Request()
                .UpdateAsync(outlookEvent)
                .GetAwaiter()
                .GetResult();
        }

        public void DeleteEvent(Event outlookEvent, string userName = null, string calendarId = null)
        {
            CalendarRequestBuilderFrom(outlookEvent, userName, calendarId)
                .Events[outlookEvent.Id]
                .Request()
                .DeleteAsync()
                .GetAwaiter()
                .GetResult();
        }

        public string LinkedWelkinEventIdFrom(Event outlookEvent)
        {
            Extension extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData == null || !extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLinkedWelkinEventIdKey))
            {
                this.logger.LogInformation($"No linked Welkin event for Outlook event {outlookEvent.ICalUId}");
                return null;
            }

            string linkedEventId = extensionForWelkin.AdditionalData[Constants.OutlookLinkedWelkinEventIdKey]?.ToString();
            if (string.IsNullOrEmpty(linkedEventId))
            {
                this.logger.LogInformation($"Null or empty linked Welkin event ID for Outlook event {outlookEvent.ICalUId}");
                return null;
            }

            return linkedEventId;
        }

        public Microsoft.Graph.Calendar RetrieveOwningUserDefaultCalendar(Event childEvent)
        {
            if (!childEvent.AdditionalData.ContainsKey(Constants.WelkinWorkerEmailKey))
            {
                return null;
            }
            
            return CalendarRequestBuilderFrom(
                childEvent, 
                childEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString())
                    .Request()
                    .GetAsync()
                    .GetAwaiter()
                    .GetResult();
        }

        public User RetrieveOwningUser(Event outlookEvent)
        {
            return RetrieveUser(outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString());
        }

        public User RetrieveUser(string email)
        {
            User retrieved;
            if (internalCache.TryGetValue(email, out retrieved))
            {
                return retrieved;
            }

            retrieved = this.graphClient.Users[email].Request().GetAsync().GetAwaiter().GetResult();

            internalCache.Set(email, retrieved, this.cacheEntryOptions);
            return retrieved;
        }

        public User FindUserCorrespondingTo(WelkinWorker welkinWorker)
        {
            ISet<string> domains = this.RetrieveAllDomainsInCompany();
            ISet<string> candidateEmails = ProducePrincipalCandidates(welkinWorker, domains);
            foreach (string email in candidateEmails)
            {
                try
                {
                    User outlookUser = this.RetrieveUser(email);
                    if (outlookUser != null)
                    {
                        return outlookUser;
                    }
                }
                catch (ServiceException ex)
                {
                    //this.logger.LogInformation($"{email}:{ex.StatusCode}");
                }
            }
            return null;
        }

        private static ISet<string> ProducePrincipalCandidates(WelkinWorker worker, ISet<string> domains)
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

        public void SetOpenExtensionPropertiesOnEvent(Event outlookEvent, IDictionary<string, object> keyValuePairs, string extensionsNamespace, string calendarId = null)
        {
            IEventExtensionsCollectionRequest request = 
                        CalendarRequestBuilderFrom(outlookEvent, null, calendarId)
                            .Events[outlookEvent.Id]
                            .Extensions
                            .Request();
            OpenTypeExtension ext = new OpenTypeExtension();
            ext.ExtensionName = extensionsNamespace;
            ext.AdditionalData = keyValuePairs;
            string parameterString = (keyValuePairs != null) ? string.Join(", ", keyValuePairs.Select(kv => kv.Key + "=" + kv.Value).ToArray()) : "NULL";

            request.AddAsync(ext).GetAwaiter().OnCompleted(() => this.logger.LogInformation($"Successfully added an extension with values {parameterString}."));
        }

        public void MergeOpenExtensionPropertiesOnEvent(Event outlookEvent, IDictionary<string, object> keyValuePairs, string extensionsNamespace)
        {
            Extension extension = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(extensionsNamespace))?.FirstOrDefault();
            if (extension?.AdditionalData != null)
            {
                extension.AdditionalData.ToList().ForEach(x => 
                {
                    if (!keyValuePairs.ContainsKey(x.Key))
                    {
                        keyValuePairs[x.Key] = x.Value;
                    }
                });
            }

            this.SetOpenExtensionPropertiesOnEvent(outlookEvent, keyValuePairs, extensionsNamespace);
        }

        public bool SetLastSyncDateTime(Event evt, DateTimeOffset? lastSync = null)
        {
            if (lastSync == null)
            {
                lastSync = DateTimeOffset.UtcNow;
            }

            IDictionary<string, object> keyValuePairs = new Dictionary<string, object>
            {
                {Constants.OutlookLastSyncDateTimeKey , lastSync.Value.ToString("o", CultureInfo.InvariantCulture)}
            };

            try
            {
                this.MergeOpenExtensionPropertiesOnEvent(evt, keyValuePairs, Constants.OutlookEventExtensionsNamespace);
            }
            catch (Exception e)
            {
                this.logger.LogError(string.Format("While setting sync date-time for event {0}", evt.ICalUId), e);
                return false;
            }

            return true;
        }

        public static DateTime? GetLastSyncDateTime(Event outlookEvent)
        {
            Extension extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
            if (extensionForWelkin?.AdditionalData != null && extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLastSyncDateTimeKey))
            {
                string lastSync = extensionForWelkin.AdditionalData[Constants.OutlookLastSyncDateTimeKey].ToString();
                return string.IsNullOrEmpty(lastSync) ? null : new DateTime?(DateTime.ParseExact(lastSync, "o", CultureInfo.InvariantCulture).ToUniversalTime());
            }

            return null;
        }

        public Event CreateOutlookEventFromWelkinEvent(WelkinEvent welkinEvent, WelkinWorker welkinUser, WelkinPatient welkinPatient, string calendarId = null)
        {
            User outlookUser = this.FindUserCorrespondingTo(welkinUser);
            if (outlookUser == null)
            {
                this.logger.LogWarning($"Couldn't find Outlook user corresponding to Welkin user {welkinUser.Email}. " +
                                       $"Can't create an Outlook event from Welkin event {welkinEvent.Id}.");
                return null;
            }
            return this.CreateOutlookEventFromWelkinEvent(welkinEvent, welkinUser, outlookUser, welkinPatient, calendarId);
        }

        public Event CreateOutlookEventFromWelkinEvent(WelkinEvent welkinEvent, WelkinWorker welkinUser, User outlookUser, WelkinPatient welkinPatient, string calendarId = null)
        {
            // Create and associate a new Outlook event
            Event outlookEvent = new Event
            {
                Subject = $"Welkin Appointment: {welkinEvent.Modality} with {welkinPatient.FirstName} {welkinPatient.LastName} for {welkinUser.FirstName} {welkinUser.LastName}",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"Event synchronized from Welkin. See Welkin calendar (user {welkinUser.Email}) for details."
                },
                IsAllDay = welkinEvent.IsAllDay,
                Start = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Day.Value.Date.ToString() // Midnight day of
                        : welkinEvent.Start.Value.ToString(), // Will be UTC
                    TimeZone = welkinUser.Timezone
                },
                End = new DateTimeTimeZone
                {
                    DateTime = welkinEvent.IsAllDay 
                        ? welkinEvent.Day.Value.Date.AddDays(1).ToString() // Midnight day after
                        : welkinEvent.End.Value.ToString(), // Will be UTC
                    TimeZone = welkinUser.Timezone
                }
            };

            Event createdEvent = CalendarRequestBuilderFrom(outlookUser.UserPrincipalName, calendarId)
                                        .Events
                                        .Request()
                                        .AddAsync(outlookEvent)
                                        .GetAwaiter()
                                        .GetResult();
            createdEvent.AdditionalData[Constants.OutlookUserObjectKey] = outlookUser;

            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = welkinEvent.Id;
            keyValuePairs[Constants.OutlookPlaceHolderEventKey] = true;
            this.SetOpenExtensionPropertiesOnEvent(createdEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);

            return createdEvent;
        }

        public Microsoft.Graph.Calendar RetrieveCalendar(string userPrincipal, string calendarId)
        {
            List<Microsoft.Graph.Calendar> calendars = new List<Microsoft.Graph.Calendar>();
            IUserCalendarsCollectionPage page = this.graphClient
                .Users[userPrincipal]
                .Calendars
                .Request()
                .GetAsync()
                .GetAwaiter().GetResult();
            calendars.AddRange(page.ToList());
            while (page.NextPageRequest != null)
            {
                page = page.NextPageRequest.GetAsync().GetAwaiter().GetResult();
                calendars.AddRange(page.ToList());
            }
            return calendars.Where(c => c.Name.ToLowerInvariant().Equals(calendarId.ToLowerInvariant())).FirstOrDefault();
        }
    }
}