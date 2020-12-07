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

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(Event outlookEvent, string userPrincipal, string calendarName = null)
        {
            if (userPrincipal == null)
            {
                userPrincipal = outlookEvent.AdditionalData[Constants.WelkinWorkerEmailKey].ToString();
            }
            
            return CalendarRequestBuilderFrom(userPrincipal, calendarName);
        }

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(string userPrincipal, string calendarName = null)
        {
            IUserRequestBuilder userBuilder = this.graphClient.Users[userPrincipal];
            
            if (calendarName != null)
            {
                return userBuilder.Calendars[calendarName];
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
            string calendarName = null)
        {
            Event found;
            if (this.internalCache.TryGetValue(guid, out found))
            {
                return found;
            }

            string filter = $"iCalUId eq '{guid}'";

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarName)
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

        public IEnumerable<Event> RetrieveEventsForUserScheduledBetween(string userPrincipal, DateTime start, DateTime end, string extensionsNamespace = null, string calendarName = null)
        {
            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("startdatetime", start.ToString("o")),
                new QueryOption("enddatetime", end.ToString("o"))
            };

            ICalendarEventsCollectionRequest request = 
                        CalendarRequestBuilderFrom(userPrincipal, calendarName)
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
        
        public Event UpdateEvent(Event outlookEvent, string userName = null, string calendarName = null)
        {
            return CalendarRequestBuilderFrom(outlookEvent, userName, calendarName)
                .Events[outlookEvent.Id]
                .Request()
                .UpdateAsync(outlookEvent)
                .GetAwaiter()
                .GetResult();
        }

        public void DeleteEvent(Event outlookEvent, string userName = null, string calendarName = null)
        {
            CalendarRequestBuilderFrom(outlookEvent, userName, calendarName)
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

        public void SetOpenExtensionPropertiesOnEvent(Event outlookEvent, IDictionary<string, object> keyValuePairs, string extensionsNamespace, string calendarName = null)
        {
            IEventExtensionsCollectionRequest request = 
                        CalendarRequestBuilderFrom(outlookEvent, calendarName)
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

        public Event CreateOutlookEventFromWelkinEvent(WelkinEvent welkinEvent, WelkinWorker welkinUser, string calendarName = null)
        {
            // TODO: Include patient info
            // Create and associate a new Outlook event
            Event outlookEvent = new Event
            {
                Subject = "Placeholder for appointment in Welkin",
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = $"See your Welkin calendar (user {welkinUser.Email}) for details."
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

            Event createdEvent = CalendarRequestBuilderFrom(welkinUser.Email, calendarName)
                                        .Events
                                        .Request()
                                        .AddAsync(outlookEvent)
                                        .GetAwaiter()
                                        .GetResult();

            Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
            keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = welkinEvent.Id;
            keyValuePairs[Constants.OutlookPlaceHolderEventKey] = true;
            this.SetOpenExtensionPropertiesOnEvent(createdEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);

            return createdEvent;
        }
    }
}