namespace OutlookWelkinSync
{
    using System;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Microsoft.Identity.Client;
    using Ninject;

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

        private ICalendarRequestBuilder CalendarRequestBuilderFrom(Event outlookEvent, string userName, string calendarName = null)
        {
            if (userName == null)
            {
                userName = outlookEvent.Calendar.Owner.Address;
            }
            
            IUserRequestBuilder userBuilder = this.graphClient.Users[userName];
            
            if (calendarName != null)
            {
                return userBuilder.Calendars[calendarName];
            }
            else
            {
                return userBuilder.Calendar;  // Use default calendar
            }
        }

        public Event Update(Event outlookEvent, string userName = null, string calendarName = null)
        {
            return CalendarRequestBuilderFrom(outlookEvent, userName, calendarName)
                .Events[outlookEvent.Id]
                .Request()
                .UpdateAsync(outlookEvent)
                .GetAwaiter()
                .GetResult();
        }

        public void Delete(Event outlookEvent, string userName = null, string calendarName = null)
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
    }
}