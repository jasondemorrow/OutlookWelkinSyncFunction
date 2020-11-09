using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Text;
using Jose;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace OutlookWelkinSyncFunction
{
    public class WelkinClient
    {
        private readonly WelkinConfig config;
        private readonly ILogger logger;
        private readonly string token;
        private readonly string dummyPatientId;

        public WelkinClient(WelkinConfig config, ILogger logger)
        {
            this.config = config;
            this.logger = logger;
            this.dummyPatientId = Environment.GetEnvironmentVariable("WelkinDummyPatientId");
            var payload = new Dictionary<string, object>()
            {
                { "iss", config.ClientId },
                { "aud", config.TokenUrl },
                { "scope", config.Scope },
                { "exp", new DateTimeOffset(DateTime.UtcNow.AddHours(1)).ToUnixTimeSeconds() }
            };

            var secretKey = Encoding.UTF8.GetBytes(config.ClientSecret);
            string assertion = JWT.Encode(payload, secretKey, JwsAlgorithm.HS256);

            string body = $"grant_type={config.GrantType}&assertion={assertion}";
            using (var httpClient = new HttpClient())
            {
                HttpResponseMessage postResponse = httpClient.PostAsync(
                    config.TokenUrl,
                    new StringContent(body, Encoding.UTF8, "application/x-www-form-urlencoded"))
                    .GetAwaiter()
                    .GetResult();
                string content = postResponse.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                dynamic resp = JObject.Parse(content);
                this.token = resp.access_token;
            }

            if (string.IsNullOrEmpty(this.token))
            {
                throw new ArgumentException($"Unable to retrieve a valid token using the credentials in env");
            }
        }

        public WelkinEvent GeneratePlaceholderEventForCalendar(string calendarId)
        {
            WelkinEvent evt = new WelkinEvent();
            evt.CalendarId = calendarId;
            evt.IsAllDay = true;
            evt.Day = DateTime.UtcNow.Date;
            evt.Modality = Constants.DefaultModality;
            evt.AppointmentType = Constants.DefaultAppointmentType;
            evt.PatientId = this.dummyPatientId;
            evt.IgnoreUnavailableTimes = true;
            evt.IgnoreWorkingHours = true;
            
            return evt;
        }

        public IEnumerable<WelkinPractitioner> GetAllPractitioners()
        {
            /*
                TODO: HAL or pagination using something like one of the following:
                    https://nugetmusthaves.com/Package/HoneyBear.HalClient
                    https://github.com/wis3guy/HalClient.Net
                    https://stackoverflow.com/questions/9164197/how-to-implement-paging-with-restsharp
            */
            var client = new RestClient(config.ApiUrl + "workers");
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();
            return JsonConvert.DeserializeObject<List<WelkinPractitioner>>(data.ToString());
        }

        public WelkinCalendar GetCalendarForPractitioner(WelkinPractitioner practitioner)
        {
            var client = new RestClient(config.ApiUrl + "calendars?worker=" + practitioner.Id);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First?.ToObject<JProperty>()?.Value.ToObject<JArray>();
            if (data == null)
            {
                return null;
            }
            JObject calendar = data.First?.ToObject<JObject>();
            if (calendar == null)
            {
                return null;
            }
            return JsonConvert.DeserializeObject<WelkinCalendar>(calendar.ToString());
        }

        public IEnumerable<WelkinEvent> GetEventsUpdatedSince(TimeSpan ago)
        {
            DateTime end = DateTime.UtcNow;
            DateTime start = end - ago;
            string url = $"{config.ApiUrl}calendar_events?page[from]={start.ToString("o")}&page[to]={end.ToString("o")}";
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();
            // Filter out placeholder events created by sync. Welkin API doesn't support querying by patient ID.
            List<WelkinEvent> events = JsonConvert.DeserializeObject<List<WelkinEvent>>(data.ToString());
            return events.Where(this.IsValid);
        }

        private bool IsValid(WelkinEvent evt)
        {
            return 
                evt != null && 
                !(evt.PatientId == null || evt.PatientId.Equals(this.dummyPatientId)) && 
                !(evt.Outcome != null && evt.Outcome.Equals(Constants.WelkinCancelledOutcome));
        }

        private T GetObject<T>(string id, string path, Dictionary<string, string> parameters = null)
        {
            string url = $"{config.ApiUrl}{path}/{id}";
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            foreach(KeyValuePair<string, string> kvp in parameters ?? Enumerable.Empty<KeyValuePair<string, string>>())
            {
                request.AddParameter(kvp.Key, kvp.Value);
            }
            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JProperty body = result.First.ToObject<JProperty>();
            return JsonConvert.DeserializeObject<T>(body.Value.ToString());
        }

        public WelkinEvent CreateOrUpdateEvent(WelkinEvent evt, string id = null)
        {
            return this.CreateOrUpdateObject(evt, Constants.CalendarEventResourceName, id);
        }

        public WelkinEvent GetEvent(string eventId)
        {
            return this.GetObject<WelkinEvent>(eventId, Constants.CalendarEventResourceName);
        }

        public void DeleteEvent(WelkinEvent evt)
        {
            this.DeleteObject(evt.Id, Constants.CalendarEventResourceName);
        }

        public WelkinExternalId CreateOrUpdateExternalId(WelkinExternalId external, string id = null)
        {
            return this.CreateOrUpdateObject(external, Constants.ExternalIdResourceName, id);
        }

        public WelkinExternalId GetExternalId(string externalId)
        {
            return this.GetObject<WelkinExternalId>(externalId, Constants.ExternalIdResourceName);
        }

        private T CreateOrUpdateObject<T>(T obj, string path, string? id = null) where T : class
        {
            string url = (id == null)? $"{config.ApiUrl}{path}" : $"{config.ApiUrl}{path}/{id}";
            var client = new RestClient(url);
            Method method = (id == null)? Method.POST : Method.PUT;
            var request = new RestRequest(method);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            request.AddParameter("application/json", JsonConvert.SerializeObject(obj), ParameterType.RequestBody);
            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }
            JObject? result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JObject? data = result?.First?.ToObject<JProperty>()?.Value.ToObject<JObject>();
            return (data == null)? default(T) : JsonConvert.DeserializeObject<T>(data.ToString());
        }

        private void DeleteObject(string id, string path)
        {
            string url = $"{config.ApiUrl}{path}/{id}";
            var client = new RestClient(url);
            Method method = Method.DELETE;
            var request = new RestRequest(method);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }
        }

        public WelkinExternalId FindExternalMappingFor(WelkinEvent internalEvent, Event externalEvent = null)
        {
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            if (externalEvent != null)
            {
                parameters["namespace"] = Constants.WelkinEventExtensionNamespacePrefix + externalEvent.ICalUId;
                string derivedGuid = Guids.FromText(externalEvent.ICalUId).ToString();
                parameters["external_id"] = derivedGuid;
            }
            parameters["resource"] = Constants.CalendarEventResourceName;
            parameters["welkin_id"] = internalEvent.Id;
            IEnumerable<WelkinExternalId> foundLinks = SearchObjects<WelkinExternalId>("external_ids", parameters);
            return foundLinks.FirstOrDefault();
        }

        public WelkinLastSyncEntry FindLastSyncEntryFor(WelkinEvent internalEvent)
        {
            // We store last sync time for an event as an external ID. This is a hack to make event types extensible.
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["resource"] = Constants.CalendarEventResourceName;
            parameters["welkin_id"] = internalEvent.Id;
            IEnumerable<WelkinExternalId> foundLinks = SearchObjects<WelkinExternalId>("external_ids", parameters);
            if (foundLinks == null || !foundLinks.Any())
            {
                return null;
            }
            WelkinExternalId externalId = 
                foundLinks
                    .Where(x => x.Namespace.StartsWith(Constants.WelkinLastSyncExtensionNamespace))
                    .FirstOrDefault();
            return new WelkinLastSyncEntry(externalId);
        }

        public bool SetLastSyncDateTimeFor(WelkinEvent internalEvent, string existingId = null, DateTimeOffset? lastSync = null)
        {
            if (lastSync == null)
            {
                lastSync = DateTimeOffset.UtcNow;
            }

            // We store last sync time for an event as an external ID namespace. 
            // This is a hack to make event types extensible.
            string isoDate = lastSync.Value.ToString("o", CultureInfo.InvariantCulture);
            string syntheticNamespace = Constants.WelkinLastSyncExtensionNamespace + ":::" + isoDate;

            WelkinExternalId welkinExternalId = new WelkinExternalId
            {
                Resource = Constants.CalendarEventResourceName,
                ExternalId = Guid.NewGuid().ToString(), // does not matter
                InternalId = internalEvent.Id,
                Namespace = syntheticNamespace
            };
            welkinExternalId = this.CreateOrUpdateExternalId(welkinExternalId, existingId);

            return welkinExternalId != null && welkinExternalId.InternalId.Equals(internalEvent.Id);
        }

        private IEnumerable<T> SearchObjects<T>(string path, Dictionary<string, string> parameters = null)
        {
            string url = $"{config.ApiUrl}{path}";
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            foreach(KeyValuePair<string, string> kvp in parameters ?? Enumerable.Empty<KeyValuePair<string, string>>())
            {
                request.AddParameter(kvp.Key, kvp.Value);
            }
            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();
            return JsonConvert.DeserializeObject<IEnumerable<T>>(data.ToString());
        }

        public bool IsPlaceHolderEvent(WelkinEvent evt)
        {
            string patientId = evt?.PatientId;
            return !string.IsNullOrEmpty(patientId) && patientId.Equals(this.dummyPatientId);
        }
    }
}