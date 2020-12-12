namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net.Http;
    using System.Text;
    using Jose;
    using Microsoft.Extensions.Caching.Memory;
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Linq;
    using Ninject;
    using RestSharp;

    public class WelkinClient
    {
        private MemoryCache internalCache = new MemoryCache(new MemoryCacheOptions()
        {
            SizeLimit = 1024
        });
        private readonly MemoryCacheEntryOptions cacheEntryOptions = 
            new MemoryCacheEntryOptions()
                .SetAbsoluteExpiration(TimeSpan.FromSeconds(180))
                .SetSize(1);
        private readonly WelkinConfig config;
        private readonly ILogger logger;
        private readonly string token;
        private readonly string dummyPatientId;

        public WelkinClient(WelkinConfig config, ILogger logger, [Named("DummyPatientId")] string dummyPatientId)
        {
            this.config = config;
            this.logger = logger;
            this.dummyPatientId = dummyPatientId;
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

        private T CreateOrUpdateObject<T>(T obj, string path, string id = null) where T : class
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

            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JObject data = result?.First?.ToObject<JProperty>()?.Value.ToObject<JObject>();
            T updated = (data == null)? default(T) : JsonConvert.DeserializeObject<T>(data.ToString());
            
            internalCache.Set(url, updated, cacheEntryOptions);
            return updated;
        }

        private T RetrieveObject<T>(string id, string path, Dictionary<string, string> parameters = null)
        {
            string url = $"{config.ApiUrl}{path}/{id}";
            T retrieved = default(T);
            if (internalCache.TryGetValue(url, out retrieved))
            {
                return retrieved;
            }

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
            retrieved = JsonConvert.DeserializeObject<T>(body.Value.ToString());

            internalCache.Set(url, retrieved, cacheEntryOptions);
            return retrieved;
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

            internalCache.Remove(url);
        }

        private IEnumerable<T> SearchObjects<T>(string path, Dictionary<string, string> parameters = null)
        {
            string url = $"{config.ApiUrl}{path}";
            string key = url + "?" + string.Join("&", parameters.Select(e => $"{e.Key}={e.Value}"));
            IEnumerable<T> found;
            if (internalCache.TryGetValue(key, out found))
            {
                return found;
            }

            var client = new RestClient(url);

            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");

            foreach(KeyValuePair<string, string> kvp in parameters ?? Enumerable.Empty<KeyValuePair<string, string>>())
            {
                request.AddParameter(kvp.Key, kvp.Value);
            }

            // Intentionally not caching this result for now. Searches will generally only be done once per run.
            var response = client.Execute(request);
            if (response.StatusCode != System.Net.HttpStatusCode.OK)
            {
                throw new Exception($"HTTP status {response.StatusCode} with message '{response.ErrorMessage}' and body '{response.Content}'");
            }

            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();

            found = JsonConvert.DeserializeObject<IEnumerable<T>>(data.ToString());

            internalCache.Set(key, found, cacheEntryOptions);
            return found;
        }

        public WelkinEvent CreateOrUpdateEvent(WelkinEvent evt, string id = null)
        {
            return this.CreateOrUpdateObject(evt, Constants.CalendarEventResourceName, id);
        }

        public WelkinEvent RetrieveEvent(string eventId)
        {
            return this.RetrieveObject<WelkinEvent>(eventId, Constants.CalendarEventResourceName);
        }

        public IEnumerable<WelkinEvent> RetrieveEventsUpdatedSince(TimeSpan ago)
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
            // Cache results for individual retrieval
            foreach (WelkinEvent welkinEvent in events)
            {
                string key = $"{config.ApiUrl}{Constants.CalendarEventResourceName}/{welkinEvent.Id}";
                internalCache.Set(key, welkinEvent, cacheEntryOptions);
            }

            return events.Where(this.IsValid);
        }

        private bool IsValid(WelkinEvent evt)
        {
            return 
                evt != null && 
                !(evt.PatientId == null || evt.PatientId.Equals(this.dummyPatientId)) && 
                !(evt.Outcome != null && evt.Outcome.Equals(Constants.WelkinCancelledOutcome));
        }

        public void DeleteEvent(WelkinEvent welkinEvent)
        {
            this.DeleteObject(welkinEvent.Id, Constants.CalendarEventResourceName);
        }

        public WelkinCalendar RetrieveCalendar(string calendarId)
        {
            return this.RetrieveObject<WelkinCalendar>(calendarId, Constants.CalendarResourceName);
        }

        public WelkinCalendar RetrieveCalendarFor(WelkinWorker worker)
        {
            string url = $"{config.ApiUrl}{Constants.CalendarResourceName}?worker={worker.Id}";
            WelkinCalendar found;
            if (internalCache.TryGetValue(url, out found))
            {
                return found;
            }
            var client = new RestClient(url);
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
            found = JsonConvert.DeserializeObject<WelkinCalendar>(calendar.ToString());
            internalCache.Set(url, found, cacheEntryOptions);
            return found;
        }

        public WelkinExternalId CreateOrUpdateExternalId(WelkinExternalId external, string id = null)
        {
            return this.CreateOrUpdateObject(external, Constants.ExternalIdResourceName, id);
        }

        public void DeleteExternalId(WelkinExternalId externalId)
        {
            this.DeleteObject(externalId.Id, Constants.ExternalIdResourceName);
        }

        public WelkinPatient RetrievePatient(string patientId)
        {
            return this.RetrieveObject<WelkinPatient>(patientId, Constants.PatientResourceName);
        }

        public WelkinWorker RetrieveWorker(string workerId)
        {
            return this.RetrieveObject<WelkinWorker>(workerId, Constants.WorkerResourceName);
        }

        public IEnumerable<WelkinWorker> RetrieveAllWorkers()
        {
            List<WelkinWorker> workers = new List<WelkinWorker>();
            var client = new RestClient(config.ApiUrl + Constants.WorkerResourceName);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            if (!result.ContainsKey("data"))
            {
                return null;
            }
            JArray data = result["data"].ToObject<JArray>();
            IEnumerable<WelkinWorker> page = JsonConvert.DeserializeObject<List<WelkinWorker>>(data.ToString());
            workers.AddRange(page);
            JObject links = result["links"]?.ToObject<JObject>();
            while (links != null && links.ContainsKey("href"))
            {
                Links href = links["href"].ToObject<Links>();
                if (string.IsNullOrEmpty(href.Next))
                {
                    break;
                }
                client = new RestClient(href.Next);
                request = new RestRequest(Method.GET);
                request.AddHeader("authorization", "Bearer " + this.token);
                request.AddHeader("cache-control", "no-cache");
                response = client.Execute(request);
                result = JsonConvert.DeserializeObject(response.Content) as JObject;
                if (result == null || !result.ContainsKey("data"))
                {
                    break;
                }
                data = result["data"].ToObject<JArray>();
                page = JsonConvert.DeserializeObject<List<WelkinWorker>>(data.ToString());
                workers.AddRange(page);
                links = result["links"]?.ToObject<JObject>();
            }
            // Cache results for individual retrieval by email or ID
            foreach (WelkinWorker worker in workers)
            {
                string key = $"{config.ApiUrl}{Constants.WorkerResourceName}/{worker.Id}";
                internalCache.Set(key, worker, cacheEntryOptions);
                internalCache.Set(worker.Email.ToLowerInvariant(), worker, cacheEntryOptions);
            }
            return workers;
        }

        public WelkinWorker FindWorker(string email)
        {
            if (string.IsNullOrEmpty(email))
            {
                return null;
            }
            WelkinWorker worker;
            if (internalCache.TryGetValue(email.ToLowerInvariant(), out worker))
            {
                return worker;
            }
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["email"] = email;
            IEnumerable<WelkinWorker> found = SearchObjects<WelkinWorker>(Constants.WorkerResourceName, parameters);
            worker = found.FirstOrDefault();
            if (worker != null)
            {
                internalCache.Set(worker.Email.ToLowerInvariant(), worker, cacheEntryOptions);
            }
            return worker;
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
            IEnumerable<WelkinExternalId> foundLinks = SearchObjects<WelkinExternalId>(Constants.ExternalIdResourceName, parameters);
            return foundLinks
                        .Where(x => x.Namespace.StartsWith(Constants.WelkinEventExtensionNamespacePrefix))
                        .FirstOrDefault();
        }

        public WelkinLastSyncEntry RetrieveLastSyncFor(WelkinEvent internalEvent)
        {
            // We store last sync time for an event as an external ID. This is a hack to make event types extensible.
            Dictionary<string, string> parameters = new Dictionary<string, string>();
            parameters["resource"] = Constants.CalendarEventResourceName;
            parameters["welkin_id"] = internalEvent.Id;
            IEnumerable<WelkinExternalId> foundLinks = SearchObjects<WelkinExternalId>(Constants.ExternalIdResourceName, parameters);
            if (foundLinks == null || !foundLinks.Any())
            {
                return null;
            }
            WelkinExternalId externalId = 
                foundLinks
                    .Where(x => x.Namespace.StartsWith(Constants.WelkinLastSyncExtensionNamespace))
                    .FirstOrDefault();
            return (externalId == null)? null : new WelkinLastSyncEntry(externalId);
        }

        public bool UpdateLastSyncFor(WelkinEvent internalEvent, string existingId = null, DateTimeOffset? lastSync = null)
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
                Id = existingId,
                Resource = Constants.CalendarEventResourceName,
                ExternalId = Guid.NewGuid().ToString(), // does not matter
                InternalId = internalEvent.Id,
                Namespace = syntheticNamespace
            };
            welkinExternalId = this.CreateOrUpdateExternalId(welkinExternalId, existingId);

            return welkinExternalId != null && welkinExternalId.InternalId.Equals(internalEvent.Id);
        }

        public WelkinEvent GeneratePlaceholderEventForCalendar(WelkinCalendar calendar)
        {
            WelkinEvent evt = new WelkinEvent();
            evt.CalendarId = calendar.Id;
            evt.IsAllDay = true;
            evt.Day = DateTime.UtcNow.Date;
            evt.Modality = Constants.DefaultModality;
            evt.AppointmentType = Constants.DefaultAppointmentType;
            evt.PatientId = this.dummyPatientId;
            evt.IgnoreUnavailableTimes = true;
            evt.IgnoreWorkingHours = true;
            
            return evt;
        }

        public bool IsPlaceHolderEvent(WelkinEvent evt)
        {
            string patientId = evt?.PatientId;
            return !string.IsNullOrEmpty(patientId) && patientId.Equals(this.dummyPatientId);
        }
    }
}