using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using Jose;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;

namespace OutlookWelkinSyncFunction
{
    public class WelkinClient // TODO: pagination
    {
        private readonly WelkinConfig config;
        private readonly ILogger logger;
        private readonly string token;

        public WelkinClient(WelkinConfig config, ILogger logger)
        {
            this.config = config;
            this.logger = logger;
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
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();
            JObject calendar = data.First.ToObject<JObject>();
            return JsonConvert.DeserializeObject<WelkinCalendar>(calendar.ToString());
        }

        public IEnumerable<WelkinEvent> GetEventsUpdatedBetween(DateTime start, DateTime end)
        {
            string url = $"{config.ApiUrl}calendar_events?page[from]={start.ToString("o")}&page[to]={end.ToString("o")}";
            var client = new RestClient(url);
            var request = new RestRequest(Method.GET);
            request.AddHeader("authorization", "Bearer " + this.token);
            request.AddHeader("cache-control", "no-cache");
            var response = client.Execute(request);
            JObject result = JsonConvert.DeserializeObject(response.Content) as JObject;
            JArray data = result.First.ToObject<JProperty>().Value.ToObject<JArray>();
            return JsonConvert.DeserializeObject<List<WelkinEvent>>(data.ToString());
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
            return JsonConvert.DeserializeObject<List<WelkinEvent>>(data.ToString());
        }
    }
}