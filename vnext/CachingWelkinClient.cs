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

    public class CachingWelkinClient
    {
        public static readonly CachingWelkinClient Instance = new CachingWelkinClient();
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

        private CachingWelkinClient()
        {
            IKernel ninjectKernel = new StandardKernel(NinjectModules.CurrentModule);
            this.config = ninjectKernel.Get<WelkinConfig>();
            this.logger = ninjectKernel.Get<ILogger>();
            this.dummyPatientId = ninjectKernel.Get<string>(NinjectModules.WelkinDummyPatientIdBinding);
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

            return JsonConvert.DeserializeObject<IEnumerable<T>>(data.ToString());
        }
    }
}