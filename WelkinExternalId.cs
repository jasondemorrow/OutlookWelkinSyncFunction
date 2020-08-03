using System;
using Newtonsoft.Json;

namespace OutlookWelkinSyncFunction
{
    public class WelkinExternalId
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("external_id")]
        public string ExternalId { get; set; }

        [JsonProperty("welkin_id")]
        public string WelkinId { get; set; }

        [JsonProperty("resource")]
        public string ResourceName { get; set; }

        [JsonProperty("namespace")]
        public string Namespace { get; set; }
    }
}