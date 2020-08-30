using System;
using Newtonsoft.Json;

namespace OutlookWelkinSyncFunction
{
    public class WelkinExternalId
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("resource")]
        public string Resource { get; set; }

        [JsonProperty("namespace")]
        public string Namespace { get; set; }

        [JsonProperty("external_id")]
        public string ExternalId { get; set; }

        [JsonProperty("internal_id")]
        public string InternalId { get; set; }
    }
}