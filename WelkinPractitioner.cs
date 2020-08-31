using System;
using Newtonsoft.Json;

namespace OutlookWelkinSyncFunction
{
    public class WelkinPractitioner
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("first_name")]
        public string FirstName { get; set; }

        [JsonProperty("last_name")]
        public string LastName { get; set; }

        [JsonProperty("email")]
        public string Email { get; set; }

        [JsonProperty("timezone")]
        public string Timezone { get; set; }
    }
}