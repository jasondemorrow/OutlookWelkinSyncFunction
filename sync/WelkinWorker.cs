namespace OutlookWelkinSync
{
    using System;
    using Newtonsoft.Json;
    
    public class WelkinWorker
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

        [JsonProperty("updated_at", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? Updated { get; set; }

        [JsonProperty("created_at", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? Created { get; set; }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}