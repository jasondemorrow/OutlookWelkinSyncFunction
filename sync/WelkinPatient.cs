namespace OutlookWelkinSync
{
    using System;
    using Newtonsoft.Json;
    
    public class WelkinPatient
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("primary_worker_id")]
        public string PrimaryWorkerId { get; set; }

        [JsonProperty("first_name")]
        public string FirstName { get; set; }

        [JsonProperty("last_name")]
        public string LastName { get; set; }

        [JsonProperty("phase")]
        public string Phase { get; set; }

        [JsonProperty("primary_language")]
        public string PrimaryLanguage { get; set; }

        [JsonProperty("timezone")]
        public string Timezone { get; set; }

        [JsonProperty("gender")]
        public string Gender { get; set; }

        [JsonProperty("is_active")]
        public bool IsActive { get; set; }

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