using System;
using Newtonsoft.Json;

namespace OutlookWelkinSync
{
    public class WelkinCalendar
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("worker_id")]
        public string PractitionerId { get; set; }

        [JsonProperty("updated_at")]
        public DateTime Updated { get; set; }

        [JsonProperty("created_at")]
        public DateTime Created { get; set; }
    }
}