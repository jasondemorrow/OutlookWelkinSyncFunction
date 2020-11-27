namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Graph;
    using Newtonsoft.Json;

    public class WelkinEvent
    {
        public bool SyncWith(Event outlookEvent)
        {
            bool keepMine = 
                (outlookEvent.LastModifiedDateTime == null) || 
                (this.Updated != null && this.Updated.Value.ToUniversalTime() > outlookEvent.LastModifiedDateTime);

            if (keepMine)
            {
                outlookEvent.IsAllDay = this.IsAllDay;
                if (this.IsAllDay)
                {
                    outlookEvent.Start.DateTime = this.Day.Value.DateTime.ToString("o");
                    outlookEvent.End.DateTime = this.Day.Value.AddDays(1).DateTime.ToString("o");
                }
                else 
                {
                    outlookEvent.Start.DateTime = this.Start.Value.DateTime.ToString("o");
                    outlookEvent.End.DateTime = this.End.Value.DateTime.ToString("o");
                }
            }
            else
            {
                this.IsAllDay = outlookEvent.IsAllDay.HasValue? outlookEvent.IsAllDay.Value : false;
                
                if (this.IsAllDay)
                {
                    this.Day = DateTime.Parse(outlookEvent.Start.DateTime).Date;
                    this.Start = this.Day; // midnight of the start date
                    this.End = this.Day.Value.AddDays(1);
                }
                else 
                {
                    this.Day = null;
                    this.Start = outlookEvent.StartUtc();
                    this.End = outlookEvent.EndUtc();
                }
            }

            return !keepMine; // was changed
        }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("is_all_day")]
        public bool IsAllDay { get; set; }

        [JsonProperty("ignore_unavailable_times", NullValueHandling=NullValueHandling.Ignore)]
        public bool IgnoreUnavailableTimes { get; set; }

        [JsonProperty("ignore_working_hours", NullValueHandling=NullValueHandling.Ignore)]
        public bool IgnoreWorkingHours { get; set; }

        [JsonProperty("calendar_id")]
        public string CalendarId { get; set; }

        [JsonProperty("patient_id")]
        public string PatientId { get; set; }

        [JsonProperty("outcome", NullValueHandling=NullValueHandling.Ignore)]
        public string Outcome { get; set; }

        [JsonProperty("modality")]
        public string Modality { get; set; }

        [JsonProperty("appointment_type")]
        public string AppointmentType { get; set; }

        [JsonProperty("updated_at", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? Updated { get; set; }

        [JsonProperty("created_at", NullValueHandling=NullValueHandling.Ignore)]
        public DateTimeOffset? Created { get; set; }

        [JsonProperty("start_time")]
        public DateTimeOffset? Start { get; set; }

        [JsonProperty("end_time")]
        public DateTimeOffset? End { get; set; }

        // If this is an all-day event, this is the date it's on
        [JsonProperty("day", NullValueHandling=NullValueHandling.Ignore)]
        [JsonConverter(typeof(JsonDateFormatConverter), "yyyy-MM-dd")]
        public DateTimeOffset? Day { get; set; }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }
    }
}