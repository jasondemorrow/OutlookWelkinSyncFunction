using System;
using Microsoft.Graph;
using Newtonsoft.Json;
using TimeZoneConverter;

namespace OutlookWelkinSyncFunction
{
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
                    outlookEvent.Start.DateTime = this.Day.Value.ToString("o");
                    outlookEvent.End.DateTime = this.Day.Value.AddDays(1).ToString("o");
                }
                else 
                {
                    outlookEvent.Start.DateTime = this.Start.Value.ToString("o");
                    outlookEvent.End.DateTime = this.End.Value.ToString("o");
                }
            }
            else
            {
                this.IsAllDay = outlookEvent.IsAllDay.HasValue? outlookEvent.IsAllDay.Value : false;
                
                if (this.IsAllDay)
                {
                    /**
                    * Outlook stores the start/end dates for an all day event in UTC. We want midnight-midnight in the timezone local to 
                    * the user, so we calculate that here. There's a bit of a trick to this, since the day of the event depends on the 
                    * user's offset from UTC. If the offset is negative, then the UTC dates will be shifted forward with respect to the 
                    * user's desired range and we want the start date's day from the Outlook event. If the offset is positive, the UTC 
                    * dates in the Outlook event will be earlier than those intended by the user, and we want the end date's day. A 
                    * corner case is if the offset is zero, in which case we want the start date's day.
                    */
                    TimeZoneInfo userTimeZone = TZConvert.GetTimeZoneInfo(outlookEvent.OriginalStartTimeZone);
                    DateTimeOffset outlookEndDateOffset = outlookEvent.EndUtc();

                    // Don't use UtcNow since there may be a daylight savings time switch between now and the event.
                    TimeSpan offset = userTimeZone.GetUtcOffset(outlookEndDateOffset);
                    if (offset < TimeSpan.Zero)
                    {
                        this.Day = DateTime.Parse(outlookEvent.End.DateTime).Date;
                    }
                    else
                    {   
                        this.Day = DateTime.Parse(outlookEvent.Start.DateTime).Date;
                    }
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
    }
}