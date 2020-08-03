using System;
using Microsoft.Graph;
using Newtonsoft.Json;

namespace OutlookWelkinSyncFunction
{
    public class WelkinEvent
    {
        public static WelkinEvent CreateDefaultForCalendar(string calendarId)
        {
            WelkinEvent evt = new WelkinEvent();
            evt.CalendarId = calendarId;
            evt.IsAllDay = true;
            evt.Day = DateTime.UtcNow.Date;
            evt.Modality = Constants.DefaultModality;
            evt.AppointmentType = Constants.DefaultAppointmentType;
            // TODO: patient ID?
            
            return evt;
        }

        public void SyncFrom(Event outlookEvent)
        {
            this.IsAllDay = outlookEvent.IsAllDay.HasValue? outlookEvent.IsAllDay.Value : false;
            
            if (this.IsAllDay)
            {
                this.Day = DateTime.Parse(outlookEvent.Start.DateTime).Date;
            }
            else 
            {
                this.Start = DateTime.Parse(outlookEvent.Start.DateTime);
                this.End = DateTime.Parse(outlookEvent.End.DateTime);
            }
        }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("is_all_day")]
        public bool IsAllDay { get; set; }

        [JsonProperty("calendar_id")]
        public string CalendarId { get; set; }

        [JsonProperty("patient_id")]
        public string PatientId { get; set; }

        [JsonProperty("outcome")]
        public string Outcome { get; set; }

        [JsonProperty("modality")]
        public string Modality { get; set; }

        [JsonProperty("appointment_type")]
        public string AppointmentType { get; set; }

        [JsonProperty("updated_at")]
        public DateTime Updated { get; set; }

        [JsonProperty("created_at")]
        public DateTime Created { get; set; }

        [JsonProperty("start_time")]
        public DateTime? Start { get; set; }

        [JsonProperty("end_time")]
        public DateTime? End { get; set; }

        // If this is an all-day event, this is the date it's on
        [JsonProperty("day")]
        public DateTime? Day { get; set; }
    }
}