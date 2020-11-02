using System;
using Microsoft.Graph;
using TimeZoneConverter;

namespace OutlookWelkinSyncFunction
{
    public static class OutlookEventExtensions
    {
        public static DateTimeOffset StartUtc(this Event outlookEvent)
        {
            return ToUtc(outlookEvent.Start);
        }

        public static DateTimeOffset EndUtc(this Event outlookEvent)
        {
            return ToUtc(outlookEvent.End);
        }

        private static DateTimeOffset ToUtc(DateTimeTimeZone dateTimeWithTimeZone)
        {
            TimeZoneInfo timezone = TZConvert.GetTimeZoneInfo(dateTimeWithTimeZone.TimeZone);
            DateTime dateTime = DateTime.Parse(dateTimeWithTimeZone.DateTime);
            DateTimeOffset dateTimeOffset = new DateTimeOffset(dateTime, timezone.GetUtcOffset(dateTime));
            DateTimeOffset utcDateTimeOffset = dateTimeOffset.ToOffset(TimeSpan.Zero);
            return utcDateTimeOffset;
        }
    }
}