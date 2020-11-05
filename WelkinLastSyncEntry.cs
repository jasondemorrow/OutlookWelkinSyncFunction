using System;
using System.Globalization;

namespace OutlookWelkinSyncFunction
{
    public class WelkinLastSyncEntry
    {
        public WelkinExternalId ExternalId { get; private set; }
        public DateTimeOffset Time { get; private set; }

        public WelkinLastSyncEntry(WelkinExternalId externalId)
        {
            this.ExternalId = externalId;
            this.Time = DateTimeOffset.ParseExact(externalId.ExternalId, "o", CultureInfo.InvariantCulture);
        }

        public Boolean IsValid()
        {
            return this.Time != null && this.ExternalId != null;
        }
    }
}