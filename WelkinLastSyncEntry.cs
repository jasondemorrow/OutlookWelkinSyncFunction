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
            int idxDateStart = 
                externalId.Namespace.IndexOf(Constants.SyncNamespaceDateSeparator) + 
                Constants.SyncNamespaceDateSeparator.Length;
            
            if (idxDateStart > Constants.SyncNamespaceDateSeparator.Length)
            {
                string dateString = externalId.Namespace.Substring(idxDateStart);
                this.Time = DateTimeOffset.ParseExact(dateString, "o", CultureInfo.InvariantCulture);
            }
        }

        public Boolean IsValid()
        {
            return this.Time != null && this.ExternalId != null;
        }
    }
}