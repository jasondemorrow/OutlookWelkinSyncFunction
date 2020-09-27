using System;
using System.Collections.Generic;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace OutlookWelkinSyncFunction
{
    public class EventLink
    {
        private readonly OutlookClient outlookClient;
        private readonly WelkinClient welkinClient;
        private readonly User outlookUser;
        private readonly WelkinPractitioner welkinUser;
        private readonly ILogger log;
        public WelkinEvent TargetWelkinEvent { get; set; } = null;
        public Event TargetOutlookEvent { get; set; } = null;
        public WelkinEvent LinkedWelkinEvent { get; private set; } = null; // might be different from target event
        public Event LinkedOutlookEvent { get; private set; } = null; // might be different from target event

        public EventLink(Event outlookEvent, WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, User outlookUser, WelkinPractitioner welkinUser, ILogger log)
        {
            Throw.IfAnyAreNull(outlookClient, welkinClient, outlookUser, welkinUser, log);
            Throw.IfAnyAreNull(outlookUser.Id, welkinUser.Id);
            this.TargetOutlookEvent = outlookEvent;
            this.TargetWelkinEvent = welkinEvent;
            this.outlookClient = outlookClient;
            this.welkinClient = welkinClient;
            this.outlookUser = outlookUser;
            this.welkinUser = welkinUser;
            this.log = log;
        }

        public static string ExistsIn(Event outlookEvent, ILogger log)
        {
                if (outlookEvent.Extensions?.AdditionalData == null || !outlookEvent.Extensions.AdditionalData.ContainsKey(Constants.LinkedWelkinEventIdKey))
                {
                    log.LogInformation($"No linked Welkin event for Outlook event {outlookEvent.ICalUId}");
                    return null;
                }

                string linkedEventId = outlookEvent.Extensions.AdditionalData[Constants.LinkedWelkinEventIdKey].ToString();
                if (string.IsNullOrEmpty(linkedEventId))
                {
                    log.LogInformation($"Null or empty linked Welkin event ID for Outlook event {outlookEvent.ICalUId}");
                    return null;
                }

                return linkedEventId;
        }

        public void Clear()
        {
            this.TargetWelkinEvent = null;
            this.TargetOutlookEvent = null;
            this.LinkedWelkinEvent = null;
            this.LinkedOutlookEvent = null;
        }

        public static WelkinExternalId ExistsIn(WelkinClient welkinClient, WelkinEvent welkinEvent, ILogger log)
        {
                WelkinExternalId externalId = welkinClient.FindExternalMappingFor(welkinEvent);

                if (externalId == null || string.IsNullOrEmpty(externalId.ExternalId))
                {
                    log.LogInformation($"No external, linked Outklook event ID associated with Welkin event {welkinEvent.Id}.");
                    return null;
                }

                return externalId;
        }

        public bool Exists(Direction direction = Direction.OutlookToWelkin | Direction.WelkinToOutlook)
        {
            if (direction.HasFlag(Direction.OutlookToWelkin) && this.TargetOutlookEvent != null &&  this.LinkedWelkinEvent == null)
            {
                string linkedEventId = ExistsIn(this.TargetOutlookEvent, this.log);
                if (string.IsNullOrEmpty(linkedEventId))
                {
                    return false;
                }

                this.log.LogInformation($"Found linked Welkin ID for Outlook event {this.TargetOutlookEvent.ICalUId}: {linkedEventId}.");
                LinkedWelkinEvent = welkinClient.GetEvent(linkedEventId);
                if (LinkedWelkinEvent == null || (this.TargetWelkinEvent != null && !linkedEventId.Equals(this.TargetWelkinEvent.Id)))
                {
                    this.log.LogInformation($"Retrieved linked Welkin event {LinkedWelkinEvent?.Id}, which was not expected.");
                    return false;
                }
            }

            if (direction.HasFlag(Direction.WelkinToOutlook) && this.TargetWelkinEvent != null && this.LinkedOutlookEvent == null)
            {
                WelkinExternalId externalId = ExistsIn(this.welkinClient, this.TargetWelkinEvent, this.log);
                if (externalId == null)
                {
                    return false;
                }

                this.log.LogInformation($"Found linked Outlook ID for Welkin event {this.TargetWelkinEvent.Id}: {externalId.ExternalId}.");
                LinkedOutlookEvent = outlookClient.GetEventForUserWithICalId(this.outlookUser, externalId.ExternalId);
                if (string.IsNullOrEmpty(LinkedOutlookEvent?.ICalUId) || 
                    (this.TargetOutlookEvent != null && !LinkedOutlookEvent.ICalUId.Equals(this.TargetOutlookEvent.ICalUId)))
                {
                    this.log.LogInformation($"Retrieved linked Outlook event {LinkedOutlookEvent?.ICalUId}, which was not expected.");
                    return false;
                }
            }

            return true;
        }

        public void Ensure(Direction direction = Direction.OutlookToWelkin | Direction.WelkinToOutlook)
        {
            Throw.IfAnyAreNull(this.TargetWelkinEvent, this.TargetOutlookEvent);

            if (direction.HasFlag(Direction.OutlookToWelkin) && !this.Exists(Direction.OutlookToWelkin))
            {
                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                keyValuePairs[Constants.LinkedWelkinEventIdKey] = this.TargetWelkinEvent.Id;
                this.outlookClient.SetOpenExtensionPropertiesOnEvent(
                    this.outlookUser, this.TargetOutlookEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);
                this.LinkedWelkinEvent = this.TargetWelkinEvent;
                this.log.LogInformation(
                    $"Successfully created link from Outlook event {this.TargetOutlookEvent.ICalUId} to Welkin event {this.TargetWelkinEvent.Id}");
            }

            if (direction.HasFlag(Direction.WelkinToOutlook) && !this.Exists(Direction.WelkinToOutlook))
            {
                WelkinExternalId welkinExternalId = new WelkinExternalId
                {
                    Resource = Constants.CalendarEventResourceName,
                    ExternalId = this.TargetOutlookEvent.ICalUId,
                    InternalId = this.TargetWelkinEvent.Id,
                    Namespace = Constants.WelkinEventExtensionNamespace
                };
                welkinExternalId = welkinClient.CreateOrUpdateExternalId(welkinExternalId, true);

                if (welkinExternalId?.Id != null && welkinExternalId.Id.Equals(this.TargetOutlookEvent.ICalUId))
                {
                    this.LinkedOutlookEvent = this.TargetOutlookEvent;
                    this.log.LogInformation(
                        $"Successfully created link from Welkin event {this.TargetWelkinEvent.Id} to Outlook event {this.TargetOutlookEvent.ICalUId}");
                }
            }
        }

        [Flags]
        public enum Direction
        {
            None = 0,
            OutlookToWelkin = 1,
            WelkinToOutlook = 2
        }
    }
}