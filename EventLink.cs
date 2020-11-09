using System;
using System.Collections.Generic;
using System.Linq;
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
                Extension extensionForWelkin = outlookEvent?.Extensions?.Where(e => e.Id.EndsWith(Constants.OutlookEventExtensionsNamespace))?.FirstOrDefault();
                if (extensionForWelkin?.AdditionalData == null || !extensionForWelkin.AdditionalData.ContainsKey(Constants.OutlookLinkedWelkinEventIdKey))
                {
                    log.LogInformation($"No linked Welkin event for Outlook event {outlookEvent.ICalUId}");
                    return null;
                }

                string linkedEventId = extensionForWelkin.AdditionalData[Constants.OutlookLinkedWelkinEventIdKey]?.ToString();
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

        public static WelkinExternalId ExistsIn(WelkinClient welkinClient, WelkinEvent welkinEvent, ILogger log, Event outlookEvent = null)
        {
                WelkinExternalId externalId = welkinClient.FindExternalMappingFor(welkinEvent, outlookEvent);

                if (externalId == null || string.IsNullOrEmpty(externalId.ExternalId))
                {
                    log.LogInformation($"No external, linked Outklook event ID associated with Welkin event {welkinEvent.Id}.");
                    return null;
                }

                return externalId;
        }

        public bool FetchAndPopulateIfExists(Direction direction = Direction.OutlookToWelkin | Direction.WelkinToOutlook)
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
                WelkinExternalId externalId = ExistsIn(this.welkinClient, this.TargetWelkinEvent, this.log, this.TargetOutlookEvent);
                if (externalId == null)
                {
                    return false;
                }

                /**
                * Outlook's UUID for calendar events, ICalUId, is not a properly formed GUID, which the Welkin External ID API expects.
                * We use a hashing method to generate a consistent, synthetic GUID for ICalUId, appending its original value to the 
                * namespace string stored in the External ID. This makes the API happy while allowing us to derive the GUID
                * from ICalUId when necessary and vice versa, in a way that is adequately collision-resistent.
                */
                string outlookICalId = externalId.Namespace.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);
                this.log.LogInformation($"Found linked Outlook ID for Welkin event {this.TargetWelkinEvent.Id}: {outlookICalId}.");

                this.LinkedOutlookEvent = outlookClient.GetEventForUserWithICalId(this.outlookUser, outlookICalId);
                if (string.IsNullOrEmpty(LinkedOutlookEvent?.ICalUId) || 
                    (this.TargetOutlookEvent != null && !LinkedOutlookEvent.ICalUId.Equals(this.TargetOutlookEvent.ICalUId)))
                {
                    this.log.LogInformation($"Retrieved linked Outlook event {LinkedOutlookEvent?.ICalUId}, which was not expected.");
                    return false;
                }
            }

            return true; // exists
        }

        public void Ensure(Direction direction = Direction.OutlookToWelkin | Direction.WelkinToOutlook)
        {
            Throw.IfAnyAreNull(this.TargetWelkinEvent, this.TargetOutlookEvent);

            if (direction.HasFlag(Direction.OutlookToWelkin) && !this.FetchAndPopulateIfExists(Direction.OutlookToWelkin))
            {
                Dictionary<string, object> keyValuePairs = new Dictionary<string, object>();
                keyValuePairs[Constants.OutlookLinkedWelkinEventIdKey] = this.TargetWelkinEvent.Id;
                this.outlookClient.MergeOpenExtensionPropertiesOnEvent(
                    this.outlookUser, this.TargetOutlookEvent, keyValuePairs, Constants.OutlookEventExtensionsNamespace);
                this.LinkedWelkinEvent = this.TargetWelkinEvent;
                this.log.LogInformation(
                    $"Successfully created link from Outlook event {this.TargetOutlookEvent.ICalUId} to Welkin event {this.TargetWelkinEvent.Id}");
            }

            if (direction.HasFlag(Direction.WelkinToOutlook) && !this.FetchAndPopulateIfExists(Direction.WelkinToOutlook))
            {
                /**
                * Outlook's UUID for calendar events, ICalUId, is not a properly formed GUID, which the Welkin External ID API expects.
                * We use a hashing method to generate a consistent, synthetic GUID for ICalUId, appending its original value to the 
                * namespace string stored in the External ID. This makes the API happy while allowing us to derive the GUID
                * from ICalUId when necessary and vice versa, in a way that is adequately collision-resistent.
                */
                string derivedGuid = Guids.FromText(this.TargetOutlookEvent.ICalUId).ToString();
                WelkinExternalId welkinExternalId = new WelkinExternalId
                {
                    Resource = Constants.CalendarEventResourceName,
                    ExternalId = derivedGuid,
                    InternalId = this.TargetWelkinEvent.Id,
                    Namespace = Constants.WelkinEventExtensionNamespacePrefix + this.TargetOutlookEvent.ICalUId
                };
                welkinExternalId = welkinClient.CreateOrUpdateExternalId(welkinExternalId);
                string outlookICalId = welkinExternalId?.Namespace?.Substring(Constants.WelkinEventExtensionNamespacePrefix.Length);

                if (outlookICalId != null && outlookICalId.Equals(this.TargetOutlookEvent.ICalUId))
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