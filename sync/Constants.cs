using System;

namespace OutlookWelkinSync
{
    public static class Constants
    {
        public const string OutlookEventExtensionsNamespace = "sync.outlook.welkinhealth.com";
        public const string WelkinWorkerEmailKey = "sync.welkin.worker.email";
        public const string OutlookUserObjectKey = "sync.welkin.outlook.user.object";
        public const string WelkinPatientExtensionNamespace = "patient_placeholders_sync_outlook_welkinhealth_com";
        public const string WelkinEventExtensionNamespacePrefix = "sync_outlook_";
        public const string WelkinLastSyncExtensionNamespace = "sync_last_datetime";
        public const string WelkinClientVersionKey = "WelkinClientVersion";
        public const string WelkinV8TenantNameKey = "WelkinV8TenantName";
        public const string WelkinV8InstanceNameKey = "WelkinV8InstanceName";
        public const string WelkinV8UseSandboxKey = "WelkinV8UseSandbox";
        public const string OutlookLinkedWelkinEventIdKey = "LinkedWelkinEventId";
        public const string OutlookPlaceHolderEventKey = "IsOutlookPlaceHolderEvent";
        public const string OutlookLastSyncDateTimeKey = "LastSyncDateTime";
        public const string OutlookUtcTimezoneLabel = "UTC";
        public const string DefaultModality = "call";
        public const string DefaultAppointmentType = "intake_call";
        public const string WelkinCancelledOutcome = "cancelled";
        public const string CalendarEventResourceName = "calendar_events";
        public const string V8CalendarEventResourceName = "calendar/events";
        public const string CalendarResourceName = "calendars";
        public const string V8CalendarResourceName = "calendar";
        public const string ExternalIdResourceName = "external_ids";
        public const string PatientResourceName = "patients";
        public const string WorkerResourceName = "workers";
        public const string SyncNamespaceDateSeparator = ":::";
        public const string DummyPatientEnvVarName = "WelkinDummyPatientId";
        public const string SharedCalUserEnvVarName = "OutlookSharedCalendarUser";
        public const string SharedCalNameEnvVarName = "OutlookSharedCalendarName";
        public const int SecondsToAccountForEventualConsistency = 3;
    }
}