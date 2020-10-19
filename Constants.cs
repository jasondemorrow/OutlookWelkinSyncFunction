using System;

namespace OutlookWelkinSyncFunction
{
    public static class Constants
    {
        public static readonly string OutlookEventExtensionsNamespace = "sync.outlook.welkinhealth.com";
        public static readonly string WelkinPatientExtensionNamespace = "patient_placeholders_sync_outlook_welkinhealth_com";
        public static readonly string WelkinEventExtensionNamespacePrefix = "sync_outlook_";
        public static readonly string WelkinLastSyncExtensionNamespace = "sync_last_datetime";
        public static readonly string OutlookLinkedWelkinEventIdKey = "LinkedWelkinEventId";
        public static readonly string OutlookPlaceHolderEventKey = "IsOutlookPlaceHolderEvent";
        public static readonly string OutlookLastSyncDateTimeKey = "LastSyncDateTime";
        public static readonly string DefaultModality = "call";
        public static readonly string DefaultAppointmentType = "intake_call";
        public static readonly string CalendarEventResourceName = "calendar_events";
    }
}