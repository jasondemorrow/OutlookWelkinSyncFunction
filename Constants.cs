using System;

namespace OutlookWelkinSyncFunction
{
    public static class Constants
    {
        public static readonly string OutlookExtensionsNamespace = "sync.outlook.welkinhealth.com";
        public static readonly string WelkinPlaceholderPatientNamespace = "patient_placeholders_sync_outlook_welkinhealth_com";
        public static readonly string WelkinPlaceholderEventNamespace = "event_placeholders_sync_outlook_welkinhealth_com";
        public static readonly string LinkedWelkinEventIdKey = "LinkedWelkinEventId";
        public static readonly string DefaultModality = "call";
        public static readonly string DefaultAppointmentType = "intake_call";
    }
}