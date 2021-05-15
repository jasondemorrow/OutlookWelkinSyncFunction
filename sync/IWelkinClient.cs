namespace OutlookWelkinSync
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Graph;

    public interface IWelkinClient
    {
        WelkinEvent CancelEvent(WelkinEvent welkinEvent);
        WelkinEvent CreateOrUpdateEvent(WelkinEvent evt, string id = null);
        WelkinExternalId CreateOrUpdateExternalId(WelkinExternalId external, string id = null);
        void DeleteEvent(WelkinEvent welkinEvent);
        void DeleteExternalId(WelkinExternalId externalId);
        IEnumerable<WelkinExternalId> FindExternalEventMappingsUpdatedBetween(DateTimeOffset start, DateTimeOffset end);
        WelkinExternalId FindExternalMappingFor(WelkinEvent internalEvent, Event externalEvent = null);
        WelkinWorker FindWorker(string email);
        WelkinEvent GeneratePlaceholderEventForCalendar(WelkinCalendar calendar);
        bool IsPlaceHolderEvent(WelkinEvent evt);
        IEnumerable<WelkinWorker> RetrieveAllWorkers();
        WelkinCalendar RetrieveCalendar(string calendarId);
        WelkinCalendar RetrieveCalendarFor(WelkinWorker worker);
        WelkinEvent RetrieveEvent(string eventId);
        IEnumerable<WelkinEvent> RetrieveEventsUpdatedSince(TimeSpan ago);
        WelkinLastSyncEntry RetrieveLastSyncFor(WelkinEvent internalEvent);
        WelkinPatient RetrievePatient(string patientId);
        WelkinWorker RetrieveWorker(string workerId);
        bool UpdateLastSyncFor(WelkinEvent internalEvent, string existingId = null, DateTimeOffset? lastSync = null);
    }
}