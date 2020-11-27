namespace OutlookWelkinSync 
{
    using Microsoft.Graph;

    /// <summary>
    /// A task for sync'ing two events according to some sync strategy. One of the events passed to it may be null, 
    /// but not both. A null value indicates that the event has not yet been created in the corresponding platform.
    /// </summary>
    public abstract class SyncTask
    {
        private WelkinEvent welkinEvent;
        private Event outlookEvent;
        private WelkinPractitioner welkinPractitioner;
        private User outlookUserForPractitioner;

        public abstract void PerformSync();
    }
}