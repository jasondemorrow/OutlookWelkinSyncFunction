namespace OutlookWelkinSync 
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Graph;

    /// <summary>
    /// A queue of sync tasks that knows if a sync task is already queued for a particular event.
    /// </summary>
    public class SyncTaskQueue
    {
        private readonly Queue<SyncTask> internalQueue = new Queue<SyncTask>();
        private readonly HashSet<string> queuedWelkinEventIds = new HashSet<string>();
        private readonly HashSet<string> queuedOutlookEventIds = new HashSet<string>();

        public void CreateAndQueueTaskIfMissing(Event outlookEvent, WelkinEvent welkinEvent, Func<Event, WelkinEvent, SyncTask> taskFactory)
        {
            
        }
    }
}