namespace OutlookWelkinSync
{
    using Microsoft.Graph;

    /// <summary>
    /// For the outlook event given, look for a linked welkin event and sync if it exists. 
    /// If not, get user that created the outlook event. If they have a welkin user with 
    /// the same user name, create a new, corresponding event in that welkin user's 
    /// schedule and link it with the outlook event.
    /// </summary>
    public class NameBasedOutlookSyncTask : OutlookSyncTask
    {
        public NameBasedOutlookSyncTask(Event outlookEvent) : base(outlookEvent)
        {
        }

        public override WelkinEvent Sync()
        {
            throw new System.NotImplementedException();
        }
    }
}