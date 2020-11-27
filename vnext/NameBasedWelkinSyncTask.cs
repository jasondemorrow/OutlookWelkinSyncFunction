namespace OutlookWelkinSync
{
    using Microsoft.Graph;
    
    /// <summary>
    /// For the welkin event given, look for a linked outlook event and sync if it exists. 
    /// If not, get user that created the welkin event. If they have an outlook user with 
    /// the same user name, create a new, corresponding event in that outlook user's 
    /// calendar and link it with the welkin event.
    /// </summary>
    public class NameBasedWelkinSyncTask : WelkinSyncTask
    {
        public NameBasedWelkinSyncTask(WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient) 
        : base(welkinEvent, outlookClient, welkinClient)
        {
        }

        public override Event Sync()
        {
            throw new System.NotImplementedException();
        }
    }
}