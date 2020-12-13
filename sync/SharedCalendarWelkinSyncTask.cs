namespace OutlookWelkinSync
{
    using Microsoft.Extensions.Logging;
    using Microsoft.Graph;
    using Ninject;

    /// <summary>
    /// For the welkin event given, look for a linked outlook event in the configured 
    /// shared calendar (by user name and calendar name) and sync if it exists. 
    /// If no corresponding event exists in the shared calendar, create it and 
    /// and link it with the welkin event.
    /// </summary>
    public class SharedCalendarWelkinSyncTask : WelkinSyncTask
    {
        private readonly string sharedCalendarUser;
        private readonly string sharedCalendarName;

        public SharedCalendarWelkinSyncTask(
            WelkinEvent welkinEvent, OutlookClient outlookClient, WelkinClient welkinClient, ILogger logger,
            [Named(Constants.SharedCalUserEnvVarName)] string sharedCalendarUser,
            [Named(Constants.SharedCalNameEnvVarName)] string sharedCalendarName
            ) : base(welkinEvent, outlookClient, welkinClient, logger)
        {
            this.sharedCalendarUser = sharedCalendarUser;
            this.sharedCalendarName = sharedCalendarName;
        }

        public override Event Sync()
        {
            throw new System.NotImplementedException();
        }
    }
}