namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Ninject.Modules;

    public class NinjectModules
    {
        public static NinjectModule CurrentModule { get; set; } = new ProdModule();
        public static ILogger CurrentLogger { get; set; } = null;

        public class ProdModule : NinjectModule
        {
            public override void Load()
            {
                Bind<ILogger>().ToConstant(CurrentLogger);
                Bind<WelkinConfig>().To<WelkinConfig>(); // just create a default instance, don't modify
                Bind<OutlookConfig>().To<OutlookConfig>(); // just create a default instance, don't modify
                Bind<string>()
                    .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.DummyPatientEnvVarName))
                    .InSingletonScope()
                    .Named(Constants.DummyPatientEnvVarName);

                string welkinClientVersion = Environment.GetEnvironmentVariable(Constants.WelkinClientVersionKey)?.ToLowerInvariant() ?? "7";

                switch(welkinClientVersion)
                {
                    case "8":
                    case "v8":
                    {
                        string sandboxMode = Environment.GetEnvironmentVariable(Constants.WelkinV8UseSandboxKey)?.ToLowerInvariant() ?? "false";
                        bool useSandbox = Boolean.Parse(sandboxMode);
                        Bind<bool>()
                            .ToConstant(useSandbox)
                            .InSingletonScope()
                            .Named(Constants.WelkinV8UseSandboxKey);
                        Bind<string>()
                            .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.WelkinV8TenantNameKey))
                            .InSingletonScope()
                            .Named(Constants.WelkinV8TenantNameKey);
                        Bind<string>()
                            .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.WelkinV8InstanceNameKey))
                            .InSingletonScope()
                            .Named(Constants.WelkinV8InstanceNameKey);
                        Bind<IWelkinClient>().To<WelkinClientV8>().InSingletonScope();
                        break;
                    }
                    default:
                    {
                        Bind<IWelkinClient>().To<WelkinClient>().InSingletonScope();
                        break;
                    }
                }

                Bind<OutlookClient>().To<OutlookClient>().InSingletonScope();
                Bind<OutlookSyncTask>().To<NameBasedOutlookSyncTask>();

                string sharedCalendarUser = Environment.GetEnvironmentVariable(Constants.SharedCalUserEnvVarName);
                string sharedCalendarName = Environment.GetEnvironmentVariable(Constants.SharedCalNameEnvVarName);

                if (!string.IsNullOrEmpty(sharedCalendarUser) && !string.IsNullOrEmpty(sharedCalendarName))
                {
                    Bind<string>()
                        .ToConstant(sharedCalendarUser)
                        .InSingletonScope()
                        .Named(Constants.SharedCalUserEnvVarName);
                    Bind<string>()
                        .ToConstant(sharedCalendarName)
                        .InSingletonScope()
                        .Named(Constants.SharedCalNameEnvVarName);
                    Bind<WelkinSyncTask>().To<SharedCalendarWelkinSyncTask>();
                    Bind<OutlookEventRetrieval>().To<SharedCalendarOutlookEventRetrieval>();
                }
                else
                {
                    Bind<WelkinSyncTask>().To<NameBasedWelkinSyncTask>();
                    Bind<OutlookEventRetrieval>().To<WelkinWorkerOutlookEventRetrieval>();
                }
            }
        }
    }
}