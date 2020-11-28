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
                if (CurrentLogger == null)
                {
                    ILoggerFactory loggerFactory = LoggerFactory.Create(builder => {
                        builder
                            .AddFilter("Microsoft", LogLevel.Warning)
                            .AddFilter("System", LogLevel.Warning)
                            .AddConsole();
                    });
                    CurrentLogger = loggerFactory.CreateLogger<ProdModule>();
                }
                Bind<ILogger>().ToConstant(CurrentLogger);
                Bind<WelkinConfig>().To<WelkinConfig>(); // just create a default instance, don't modify
                Bind<OutlookConfig>().To<OutlookConfig>(); // just create a default instance, don't modify
                Bind<string>()
                    .ToMethod((context) => Environment.GetEnvironmentVariable(Constants.DummyPatientEnvVarName))
                    .InSingletonScope()
                    .Named("DummyPatientId");
                Bind<OutlookClient>().To<OutlookClient>().InSingletonScope();
                Bind<WelkinClient>().To<WelkinClient>().InSingletonScope();
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
                }
                else
                {
                    Bind<WelkinSyncTask>().To<NameBasedWelkinSyncTask>();
                }
            }
        }
    }
}