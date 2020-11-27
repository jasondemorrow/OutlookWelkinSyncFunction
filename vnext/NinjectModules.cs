namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Logging;
    using Ninject.Modules;

    public class NinjectModules
    {
        public static NinjectModule CurrentModule { get; set; } = new ProdModule();
        public static ILogger CurrentLogger { get; set; } = null;
        public static readonly string WelkinDummyPatientIdBinding = "WelkinDummyPatientIdBinding";

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
                    .ToMethod((context) => Environment.GetEnvironmentVariable("WelkinDummyPatientId"))
                    .Named(WelkinDummyPatientIdBinding);
            }
        }
    }
}