using System;

namespace OutlookWelkinSyncFunction
{
    public class WelkinConfig
    {
        /// <summary>
        /// Welkin API endpoint
        /// </summary>
        public string ApiUrl { get; set; } = "https://api.welkinhealth.com/v1/";

        /// <summary>
        /// Welkin token endpoint
        /// </summary>
        public string TokenUrl { get; set; } = "https://api.welkinhealth.com/v1/token";

        /// <summary>
        /// Guid used by the application to uniquely identify itself
        /// </summary>
        public string ClientId { get; set; }

        /// <summary>
        /// Client secret (application password)
        /// </summary>
        public string ClientSecret { get; set; }

        /// <summary>
        /// Scope
        /// </summary>
        public string Scope { get; set; } = "workers.read"; //"calendars.all calendar_events.all workers.read";

        /// <summary>
        /// Grant type
        /// </summary>
        public string GrantType { get; set; } = "urn:ietf:params:oauth:grant-type:jwt-bearer";

        /// <summary>
        /// Reads the configuration from a json file
        /// </summary>
        /// <param name="path">Path to the configuration json file</param>
        /// <returns>WelkinConfig read from the json file</returns>
        public WelkinConfig()
        {
            this.ClientId = Environment.GetEnvironmentVariable("WelkinClientId");
            this.ClientSecret = Environment.GetEnvironmentVariable("WelkinClientSecret");
            this.Scope = Environment.GetEnvironmentVariable("WelkinScope");
        }
    }
}