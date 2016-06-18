using System;
using System.Configuration;

namespace AutomaticSchedule
{
    public class AppSettings
    {
        public static class Google
        {
            public static string ClientId => ConfigHelper.GetConfigurationFromAppSettings("Google.ClientId");
            public static string ClientSecret => ConfigHelper.GetConfigurationFromAppSettings("Google.ClientSecret");
            public static string API_KEY => ConfigHelper.GetConfigurationFromAppSettings("Google.API_KEY");

            public static string CalendarIdentifyer => ConfigHelper.GetConfigurationFromAppSettings("Google.CalendarIdentifyer");
            public static string DefaultCalendar => ConfigHelper.GetConfigurationFromAppSettings("Google.DefaultCalendar");
        }

        public static string WantedExtention => ConfigHelper.GetConfigurationFromAppSettings("WantedExtention");
        public static string WantedWorker => ConfigHelper.GetConfigurationFromAppSettings("WantedWorker");

        public static bool SuggestTrainigs => ConfigHelper.GetBooleanFromConfiguration(false, "SuggestTrainigs");
        public static bool SendMailNotification => ConfigHelper.GetBooleanFromConfiguration(false, "SendMailNotification");
        public static bool IsLocalWork => ConfigHelper.GetBooleanFromConfiguration(true, "IsLocalWork");
    }

    public class ConfigHelper
    {
        public static string GetConfigurationFromAppSettings(params string[] configNames)
        {
            foreach (var str in configNames)
            {
                var value = ConfigurationManager.AppSettings[str];
                if (value.IsNotNullOrEmpty())
                {
                    return value;
                }
            }
            return null;
        }

        public static T GetEnumFromConfiguration<T>(params string[] configNames)
        {
            var value = GetConfigurationFromAppSettings(configNames);
            if (string.IsNullOrEmpty(value))
                return default(T);
            try
            {
                return (T)Enum.Parse(typeof(T), value, true);
            }
            catch(Exception ex)
            {
                Utils.WriteErrorLog(ex);
                throw new Exception("Could not convert string value " + value + " to enum " + typeof(T).FullName);
            }
        }

        public static bool GetBooleanFromConfiguration(bool defaultValue, params string[] configNames)
        {
            var value = GetConfigurationFromAppSettings(configNames);
            return value.IsNullOrEmpty() ? defaultValue : value.EqualsIgnoreCase("true");
        }
    }
}
