using Ascon.Pilot.SDK;
using UpdateFieldsDocumentActivity.Settings.Providers;

namespace UpdateFieldsDocumentActivity.Settings.Factories
{
    public class SettingProviderFactory<T>
    {
        public static ISettingProvider<T> GetSettingProvider(ISettingValueProvider settingValueProvider)
        {
            return new SettingProvider<T>(new SettingValueProviderProxy<T>(settingValueProvider));
        }

        public static ISettingProvider<T> GetSettingProvider(IPersonalSettings personalSettings, string settingKey)
        {
            return new SettingProvider<T>(new PersonalSettingsProxy<T>(personalSettings, settingKey));
        }
    }
}