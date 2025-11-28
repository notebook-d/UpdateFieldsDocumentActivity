using Ascon.Pilot.SDK;
using UpdateFieldsDocumentActivity.Services;

namespace UpdateFieldsDocumentActivity.Settings.Providers
{
    class SettingValueProviderProxy<T> : ISettingProxy<T>
    {
        private ISettingValueProvider _settingValueProvider;

        public SettingValueProviderProxy(ISettingValueProvider settingValueProvider)
        {
            _settingValueProvider = settingValueProvider;
        }

        public T Settings
        {
            get
            {
                var setting = _settingValueProvider.GetValue();
                if (string.IsNullOrEmpty(setting))
                    return default(T);
                if (setting is T)
                    return (T)(object)setting;
                return SerializeService.DeserializeFromString<T>(setting);
            }
        }

        public void SaveSettings(T settings)
        {
            var value = SerializeService.SerializeToString(settings);
            _settingValueProvider.SetValue(value);
        }


    }
}