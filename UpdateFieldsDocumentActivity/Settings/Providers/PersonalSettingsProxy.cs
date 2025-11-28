using Ascon.Pilot.SDK;
using System;
using System.Collections.Generic;
using UpdateFieldsDocumentActivity.Services;

namespace UpdateFieldsDocumentActivity.Settings.Providers
{
    class PersonalSettingsProxy<T> : ISettingProxy<T>, IObserver<KeyValuePair<string, string>>
    {
        private IPersonalSettings _personalSettings;
        private string _settingKey;
        public PersonalSettingsProxy(IPersonalSettings personalSettings, string settingKey)
        {
            _personalSettings = personalSettings;
            _settingKey = settingKey;
            _personalSettings.SubscribeSetting(_settingKey).Subscribe(this);
        }
        public T Settings { get; private set; }
        public void OnCompleted(){}
        public void OnError(Exception error){}
        public void OnNext(KeyValuePair<string, string> value)
        {
            if (value.Key != _settingKey) return;
            if (value.Value is T)
            {
                Settings = (T)(object)value.Value;
                return;
            }
            Settings = SerializeService.DeserializeFromString<T>(value.Value);
        }

        public void SaveSettings(T settings)
        {
            var value = SerializeService.SerializeToString(settings);
            _personalSettings.ChangeSettingValue(_settingKey, value);
            Settings = settings;
        }
    }
}
