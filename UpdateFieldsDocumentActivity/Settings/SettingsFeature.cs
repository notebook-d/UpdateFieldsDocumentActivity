/*
  Copyright © 2018 ASCON-Design Systems LLC. All rights reserved.
  This sample is licensed under the MIT License.
*/
using Ascon.Pilot.SDK;
using System.ComponentModel.Composition;
using System.Windows;

namespace UpdateFieldsDocumentActivity.Settings
{
    [Export(typeof(ISettingsFeature2))]
    public class SampleSettingsFeature : ISettingsFeature2
    {
        public string Key => SettingsFeatureKeys.SettingKey;

        public string Title => "UpdateFieldsDocumentActivity";

        public FrameworkElement Editor => null;

        public bool IsValid(string settingsItemValue)
        {
            return true;
        }

        public void SetValueProvider(ISettingValueProvider settingValueProvider)
        {
            
        }
    }
}
