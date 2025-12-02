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
    public class MaskNameTemplateSettingsFeature : ISettingsFeature2
    {
        public string Key => SettingsFeatureKeys.MaskNameTemplateSettingsFeatureKey;

        public string Title => "Разработка настройки ECM-документы — маски имен документов по шаблону";

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
