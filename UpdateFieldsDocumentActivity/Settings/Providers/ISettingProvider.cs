

namespace UpdateFieldsDocumentActivity.Settings.Providers
{
    public interface ISettingProvider<T>
    {
        T GetSettings();
        void SaveSettings(T settings);
    }
}
