
namespace UpdateFieldsDocumentActivity.Settings.Providers
{
    public interface ISettingProxy<T>
    {
        void SaveSettings(T settings);
        T Settings { get; }
    }
}