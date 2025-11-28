
namespace UpdateFieldsDocumentActivity.Settings.Providers
{
	public class SettingProvider<T> : ISettingProvider<T>
	{
		private ISettingProxy<T> _settingProxy;

		public SettingProvider(ISettingProxy<T> settingProxy)
		{
			_settingProxy = settingProxy;
		}

		public T GetSettings()
		{
			return _settingProxy.Settings;
		}

		public void SaveSettings(T settings)
		{
			_settingProxy.SaveSettings(settings);
		}
	}
}