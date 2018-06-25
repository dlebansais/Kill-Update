using Microsoft.Win32;
using System;

namespace TaskbarIconHost
{
    public class PluginSettings : IPluginSettings, IDisposable
    {
        #region Init
        public PluginSettings(string pluginName, IPluginLogger logger)
        {
            Logger = logger;

            try
            {
                Logger.AddLog("InitSettings starting");

                RegistryKey Key = Registry.CurrentUser.OpenSubKey(@"Software", true);
                Key = Key.CreateSubKey("TaskbarIconHost");

                if (pluginName != null)
                    SettingKey = Key.CreateSubKey("Settings-" + pluginName);
                else
                    SettingKey = Key.CreateSubKey("Main Settings");

                Logger.AddLog("InitSettings done");
            }
            catch (Exception e)
            {
                Logger.AddLog($"(from InitSettings) {e.Message}");
            }
        }
        #endregion

        #region Settings
        public bool IsBoolKeySet(string valueName)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue;
        }

        public bool GetSettingBool(string valueName, bool defaultValue)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue ? (value.Value != 0) : defaultValue;
        }

        public void SetSettingBool(string valueName, bool value)
        {
            SetSettingKey(valueName, value ? 1 : 0, RegistryValueKind.DWord);
        }

        public int GetSettingInt(string valueName, int defaultValue)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue ? value.Value : defaultValue;
        }

        public void SetSettingInt(string valueName, int value)
        {
            SetSettingKey(valueName, value, RegistryValueKind.DWord);
        }

        public string GetSettingString(string valueName, string defaultValue)
        {
            string value = GetSettingKey(valueName) as string;
            return value != null ? value : defaultValue;
        }

        public void SetSettingString(string valueName, string value)
        {
            if (value == null)
                DeleteSetting(valueName);
            else
                SetSettingKey(valueName, value, RegistryValueKind.String);
        }

        public double GetSettingDouble(string valueName, double defaultValue)
        {
            string StringValue = GetSettingKey(valueName) as string;
            if (StringValue != null && double.TryParse(StringValue, out double value))
                return value;
            else
                return defaultValue;
        }

        public void SetSettingDouble(string valueName, double value)
        {
            string StringValue = value.ToString();
            SetSettingKey(valueName, StringValue, RegistryValueKind.String);
        }

        private object GetSettingKey(string valueName)
        {
            try
            {
                return SettingKey?.GetValue(valueName);
            }
            catch
            {
                return null;
            }
        }

        private void SetSettingKey(string valueName, object value, RegistryValueKind kind)
        {
            try
            {
                SettingKey?.SetValue(valueName, value, kind);
            }
            catch
            {
            }
        }

        private void DeleteSetting(string valueName)
        {
            try
            {
                SettingKey?.DeleteValue(valueName, false);
            }
            catch
            {
            }
        }

        private IPluginLogger Logger;
        private RegistryKey SettingKey;
        #endregion

        #region Implementation of IDisposable
        protected virtual void Dispose(bool isDisposing)
        {
            if (isDisposing)
                DisposeNow();
        }

        private void DisposeNow()
        {
            if (SettingKey != null)
            {
                SettingKey.Dispose();
                SettingKey = null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~PluginSettings()
        {
            Dispose(false);
        }
        #endregion
    }
}
