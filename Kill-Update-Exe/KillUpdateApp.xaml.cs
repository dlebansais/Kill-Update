namespace KillUpdateExe
{
    using System;
    using System.Windows;
    using KillUpdate;
    using TaskbarIconHost;

    public partial class KillUpdateApp : Application, IDisposable
    {
        #region Init
        public KillUpdateApp()
        {
            Plugin = new KillUpdatePlugin();
            PluginApp = new App(this, new PluginDetails(Plugin.Name, Plugin.Guid, "Kill-Update-Plugin"));
        }

        private KillUpdatePlugin Plugin;
        private App PluginApp;
        #endregion

        #region Implementation of IDisposable
        protected virtual void Dispose(bool isDisposing)
        {
            if (isDisposing)
                DisposeNow();
        }

        private void DisposeNow()
        {
            using (Plugin)
            {
            }

            using (PluginApp)
            {
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~KillUpdateApp()
        {
            Dispose(false);
        }
        #endregion
    }
}
