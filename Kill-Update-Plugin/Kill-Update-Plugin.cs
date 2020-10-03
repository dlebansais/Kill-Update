namespace KillUpdate
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Reflection;
    using System.Windows.Input;
    using System.Windows.Threading;
    using RegistryTools;
    using TaskbarIconHost;
    using Tracing;

    public class KillUpdatePlugin : IPluginClient, IDisposable
    {
        #region Plugin
        public string Name
        {
            get { return PluginDetails2.Name; }
        }

        public Guid Guid
        {
            get { return PluginDetails2.Guid; }
        }

        public bool RequireElevated
        {
            get { return true; }
        }

        public bool HasClickHandler
        {
            get { return false; }
        }

        private KillUpdateCore Core = null !;

        public void Initialize(bool isElevated, Dispatcher dispatcher, Settings settings, ITracer logger)
        {
            Core = new KillUpdateCore(isElevated, settings, logger);
            IsElevated = isElevated;
            Dispatcher = dispatcher;
            Settings = settings;
            Logger = logger;

            Core.InitTimer(dispatcher);
            Core.InitServiceManager();

            InitializeCommand("Locked",
                              isVisibleHandler: () => true,
                              isEnabledHandler: () => Core.StartType.HasValue && IsElevated,
                              isCheckedHandler: () => Core.IsLockEnabled,
                              commandHandler: OnCommandLock);

            Core.InitZombification();
        }

        private void InitializeCommand(string header, Func<bool> isVisibleHandler, Func<bool> isEnabledHandler, Func<bool> isCheckedHandler, Action commandHandler)
        {
            ICommand Command = new RoutedUICommand();
            CommandList.Add(Command);
            MenuHeaderTable.Add(Command, header);
            MenuIsVisibleTable.Add(Command, isVisibleHandler);
            MenuIsEnabledTable.Add(Command, isEnabledHandler);
            MenuIsCheckedTable.Add(Command, isCheckedHandler);
            MenuHandlerTable.Add(Command, commandHandler);
        }

        public List<ICommand> CommandList { get; private set; } = new List<ICommand>();

        public bool GetIsMenuChanged(bool beforeMenuOpening)
        {
            bool Result = IsMenuChanged;
            IsMenuChanged = false;

            return Result;
        }

        public string GetMenuHeader(ICommand command)
        {
            return MenuHeaderTable[command];
        }

        public bool GetMenuIsVisible(ICommand command)
        {
            return MenuIsVisibleTable[command]();
        }

        public bool GetMenuIsEnabled(ICommand command)
        {
            return MenuIsEnabledTable[command]();
        }

        public bool GetMenuIsChecked(ICommand command)
        {
            return MenuIsCheckedTable[command]();
        }

        public Bitmap? GetMenuIcon(ICommand command)
        {
            return null;
        }

        public void OnMenuOpening()
        {
        }

        public void OnExecuteCommand(ICommand command)
        {
            MenuHandlerTable[command]();
        }

        private void OnCommandLock()
        {
            Core.ToggleLock();
            IsMenuChanged = true;
        }

        public bool GetIsIconChanged()
        {
            bool Result = Core.IsLockModeChanged;
            Core.IsLockModeChanged = false;

            return Result;
        }

        public Icon Icon
        {
            get
            {
                return LoadCurrentIcon(IsElevated, Core.IsLockEnabled);
            }
        }

        #region Icon & Bitmap
        public Icon LoadCurrentIcon(bool isElevated, bool isLockEnabled)
        {
            string ResourceName;

            if (isElevated)
                if (isLockEnabled)
                    ResourceName = "Locked-Enabled.ico";
                else
                    ResourceName = "Unlocked-Enabled.ico";
            else
                if (isLockEnabled)
                ResourceName = "Locked-Disabled.ico";
            else
                ResourceName = "Unlocked-Disabled.ico";

            if (ResourceTools.LoadEmbeddedResource(PluginDetails2.AssemblyName, ResourceName, out Icon Icon))
                Logger.Write(Category.Debug, $"Resource {ResourceName} loaded");
            else
                Logger.Write(Category.Debug, $"Resource {ResourceName} not found");

            return Icon;
        }
        #endregion

        public Bitmap SelectionBitmap
        {
            get
            {
                string ResourceName = "Kill-Update.png";

                if (ResourceTools.LoadEmbeddedResource(PluginDetails2.AssemblyName, ResourceName, out Bitmap Bitmap))
                    Logger.Write(Category.Debug, $"Resource {ResourceName} loaded");
                else
                    Logger.Write(Category.Error, $"Resource {ResourceName} not found");

                return Bitmap;
            }
        }

        public void OnIconClicked()
        {
        }

        public bool GetIsToolTipChanged()
        {
            return false;
        }

        public string ToolTip { get { return Core.ToolTip; } }

        public void OnActivated()
        {
        }

        public void OnDeactivated()
        {
        }

        public bool CanClose(bool canClose)
        {
            return true;
        }

        public void BeginClose()
        {
            Core.ExitZombification();
            Core.StopTimer();
        }

        public bool IsClosed
        {
            get { return true; }
        }

        public bool IsElevated { get; private set; }
        public Dispatcher Dispatcher { get; private set; } = null !;
        public Settings Settings { get; private set; } = null !;
        public ITracer Logger { get; private set; } = null !;

        private Dictionary<ICommand, string> MenuHeaderTable = new Dictionary<ICommand, string>();
        private Dictionary<ICommand, Func<bool>> MenuIsVisibleTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsEnabledTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsCheckedTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Action> MenuHandlerTable = new Dictionary<ICommand, Action>();
        private bool IsMenuChanged;
        #endregion

        #region Implementation of IDisposable
        /// <summary>
        /// Called when an object should release its resources.
        /// </summary>
        /// <param name="isDisposing">Indicates if resources must be disposed now.</param>
        protected virtual void Dispose(bool isDisposing)
        {
            if (!IsDisposed)
            {
                IsDisposed = true;

                if (isDisposing)
                    DisposeNow();
            }
        }

        /// <summary>
        /// Called when an object should release its resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Finalizes an instance of the <see cref="KillUpdatePlugin"/> class.
        /// </summary>
        ~KillUpdatePlugin()
        {
            Dispose(false);
        }

        /// <summary>
        /// True after <see cref="Dispose(bool)"/> has been invoked.
        /// </summary>
        private bool IsDisposed;

        /// <summary>
        /// Disposes of every reference that must be cleaned up.
        /// </summary>
        private void DisposeNow()
        {
            using (Core)
            {
            }
        }
        #endregion
    }
}
