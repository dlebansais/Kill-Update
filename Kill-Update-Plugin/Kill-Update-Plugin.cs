using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.ServiceProcess;
using System.Threading;
using System.Windows.Input;
using System.Windows.Threading;
using TaskbarIconHost;

namespace KillUpdate
{
    public class KillUpdatePlugin : MarshalByRefObject, TaskbarIconHost.IPluginClient
    {
        #region Plugin
        public string Name
        {
            get { return PluginDetails.Name; }
        }

        public Guid Guid
        {
            get { return PluginDetails.Guid; }
        }

        public bool RequireElevated
        {
            get { return true; }
        }

        public bool HasClickHandler
        {
            get { return false; }
        }

        public void Initialize(bool isElevated, Dispatcher dispatcher, TaskbarIconHost.IPluginSettings settings, TaskbarIconHost.IPluginLogger logger)
        {
            IsElevated = isElevated;
            Dispatcher = dispatcher;
            Settings = settings;
            Logger = logger;

            InitServiceManager();

            InitializeCommand("Locked",
                              isVisibleHandler: () => true,
                              isEnabledHandler: () => StartType.HasValue && IsElevated,
                              isCheckedHandler: () => IsLockEnabled,
                              commandHandler: OnCommandLock);
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

        public bool GetIsMenuChanged()
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

        public Bitmap GetMenuIcon(ICommand command)
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
            bool LockIt = !IsLockEnabled;

            Settings.SetSettingBool(LockedSettingName, LockIt);
            ChangeLockMode(LockIt);
            IsMenuChanged = true;
        }

        public bool GetIsIconChanged()
        {
            bool Result = IsIconChanged;
            IsIconChanged = false;

            return Result;
        }

        public Icon Icon
        {
            get
            {
                if (IsElevated)
                    if (IsLockEnabled)
                        return LoadEmbeddedResource<Icon>("Locked-Enabled.ico");
                    else
                        return LoadEmbeddedResource<Icon>("Unlocked-Enabled.ico");
                else
                    if (IsLockEnabled)
                        return LoadEmbeddedResource<Icon>("Locked-Disabled.ico");
                    else
                        return LoadEmbeddedResource<Icon>("Unlocked-Disabled.ico");
            }
        }

        public Bitmap SelectionBitmap
        {
            get { return LoadEmbeddedResource<Bitmap>("Kill-Update.png"); }
        }

        public void OnIconClicked()
        {
        }

        public bool GetIsToolTipChanged()
        {
            return false;
        }

        public string ToolTip
        {
            get
            {
                if (IsElevated)
                    return "Lock/Unlock Windows updates";
                else
                    return "Lock/Unlock Windows updates (Requires administrator mode)";
            }
        }

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
            StopServiceManager();
        }

        public bool IsClosed
        {
            get { return true; }
        }

        public bool IsElevated { get; private set; }
        public Dispatcher Dispatcher { get; private set; }
        public TaskbarIconHost.IPluginSettings Settings { get; private set; }
        public TaskbarIconHost.IPluginLogger Logger { get; private set; }

        private T LoadEmbeddedResource<T>(string resourceName)
        {
            // Loads an "Embedded Resource" of type T (ex: Bitmap for a PNG file).
            foreach (string ResourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                if (ResourceName.EndsWith(resourceName))
                {
                    using (Stream rs = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
                    {
                        T Result = (T)Activator.CreateInstance(typeof(T), rs);
                        Logger.AddLog($"Resource {resourceName} loaded");

                        return Result;
                    }
                }

            Logger.AddLog($"Resource {resourceName} not found");
            return default(T);
        }

        private Dictionary<ICommand, string> MenuHeaderTable = new Dictionary<ICommand, string>();
        private Dictionary<ICommand, Func<bool>> MenuIsVisibleTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsEnabledTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsCheckedTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Action> MenuHandlerTable = new Dictionary<ICommand, Action>();
        private bool IsIconChanged;
        private bool IsMenuChanged;
        #endregion

        #region Service Manager
        private void InitServiceManager()
        {
            Logger.AddLog("InitServiceManager starting");

            if (Settings.IsBoolKeySet(LockedSettingName))
                if (IsSettingLock)
                    StartType = ServiceStartMode.Disabled;
                else
                    StartType = ServiceStartMode.Manual;

            UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
            UpdateWatch = new Stopwatch();

            OnUpdate();

            UpdateTimer.Change(CheckInterval, CheckInterval);
            UpdateWatch.Start();

            Logger.AddLog("InitServiceManager done");
        }

        private bool IsLockEnabled { get { return (StartType.HasValue && StartType == ServiceStartMode.Disabled); } }
        private bool IsSettingLock { get { return Settings.GetSettingBool(LockedSettingName, true); } }

        private void StopServiceManager()
        {
            UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            UpdateTimer = null;
        }

        private void UpdateTimerCallback(object parameter)
        {
            // Protection against reentering too many times after a sleep/wake up.
            if (UpdateWatch.Elapsed.Seconds < CheckInterval.Seconds / 2)
                return;

            Dispatcher.BeginInvoke(new OnUpdateHandler(OnUpdate));
        }

        private delegate void OnUpdateHandler();
        private void OnUpdate()
        {
            UpdateWatch.Restart();

            try
            {
                ServiceStartMode? PreviousStartType = StartType;
                bool LockIt = IsSettingLock;

                ServiceController[] Services = ServiceController.GetServices();

                foreach (ServiceController Service in Services)
                    if (Service.ServiceName == WindowsUpdateServiceName)
                    {
                        StartType = Service.StartType;

                        if (PreviousStartType.HasValue && PreviousStartType.Value != StartType.Value)
                        {
                            Logger.AddLog("Start type changed");

                            ChangeLockMode(Service, LockIt);
                        }

                        StopIfRunning(Service, LockIt);

                        PreviousStartType = StartType;
                        break;
                    }
            }
            catch (Exception e)
            {
                Logger.AddLog($"(from OnUpdate) {e.Message}");
            }
        }

        private void ChangeLockMode(bool lockIt)
        {
            Logger.AddLog("ChangeLockMode starting");

            try
            {
                ServiceController Service = new ServiceController(WindowsUpdateServiceName);

                if (IsElevated)
                    ChangeLockMode(Service, lockIt);
                else
                    Logger.AddLog("Not elevated, cannot change");
            }
            catch (Exception e)
            {
                Logger.AddLog($"(from ChangeLockMode) {e.Message}");
            }
        }

        private void ChangeLockMode(ServiceController Service, bool lockIt)
        {
            IsIconChanged = true;
            ServiceStartMode NewStartType = lockIt ? ServiceStartMode.Disabled : ServiceStartMode.Manual;
            ServiceHelper.ChangeStartMode(Service, NewStartType);

            StartType = NewStartType;
            Logger.AddLog($"Service type={StartType}");
        }

        private void StopIfRunning(ServiceController Service, bool lockIt)
        {
            if (lockIt && Service.Status == ServiceControllerStatus.Running && Service.CanStop)
            {
                Service.Stop();
                Logger.AddLog("Service stopped");
            }
        }

        private static readonly string WindowsUpdateServiceName = "wuauserv";
        private static readonly string LockedSettingName = "Locked";
        private readonly TimeSpan CheckInterval = TimeSpan.FromSeconds(10);
        private ServiceStartMode? StartType;
        private Timer UpdateTimer;
        private Stopwatch UpdateWatch;
        #endregion
    }
}
