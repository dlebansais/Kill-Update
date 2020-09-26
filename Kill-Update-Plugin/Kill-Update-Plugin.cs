using RegistryTools;
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
using Tracing;

namespace KillUpdate
{
    public class KillUpdatePlugin : IPluginClient, IDisposable
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

        public void Initialize(bool isElevated, Dispatcher dispatcher, Settings settings, ITracer logger)
        {
            IsElevated = isElevated;
            Dispatcher = dispatcher;
            Settings = settings;
            Logger = logger;
            KillUpdateCore.Logger = Logger.Write;

            InitTimer();
            InitServiceManager();

            InitializeCommand("Locked",
                              isVisibleHandler: () => true,
                              isEnabledHandler: () => StartType.HasValue && IsElevated,
                              isCheckedHandler: () => IsLockEnabled,
                              commandHandler: OnCommandLock);

            KillUpdateCore.InitZombification();
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
            bool LockIt = !IsLockEnabled;

            Settings.SetBool(LockedSettingName, LockIt);
            ChangeLockMode(LockIt);
            IsMenuChanged = true;
        }

        public bool GetIsIconChanged()
        {
            bool Result = IsLockModeChanged;
            IsLockModeChanged = false;

            return Result;
        }

        public Icon Icon { get { return KillUpdateCore.LoadCurrentIcon(IsElevated, IsLockEnabled); } }

        public Bitmap SelectionBitmap
        {
            get 
            {
                string ResourceName = "Kill-Update.png";

                if (KillUpdateCore.LoadEmbeddedResource(ResourceName, out Bitmap Bitmap))
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
            KillUpdateCore.ExitZombification();
            StopTimer();
        }

        public bool IsClosed
        {
            get { return true; }
        }

        public bool IsElevated { get; private set; }
        public Dispatcher Dispatcher { get; private set; } = null !;
        public RegistryTools.Settings Settings { get; private set; } = null !;
        public ITracer Logger { get; private set; } = null !;

        private Dictionary<ICommand, string> MenuHeaderTable = new Dictionary<ICommand, string>();
        private Dictionary<ICommand, Func<bool>> MenuIsVisibleTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsEnabledTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Func<bool>> MenuIsCheckedTable = new Dictionary<ICommand, Func<bool>>();
        private Dictionary<ICommand, Action> MenuHandlerTable = new Dictionary<ICommand, Action>();
        private bool IsMenuChanged;
        #endregion

        #region Service Manager
        private void InitServiceManager()
        {
            Logger.Write(Category.Debug, "InitServiceManager starting");

            if (Settings.IsValueSet(LockedSettingName))
                if (IsSettingLock)
                    StartType = ServiceStartMode.Disabled;
                else
                    StartType = ServiceStartMode.Manual;

            OnUpdate();

            UpdateTimer.Change(CheckInterval, CheckInterval);
            FullRestartTimer.Change(FullRestartInterval, Timeout.InfiniteTimeSpan);
            UpdateWatch.Start();

            Logger.Write(Category.Debug, "InitServiceManager done");
        }

        private bool IsLockEnabled 
        { 
            get 
            { 
                return (StartType.HasValue && StartType == ServiceStartMode.Disabled); 
            }
        }

        private bool IsSettingLock 
        { 
            get 
            {
                Settings.GetBool(LockedSettingName, true, out bool Result);
                return Result;
            }
        }

        private void UpdateService(ServiceStartMode? previousStartType, bool lockIt)
        {
            ServiceController[] Services = ServiceController.GetServices();
            if (Services == null)
                Logger.Write(Category.Error, "Failed to get services");
            else
            {
                Logger.Write(Category.Information, $"Found {Services.Length} service(s)");

                foreach (ServiceController Service in Services)
                    if (Service.ServiceName == WindowsUpdateServiceName)
                    {
                        Logger.Write(Category.Debug, $"Checking {Service.ServiceName}");

                        StartType = Service.StartType;

                        Logger.Write(Category.Debug, $"Current start type: {StartType}");

                        if (previousStartType.HasValue && previousStartType.Value != StartType.Value)
                        {
                            Logger.Write(Category.Information, "Start type changed");

                            ChangeLockMode(Service, lockIt);
                        }

                        StopIfRunning(Service, lockIt);

                        previousStartType = StartType;
                        break;
                    }
            }
        }

        private void ChangeLockMode(bool lockIt)
        {
            Logger.Write(Category.Debug, "ChangeLockMode starting");

            try
            {
                using ServiceController Service = new ServiceController(WindowsUpdateServiceName);

                if (IsElevated)
                    ChangeLockMode(Service, lockIt);
                else
                    Logger.Write(Category.Warning, "Not elevated, cannot change");
            }
            catch (Exception e)
            {
                Logger.Write(Category.Debug, $"(from ChangeLockMode) {e.Message}");
            }
        }

        private void ChangeLockMode(ServiceController Service, bool lockIt)
        {
            IsLockModeChanged = true;
            ServiceStartMode NewStartType = lockIt ? ServiceStartMode.Disabled : ServiceStartMode.Manual;
            NativeMethods.ChangeStartMode(Service, NewStartType, out _);

            StartType = NewStartType;
            Logger.Write(Category.Debug, $"Service type={StartType}");
        }

        private void StopIfRunning(ServiceController Service, bool lockIt)
        {
            if (lockIt && Service.Status == ServiceControllerStatus.Running && Service.CanStop)
            {
                Logger.Write(Category.Debug, "Stopping service");
                Service.Stop();
                Logger.Write(Category.Debug, "Service stopped");
            }
        }

        private const string WindowsUpdateServiceName = "wuauserv";
        private const string LockedSettingName = "Locked";
        private readonly TimeSpan CheckInterval = TimeSpan.FromSeconds(15);
        private readonly TimeSpan FullRestartInterval = TimeSpan.FromHours(1);
        private ServiceStartMode? StartType;
        private bool IsLockModeChanged;
        #endregion

        #region Timer
        private void InitTimer()
        {
            UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
            FullRestartTimer = new Timer(new TimerCallback(FullRestartTimerCallback));
        }

        private void UpdateTimerCallback(object parameter)
        {
            // Protection against reentering too many times after a sleep/wake up.
            // There must be at most two pending calls to OnUpdate in the dispatcher.
            int NewTimerDispatcherCount = Interlocked.Increment(ref TimerDispatcherCount);
            if (NewTimerDispatcherCount > 2)
            {
                Interlocked.Decrement(ref TimerDispatcherCount);
                return;
            }

            // For debug purpose.
            LastTotalElapsed = Math.Round(UpdateWatch.Elapsed.TotalSeconds, 0);

            Dispatcher.BeginInvoke(new OnUpdateHandler(OnUpdate));
        }

        private void FullRestartTimerCallback(object parameter)
        {
            if (UpdateTimer != null)
            {
                Logger.Write(Category.Debug, "Restarting the timer");

                // Restart the update timer from scratch.
                UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);

                UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
                UpdateTimer.Change(CheckInterval, CheckInterval);

                Logger.Write(Category.Debug, "Timer restarted");
            }
            else
                Logger.Write(Category.Warning, "No timer to restart");

            FullRestartTimer.Change(FullRestartInterval, Timeout.InfiniteTimeSpan);
            Logger.Write(Category.Debug, $"Next check scheduled at {DateTime.UtcNow + FullRestartInterval}");
        }

        private int TimerDispatcherCount = 1;
        private double LastTotalElapsed = double.NaN;

        private delegate void OnUpdateHandler();
        private void OnUpdate()
        {
            try
            {
                Logger.Write(Category.Debug, "%% Running timer callback");

                int LastTimerDispatcherCount = Interlocked.Decrement(ref TimerDispatcherCount);
                UpdateWatch.Restart();

                Logger.Write(Category.Debug, $"Watch restarted, Elapsed = {LastTotalElapsed}, pending count = {LastTimerDispatcherCount}");

                Settings.RenewKey();

                Logger.Write(Category.Debug, "Key renewed");

                ServiceStartMode? PreviousStartType = StartType;
                bool LockIt = IsSettingLock;

                Logger.Write(Category.Debug, "Lock setting read");

                UpdateService(PreviousStartType, LockIt);

                ZombifyMe.Zombification.SetAlive();

                Logger.Write(Category.Debug, "%% Timer callback completed");
            }
            catch (Exception e)
            {
                Logger.Write(Category.Debug, $"(from OnUpdate) {e.Message}");
            }
        }

        private void StopTimer()
        {
            FullRestartTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        }

        private Timer UpdateTimer = null!;
        private Timer FullRestartTimer = null!;
        private Stopwatch UpdateWatch = new Stopwatch();
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
            using (UpdateTimer)
            {
            }

            using (FullRestartTimer)
            {
            }
        }
        #endregion
    }
}
