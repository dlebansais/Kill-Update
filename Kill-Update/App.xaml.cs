namespace KillUpdate
{
    using SchedulerTools;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Reflection;
    using System.Security.Principal;
    using System.ServiceProcess;
    using System.Threading;
    using System.Windows;
    using System.Drawing;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Threading;
    using TaskbarTools;
    using System.Diagnostics;
    using System.Globalization;
    using Tracing;

    public partial class App : Application, IDisposable
    {
        #region Init
        static App()
        {
            KillUpdateCore.Logger = Logger.Write;

            Logger.Write(Category.Debug, "Starting");
        }

        private static ITracer Logger = Tracer.Create("Kill-Update");

        public App()
        {
            // Ensure only one instance is running at a time.
            Logger.Write(Category.Debug, "Checking uniqueness");
            try
            {
                bool createdNew;
                InstanceEvent = new EventWaitHandle(false, EventResetMode.ManualReset, "{5293D078-E9B9-4E5D-AD4C-489A017748A5}", out createdNew);
                if (!createdNew)
                {
                    Logger.Write(Category.Information, "Another instance is already running");
                    InstanceEvent.Close();
                    InstanceEvent = null;
                    Shutdown();
                    return;
                }
            }
            catch (Exception e)
            {
                Logger.Write(Category.Error, $"(from App) {e.Message}");

                Shutdown();
                return;
            }

            Startup += OnStartup;
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
        }

        private EventWaitHandle? InstanceEvent;
        #endregion

        #region Properties
        public bool IsElevated
        {
            get
            {
                if (_IsElevated == null)
                {
                    WindowsIdentity wi = WindowsIdentity.GetCurrent();
                    if (wi != null)
                    {
                        WindowsPrincipal wp = new WindowsPrincipal(wi);
                        if (wp != null)
                            _IsElevated = wp.IsInRole(WindowsBuiltInRole.Administrator);
                        else
                            _IsElevated = false;
                    }
                    else
                        _IsElevated = false;

                    Logger.Write(Category.Information, $"IsElevated={_IsElevated}");
                }

                return _IsElevated.Value;
            }
        }
        private bool? _IsElevated;

        public string ToolTipText
        {
            get
            {
                if (IsElevated)
                    return "Lock/Unlock Windows updates";
                else
                    return "Lock/Unlock Windows updates (Requires administrator mode)";
            }
        }
        #endregion

        #region Settings
        private readonly RegistryTools.Settings Settings = new RegistryTools.Settings("KillUpdate", "Settings");
        #endregion

        #region Taskbar Icon
        private void InitTaskbarIcon()
        {
            Logger.Write(Category.Debug, "InitTaskbarIcon starting");

            LoadAtStartupCommand = InitMenuCommand("LoadAtStartupCommand", LoadAtStartupHeader, OnCommandLoadAtStartup);
            LockCommand = InitMenuCommand("LockCommand", "Locked", OnCommandLock);
            ExitCommand = InitMenuCommand("ExitCommand", "Exit", OnCommandExit);

            LoadContextMenu(out ContextMenu ContextMenu, out Icon Icon);
            AppIcon = Icon;

            AppTaskbarIcon = TaskbarIcon.Create(AppIcon, ToolTipText, ContextMenu, ContextMenu);
            AppTaskbarIcon.MenuOpening += OnMenuOpening;

            Logger.Write(Category.Debug, "InitTaskbarIcon done");
        }

        private ICommand InitMenuCommand(string commandName, string header, ExecutedRoutedEventHandler executed)
        {
            ICommand Command = (ICommand)FindResource(commandName);
            MenuHeaderTable.Add(Command, header);

            // Bind the command to the corresponding handler. Requires the menu to be the target of notifications in TaskbarIcon.
            CommandManager.RegisterClassCommandBinding(typeof(ContextMenu), new CommandBinding(Command, executed));

            return Command;
        }

        private void LoadContextMenu(out ContextMenu contextMenu, out Icon Icon)
        {
            contextMenu = new ContextMenu();

            MenuItem LoadAtStartup;
            bool UseUacIcon = false;
            string BitmapName = string.Empty;

            string ExeName = Assembly.GetExecutingAssembly().Location;
            if (Scheduler.IsTaskActive(ExeName))
            {
                if (IsElevated)
                {
                    LoadAtStartup = LoadNotificationMenuItem(LoadAtStartupCommand);
                    LoadAtStartup.IsChecked = true;
                }
                else
                {
                    LoadAtStartup = LoadNotificationMenuItem(LoadAtStartupCommand, RemoveFromStartupHeader);
                    BitmapName = "UAC-16.png";
                }
            }
            else
            {
                LoadAtStartup = LoadNotificationMenuItem(LoadAtStartupCommand);

                if (!IsElevated)
                    BitmapName = "UAC-16.png";
            }

            if (UseUacIcon)
            {
                if (KillUpdateCore.LoadEmbeddedResource(BitmapName, out Bitmap Bitmap))
                {
                    Logger.Write(Category.Debug, $"Resource {BitmapName} loaded");
                    LoadAtStartup.Icon = Bitmap;
                }
                else
                    Logger.Write(Category.Error, $"Resource {BitmapName} not found");
            }

            MenuItem LockMenu = LoadNotificationMenuItem(LockCommand);
            MenuItem ExitMenu = LoadNotificationMenuItem(ExitCommand);

            bool IsLockEnabled;

            if (!StartType.HasValue)
                IsLockEnabled = false;
            else
            {
                IsLockEnabled = IsElevated;
                if (StartType == ServiceStartMode.Disabled)
                    LockMenu.IsChecked = true;
            }

            AddContextMenu(contextMenu, LoadAtStartup, true, IsElevated);
            AddContextMenu(contextMenu, LockMenu, true, IsLockEnabled);
            AddContextMenuSeparator(contextMenu);
            AddContextMenu(contextMenu, ExitMenu, true, true);

            Icon = KillUpdateCore.LoadCurrentIcon(IsElevated, LockMenu.IsChecked);

            Logger.Write(Category.Debug, "Menu created");
        }

        private MenuItem LoadNotificationMenuItem(ICommand command)
        {
            MenuItem Result = new MenuItem();
            Result.Header = MenuHeaderTable[command];
            Result.Command = command;
            Result.Icon = null;

            return Result;
        }

        private static MenuItem LoadNotificationMenuItem(ICommand command, string header)
        {
            MenuItem Result = new MenuItem();
            Result.Header = header;
            Result.Command = command;
            Result.Icon = null;

            return Result;
        }

        private static void AddContextMenu(ContextMenu menu, MenuItem item, bool isVisible, bool isEnabled)
        {
            TaskbarIcon.PrepareMenuItem(item, isVisible, isEnabled);
            menu.Items.Add(item);
        }

        private static void AddContextMenuSeparator(ContextMenu menu)
        {
            menu.Items.Add(new Separator());
        }

        private void OnMenuOpening(object sender, EventArgs e)
        {
            Logger.Write(Category.Debug, "OnMenuOpening");

            string ExeName = Assembly.GetExecutingAssembly().Location;

            if (IsElevated)
                TaskbarIcon.SetMenuCheck(LoadAtStartupCommand, Scheduler.IsTaskActive(ExeName));
            else
            {
                if (Scheduler.IsTaskActive(ExeName))
                    TaskbarIcon.SetMenuText(LoadAtStartupCommand, RemoveFromStartupHeader);
                else
                    TaskbarIcon.SetMenuText(LoadAtStartupCommand, LoadAtStartupHeader);
            }
        }

        private TaskbarIcon AppTaskbarIcon = TaskbarIcon.Empty;
        private Icon AppIcon = null !;
        private string LoadAtStartupHeader { get { return (string)TryFindResource("LoadAtStartupHeader"); } }
        private string RemoveFromStartupHeader { get { return (string)TryFindResource("RemoveFromStartupHeader"); } }
        private ICommand LoadAtStartupCommand = null!;
        private ICommand LockCommand = null!;
        private ICommand ExitCommand = null!;
        private Dictionary<ICommand, string> MenuHeaderTable = new Dictionary<ICommand, string>();
        #endregion

        #region Events
        private void OnStartup(object sender, StartupEventArgs e)
        {
            Logger.Write(Category.Debug, "OnStartup");

            InitTimer();
            InitServiceManager();
            InitTaskbarIcon();
            KillUpdateCore.InitZombification();

            Exit += OnExit;
        }

        private void OnCommandLoadAtStartup(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandLoadAtStartup");

            TaskbarIcon.ToggleMenuCheck(LoadAtStartupCommand, out bool Install);
            InstallLoad(Install);
        }

        private void OnCommandLock(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandLock");

            TaskbarIcon.ToggleMenuCheck(LockCommand, out bool LockIt);
            Settings.SetBool(LockedSettingName, LockIt);
            ChangeLockMode(LockIt);

            if (IsLockModeChanged)
            {
                IsLockModeChanged = false;

                using Icon OldIcon = AppIcon;

                AppIcon = KillUpdateCore.LoadCurrentIcon(IsElevated, LockIt);
                AppTaskbarIcon.UpdateIcon(AppIcon);
            }

            Logger.Write(Category.Information, $"Lock mode: {LockIt}");
        }

        private void OnCommandExit(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandExit");

            Shutdown();
        }

        private void OnExit(object sender, ExitEventArgs e)
        {
            Logger.Write(Category.Debug, "Exiting application");

            KillUpdateCore.ExitZombification();

            StopTimer();

            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

            using (AppTaskbarIcon)
            {
            }

            Logger.Write(Category.Debug, "Done");
        }
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
                Logger.Write(Category.Debug, $"Found {Services.Length} service(s)");

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
                Logger.Write(Category.Error, $"(from ChangeLockMode) {e.Message}");
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

        private static void StopIfRunning(ServiceController Service, bool lockIt)
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
            if (UpdateWatch.Elapsed.Seconds < CheckInterval.Seconds / 2)
                return;

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

        private delegate void OnUpdateHandler();
        private void OnUpdate()
        {
            UpdateWatch.Restart();

            try
            {
                ServiceStartMode? PreviousStartType = StartType;
                Settings.GetBool(LockedSettingName, true, out bool LockIt);

                UpdateService(PreviousStartType, LockIt);

                ZombifyMe.Zombification.SetAlive();
            }
            catch (Exception e)
            {
                Logger.Write(Category.Error, $"(from OnUpdate) {e.Message}");
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

        #region Load at startup
        private static void InstallLoad(bool isInstalled)
        {
            string ExeName = Assembly.GetExecutingAssembly().Location;

            if (isInstalled)
                Scheduler.AddTask("Kill Windows Update", ExeName, TaskRunLevel.Highest);
            else
                Scheduler.RemoveTask(ExeName, out bool IsFound);
        }
        #endregion

        #region Debugging
        /*
        public static void AddLog(string logText)
        {
            DateTime UtcNow = DateTime.UtcNow;
            string TimeLog = UtcNow.ToString(CultureInfo.InvariantCulture) + UtcNow.Millisecond.ToString("D3", CultureInfo.InvariantCulture);
            Debug.WriteLine($"KillUpdate - {TimeLog}: {logText}");
        }
        */
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
        /// Finalizes an instance of the <see cref="App"/> class.
        /// </summary>
        ~App()
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

            using (InstanceEvent)
            {
            }

            using (AppTaskbarIcon)
            {
            }

            using (AppIcon)
            {
            }

            using (Settings)
            {
            }
        }
        #endregion
    }
}
