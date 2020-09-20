using Microsoft.Win32;
using Microsoft.Win32.TaskScheduler;
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

namespace KillUpdate
{
    public partial class App : Application, IDisposable
    {
        #region Init
        static App()
        {
            App.AddLog("Starting");

            InitSettings();
        }

#pragma warning disable CS8618 // Non-nullable property is uninitialized
        public App()
        {
            // Ensure only one instance is running at a time.
            App.AddLog("Checking uniqueness");
            try
            {
                bool createdNew;
                InstanceEvent = new EventWaitHandle(false, EventResetMode.ManualReset, "{5293D078-E9B9-4E5D-AD4C-489A017748A5}", out createdNew);
                if (!createdNew)
                {
                    App.AddLog("Another instance is already running");
                    InstanceEvent.Close();
                    InstanceEvent = null;
                    Shutdown();
                    return;
                }
            }
            catch (Exception e)
            {
                App.AddLog($"(from App) {e.Message}");

                Shutdown();
                return;
            }

            Startup += OnStartup;
            ShutdownMode = ShutdownMode.OnExplicitShutdown;
        }
#pragma warning restore CS8618 // Non-nullable property is uninitialized

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

                    App.AddLog($"IsElevated={_IsElevated}");
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
        private static void InitSettings()
        {
            try
            {
                App.AddLog("InitSettings starting");

                RegistryKey Key = Registry.CurrentUser.OpenSubKey(@"Software", true);
                Key = Key.CreateSubKey("KillUpdate");
                SettingKey = Key.CreateSubKey("Settings");

                App.AddLog("InitSettings done");
            }
            catch (Exception e)
            {
                App.AddLog($"(from InitSettings) {e.Message}");
            }
        }

        private static object? GetSettingKey(string valueName)
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

        private static void SetSettingKey(string valueName, object value, RegistryValueKind kind)
        {
            try
            {
                SettingKey?.SetValue(valueName, value, kind);
            }
            catch
            {
            }
        }

        private static void DeleteSetting(string valueName)
        {
            try
            {
                SettingKey?.DeleteValue(valueName, false);
            }
            catch
            {
            }
        }

        public static bool IsBoolKeySet(string valueName)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue;
        }

        public static bool GetSettingBool(string valueName, bool defaultValue)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue ? (value.Value != 0) : defaultValue;
        }

        public static void SetSettingBool(string valueName, bool value)
        {
            SetSettingKey(valueName, value ? 1 : 0, RegistryValueKind.DWord);
        }

        public static int GetSettingInt(string valueName, int defaultValue)
        {
            int? value = GetSettingKey(valueName) as int?;
            return value.HasValue ? value.Value : defaultValue;
        }

        public static void SetSettingInt(string valueName, int value)
        {
            SetSettingKey(valueName, value, RegistryValueKind.DWord);
        }

        public static string GetSettingString(string valueName, string defaultValue)
        {
            string? value = GetSettingKey(valueName) as string;
            return value ?? defaultValue;
        }

        public static void SetSettingString(string valueName, string value)
        {
            if (value == null)
                DeleteSetting(valueName);
            else
                SetSettingKey(valueName, value, RegistryValueKind.String);
        }

        private static RegistryKey? SettingKey;
        #endregion

        #region Taskbar Icon
        private void InitTaskbarIcon()
        {
            App.AddLog("InitTaskbarIcon starting");

            MenuHeaderTable = new Dictionary<ICommand, string>();
            LoadAtStartupCommand = InitMenuCommand("LoadAtStartupCommand", LoadAtStartupHeader, OnCommandLoadAtStartup);
            LockCommand = InitMenuCommand("LockCommand", "Locked", OnCommandLock);
            ExitCommand = InitMenuCommand("ExitCommand", "Exit", OnCommandExit);

            ContextMenu ContextMenu = LoadContextMenu(out Icon Icon);

            TaskbarIcon = TaskbarIcon.Create(Icon, ToolTipText, ContextMenu, ContextMenu);
            TaskbarIcon.MenuOpening += OnMenuOpening;

            App.AddLog("InitTaskbarIcon done");
        }

        private ICommand InitMenuCommand(string commandName, string header, ExecutedRoutedEventHandler executed)
        {
            ICommand Command = (ICommand)FindResource(commandName);
            MenuHeaderTable.Add(Command, header);

            // Bind the command to the corresponding handler. Requires the menu to be the target of notifications in TaskbarIcon.
            CommandManager.RegisterClassCommandBinding(typeof(ContextMenu), new CommandBinding(Command, executed));

            return Command;
        }

        private static Icon LoadIcon(string iconName)
        {
            string ResourceName = GetResourceName(iconName);

            using (Stream ResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                Icon Result = new Icon(ResourceStream);
                App.AddLog($"Resource {iconName} loaded");

                return Result;
            }
        }

        private static Bitmap LoadBitmap(string bitmapName)
        {
            string ResourceName = GetResourceName(bitmapName);

            using (Stream ResourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
            {
                Bitmap Result = new Bitmap(ResourceStream);
                App.AddLog($"Resource {bitmapName} loaded");

                return Result;
            }
        }

        private static string GetResourceName(string name)
        {
            string ResourceName = string.Empty;

            // Loads an "Embedded Resource".
            foreach (string Item in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                if (Item.EndsWith(name, StringComparison.InvariantCulture))
                    ResourceName = Item;

            return ResourceName;
        }

        private ContextMenu LoadContextMenu(out Icon Icon)
        {
            ContextMenu Result = new ContextMenu();

            MenuItem LoadAtStartup;
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
                    LoadAtStartup.Icon = LoadBitmap("UAC-16.png");
                }
            }
            else
            {
                LoadAtStartup = LoadNotificationMenuItem(LoadAtStartupCommand);

                if (!IsElevated)
                    LoadAtStartup.Icon = LoadBitmap("UAC-16.png");
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

            AddContextMenu(Result, LoadAtStartup, true, IsElevated);
            AddContextMenu(Result, LockMenu, true, IsLockEnabled);
            AddContextMenuSeparator(Result);
            AddContextMenu(Result, ExitMenu, true, true);

            Icon = LoadCurrentIcon(LockMenu.IsChecked);

            App.AddLog("Menu created");

            return Result;
        }

        private Icon LoadCurrentIcon(bool isLockEnabled)
        {
            if (IsElevated)
                if (isLockEnabled)
                    return LoadIcon("Locked-Enabled.ico");
                else
                    return LoadIcon("Unlocked-Enabled.ico");
            else
                if (isLockEnabled)
                    return LoadIcon("Locked-Disabled.ico");
                else
                    return LoadIcon("Unlocked-Disabled.ico");
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
            App.AddLog("OnMenuOpening");

            string ExeName = Assembly.GetExecutingAssembly().Location;

            if (IsElevated)
                TaskbarIcon.SetMenuIsChecked(LoadAtStartupCommand, Scheduler.IsTaskActive(ExeName));
            else
            {
                if (Scheduler.IsTaskActive(ExeName))
                    TaskbarIcon.SetMenuHeader(LoadAtStartupCommand, RemoveFromStartupHeader);
                else
                    TaskbarIcon.SetMenuHeader(LoadAtStartupCommand, LoadAtStartupHeader);
            }
        }

        internal TaskbarIcon TaskbarIcon { get; private set; }
        private string LoadAtStartupHeader { get { return (string)TryFindResource("LoadAtStartupHeader"); } }
        private string RemoveFromStartupHeader { get { return (string)TryFindResource("RemoveFromStartupHeader"); } }
        private ICommand LoadAtStartupCommand;
        private ICommand LockCommand;
        private ICommand ExitCommand;
        private Dictionary<ICommand, string> MenuHeaderTable;
        #endregion

        #region Zombification
        private static bool IsRestart { get { return ZombifyMe.Zombification.IsRestart; } }

        private void InitZombification()
        {
            App.AddLog("InitZombification starting");

            if (IsRestart)
                App.AddLog("This process has been restarted");

            Zombification = new ZombifyMe.Zombification("Kill-Update");
            Zombification.Delay = TimeSpan.FromMinutes(1);
            Zombification.WatchingMessage = string.Empty;
            Zombification.RestartMessage = string.Empty;
            Zombification.Flags = ZombifyMe.Flags.NoWindow | ZombifyMe.Flags.ForwardArguments;
            Zombification.IsSymmetric = true;
            Zombification.AliveTimeout = TimeSpan.FromMinutes(1);
            Zombification.ZombifyMe();

            App.AddLog("InitZombification done");
        }

        private void ExitZombification()
        {
            App.AddLog("ExitZombification starting");

            Zombification.Cancel();

            App.AddLog("ExitZombification done");
        }

        private ZombifyMe.Zombification Zombification;
        #endregion

        #region Events
        private void OnStartup(object sender, StartupEventArgs e)
        {
            App.AddLog("OnStartup");

            InitServiceManager();
            InitTaskbarIcon();
            InitZombification();

            Exit += OnExit;
        }

        private void OnCommandLoadAtStartup(object sender, ExecutedRoutedEventArgs e)
        {
            App.AddLog("OnCommandLoadAtStartup");

            TaskbarIcon.ToggleMenuIsChecked(LoadAtStartupCommand, out bool Install);
            InstallLoad(Install);
        }

        private void OnCommandLock(object sender, ExecutedRoutedEventArgs e)
        {
            App.AddLog("OnCommandLock");

            TaskbarIcon.ToggleMenuIsChecked(LockCommand, out bool LockIt);
            SetSettingBool("Locked", LockIt);
            ChangeLockMode(LockIt);

            Icon Icon = LoadCurrentIcon(LockIt);
            TaskbarIcon.UpdateIcon(Icon);

            App.AddLog($"Lock mode: {LockIt}");
        }

        private void OnCommandExit(object sender, ExecutedRoutedEventArgs e)
        {
            App.AddLog("OnCommandExit");

            Shutdown();
        }

        private void OnExit(object sender, ExitEventArgs e)
        {
            App.AddLog("Exiting application");

            ExitZombification();

            StopServiceManager();

            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

            using (TaskbarIcon)
            {
            }

            App.AddLog("Done");
        }
        #endregion

        #region Service Manager
        private void InitServiceManager()
        {
            App.AddLog("InitServiceManager starting");

            if (IsBoolKeySet("Locked"))
                if (GetSettingBool("Locked", true))
                    StartType = ServiceStartMode.Disabled;
                else
                    StartType = ServiceStartMode.Manual;

            OnUpdate();

            UpdateTimer = new Timer(UpdateTimerCallback, this, CheckInterval, CheckInterval);
            UpdateWatch.Start();

            App.AddLog("InitServiceManager done");
        }

        private void StopServiceManager()
        {
            UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
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
                bool LockIt = GetSettingBool("Locked", true);

                ServiceController[] Services = ServiceController.GetServices();

                foreach (ServiceController Service in Services)
                    if (Service.ServiceName == WindowsUpdateServiceName)
                    {
                        StartType = Service.StartType;

                        if (PreviousStartType.HasValue && PreviousStartType.Value != StartType.Value)
                        {
                            App.AddLog("Start type changed");

                            ChangeLockMode(Service, LockIt);
                        }

                        StopIfRunning(Service, LockIt);

                        PreviousStartType = StartType;
                        break;
                    }

                ZombifyMe.Zombification.SetAlive();
            }
            catch (Exception e)
            {
                App.AddLog($"(from OnUpdate) {e.Message}");
            }
        }

        private void ChangeLockMode(bool lockIt)
        {
            App.AddLog("ChangeLockMode starting");

            try
            {
                using ServiceController Service = new ServiceController(WindowsUpdateServiceName);

                if (IsElevated)
                    ChangeLockMode(Service, lockIt);
                else
                    App.AddLog("Not elevated, cannot change");
            }
            catch (Exception e)
            {
                App.AddLog($"(from ChangeLockMode) {e.Message}");
            }
        }

        private void ChangeLockMode(ServiceController Service, bool lockIt)
        {
            ServiceStartMode NewStartType = lockIt ? ServiceStartMode.Disabled : ServiceStartMode.Manual;
            NativeMethods.ChangeStartMode(Service, NewStartType, out _);

            StartType = NewStartType;
            App.AddLog($"Service type={StartType}");
        }

        private static void StopIfRunning(ServiceController Service, bool lockIt)
        {
            if (lockIt && Service.Status == ServiceControllerStatus.Running && Service.CanStop)
            {
                Service.Stop();
                App.AddLog("Service stopped");
            }
        }

        private const string WindowsUpdateServiceName = "wuauserv";
        private readonly TimeSpan CheckInterval = TimeSpan.FromSeconds(10);
        private ServiceStartMode? StartType;
        private Timer UpdateTimer = new Timer((object parameter) => { });
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
        public static void AddLog(string logText)
        {
            DateTime UtcNow = DateTime.UtcNow;
            string TimeLog = UtcNow.ToString(CultureInfo.InvariantCulture) + UtcNow.Millisecond.ToString("D3", CultureInfo.InvariantCulture);
            Debug.WriteLine($"KillUpdate - {TimeLog}: {logText}");
        }
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

            using (InstanceEvent)
            {
            }
        }
        #endregion
    }
}
