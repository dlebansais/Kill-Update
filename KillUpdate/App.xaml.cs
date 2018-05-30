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
    public partial class App : Application
    {
        #region Init
        static App()
        {
            App.AddLog("Starting");

            InitSettings();
        }

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

        private EventWaitHandle InstanceEvent;
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

        private static object GetSettingKey(string valueName)
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
            string value = GetSettingKey(valueName) as string;
            return value != null ? value : defaultValue;
        }

        public static void SetSettingString(string valueName, string value)
        {
            if (value == null)
                DeleteSetting(valueName);
            else
                SetSettingKey(valueName, value, RegistryValueKind.String);
        }

        private static RegistryKey SettingKey = null;
        #endregion

        #region Taskbar Icon
        private void InitTaskbarIcon()
        {
            App.AddLog("InitTaskbarIcon starting");

            MenuHeaderTable = new Dictionary<ICommand, string>();
            LoadAtStartupCommand = InitMenuCommand("LoadAtStartupCommand", LoadAtStartupHeader, OnCommandLoadAtStartup);
            LockCommand = InitMenuCommand("LockCommand", "Locked", OnCommandLock);
            ExitCommand = InitMenuCommand("ExitCommand", "Exit", OnCommandExit);

            Icon Icon;
            ContextMenu ContextMenu = LoadContextMenu(out Icon);

            TaskbarIcon = TaskbarIcon.Create(Icon, ToolTipText, ContextMenu, ContextMenu);
            TaskbarIcon.MenuOpening += OnMenuOpening;

            App.AddLog("InitTaskbarIcon done");
        }

        private ICommand InitMenuCommand(string commandName, string header, ExecutedRoutedEventHandler executed)
        {
            ICommand Command = FindResource(commandName) as ICommand;
            MenuHeaderTable.Add(Command, header);

            // Bind the command to the corresponding handler. Requires the menu to be the target of notifications in TaskbarIcon.
            CommandManager.RegisterClassCommandBinding(typeof(ContextMenu), new CommandBinding(Command, executed));

            return Command;
        }

        private Icon LoadIcon(string iconName)
        {
            // Loads an "Embedded Resource" icon.
            foreach (string ResourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                if (ResourceName.EndsWith(iconName))
                {
                    using (Stream rs = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
                    {
                        Icon Result = new Icon(rs);
                        App.AddLog($"Resource {iconName} loaded");

                        return Result;
                    }
                }

            App.AddLog($"Resource {iconName} not found");
            return null;
        }

        private Bitmap LoadBitmap(string bitmapName)
        {
            // Loads an "Embedded Resource" bitmap.
            foreach (string ResourceName in Assembly.GetExecutingAssembly().GetManifestResourceNames())
                if (ResourceName.EndsWith(bitmapName))
                {
                    using (Stream rs = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName))
                    {
                        Bitmap Result = new Bitmap(rs);
                        App.AddLog($"Resource {bitmapName} loaded");

                        return Result;
                    }
                }

            App.AddLog($"Resource {bitmapName} not found");
            return null;
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

            Icon = LoadCurrentIcon(IsLockEnabled);

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

        private MenuItem LoadNotificationMenuItem(ICommand command, string header)
        {
            MenuItem Result = new MenuItem();
            Result.Header = header;
            Result.Command = command;
            Result.Icon = null;

            return Result;
        }

        private void AddContextMenu(ContextMenu menu, MenuItem item, bool isVisible, bool isEnabled)
        {
            TaskbarIcon.PrepareMenuItem(item, isVisible, isEnabled);
            menu.Items.Add(item);
        }

        private void AddContextMenuSeparator(ContextMenu menu)
        {
            menu.Items.Add(new Separator());
        }

        private void OnMenuOpening(object sender, EventArgs e)
        {
            App.AddLog("OnMenuOpening");

            TaskbarIcon SenderIcon = sender as TaskbarIcon;
            string ExeName = Assembly.GetExecutingAssembly().Location;

            if (IsElevated)
                SenderIcon.SetCheck(LoadAtStartupCommand, Scheduler.IsTaskActive(ExeName));
            else
            {
                if (Scheduler.IsTaskActive(ExeName))
                    SenderIcon.SetText(LoadAtStartupCommand, RemoveFromStartupHeader);
                else
                    SenderIcon.SetText(LoadAtStartupCommand, LoadAtStartupHeader);
            }
        }

        public TaskbarIcon TaskbarIcon { get; private set; }
        private static readonly string LoadAtStartupHeader = "Load at startup";
        private static readonly string RemoveFromStartupHeader = "Remove from startup";
        private ICommand LoadAtStartupCommand;
        private ICommand LockCommand;
        private ICommand ExitCommand;
        private Dictionary<ICommand, string> MenuHeaderTable;
        #endregion

        #region Events
        private void OnStartup(object sender, StartupEventArgs e)
        {
            App.AddLog("OnStartup");

            InitServiceManager();
            InitTaskbarIcon();

            Exit += OnExit;
        }

        private void OnCommandLoadAtStartup(object sender, ExecutedRoutedEventArgs e)
        {
            App.AddLog("OnCommandLoadAtStartup");

            TaskbarIcon.ToggleChecked(LoadAtStartupCommand, out bool Install);
            InstallLoad(Install);
        }

        private void OnCommandLock(object sender, ExecutedRoutedEventArgs e)
        {
            App.AddLog("OnCommandLock");

            TaskbarIcon.ToggleChecked(LockCommand, out bool LockIt);
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

            StopServiceManager();

            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

            using (TaskbarIcon Icon = TaskbarIcon)
            {
                TaskbarIcon = null;
            }

            App.AddLog("Done");
        }
        #endregion

        #region Service Manager
        private void InitServiceManager()
        {
            App.AddLog("InitServiceManager starting");

            if (GetSettingBool("Locked", true))
                StartType = ServiceStartMode.Disabled;

            UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
            UpdateWatch = new Stopwatch();

            OnUpdate();

            UpdateTimer.Change(CheckInterval, CheckInterval);
            UpdateWatch.Start();

            App.AddLog("InitServiceManager done");
        }

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

                ServiceController[] Services = ServiceController.GetServices();

                foreach (ServiceController Service in Services)
                    if (Service.ServiceName == WindowsUpdateServiceName)
                    {
                        StartType = Service.StartType;

                        if (PreviousStartType.HasValue && PreviousStartType.Value != StartType.Value)
                        {
                            App.AddLog("Changing lock mode because service start type doesn't match");
                            ChangeLockMode(GetSettingBool("Locked", true));
                        }

                        PreviousStartType = StartType;
                        break;
                    }
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
                ServiceController Service = new ServiceController(WindowsUpdateServiceName);

                if (IsElevated)
                {
                    ServiceStartMode NewStartType = lockIt ? ServiceStartMode.Disabled : ServiceStartMode.Manual;
                    ServiceHelper.ChangeStartMode(Service, NewStartType);

                    StartType = NewStartType;
                    App.AddLog($"Service type={StartType}");
                }
                else
                {
                    App.AddLog("Not elevated, cannot change");
                }
            }
            catch (Exception e)
            {
                App.AddLog($"(from ChangeLockMode) {e.Message}");
            }
        }

        private static readonly string WindowsUpdateServiceName = "wuauserv";
        private readonly TimeSpan CheckInterval = TimeSpan.FromSeconds(10);
        private ServiceStartMode? StartType;
        private Timer UpdateTimer;
        private Stopwatch UpdateWatch;
        #endregion

        #region Load at startup
        private void InstallLoad(bool isInstalled)
        {
            string ExeName = Assembly.GetExecutingAssembly().Location;

            if (isInstalled)
                Scheduler.AddTask("Kill Windows Update", ExeName, TaskRunLevel.Highest);
            else
                Scheduler.RemoveTask(ExeName, out bool IsFound);
        }
        #endregion

        #region Debugging
        public static void AddLog(string text)
        {
            DateTime UtcNow = DateTime.UtcNow;
            string TimeLog = UtcNow.ToString(CultureInfo.InvariantCulture) + UtcNow.Millisecond.ToString("D3");
            Debug.WriteLine($"KillUpdate - {TimeLog}: {text}");
        }
        #endregion
    }
}
