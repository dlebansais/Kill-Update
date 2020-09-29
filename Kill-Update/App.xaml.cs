namespace KillUpdate
{
    using RegistryTools;
    using SchedulerTools;
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.Reflection;
    using System.Security.Principal;
    using System.ServiceProcess;
    using System.Threading;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using TaskbarTools;
    using Tracing;

    public partial class App : Application, IDisposable
    {
        #region Init
        public App()
        {
            Logger = Tracer.Create("Kill-Update");
            Logger.Write(Category.Debug, "Starting");
            Logger.Write(Category.Information, $"IsElevated={_IsElevated}");

            Settings = new Settings("KillUpdate", "Settings");
            Core = new KillUpdateCore(IsElevated, Settings, Logger);


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

        private void OnStartup(object sender, StartupEventArgs e)
        {
            Logger.Write(Category.Debug, "OnStartup");

            Core.InitTimer(Dispatcher);
            Core.InitServiceManager();
            InitTaskbarIcon();
            Core.InitZombification();

            Exit += OnExit;
        }

        private ITracer Logger;
        private readonly Settings Settings;
        private KillUpdateCore Core;
        private EventWaitHandle? InstanceEvent;
        #endregion

        #region Properties
        public static bool IsElevated
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
                }

                return _IsElevated.Value;
            }
        }
        private static bool? _IsElevated;

        public string ToolTipText { get { return Core.ToolTip; } }
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
                if (Core.LoadEmbeddedResource(BitmapName, out Bitmap Bitmap))
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

            if (!Core.StartType.HasValue)
                IsLockEnabled = false;
            else
            {
                IsLockEnabled = IsElevated;
                if (Core.StartType == ServiceStartMode.Disabled)
                    LockMenu.IsChecked = true;
            }

            AddContextMenu(contextMenu, LoadAtStartup, true, IsElevated);
            AddContextMenu(contextMenu, LockMenu, true, IsLockEnabled);
            AddContextMenuSeparator(contextMenu);
            AddContextMenu(contextMenu, ExitMenu, true, true);

            Icon = Core.LoadCurrentIcon(IsElevated, LockMenu.IsChecked);

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

        private void OnCommandLoadAtStartup(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandLoadAtStartup");

            TaskbarIcon.ToggleMenuCheck(LoadAtStartupCommand, out bool Install);
            InstallLoad(Install);
        }

        private void OnCommandLock(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandLock");

            Core.ToggleLock();

            bool IsLockEnabled = Core.IsLockEnabled;
            TaskbarIcon.SetMenuCheck(LockCommand, Core.IsLockEnabled);

            if (Core.IsLockModeChanged)
            {
                Core.IsLockModeChanged = false;

                using Icon OldIcon = AppIcon;

                AppIcon = Core.LoadCurrentIcon(IsElevated, IsLockEnabled);
                AppTaskbarIcon.UpdateIcon(AppIcon);
            }

            Logger.Write(Category.Information, $"Lock mode: {IsLockEnabled}");
        }

        private void OnCommandExit(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.Write(Category.Debug, "OnCommandExit");

            Shutdown();
        }

        private void OnExit(object sender, ExitEventArgs e)
        {
            Logger.Write(Category.Debug, "Exiting application");

            Core.ExitZombification();

            Core.StopTimer();

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
            using (Core)
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
