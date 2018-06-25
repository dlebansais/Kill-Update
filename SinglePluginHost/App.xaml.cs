using Microsoft.Win32.TaskScheduler;
using SchedulerTools;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using TaskbarTools;

namespace TaskbarIconHost
{
    public partial class App : Application, IDisposable
    {
        #region Init
        static App()
        {
            InitLogger();
            Logger.AddLog("Starting");
        }

        public App()
        {
            // Ensure only one instance is running at a time.
            Logger.AddLog("Checking uniqueness");

            try
            {
                Guid AppGuid = PluginDetails.Guid;
                if (AppGuid == Guid.Empty)
                {
                    // In case the guid is provided by the project settings and not source code.
                    GuidAttribute AppGuidAttribute = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>();
                    AppGuid = Guid.Parse(AppGuidAttribute.Value);
                }

                string AppUniqueId = AppGuid.ToString("B").ToUpper();

                // Try to create a global named event with a unique name. If we can create it we are first, otherwise there is another instance.
                // In that case, we just abort.
                bool createdNew;
                InstanceEvent = new EventWaitHandle(false, EventResetMode.ManualReset, AppUniqueId, out createdNew);
                if (!createdNew)
                {
                    Logger.AddLog("Another instance is already running");
                    InstanceEvent.Close();
                    InstanceEvent = null;
                    Shutdown();
                    return;
                }
            }
            catch (Exception e)
            {
                Logger.AddLog($"(from App) {e.Message}");

                Shutdown();
                return;
            }

            // This code is here mostly to make sure that the Taskbar static class is initialized ASAP.
            // The taskbar rectangle is never empty. And if it is, we have no purpose.
            if (Taskbar.ScreenBounds.IsEmpty)
                Shutdown();
            else
            {
                Startup += OnStartup;

                // Make sure we stop only on a call to Shutdown. This is for plugins that have a main window, we don't want to exit when it's closed.
                ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }
        }

        private void OnStartup(object sender, StartupEventArgs e)
        {
            Logger.AddLog("OnStartup");

            InitTimer();

            // The plugin manager can fail for various reasons. If it does, we just abort.
            if (InitPlugInManager())
            {
                // Install the taskbar icon and create its menu.
                InitTaskbarIcon();

                // Get a notification when we get or loose the focus.
                Activated += OnActivated;
                Deactivated += OnDeactivated;
                Exit += OnExit;
            }
            else
            {
                CleanupTimer();
                Shutdown();
            }
        }

        // The taskbar got the focus.
        private void OnActivated(object sender, EventArgs e)
        {
            PluginManager.OnActivated();
        }

        // The taskbar lost the focus.
        private void OnDeactivated(object sender, EventArgs e)
        {
            PluginManager.OnDeactivated();
        }

        // Someone called Exit on the application. Time to clean things up.
        private void OnExit(object sender, ExitEventArgs e)
        {
            Logger.AddLog("Exiting application");

            // Set this flag to minimize asynchronous activities.
            IsExiting = true;

            StopPlugInManager();
            CleanupTaskbarIcon();
            CleanupTimer();

            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

            // Explicit display of the last message since timed debug is not running anymore.
            Logger.AddLog("Done");
            UpdateLogger();
        }

        private bool IsExiting;
        private EventWaitHandle InstanceEvent;
        #endregion

        #region Properties
        public bool IsElevated
        {
            get
            {
                // Evaluate this property once, and return the same value after that.
                // This elevated status is the administrator mode, it never changes during the lifetime of the application.
                if (!_IsElevated.HasValue)
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

                    Logger.AddLog($"IsElevated={_IsElevated}");
                }

                return _IsElevated.Value;
            }
        }
        private bool? _IsElevated;
        #endregion

        #region Plugin Manager
        private bool InitPlugInManager()
        {
            if (!PluginManager.Init(IsElevated, PluginDetails.AssemblyName, PluginDetails.Guid, Dispatcher, Logger))
                return false;

            // In the case of a single plugin version, this code won't do anything.
            // However, if several single plugin versions run concurrently, the last one to run will be the preferred one for another plugin host.
            GlobalSettings = new PluginSettings(null, Logger);

            try
            {
                // Assign the guid with a value taken from the registry. The try/catch blocks allows us to ignore invalid ones.
                PluginManager.PreferredPluginGuid = new Guid(GlobalSettings.GetSettingString(PreferredPluginSettingName, PluginManager.GuidToString(Guid.Empty)));
            }
            catch
            {
            }

            return true;
        }

        private void StopPlugInManager()
        {
            // Save this plugin guid so that the last saved will be the preferred one if there is another plugin host.
            GlobalSettings.SetSettingString(PreferredPluginSettingName, PluginManager.GuidToString(PluginManager.PreferredPluginGuid));
            PluginManager.Shutdown();

            CleanupPlugInManager();
        }

        private void CleanupPlugInManager()
        {
            using (PluginSettings PluginSettings = GlobalSettings)
            {
                GlobalSettings = null;
            }
        }

        private static readonly string PreferredPluginSettingName = "PreferredPlugin";
        private PluginSettings GlobalSettings;
        #endregion

        #region Taskbar Icon
        private void InitTaskbarIcon()
        {
            Logger.AddLog("InitTaskbarIcon starting");

            // Create and bind the load at startip/remove from startup command.
            LoadAtStartupCommand = new RoutedUICommand();
            AddMenuCommand(LoadAtStartupCommand, OnCommandLoadAtStartup);

            // Create and bind the exit command.
            ExitCommand = new RoutedUICommand();
            AddMenuCommand(ExitCommand, OnCommandExit);

            // Do the same with all plugin commands.
            foreach (KeyValuePair<List<ICommand>, string> Entry in PluginManager.FullCommandList)
            {
                List<ICommand> FullPluginCommandList = Entry.Key;
                foreach (ICommand Command in FullPluginCommandList)
                    AddMenuCommand(Command, OnPluginCommand);
            }

            // Get the preferred icon and tooltip, and build the taskbar context menu.
            Icon Icon = PluginManager.Icon;
            string ToolTip = PluginManager.ToolTip;
            ContextMenu ContextMenu = LoadContextMenu();

            // Install the taskbar icon.
            TaskbarIcon = TaskbarIcon.Create(Icon, ToolTip, ContextMenu, ContextMenu);
            TaskbarIcon.MenuOpening += OnMenuOpening;
            TaskbarIcon.IconClicked += OnIconClicked;

            Logger.AddLog("InitTaskbarIcon done");
        }

        private void CleanupTaskbarIcon()
        {
            using (TaskbarIcon Icon = TaskbarIcon)
            {
                TaskbarIcon = null;
            }
        }

        private void AddMenuCommand(ICommand command, ExecutedRoutedEventHandler executed)
        {
            // The command can be null if a separator is intended.
            if (command == null)
                return;

            // Bind the command to the corresponding handler. Requires the menu to be the target of notifications in TaskbarIcon.
            CommandManager.RegisterClassCommandBinding(typeof(ContextMenu), new CommandBinding(command, executed));
        }

        private T LoadEmbeddedResource<T>(string resourceName)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            // Loads an "Embedded Resource" of type T (ex: Bitmap for a PNG file).
            // Make sure the resource is tagged as such in the resource properties.
            foreach (string ResourceName in assembly.GetManifestResourceNames())
                if (ResourceName.EndsWith(resourceName))
                {
                    using (Stream rs = assembly.GetManifestResourceStream(ResourceName))
                    {
                        T Result = (T)Activator.CreateInstance(typeof(T), rs);
                        Logger.AddLog($"Resource {resourceName} loaded");

                        return Result;
                    }
                }

            // If not found, it could be because it's not tagged as "Embedded Resource".
            Logger.AddLog($"Resource {resourceName} not found");
            return default(T);
        }

        private ContextMenu LoadContextMenu()
        {
            // Create the taskbar context menu and populate it with menu items, submenus and separators.
            ContextMenu Result = new ContextMenu();
            ItemCollection Items = Result.Items;

            MenuItem LoadAtStartup;
            string ExeName = Assembly.GetExecutingAssembly().Location;

            // Create a menu item for the load at startup/remove from startup command, depending on the current situation.
            // UAC-16.png is the recommended 'shield' icon to indicate administrator mode is required for the operation.
            if (Scheduler.IsTaskActive(ExeName))
            {
                if (IsElevated)
                    LoadAtStartup = CreateMenuItem(LoadAtStartupCommand, LoadAtStartupHeader, true, null);
                else
                    LoadAtStartup = CreateMenuItem(LoadAtStartupCommand, RemoveFromStartupHeader, false, LoadEmbeddedResource<Bitmap>("UAC-16.png"));
            }
            else
            {
                if (IsElevated)
                    LoadAtStartup = CreateMenuItem(LoadAtStartupCommand, LoadAtStartupHeader, false, null);
                else
                    LoadAtStartup = CreateMenuItem(LoadAtStartupCommand, LoadAtStartupHeader, false, LoadEmbeddedResource<Bitmap>("UAC-16.png"));
            }

            TaskbarIcon.PrepareMenuItem(LoadAtStartup, true, true);
            Items.Add(LoadAtStartup);

            // Below load at startup, we add plugin menus.
            // Separate them in two categories, small and large. Small menus are always directly visible in the main context menu.
            // Large plugin menus, if there is more than one, have their own submenu. If there is just one plugin with a large menu we don't bother.

            Dictionary<List<MenuItem>, string> FullPluginMenuList = new Dictionary<List<MenuItem>, string>();
            int LargePluginMenuCount = 0;

            foreach (KeyValuePair<List<ICommand>, string> Entry in PluginManager.FullCommandList)
            {
                List<ICommand> FullPluginCommandList = Entry.Key;
                string PluginName = Entry.Value;

                List<MenuItem> PluginMenuList = new List<MenuItem>();
                int VisiblePluginMenuCount = 0;

                foreach (ICommand Command in FullPluginCommandList)
                    if (Command == null)
                        PluginMenuList.Add(null); // This will result in the creation of a separator.
                    else
                    {
                        string MenuHeader = PluginManager.GetMenuHeader(Command);
                        bool MenuIsVisible = PluginManager.GetMenuIsVisible(Command);
                        bool MenuIsEnabled = PluginManager.GetMenuIsEnabled(Command);
                        bool MenuIsChecked = PluginManager.GetMenuIsChecked(Command);
                        Bitmap MenuIcon = PluginManager.GetMenuIcon(Command);

                        MenuItem PluginMenu = CreateMenuItem(Command, MenuHeader, MenuIsChecked, MenuIcon);
                        TaskbarIcon.PrepareMenuItem(PluginMenu, MenuIsVisible, MenuIsEnabled);

                        PluginMenuList.Add(PluginMenu);

                        // Count how many visible items to decide if the menu is large or small.
                        if (MenuIsVisible)
                            VisiblePluginMenuCount++;
                    }

                if (VisiblePluginMenuCount > 1)
                    LargePluginMenuCount++;

                FullPluginMenuList.Add(PluginMenuList, PluginName);
            }

            // Add small menus, then large menus.
            AddPluginMenuItems(Items, FullPluginMenuList, false, false);
            AddPluginMenuItems(Items, FullPluginMenuList, true, LargePluginMenuCount > 1);

            // If there are more than one plugin capable of receiving click notification, we must give the user a way to choose which.
            // For this purpose, create a "Icons" menu with a choice of plugins, with their name and preferred icon.
            if (PluginManager.ConsolidatedPluginList.Count > 1)
            {
                Items.Add(new Separator());

                MenuItem IconSubmenu = new MenuItem();
                IconSubmenu.Header = "Icons";

                foreach (IPluginClient Plugin in PluginManager.ConsolidatedPluginList)
                {
                    Guid SubmenuGuid = Plugin.Guid;
                    if (!IconSelectionTable.ContainsKey(SubmenuGuid)) // Protection against plugins reusing a guid...
                    {
                        RoutedUICommand SubmenuCommand = new RoutedUICommand();
                        SubmenuCommand.Text = PluginManager.GuidToString(SubmenuGuid);

                        string SubmenuHeader = Plugin.Name;
                        bool SubmenuIsChecked = (SubmenuGuid == PluginManager.PreferredPluginGuid); // The currently preferred plugin will be checked as so.
                        Bitmap SubmenuIcon = Plugin.SelectionBitmap;

                        AddMenuCommand(SubmenuCommand, OnCommandSelectPreferred);
                        MenuItem PluginMenu = CreateMenuItem(SubmenuCommand, SubmenuHeader, SubmenuIsChecked, SubmenuIcon);
                        TaskbarIcon.PrepareMenuItem(PluginMenu, true, true);
                        IconSubmenu.Items.Add(PluginMenu);

                        IconSelectionTable.Add(SubmenuGuid, SubmenuCommand);
                    }
                }

                // Add this "Icons" menu to the main context menu.
                Items.Add(IconSubmenu);
            }

            // Always add a separator above the exit menu.
            Items.Add(new Separator());

            MenuItem ExitMenu = CreateMenuItem(ExitCommand, "Exit", false, null);
            TaskbarIcon.PrepareMenuItem(ExitMenu, true, true);
            Items.Add(ExitMenu);

            Logger.AddLog("Menu created");

            return Result;
        }

        private void AddPluginMenuItems(ItemCollection Items, Dictionary<List<MenuItem>, string> FullPluginMenuList, bool largeSubmenu, bool useSubmenus)
        {
            bool AddSeparator = true;

            foreach (KeyValuePair<List<MenuItem>, string> Entry in FullPluginMenuList)
            {
                List<MenuItem> PluginMenuList = Entry.Key;

                // Only add the category of plugin menu targetted by this call.
                if ((PluginMenuList.Count <= 1 && largeSubmenu) || (PluginMenuList.Count > 1 && !largeSubmenu))
                    continue;

                string PluginName = Entry.Value;

                if (AddSeparator)
                    Items.Add(new Separator());

                ItemCollection SubmenuItems;
                if (useSubmenus)
                {
                    AddSeparator = false;
                    MenuItem PluginSubmenu = new MenuItem();
                    PluginSubmenu.Header = PluginName;
                    SubmenuItems = PluginSubmenu.Items;
                    Items.Add(PluginSubmenu);
                }
                else
                {
                    AddSeparator = true;
                    SubmenuItems = Items;
                }

                // null in the plugin menu means separator.
                foreach (MenuItem MenuItem in PluginMenuList)
                    if (MenuItem != null)
                        SubmenuItems.Add(MenuItem);
                    else
                        SubmenuItems.Add(new Separator());
            }
        }

        private MenuItem CreateMenuItem(ICommand command, string header, bool isChecked, Bitmap icon)
        {
            MenuItem Result = new MenuItem();
            Result.Header = header;
            Result.Command = command;
            Result.IsChecked = isChecked;
            Result.Icon = icon;

            return Result;
        }

        private void OnMenuOpening(object sender, EventArgs e)
        {
            Logger.AddLog("OnMenuOpening");

            TaskbarIcon SenderIcon = sender as TaskbarIcon;
            string ExeName = Assembly.GetExecutingAssembly().Location;

            // Update the load at startup menu with the current state (the user can change it directly in the Task Scheduler at any time).
            if (IsElevated)
                SenderIcon.SetMenuIsChecked(LoadAtStartupCommand, Scheduler.IsTaskActive(ExeName));
            else
            {
                if (Scheduler.IsTaskActive(ExeName))
                    SenderIcon.SetMenuHeader(LoadAtStartupCommand, RemoveFromStartupHeader);
                else
                    SenderIcon.SetMenuHeader(LoadAtStartupCommand, LoadAtStartupHeader);
            }

            // Update the menu with latest news from plugins.
            UpdateMenu(true);

            PluginManager.OnMenuOpening();
        }

        private void UpdateMenu(bool beforeMenuOpening)
        {
            List<ICommand> ChangedCommandList = PluginManager.GetChangedCommands(beforeMenuOpening);

            foreach (ICommand Command in ChangedCommandList)
            {
                // Update changed menus with their new state.
                bool MenuIsVisible = PluginManager.GetMenuIsVisible(Command);
                if (MenuIsVisible)
                {
                    TaskbarIcon.SetMenuIsVisible(Command, true);
                    TaskbarIcon.SetMenuHeader(Command, PluginManager.GetMenuHeader(Command));
                    TaskbarIcon.SetMenuIsEnabled(Command, PluginManager.GetMenuIsEnabled(Command));

                    Bitmap MenuIcon = PluginManager.GetMenuIcon(Command);
                    if (MenuIcon != null)
                    {
                        TaskbarIcon.SetMenuIsChecked(Command, false);
                        TaskbarIcon.SetMenuIcon(Command, MenuIcon);
                    }
                    else
                        TaskbarIcon.SetMenuIsChecked(Command, PluginManager.GetMenuIsChecked(Command));
                }
                else
                    TaskbarIcon.SetMenuIsVisible(Command, false);
            }
        }

        private void OnIconClicked(object sender, EventArgs e)
        {
            PluginManager.OnIconClicked();
        }

        public TaskbarIcon TaskbarIcon { get; private set; }
        private static readonly string LoadAtStartupHeader = "Load at startup";
        private static readonly string RemoveFromStartupHeader = "Remove from startup";
        private ICommand LoadAtStartupCommand;
        private ICommand ExitCommand;
        private Dictionary<Guid, ICommand> IconSelectionTable = new Dictionary<Guid, ICommand>();
        private bool IsIconChanged;
        private bool IsToolTipChanged;
        #endregion

        #region Events
        private void OnCommandLoadAtStartup(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.AddLog("OnCommandLoadAtStartup");

            if (IsElevated)
            {
                // The user is changing the state.
                TaskbarIcon.ToggleMenuIsChecked(LoadAtStartupCommand, out bool Install);
                InstallLoad(Install, PluginDetails.Name);
            }
            else
            {
                // The user would like to change the state, we show to them a dialog box that explains how to do it.
                string ExeName = Assembly.GetExecutingAssembly().Location;

                if (Scheduler.IsTaskActive(ExeName))
                {
                    RemoveFromStartupWindow Dlg = new RemoveFromStartupWindow(PluginDetails.Name);
                    Dlg.ShowDialog();
                }
                else
                {
                    LoadAtStartupWindow Dlg = new LoadAtStartupWindow(PluginManager.RequireElevated, PluginDetails.Name);
                    Dlg.ShowDialog();
                }
            }
        }

        private bool GetIsIconOrToolTipChanged()
        {
            IsIconChanged |= PluginManager.GetIsIconChanged();
            IsToolTipChanged |= PluginManager.GetIsToolTipChanged();

            return IsIconChanged || IsToolTipChanged;
        }

        private void UpdateIconAndToolTip()
        {
            if (IsIconChanged)
            {
                IsIconChanged = false;
                Icon Icon = PluginManager.Icon;
                TaskbarIcon.UpdateIcon(Icon);
            }

            if (IsToolTipChanged)
            {
                IsToolTipChanged = false;
                string ToolTip = PluginManager.ToolTip;
                TaskbarIcon.UpdateToolTip(ToolTip);
            }
        }

        private void OnPluginCommand(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.AddLog("OnPluginCommand");

            PluginManager.OnExecuteCommand(e.Command);

            // After a command is executed, update the menu with the new state.
            // This allows us to do it in the background, instead of when the menu is being opened. It's smoother (or should be).
            if (GetIsIconOrToolTipChanged())
                UpdateIconAndToolTip();

            UpdateMenu(false);
        }

        private void OnCommandSelectPreferred(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.AddLog("OnCommandSelectPreferred");

            RoutedUICommand SubmenuCommand = e.Command as RoutedUICommand;

            Guid NewSelectedGuid = new Guid(SubmenuCommand.Text);
            Guid OldSelectedGuid = PluginManager.PreferredPluginGuid;
            if (NewSelectedGuid != OldSelectedGuid)
            {
                PluginManager.PreferredPluginGuid = NewSelectedGuid;

                // If the preferred plugin changed, make sure the icon and tooltip reflect the change.
                IsIconChanged = true;
                IsToolTipChanged = true;
                UpdateIconAndToolTip();

                // Check the new plugin in the menu, and uncheck the previous one.
                if (IconSelectionTable.ContainsKey(NewSelectedGuid))
                    TaskbarIcon.SetMenuIsChecked(IconSelectionTable[NewSelectedGuid], true);

                if (IconSelectionTable.ContainsKey(OldSelectedGuid))
                    TaskbarIcon.SetMenuIsChecked(IconSelectionTable[OldSelectedGuid], false);
            }
        }

        private void OnCommandExit(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.AddLog("OnCommandExit");

            Shutdown();
        }
        #endregion

        #region Timer
        private void InitTimer()
        {
            // Create a timer to display traces asynchrousnously.
            AppTimer = new Timer(new TimerCallback(AppTimerCallback));
            AppTimer.Change(CheckInterval, CheckInterval);
        }

        private void AppTimerCallback(object parameter)
        {
            // If a shutdown is started, don't show traces anymore so the shutdown can complete smoothly.
            if (IsExiting)
                return;

            // Print traces asynchronously from the timer thread.
            UpdateLogger();

            // Also, schedule an update of the icon and tooltip if they changed, or the first time.
            if (AppTimerOperation == null || (AppTimerOperation.Status == DispatcherOperationStatus.Completed && GetIsIconOrToolTipChanged()))
                AppTimerOperation = Dispatcher.BeginInvoke(DispatcherPriority.Normal, new System.Action(OnAppTimer));
        }

        private void OnAppTimer()
        {
            // If a shutdown is started, don't update the taskbar anymore so the shutdown can complete smoothly.
            if (IsExiting)
                return;

            UpdateIconAndToolTip();
        }

        private void CleanupTimer()
        {
            using (Timer Timer = AppTimer)
            {
                AppTimer = null;
            }
        }

        private Timer AppTimer;
        private DispatcherOperation AppTimerOperation;
        private TimeSpan CheckInterval = TimeSpan.FromSeconds(0.1);
        #endregion

        #region Logger
        private static void InitLogger()
        {
            Logger = new PluginLogger();
        }

        private static void UpdateLogger()
        {
            Logger.PrintLog();
        }

        private static PluginLogger Logger;
        #endregion

        #region Load at startup
        private void InstallLoad(bool isInstalled, string appName)
        {
            string ExeName = Assembly.GetExecutingAssembly().Location;

            // Create or delete a task in the Task Scheduler.
            if (isInstalled)
            {
                TaskRunLevel RunLevel = PluginManager.RequireElevated ? TaskRunLevel.Highest : TaskRunLevel.LUA;
                Scheduler.AddTask(appName, ExeName, RunLevel, Logger);
            }
            else
                Scheduler.RemoveTask(ExeName, out bool IsFound); // Ignore it if the task was not found.
        }
        #endregion

        #region Implementation of IDisposable
        protected virtual void Dispose(bool isDisposing)
        {
            if (isDisposing)
                DisposeNow();
        }

        private void DisposeNow()
        {
            CleanupTimer();
            CleanupPlugInManager();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~App()
        {
            Dispose(false);
        }
        #endregion
    }
}
