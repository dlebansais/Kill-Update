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
                    GuidAttribute AppGuidAttribute = Assembly.GetExecutingAssembly().GetCustomAttribute<GuidAttribute>();
                    AppGuid = Guid.Parse(AppGuidAttribute.Value);
                }

                string AppUniqueId = AppGuid.ToString("B").ToUpper();

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
            if (Taskbar.ScreenBounds.IsEmpty)
                Shutdown();
            else
            {
                Startup += OnStartup;
                ShutdownMode = ShutdownMode.OnExplicitShutdown;
            }
        }

        private void OnStartup(object sender, StartupEventArgs e)
        {
            Logger.AddLog("OnStartup");

            InitTimer();

            if (InitPlugInManager())
            {
                InitTaskbarIcon();

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

        private void OnActivated(object sender, EventArgs e)
        {
            PluginManager.OnActivated();
        }

        private void OnDeactivated(object sender, EventArgs e)
        {
            PluginManager.OnDeactivated();
        }

        private void OnExit(object sender, ExitEventArgs e)
        {
            Logger.AddLog("Exiting application");

            IsExiting = true;
            StopPlugInManager();
            CleanupTaskbarIcon();
            CleanupTimer();

            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

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

            GlobalSettings = new PluginSettings(null, Logger);

            try
            {
                PluginManager.PreferredPluginGuid = new Guid(GlobalSettings.GetSettingString(PreferredPluginSettingName, PluginManager.GuidToString(Guid.Empty)));
            }
            catch
            {
            }

            return true;
        }

        private void StopPlugInManager()
        {
            try
            {
                GlobalSettings.SetSettingString(PreferredPluginSettingName, PluginManager.GuidToString(PluginManager.PreferredPluginGuid));
            }
            catch
            {
            }

            PluginManager.Shutdown();
        }

        private static readonly string PreferredPluginSettingName = "PreferredPlugin";
        private PluginSettings GlobalSettings;
        #endregion

        #region Taskbar Icon
        private void InitTaskbarIcon()
        {
            Logger.AddLog("InitTaskbarIcon starting");

            LoadAtStartupCommand = new RoutedUICommand();
            AddMenuCommand(LoadAtStartupCommand, OnCommandLoadAtStartup);

            ExitCommand = new RoutedUICommand();
            AddMenuCommand(ExitCommand, OnCommandExit);

            foreach (KeyValuePair<List<ICommand>, string> Entry in PluginManager.FullCommandList)
            {
                List<ICommand> FullPluginCommandList = Entry.Key;
                foreach (ICommand Command in FullPluginCommandList)
                    AddMenuCommand(Command, OnPluginCommand);
            }

            Icon Icon = PluginManager.Icon;
            string ToolTip = PluginManager.ToolTip;
            ContextMenu ContextMenu = LoadContextMenu();

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

        private void AddMenuCommand(ICommand Command, ExecutedRoutedEventHandler executed)
        {
            if (Command == null)
                return;

            // Bind the command to the corresponding handler. Requires the menu to be the target of notifications in TaskbarIcon.
            CommandManager.RegisterClassCommandBinding(typeof(ContextMenu), new CommandBinding(Command, executed));
        }

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

        private ContextMenu LoadContextMenu()
        {
            ContextMenu Result = new ContextMenu();
            ItemCollection Items = Result.Items;

            MenuItem LoadAtStartup;
            string ExeName = Assembly.GetExecutingAssembly().Location;

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

            Dictionary<List<MenuItem>, string> FullPluginMenuList = new Dictionary<List<MenuItem>, string>();
            int VisiblePluginMenuCount = 0;

            foreach (KeyValuePair<List<ICommand>, string> Entry in PluginManager.FullCommandList)
            {
                List<ICommand> FullPluginCommandList = Entry.Key;
                string PluginName = Entry.Value;

                List<MenuItem> PluginMenuList = new List<MenuItem>();

                foreach (ICommand Command in FullPluginCommandList)
                    if (Command == null)
                        PluginMenuList.Add(null);
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

                        if (MenuIsVisible)
                            VisiblePluginMenuCount++;
                    }

                FullPluginMenuList.Add(PluginMenuList, PluginName);
            }

            bool AddSeparator = VisiblePluginMenuCount > 1;

            foreach (KeyValuePair<List<MenuItem>, string> Entry in FullPluginMenuList)
            {
                List<MenuItem> PluginMenuList = Entry.Key;
                string PluginName = Entry.Value;

                if (AddSeparator)
                    Items.Add(new Separator());

                AddSeparator = true;

                ItemCollection SubmenuItems;
                if (PluginMenuList.Count > 1 && FullPluginMenuList.Count > 1)
                {
                    MenuItem PluginSubmenu = new MenuItem();
                    PluginSubmenu.Header = PluginName;
                    SubmenuItems = PluginSubmenu.Items;
                    Items.Add(PluginSubmenu);
                }
                else
                    SubmenuItems = Items;

                foreach (MenuItem MenuItem in PluginMenuList)
                    if (MenuItem != null)
                        SubmenuItems.Add(MenuItem);
                    else
                        SubmenuItems.Add(new Separator());
            }

            Items.Add(new Separator());

            if (PluginManager.ConsolidatedPluginList.Count > 1)
            {
                MenuItem IconSubmenu = new MenuItem();
                IconSubmenu.Header = "Icons";

                foreach (IPluginClient Plugin in PluginManager.ConsolidatedPluginList)
                {
                    Guid SubmenuGuid = Plugin.Guid;
                    if (!IconSelectionTable.ContainsKey(SubmenuGuid))
                    {
                        RoutedUICommand SubmenuCommand = new RoutedUICommand();
                        SubmenuCommand.Text = PluginManager.GuidToString(SubmenuGuid);

                        string SubmenuHeader = Plugin.Name;
                        bool SubmenuIsChecked = (SubmenuGuid == PluginManager.PreferredPluginGuid);
                        Bitmap SubmenuIcon = Plugin.SelectionBitmap;

                        AddMenuCommand(SubmenuCommand, OnCommandSelectPreferred);
                        MenuItem PluginMenu = CreateMenuItem(SubmenuCommand, SubmenuHeader, SubmenuIsChecked, SubmenuIcon);
                        IconSubmenu.Items.Add(PluginMenu);

                        IconSelectionTable.Add(SubmenuGuid, SubmenuCommand);
                    }
                }

                Items.Add(IconSubmenu);
                Items.Add(new Separator());
            }

            MenuItem ExitMenu = CreateMenuItem(ExitCommand, "Exit", false, null);
            Items.Add(ExitMenu);

            Logger.AddLog("Menu created");

            return Result;
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

            if (IsElevated)
                SenderIcon.SetMenuIsChecked(LoadAtStartupCommand, Scheduler.IsTaskActive(ExeName));
            else
            {
                if (Scheduler.IsTaskActive(ExeName))
                    SenderIcon.SetMenuHeader(LoadAtStartupCommand, RemoveFromStartupHeader);
                else
                    SenderIcon.SetMenuHeader(LoadAtStartupCommand, LoadAtStartupHeader);
            }

            UpdateMenu();

            PluginManager.OnMenuOpening();
        }

        private void UpdateMenu()
        {
            foreach (ICommand Command in PluginManager.GetChangedCommands())
            {
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
                TaskbarIcon.ToggleMenuIsChecked(LoadAtStartupCommand, out bool Install);
                InstallLoad(Install, PluginDetails.Name);
            }
            else
            {
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

            if (GetIsIconOrToolTipChanged())
                UpdateIconAndToolTip();

            UpdateMenu();
        }

        private void OnCommandSelectPreferred(object sender, ExecutedRoutedEventArgs e)
        {
            Logger.AddLog("OnCommandSelectPreferred");

            if (e.Command is RoutedUICommand SubmenuCommand)
            {
                Guid NewSelectedGuid = new Guid(SubmenuCommand.Text);
                Guid OldSelectedGuid = PluginManager.PreferredPluginGuid;
                if (NewSelectedGuid != OldSelectedGuid)
                {
                    PluginManager.PreferredPluginGuid = NewSelectedGuid;

                    IsIconChanged = true;
                    IsToolTipChanged = true;
                    UpdateIconAndToolTip();

                    if (IconSelectionTable.ContainsKey(NewSelectedGuid))
                        TaskbarIcon.SetMenuIsChecked(IconSelectionTable[NewSelectedGuid], true);

                    if (IconSelectionTable.ContainsKey(OldSelectedGuid))
                        TaskbarIcon.SetMenuIsChecked(IconSelectionTable[OldSelectedGuid], false);
                }
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
            AppTimer = new Timer(new TimerCallback(AppTimerCallback));
            AppTimer.Change(CheckInterval, CheckInterval);
        }

        private void AppTimerCallback(object parameter)
        {
            if (IsExiting)
                return;

            UpdateLogger();

            if (AppTimerOperation == null || (AppTimerOperation.Status == DispatcherOperationStatus.Completed && GetIsIconOrToolTipChanged()))
                AppTimerOperation = Dispatcher.BeginInvoke(DispatcherPriority.Normal, new System.Action(OnAppTimer));
        }

        private void OnAppTimer()
        {
            if (IsExiting)
                return;

            UpdateIconAndToolTip();
        }

        private void CleanupTimer()
        {
            if (AppTimer != null)
            {
                AppTimer.Dispose();
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

            if (isInstalled)
            {
                TaskRunLevel RunLevel = PluginManager.RequireElevated ? TaskRunLevel.Highest : TaskRunLevel.LUA;
                Scheduler.AddTask(appName, ExeName, RunLevel, Logger);
            }
            else
                Scheduler.RemoveTask(ExeName, out bool IsFound);
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
