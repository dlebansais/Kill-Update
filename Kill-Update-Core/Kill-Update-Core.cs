namespace KillUpdate
{
    using RegistryTools;
    using System;
    using System.Diagnostics;
    using System.Drawing;
    using System.IO;
    using System.Reflection;
    using System.ServiceProcess;
    using System.Threading;
    using System.Windows.Threading;
    using Tracing;

    public class KillUpdateCore : IDisposable
    {
        public KillUpdateCore(bool isElevated, Settings settings, ITracer logger)
        {
            IsElevated = isElevated;
            Settings = settings;
            Logger = logger;
        }

        public bool IsElevated { get; }
        public Settings Settings { get; }
        public ITracer Logger { get; }
        public Dispatcher Dispatcher { get; private set; } = null!;

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

            if (LoadEmbeddedResource(ResourceName, out Icon Result))
                Logger.Write(Category.Debug, $"Resource {ResourceName} loaded");
            else
                Logger.Write(Category.Debug, $"Resource {ResourceName} not found");

            return Result;
        }

        public bool LoadEmbeddedResource<T>(string resourceName, out T resource)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string ResourcePath = string.Empty;

            // Loads an "Embedded Resource" of type T (ex: Bitmap for a PNG file).
            // Make sure the resource is tagged as such in the resource properties.
            foreach (string Item in assembly.GetManifestResourceNames())
                if (Item.EndsWith(resourceName, StringComparison.InvariantCulture))
                    ResourcePath = Item;

            // If not found, it could be because it's not tagged as "Embedded Resource".
            if (ResourcePath.Length == 0)
            {
                resource = default!;
                return false;
            }

            using Stream rs = assembly.GetManifestResourceStream(ResourcePath);
            resource = (T)Activator.CreateInstance(typeof(T), rs);

            return true;
        }
        #endregion

        #region Tooltip
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
        #endregion

        #region Zombification
        public bool IsRestart { get { return ZombifyMe.Zombification.IsRestart; } }

        public void InitZombification()
        {
            Logger.Write(Category.Debug, "InitZombification starting");

            if (IsRestart)
                Logger.Write(Category.Information, "This process has been restarted");

            Zombification = new ZombifyMe.Zombification("Kill-Update");
            Zombification.Delay = TimeSpan.FromMinutes(1);
            Zombification.WatchingMessage = string.Empty;
            Zombification.RestartMessage = string.Empty;
            Zombification.Flags = ZombifyMe.Flags.NoWindow | ZombifyMe.Flags.ForwardArguments;
            Zombification.IsSymmetric = true;
            Zombification.AliveTimeout = TimeSpan.FromMinutes(1);
            Zombification.ZombifyMe();

            Logger.Write(Category.Debug, "InitZombification done");
        }

        public void ExitZombification()
        {
            Logger.Write(Category.Debug, "ExitZombification starting");

            Zombification.Cancel();

            Logger.Write(Category.Debug, "ExitZombification done");
        }

        private ZombifyMe.Zombification Zombification = null!;
        #endregion

        #region Service Manager
        public void InitServiceManager()
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

        public bool IsLockEnabled
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

        public void ToggleLock()
        {
            bool LockIt = !IsLockEnabled;

            Settings.SetBool(LockedSettingName, LockIt);
            ChangeLockMode(LockIt);
        }

        public void ChangeLockMode(bool lockIt)
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

        private void StopIfRunning(ServiceController Service, bool lockIt)
        {
            if (lockIt && Service.Status == ServiceControllerStatus.Running && Service.CanStop)
            {
                Logger.Write(Category.Debug, "Stopping service");
                Service.Stop();
                Logger.Write(Category.Debug, "Service stopped");
            }
        }

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

        private const string WindowsUpdateServiceName = "wuauserv";
        private const string LockedSettingName = "Locked";
        private readonly TimeSpan CheckInterval = TimeSpan.FromSeconds(15);
        private readonly TimeSpan FullRestartInterval = TimeSpan.FromHours(1);
        public ServiceStartMode? StartType;
        public bool IsLockModeChanged { get; set; }
        #endregion

        #region Timer
        public void InitTimer(Dispatcher dispatcher)
        {
            Dispatcher = dispatcher;
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

            Dispatcher.BeginInvoke(new Action(OnUpdate));
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

        public void StopTimer()
        {
            FullRestartTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
            UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        }

        private Timer UpdateTimer = null!;
        private Timer FullRestartTimer = null!;
        private Stopwatch UpdateWatch = new Stopwatch();
        private int TimerDispatcherCount = 1;
        private double LastTotalElapsed = double.NaN;
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
        /// Finalizes an instance of the <see cref="KillUpdateCore"/> class.
        /// </summary>
        ~KillUpdateCore()
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
