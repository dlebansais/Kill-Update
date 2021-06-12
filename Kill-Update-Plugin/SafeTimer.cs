namespace KillUpdate
{
    using System;
    using System.Diagnostics;
    using System.Threading;
    using System.Windows.Threading;
    using Tracing;

    /// <summary>
    /// Execute an action regularly at a given time interval, with safeguards.
    /// </summary>
    public class SafeTimer : IDisposable
    {
        #region Init
        /// <summary>
        /// Creates a new instance of the <see cref="SafeTimer"/> class.
        /// </summary>
        /// <param name="action">The action to execute.</param>
        /// <param name="timeInterval">The time interval.</param>
        /// <param name="logger">an interface to log events asynchronously.</param>
        public static SafeTimer Create(Action action, TimeSpan timeInterval, ITracer logger)
        {
            return new SafeTimer(action, timeInterval, logger);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SafeTimer"/> class.
        /// </summary>
        /// <param name="action">The action to execute.</param>
        /// <param name="timeInterval">The time interval.</param>
        /// <param name="logger">an interface to log events asynchronously.</param>
        protected SafeTimer(Action action, TimeSpan timeInterval, ITracer logger)
        {
            Action = action;
            TimeInterval = timeInterval;
            Logger = logger;

            UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
            FullRestartTimer = new Timer(new TimerCallback(FullRestartTimerCallback));
            UpdateWatch = new Stopwatch();
            UpdateWatch.Start();

            Action();

            UpdateTimer.Change(TimeInterval, TimeInterval);
            FullRestartTimer.Change(FullRestartInterval, Timeout.InfiniteTimeSpan);
        }
        #endregion

        #region Cleanup
        /// <summary>
        /// Destroys an instance of the <see cref="SafeTimer"/> class.
        /// </summary>
        /// <param name="instance">The instance to destroy.</param>
        public static void Destroy(ref SafeTimer? instance)
        {
            if (instance != null)
            {
                using SafeTimer DisposedObject = instance;
                instance = null;
            }
        }
        #endregion

        #region Properties
        /// <summary>
        /// Gets the action to execute.
        /// </summary>
        public Action Action { get; }

        /// <summary>
        /// Gets the time interval.
        /// </summary>
        public TimeSpan TimeInterval { get; }

        /// <summary>
        /// Gets an interface to log events asynchronously.
        /// </summary>
        public ITracer Logger { get; }
        #endregion

        #region Client Interface
        /// <summary>
        /// To call from the action callback.
        /// </summary>
        public void NotifyCallbackCalled()
        {
            int LastTimerDispatcherCount = Interlocked.Decrement(ref TimerDispatcherCount);
            UpdateWatch.Restart();

            AddLog($"Watch restarted, Elapsed = {LastTotalElapsed}, pending count = {LastTimerDispatcherCount}");
        }
        #endregion

        #region Implementation
        private void UpdateTimerCallback(object? parameter)
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

            System.Windows.Application.Current.Dispatcher.BeginInvoke(Action);
        }

        private void FullRestartTimerCallback(object? parameter)
        {
            if (UpdateTimer != null)
            {
                AddLog("Restarting the timer");

                // Restart the update timer from scratch.
                UpdateTimer.Change(Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);

                UpdateTimer = new Timer(new TimerCallback(UpdateTimerCallback));
                UpdateTimer.Change(TimeInterval, TimeInterval);

                AddLog("Timer restarted");
            }
            else
                AddLog("No timer to restart");

            FullRestartTimer?.Change(FullRestartInterval, Timeout.InfiniteTimeSpan);
            AddLog($"Next check scheduled at {DateTime.UtcNow + FullRestartInterval}");
        }

        private void AddLog(string message)
        {
            Logger.Write(Category.Information, message);
        }

        private readonly TimeSpan FullRestartInterval = TimeSpan.FromHours(1);
        private Timer UpdateTimer;
        private Timer FullRestartTimer;
        private Stopwatch UpdateWatch;
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
        /// Finalizes an instance of the <see cref="SafeTimer"/> class.
        /// </summary>
        ~SafeTimer()
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
            using (FullRestartTimer)
            {
            }

            using (UpdateTimer)
            {
            }
        }
        #endregion
    }
}
