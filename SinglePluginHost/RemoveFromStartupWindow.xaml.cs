using System.Diagnostics;
using System.Windows;
using System.Windows.Input;

namespace TaskbarIconHost
{
    public partial class RemoveFromStartupWindow : Window
    {
        #region Init
        public RemoveFromStartupWindow(string appName)
        {
            InitializeComponent();
            DataContext = this;

            Title = appName;
            TaskSelectiontext = $"From the Task Scheduler, select 'Task Scheduler Library' and search for the task called '{appName}'.";
        }

        public string TaskSelectiontext { get; private set; }
        #endregion

        #region Events
        private void OnLaunch(object sender, ExecutedRoutedEventArgs e)
        {
            // Launch the Windows Task Scheduler.
            Process ControlProcess = new Process();
            ControlProcess.StartInfo.FileName = "control.exe";
            ControlProcess.StartInfo.Arguments = "schedtasks";
            ControlProcess.StartInfo.UseShellExecute = true;

            ControlProcess.Start();
        }

        private void OnClose(object sender, ExecutedRoutedEventArgs e)
        {
            Close();
        }
        #endregion
    }
}
