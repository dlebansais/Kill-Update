using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace TaskbarIconHost
{
    public partial class RemoveFromStartupWindow : Window
    {
        #region Init
        public RemoveFromStartupWindow()
        {
            InitializeComponent();
            DataContext = this;
        }
        #endregion

        #region Events
        private void OnLaunch(object sender, ExecutedRoutedEventArgs e)
        {
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
