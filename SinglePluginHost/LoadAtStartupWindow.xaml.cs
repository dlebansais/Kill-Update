using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows;
using System.Windows.Input;

namespace TaskbarIconHost
{
    public partial class LoadAtStartupWindow : Window
    {
        #region Init
        public LoadAtStartupWindow(bool requireElevated)
        {
            InitializeComponent();
            DataContext = this;

            RequireElevated = requireElevated;

            try
            {
                string ApplicationFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TaskbarIconHost");

                if (!Directory.Exists(ApplicationFolder))
                    Directory.CreateDirectory(ApplicationFolder);

                TaskFile = Path.Combine(ApplicationFolder, "TaskbarIconHost.xml");

                if (!File.Exists(TaskFile))
                {
                    Assembly ExecutingAssembly = Assembly.GetExecutingAssembly();

                    foreach (string ResourceName in ExecutingAssembly.GetManifestResourceNames())
                        if (ResourceName.EndsWith("TaskbarIconHost.xml"))
                        {
                            using (Stream rs = ExecutingAssembly.GetManifestResourceStream(ResourceName))
                            {
                                using (FileStream fs = new FileStream(TaskFile, FileMode.Create, FileAccess.Write, FileShare.None))
                                {
                                    using (StreamReader sr = new StreamReader(rs))
                                    {
                                        using (StreamWriter sw = new StreamWriter(fs))
                                        {
                                            string Content = sr.ReadToEnd();
                                            string Location = Path.GetDirectoryName(ExecutingAssembly.Location);
                                            Content = Content.Replace("%PATH%", Location);
                                            sw.WriteLine(Content);
                                        }
                                    }
                                }
                            }

                            break;
                        }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        public bool RequireElevated { get; private set; }

        private string TaskFile;
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

        private void OnCopy(object sender, ExecutedRoutedEventArgs e)
        {
            Clipboard.SetText(TaskFile);
        }

        private void OnClose(object sender, ExecutedRoutedEventArgs e)
        {
            Close();
        }
        #endregion
    }
}
