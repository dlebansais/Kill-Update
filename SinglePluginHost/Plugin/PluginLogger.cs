using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace TaskbarIconHost
{
    public class PluginLogger : IPluginLogger
    {
        public PluginLogger()
        {
#if DEBUG
            IsLogOn = true;
#endif

            try
            {
                string Location = Assembly.GetExecutingAssembly().Location;
                string SettingFilePath = Path.Combine(Path.GetDirectoryName(Location), "settings.txt");

                if (File.Exists(SettingFilePath))
                {
                    using (FileStream fs = new FileStream(SettingFilePath, FileMode.Open, FileAccess.Read))
                    {
                        using (StreamReader sr = new StreamReader(fs))
                        {
                            TraceFilePath = sr.ReadLine();
                        }
                    }
                }

                if (!string.IsNullOrEmpty(TraceFilePath))
                {
                    bool IsFirstTraceWritten = false;

                    using (FileStream fs = new FileStream(TraceFilePath, FileMode.Append, FileAccess.Write))
                    {
                        using (StreamWriter sw = new StreamWriter(fs))
                        {
                            sw.WriteLine("** Log started **");
                            IsFirstTraceWritten = true;
                        }
                    }

                    if (IsFirstTraceWritten)
                    {
                        IsLogOn = true;
                        IsFileLogOn = true;
                    }
                }
            }
            catch (Exception e)
            {
                PrintLine("Unable to start logging traces.");
                PrintLine(e.Message);
            }
        }

        public void AddLog(string text)
        {
            AddLog(text, false);
        }

        public void AddLog(string text, bool showNow)
        {
            if (IsLogOn)
            {
                lock (GlobalLock)
                {
                    DateTime UtcNow = DateTime.UtcNow;
                    string TimeLog = UtcNow.ToString(CultureInfo.InvariantCulture) + UtcNow.Millisecond.ToString("D3");

                    string Line = $"TaskbarIconHost - {TimeLog}: {text}\n";

                    if (LogLines == null)
                        LogLines = Line;
                    else
                        LogLines += Line;
                }
            }

            if (showNow)
                PrintLog();
        }

        public void PrintLog()
        {
            if (IsLogOn)
            {
                lock (GlobalLock)
                {
                    if (LogLines != null)
                    {
                        string[] Lines = LogLines.Split('\n');
                        foreach (string Line in Lines)
                            PrintLine(Line);

                        LogLines = null;
                    }
                }
            }
        }

        private void PrintLine(string line)
        {
            OutputDebugString(line);

            if (IsFileLogOn)
                WriteLineToTraceFile(line);
        }

        private void WriteLineToTraceFile(string line)
        {
            try
            {
                if (line.Length == 0)
                    return;

                using (FileStream fs = new FileStream(TraceFilePath, FileMode.Append, FileAccess.Write))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        sw.WriteLine(line);
                    }
                }
            }
            catch
            {
            }
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        public static extern void OutputDebugString([In][MarshalAs(UnmanagedType.LPWStr)] string message);

        private string LogLines = null;
        private object GlobalLock = "";
        private bool IsLogOn;
        private bool IsFileLogOn;
        private string TraceFilePath;
    }
}
