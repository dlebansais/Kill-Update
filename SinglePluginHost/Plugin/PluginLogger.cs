using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Tracing;

namespace TaskbarIconHost
{
    public class PluginLogger : ITracer
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

                if (TraceFilePath != null && TraceFilePath.Length > 0)
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

        public void Write(Category category, string message, params object[] arguments)
        {
            if (arguments.Length > 0)
                AddLog(string.Format(CultureInfo.InvariantCulture, message, arguments));
            else
                AddLog(message);
        }

        public void Write(Category category, Exception exception, string message, params object[] arguments)
        {
            if (arguments.Length > 0)
                AddLog(string.Format(CultureInfo.InvariantCulture, message, arguments));
            else
                AddLog(message);

            if (exception != null)
                AddLog(exception.Message);
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
                    string TimeLog = UtcNow.ToString(CultureInfo.InvariantCulture) + UtcNow.Millisecond.ToString("D3", CultureInfo.InvariantCulture);

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
            NativeMethods.OutputDebugString(line);

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

        private string? LogLines;
        private object GlobalLock = string.Empty;
        private bool IsLogOn;
        private bool IsFileLogOn;
        private string? TraceFilePath;
    }
}
