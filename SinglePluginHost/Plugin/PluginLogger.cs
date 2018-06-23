using System;
using System.Globalization;
using System.Runtime.InteropServices;

namespace TaskbarIconHost
{
    public class PluginLogger : MarshalByRefObject, IPluginLogger
    {
        public void AddLog(string text)
        {
            AddLog(text, false);
        }

        public void AddLog(string text, bool showNow)
        {
#if DEBUG
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
#endif
            if (showNow)
                PrintLog();
        }

        public void PrintLog()
        {
#if DEBUG
            lock (GlobalLock)
            {
                if (LogLines != null)
                {
                    string[] Lines = LogLines.Split('\n');
                    foreach (string Line in Lines)
                        OutputDebugString(Line);

                    LogLines = null;
                }
            }
#endif
        }

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        public static extern void OutputDebugString([In][MarshalAs(UnmanagedType.LPWStr)] string message);

#if DEBUG
        private string LogLines = null;
#endif
        private object GlobalLock = "";
    }
}
