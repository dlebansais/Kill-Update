using System;

namespace TaskbarIconHost
{
    public class PluginException : Exception
    {
        public PluginException(string message)
            : base(message)
        {
        }
    }
}
