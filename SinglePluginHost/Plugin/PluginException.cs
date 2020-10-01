namespace TaskbarIconHost
{
    using System;

#pragma warning disable CA1032 // Implement standard exception constructors
#pragma warning disable CA2237 // Mark ISerializable types with serializable
    public class PluginException : Exception
#pragma warning restore CA2237 // Mark ISerializable types with serializable
#pragma warning restore CA1032 // Implement standard exception constructors
    {
        public PluginException(string message)
            : base(message)
        {
        }
    }
}
