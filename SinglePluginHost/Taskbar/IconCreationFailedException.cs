namespace TaskbarTools
{
    using System;

#pragma warning disable CA1032 // Implement standard exception constructors
#pragma warning disable CA2237 // Mark ISerializable types with SerializableAttribute
    public class IconCreationFailedException : Exception
    {
        public IconCreationFailedException(Exception originalException) { OriginalException = originalException; }
        public Exception OriginalException { get; }
    }
#pragma warning restore CA2237 // Mark ISerializable types with SerializableAttribute
#pragma warning restore CA1032 // Implement standard exception constructors
}
