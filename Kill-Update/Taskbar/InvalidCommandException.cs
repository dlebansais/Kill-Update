namespace TaskbarTools
{
    using System;
    using System.Windows.Input;

#pragma warning disable CA1032 // Implement standard exception constructors
#pragma warning disable CA2237 // Mark ISerializable types with SerializableAttribute
    public class InvalidCommandException : Exception
    {
        public InvalidCommandException(ICommand command) { Command = command; }
        public ICommand Command { get; }
    }
#pragma warning restore CA2237 // Mark ISerializable types with SerializableAttribute
#pragma warning restore CA1032 // Implement standard exception constructors
}
