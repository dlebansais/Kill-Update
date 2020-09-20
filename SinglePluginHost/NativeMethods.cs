namespace TaskbarIconHost
{
    using System.Runtime.InteropServices;

    public static class NativeMethods
    {
        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        internal static extern void OutputDebugString([In][MarshalAs(UnmanagedType.LPWStr)] string message);
    }
}
