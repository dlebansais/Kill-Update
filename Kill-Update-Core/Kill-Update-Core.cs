using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using Tracing;

namespace KillUpdate
{
    public class KillUpdateCore
    {
        public delegate void LoggerHandler(Category category, string message, params object[] arguments);

        public static LoggerHandler Logger { private get; set; } = null !;

        #region Icon & Bitmap
        public static Icon LoadCurrentIcon(bool isElevated, bool isLockEnabled)
        {
            string ResourceName;

            if (isElevated)
                if (isLockEnabled)
                    ResourceName = "Locked-Enabled.ico";
                else
                    ResourceName = "Unlocked-Enabled.ico";
            else
                if (isLockEnabled)
                    ResourceName = "Locked-Disabled.ico";
                else
                    ResourceName = "Unlocked-Disabled.ico";

            if (LoadEmbeddedResource(ResourceName, out Icon Result))
                Logger(Category.Debug, $"Resource {ResourceName} loaded");
            else
                Logger(Category.Debug, $"Resource {ResourceName} not found");

            return Result;
        }

        public static bool LoadEmbeddedResource<T>(string resourceName, out T resource)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string ResourcePath = string.Empty;

            // Loads an "Embedded Resource" of type T (ex: Bitmap for a PNG file).
            // Make sure the resource is tagged as such in the resource properties.
            foreach (string Item in assembly.GetManifestResourceNames())
                if (Item.EndsWith(resourceName, StringComparison.InvariantCulture))
                    ResourcePath = Item;

            // If not found, it could be because it's not tagged as "Embedded Resource".
            if (ResourcePath.Length == 0)
            {
                resource = default!;
                return false;
            }

            using Stream rs = assembly.GetManifestResourceStream(ResourcePath);
            resource = (T)Activator.CreateInstance(typeof(T), rs);

            return true;
        }
        #endregion

        #region Zombification
        public static bool IsRestart { get { return ZombifyMe.Zombification.IsRestart; } }

        public static void InitZombification()
        {
            Logger(Category.Debug, "InitZombification starting");

            if (IsRestart)
                Logger(Category.Information, "This process has been restarted");

            Zombification = new ZombifyMe.Zombification("Kill-Update");
            Zombification.Delay = TimeSpan.FromMinutes(1);
            Zombification.WatchingMessage = string.Empty;
            Zombification.RestartMessage = string.Empty;
            Zombification.Flags = ZombifyMe.Flags.NoWindow | ZombifyMe.Flags.ForwardArguments;
            Zombification.IsSymmetric = true;
            Zombification.AliveTimeout = TimeSpan.FromMinutes(1);
            Zombification.ZombifyMe();

            Logger(Category.Debug, "InitZombification done");
        }

        public static void ExitZombification()
        {
            Logger(Category.Debug, "ExitZombification starting");

            Zombification.Cancel();

            Logger(Category.Debug, "ExitZombification done");
        }

        private static ZombifyMe.Zombification Zombification = null!;
        #endregion
    }
}
