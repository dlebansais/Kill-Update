using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Windows.Input;
using System.Windows.Threading;
using System.Security.Cryptography.X509Certificates;

namespace TaskbarIconHost
{
    public static class PluginManager
    {
        public static bool Init(bool isElevated, string embeddedPluginName, Guid embeddedPluginGuid, Dispatcher dispatcher, IPluginLogger logger)
        {
            PluginInterfaceType = typeof(IPluginClient);

            Assembly CurrentAssembly = Assembly.GetExecutingAssembly();
            string Location = CurrentAssembly.Location;
            string AppFolder = Path.GetDirectoryName(Location);
            int AssemblyCount = 0;
            int CompatibleAssemblyCount = 0;

            Dictionary<Assembly, List<Type>> PluginClientTypeTable = new Dictionary<Assembly, List<Type>>();
            Assembly PluginAssembly;
            List<Type> PluginClientTypeList;

            if (embeddedPluginName != null)
            {
                AssemblyName[] AssemblyNames = CurrentAssembly.GetReferencedAssemblies();
                foreach (AssemblyName name in AssemblyNames)
                    if (name.Name == embeddedPluginName)
                    {
                        FindPluginClientTypesByName(name, out PluginAssembly, out PluginClientTypeList);
                        if (PluginAssembly != null && PluginClientTypeList != null && PluginClientTypeList.Count > 0)
                            PluginClientTypeTable.Add(PluginAssembly, PluginClientTypeList);
                    }
            }

            string[] Assemblies = Directory.GetFiles(AppFolder, "*.dll");
            foreach (string AssemblyPath in Assemblies)
            {
                FindPluginClientTypesByPath(AssemblyPath, out PluginAssembly, out PluginClientTypeList);
                if (PluginAssembly != null && PluginClientTypeList != null)
                    PluginClientTypeTable.Add(PluginAssembly, PluginClientTypeList);
            }

            foreach (KeyValuePair<Assembly, List<Type>> Entry in PluginClientTypeTable)
            {
                AssemblyCount++;

                PluginAssembly = Entry.Key;
                PluginClientTypeList = Entry.Value;

                if (PluginClientTypeList.Count > 0)
                {
                    CompatibleAssemblyCount++;

                    CreatePluginList(PluginAssembly, PluginClientTypeList, embeddedPluginGuid, logger, out List<IPluginClient> PluginList);
                    if (PluginList.Count > 0)
                        LoadedPluginTable.Add(PluginAssembly, PluginList);
                }
            }

            if (LoadedPluginTable.Count > 0)
            {
                foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                    foreach (IPluginClient Plugin in Entry.Value)
                    {
                        IPluginSettings Settings = new PluginSettings(GuidToString(Plugin.Guid), logger);
                        Plugin.Initialize(isElevated, dispatcher, Settings, logger);

                        if (Plugin.RequireElevated)
                            RequireElevated = true;
                    }

                foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                    foreach (IPluginClient Plugin in Entry.Value)
                    {
                        List<ICommand> PluginCommandList = Plugin.CommandList;
                        if (PluginCommandList != null)
                        {
                            List<ICommand> FullPluginCommandList = new List<ICommand>();
                            FullCommandList.Add(FullPluginCommandList, Plugin.Name);

                            foreach (ICommand Command in PluginCommandList)
                            {
                                FullPluginCommandList.Add(Command);

                                if (Command != null)
                                    CommandTable.Add(Command, Plugin);
                            }
                        }

                        Icon PluginIcon = Plugin.Icon;
                        if (PluginIcon != null)
                            ConsolidatedPluginList.Add(Plugin);
                    }

                foreach (IPluginClient Plugin in ConsolidatedPluginList)
                    if (Plugin.HasClickHandler)
                    {
                        PreferredPlugin = Plugin;
                        break;
                    }

                if (PreferredPlugin == null && ConsolidatedPluginList.Count > 0)
                    PreferredPlugin = ConsolidatedPluginList[0];

                return true;
            }
            else
            {
                logger.AddLog($"Could not load plugins, {AssemblyCount} assemblies found, {CompatibleAssemblyCount} are compatible.");
                return false;
            }
        }

        private static bool IsReferencingSharedAssembly(Assembly assembly, out AssemblyName SharedAssemblyName)
        {
            AssemblyName[] AssemblyNames = assembly.GetReferencedAssemblies();
            foreach (AssemblyName AssemblyName in AssemblyNames)
                if (AssemblyName.Name == SharedPluginAssemblyName)
                {
                    SharedAssemblyName = AssemblyName;
                    return true;
                }

            SharedAssemblyName = null;
            return false;
        }

        private static void FindPluginClientTypesByPath(string assemblyPath, out Assembly PluginAssembly, out List<Type> PluginClientTypeList)
        {
            try
            {
                PluginAssembly = Assembly.LoadFrom(assemblyPath);
                FindPluginClientTypes(PluginAssembly, out PluginClientTypeList);
            }
            catch
            {
                PluginAssembly = null;
                PluginClientTypeList = null;
            }
        }

        private static void FindPluginClientTypesByName(AssemblyName name, out Assembly PluginAssembly, out List<Type> PluginClientTypeList)
        {
            try
            {
                PluginAssembly = Assembly.Load(name);
                FindPluginClientTypes(PluginAssembly, out PluginClientTypeList);
            }
            catch
            {
                PluginAssembly = null;
                PluginClientTypeList = null;
            }
        }

        private static void FindPluginClientTypes(Assembly assembly, out List<Type> PluginClientTypeList)
        {
            PluginClientTypeList = null;

            try
            {
#if !DEBUG
                if (!string.IsNullOrEmpty(assembly.Location) && !IsAssemblySigned(assembly))
                    return;
#endif
                PluginClientTypeList = new List<Type>();

                if (IsReferencingSharedAssembly(assembly, out AssemblyName SharedAssemblyName))
                {
                    Type[] AssemblyTypes;
                    try
                    {
                        AssemblyTypes = assembly.GetTypes();
                    }
                    catch (ReflectionTypeLoadException LoaderException)
                    {
                        AssemblyTypes = LoaderException.Types;
                    }
                    catch
                    {
                        AssemblyTypes = new Type[0];
                    }

                    foreach (Type ClientType in AssemblyTypes)
                    {
                        if (!ClientType.IsPublic || ClientType.IsInterface || !ClientType.IsClass || ClientType.IsAbstract)
                            continue;

                        Type InterfaceType = ClientType.GetInterface(PluginInterfaceType.FullName);
                        if (InterfaceType != null)
                            PluginClientTypeList.Add(ClientType);
                    }
                }
            }
            catch
            {
            }
        }

        private static bool IsAssemblySigned(Assembly assembly)
        {
            foreach (Module Module in assembly.GetModules())
                return IsModuleSigned(Module);

            return false;
        }

        private static bool IsModuleSigned(Module module)
        {
            for (int i = 0; i < 3; i++)
                if (IsModuleSignedOneTry(module))
                    return true;
                else if (!IsWakeUpDelayElapsed)
                {
                    Thread.Sleep(5000);
                    IsWakeUpDelayElapsed = true;
                }
                else
                    return false;

            return false;
        }

        private static bool IsModuleSignedOneTry(Module module)
        {
            try
            {
                X509Certificate certificate = module.GetSignerCertificate();
                if (certificate == null)
                {
                    // File is not signed.
                    return false;
                }

                X509Certificate2 certificate2 = new X509Certificate2(certificate);

                using (X509Chain CertificateChain = X509Chain.Create())
                {
                    CertificateChain.ChainPolicy.RevocationFlag = X509RevocationFlag.EndCertificateOnly;
                    CertificateChain.ChainPolicy.UrlRetrievalTimeout = TimeSpan.FromSeconds(60);
                    CertificateChain.ChainPolicy.RevocationMode = X509RevocationMode.Online;
                    bool IsEndCertificateValid = CertificateChain.Build(certificate2);

                    if (!IsEndCertificateValid)
                        return false;

                    CertificateChain.ChainPolicy.RevocationFlag = X509RevocationFlag.EntireChain;
                    CertificateChain.ChainPolicy.UrlRetrievalTimeout = TimeSpan.FromSeconds(60);
                    CertificateChain.ChainPolicy.RevocationMode = X509RevocationMode.Online;
                    CertificateChain.ChainPolicy.VerificationFlags = X509VerificationFlags.IgnoreCertificateAuthorityRevocationUnknown;
                    bool IsCertificateChainValid = CertificateChain.Build(certificate2);

                    if (!IsCertificateChainValid)
                        return false;

                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static void CreatePluginList(Assembly pluginAssembly, List<Type> PluginClientTypeList, Guid embeddedPluginGuid, IPluginLogger logger, out List<IPluginClient> PluginList)
        {
            PluginList = new List<IPluginClient>();

            foreach (Type ClientType in PluginClientTypeList)
            {
                try
                {
                    object PluginHandle = pluginAssembly.CreateInstance(ClientType.FullName);
                    if (PluginHandle != null)
                    {
                        string PluginName = PluginHandle.GetType().InvokeMember(nameof(IPluginClient.Name), BindingFlags.Default | BindingFlags.GetProperty, null, PluginHandle, null) as string;
                        Guid PluginGuid = (Guid)PluginHandle.GetType().InvokeMember(nameof(IPluginClient.Guid), BindingFlags.Default | BindingFlags.GetProperty, null, PluginHandle, null);
                        bool PluginRequireElevated = (bool)PluginHandle.GetType().InvokeMember(nameof(IPluginClient.RequireElevated), BindingFlags.Default | BindingFlags.GetProperty, null, PluginHandle, null);
                        bool PluginHasClickHandler = (bool)PluginHandle.GetType().InvokeMember(nameof(IPluginClient.HasClickHandler), BindingFlags.Default | BindingFlags.GetProperty, null, PluginHandle, null);

                        if (!string.IsNullOrEmpty(PluginName) && PluginGuid != Guid.Empty)
                        {
                            bool createdNew;
                            EventWaitHandle InstanceEvent;

                            if (PluginGuid != embeddedPluginGuid)
                                InstanceEvent = new EventWaitHandle(false, EventResetMode.ManualReset, GuidToString(PluginGuid), out createdNew);
                            else
                            {
                                createdNew = true;
                                InstanceEvent = null;
                            }

                            if (createdNew)
                            {
                                IPluginClient NewPlugin = new PluginClient(PluginHandle, PluginName, PluginGuid, PluginRequireElevated, PluginHasClickHandler, InstanceEvent);
                                PluginList.Add(NewPlugin);
                            }
                            else
                            {
                                logger.AddLog("Another instance of a plugin is already running");
                                InstanceEvent.Close();
                                InstanceEvent = null;
                            }
                        }
                    }
                }
                catch
                {
                }
            }
        }

        public static Guid PreferredPluginGuid
        {
            get { return PreferredPlugin != null ? PreferredPlugin.Guid : Guid.Empty; }
            set
            {
                foreach (IPluginClient Plugin in ConsolidatedPluginList)
                    if (Plugin.Guid == value)
                    {
                        PreferredPlugin = Plugin;
                        break;
                    }
            }
        }

        public static List<IPluginClient> ConsolidatedPluginList = new List<IPluginClient>();
        private static IPluginClient PreferredPlugin;

        public static T PluginProperty<T>(object pluginHandle, string propertyName)
        {
            return (T)pluginHandle.GetType().InvokeMember(propertyName, BindingFlags.Default | BindingFlags.GetProperty, null, pluginHandle, null);
        }

        public static void ExecutePluginMethod(object pluginHandle, string methodName, params object[] args)
        {
            pluginHandle.GetType().InvokeMember(methodName, BindingFlags.Default | BindingFlags.InvokeMethod, null, pluginHandle, args);
        }

        public static T GetPluginFunctionValue<T>(object pluginHandle, string functionName, params object[] args)
        {
            return (T)pluginHandle.GetType().InvokeMember(functionName, BindingFlags.Default | BindingFlags.InvokeMethod, null, pluginHandle, args);
        }

        public static bool RequireElevated { get; private set; }
        public static Dictionary<ICommand, IPluginClient> CommandTable { get; } = new Dictionary<ICommand, IPluginClient>();
        public static Dictionary<List<ICommand>, string> FullCommandList { get; } = new Dictionary<List<ICommand>, string>();

        public static List<ICommand> GetChangedCommands(bool beforeMenuOpening)
        {
            List<ICommand> Result = new List<ICommand>();

            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    if (Plugin.GetIsMenuChanged(false))
                    {
                        foreach (KeyValuePair<ICommand, IPluginClient> CommandEntry in CommandTable)
                            if (CommandEntry.Value == Plugin)
                                Result.Add(CommandEntry.Key);
                    }

            return Result;
        }

        public static string GetMenuHeader(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            return Plugin.GetMenuHeader(Command);
        }

        public static bool GetMenuIsVisible(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            return Plugin.GetMenuIsVisible(Command);
        }

        public static bool GetMenuIsEnabled(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            return Plugin.GetMenuIsEnabled(Command);
        }

        public static bool GetMenuIsChecked(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            return Plugin.GetMenuIsChecked(Command);
        }

        public static Bitmap GetMenuIcon(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            return Plugin.GetMenuIcon(Command);
        }

        public static void OnMenuOpening()
        {
            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    Plugin.OnMenuOpening();
        }

        public static void OnExecuteCommand(ICommand Command)
        {
            IPluginClient Plugin = CommandTable[Command];
            Plugin.OnExecuteCommand(Command);
        }

        public static bool GetIsIconChanged()
        {
            return PreferredPlugin != null ? PreferredPlugin.GetIsIconChanged() : false;
        }

        public static Icon Icon
        {
            get { return PreferredPlugin != null ? PreferredPlugin.Icon : null; }
        }

        public static void OnIconClicked()
        {
            if (PreferredPlugin != null)
                PreferredPlugin.OnIconClicked();
        }

        public static bool GetIsToolTipChanged()
        {
            return PreferredPlugin != null ? PreferredPlugin.GetIsToolTipChanged() : false;
        }

        public static string ToolTip
        {
            get { return PreferredPlugin != null ? PreferredPlugin.ToolTip : null; }
        }

        public static void OnActivated()
        {
            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    Plugin.OnActivated();
        }

        public static void OnDeactivated()
        {
            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    Plugin.OnDeactivated();
        }

        public static void Shutdown()
        {
            bool CanClose = true;
            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    CanClose &= Plugin.CanClose(CanClose);

            if (!CanClose)
                return;

            foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                foreach (IPluginClient Plugin in Entry.Value)
                    Plugin.BeginClose();

            bool IsClosed;

            do
            {
                IsClosed = true;

                foreach (KeyValuePair<Assembly, List<IPluginClient>> Entry in LoadedPluginTable)
                    foreach (IPluginClient Plugin in Entry.Value)
                        IsClosed &= Plugin.IsClosed;

                if (!IsClosed)
                    Thread.Sleep(100);
            }
            while (!IsClosed);

            FullCommandList.Clear();
            LoadedPluginTable.Clear();
        }

        public static string GuidToString(Guid guid)
        {
            return guid.ToString("B").ToUpper();
        }

        private static readonly string SharedPluginAssemblyName = "TaskbarIconShared";
        private static Type PluginInterfaceType;
        private static Dictionary<Assembly, List<IPluginClient>> LoadedPluginTable = new Dictionary<Assembly, List<IPluginClient>>();
        private static bool IsWakeUpDelayElapsed;
    }
}
