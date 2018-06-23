using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading;
using System.Windows.Input;
using System.Windows.Threading;

namespace TaskbarIconHost
{
    public class PluginClient : IPluginClient
    {
        public PluginClient(object pluginHandle, string name, Guid guid, bool requireElevated, bool hasClickHandler, EventWaitHandle instanceEvent)
        {
            PluginHandle = pluginHandle;
            Name = name;
            Guid = guid;
            RequireElevated = requireElevated;
            HasClickHandler = hasClickHandler;
            InstanceEvent = instanceEvent;
        }

        public object PluginHandle { get; private set; }
        public string Name { get; private set; }
        public Guid Guid { get; private set; }
        public bool RequireElevated { get; private set; }
        public bool HasClickHandler { get; private set; }
        public EventWaitHandle InstanceEvent { get; private set; }

        public void Initialize(bool isElevated, Dispatcher dispatcher, IPluginSettings settings, IPluginLogger logger)
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.Initialize), isElevated, dispatcher, settings, logger);
        }

        public List<ICommand> CommandList { get { return PluginManager.PluginProperty<List<ICommand>>(PluginHandle, nameof(IPluginClient.CommandList)); } }

        public bool GetIsMenuChanged()
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetIsMenuChanged));
        }

        public string GetMenuHeader(ICommand Command)
        {
            return PluginManager.GetPluginFunctionValue<string>(PluginHandle, nameof(IPluginClient.GetMenuHeader), Command);
        }

        public bool GetMenuIsVisible(ICommand Command)
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetMenuIsVisible), Command);
        }

        public bool GetMenuIsEnabled(ICommand Command)
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetMenuIsEnabled), Command);
        }

        public bool GetMenuIsChecked(ICommand Command)
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetMenuIsChecked), Command);
        }

        public Bitmap GetMenuIcon(ICommand Command)
        {
            return PluginManager.GetPluginFunctionValue<Bitmap>(PluginHandle, nameof(IPluginClient.GetMenuIcon), Command);
        }

        public void OnMenuOpening()
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.OnMenuOpening));
        }

        public void OnExecuteCommand(ICommand Command)
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.OnExecuteCommand), Command);
        }

        public bool GetIsIconChanged()
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetIsIconChanged));
        }

        public Icon Icon { get { return PluginManager.PluginProperty<Icon>(PluginHandle, nameof(IPluginClient.Icon)); } }
        public Bitmap SelectionBitmap { get { return PluginManager.PluginProperty<Bitmap>(PluginHandle, nameof(IPluginClient.SelectionBitmap)); } }

        public void OnIconClicked()
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.OnIconClicked));
        }

        public bool GetIsToolTipChanged()
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.GetIsToolTipChanged));
        }

        public string ToolTip { get { return PluginManager.PluginProperty<string>(PluginHandle, nameof(IPluginClient.ToolTip)); } }

        public void OnActivated()
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.OnActivated));
        }

        public void OnDeactivated()
        {
            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.OnDeactivated));
        }

        public bool CanClose(bool canClose)
        {
            return PluginManager.GetPluginFunctionValue<bool>(PluginHandle, nameof(IPluginClient.CanClose), canClose);
        }

        public void BeginClose()
        {
            if (InstanceEvent != null)
            {
                InstanceEvent.Close();
                InstanceEvent = null;
            }

            PluginManager.ExecutePluginMethod(PluginHandle, nameof(IPluginClient.BeginClose));
        }

        public bool IsClosed { get { return PluginManager.PluginProperty<bool>(PluginHandle, nameof(IPluginClient.IsClosed)); } }
    }
}
