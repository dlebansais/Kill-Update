using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Forms;
using System.Windows.Input;

namespace TaskbarTools
{
    public class TaskbarIcon : IDisposable
    {
        #region Init
        protected TaskbarIcon(NotifyIcon notifyIcon, IInputElement target)
        {
            NotifyIcon = notifyIcon;
            Target = target;
        }

        protected static List<TaskbarIcon> ActiveIconList { get; private set; } = new List<TaskbarIcon>();
        private NotifyIcon NotifyIcon;
        private IInputElement Target;
        #endregion

        #region Client Interface
        /// <summary>
        /// Create and display a taskbar icon.
        /// </summary>
        /// <param name="icon">The icon displayed</param>
        /// <param name="toolTipText">The text shown when the mouse is over the icon, can be null</param>
        /// <param name="menu">The menu that pops up when the user left click the icon, can be null</param>
        /// <param name="target">The object that receives command notifications, can be null</param>
        /// <returns>The created taskbar icon object</returns>
        public static TaskbarIcon Create(Icon icon, string toolTipText, System.Windows.Controls.ContextMenu menu, IInputElement target)
        {
            try
            {
                NotifyIcon NotifyIcon = new NotifyIcon();
                NotifyIcon.Icon = icon;
                NotifyIcon.Text = "";
                NotifyIcon.Click += OnClick;

                TaskbarIcon NewTaskbarIcon = new TaskbarIcon(NotifyIcon, target);
                NotifyIcon.ContextMenuStrip = NewTaskbarIcon.MenuToMenuStrip(menu);
                ActiveIconList.Add(NewTaskbarIcon);
                NewTaskbarIcon.UpdateToolTip(toolTipText);
                NotifyIcon.Visible = true;

                return NewTaskbarIcon;
            }
            catch (Exception e)
            {
                throw new IconCreationFailedException(e);
            }
        }

        /// <summary>
        /// Toggles the check mark of a menu item.
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="isChecked">The new value of the check mark</param>
        public void ToggleMenuIsChecked(ICommand command, out bool isChecked)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            isChecked = !MenuItem.Checked;
            MenuItem.Checked = isChecked;
        }

        /// <summary>
        /// Returns the current check mark of a menu item.
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <returns>True if the menu item has a check mark, false otherwise</returns>
        public bool IsMenuChecked(ICommand command)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            return MenuItem.Checked;
        }

        /// <summary>
        /// Set the check mark of a menu item. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="isChecked">True if the menu item must have a check mark, false otherwise</param>
        public void SetMenuIsChecked(ICommand command, bool isChecked)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Checked = isChecked;
        }

        /// <summary>
        /// Set the text of menu item. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="header">The new menu item text</param>
        public void SetMenuHeader(ICommand command, string header)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Text = header;
        }

        /// <summary>
        /// Enable or disable the menu item. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="isEnabled">True if enabled</param>
        public void SetMenuIsEnabled(ICommand command, bool isEnabled)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Enabled = isEnabled;
        }

        /// <summary>
        /// Show or hide the menu item. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="isVisible">True to show the menu item</param>
        public void SetMenuIsVisible(ICommand command, bool isVisible)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Visible = isVisible;
        }

        /// <summary>
        /// Set the menu item icon. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="icon">The icon to set, null for no icon</param>
        public void SetMenuIcon(ICommand command, Icon icon)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Image = icon.ToBitmap();
        }

        /// <summary>
        /// Set the menu item icon. This can be called within a handler of the <see cref="MenuOpening"/> event, the change is applied as the menu pops up. 
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="bitmap">The icon to set, as a bitmap, null for no icon</param>
        public void SetMenuIcon(ICommand command, Bitmap bitmap)
        {
            ToolStripMenuItem MenuItem = GetMenuItemFromCommand(command);
            MenuItem.Image = bitmap;
        }

        /// <summary>
        /// Change the taskbar icon.
        /// </summary>
        /// <param name="command">The command associated to the menu item</param>
        /// <param name="text">The new menu item text</param>
        public void UpdateIcon(Icon icon)
        {
            SetNotifyIcon(NotifyIcon, icon);
        }

        /// <summary>
        /// Set the tool tip text displayed when the mouse is over the taskbar icon.
        /// </summary>
        /// <param name="toolTipText">The new tool tip text</param>
        public void UpdateToolTip(string toolTipText)
        {
            // Various versions of windows have length limitations (documented as usual).
            // We remove extra lines until it works...
            for (;;)
            {
                try
                {
                    SetNotifyIconText(NotifyIcon, toolTipText);
                    return;
                }
                catch
                {
                    if (string.IsNullOrEmpty(toolTipText))
                        throw;
                    else
                    {
                        string[] Split = toolTipText.Split('\r');

                        toolTipText = "";
                        for (int i = 0; i + 1 < Split.Length; i++)
                        {
                            if (i > 0)
                                toolTipText += "\r";

                            toolTipText += Split[i];
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Prepares a menu item before is is added to a menu, before calling <see cref="Create"/>.
        /// This method is required only if either <paramref>IsVisible</paramref> or <paramref>IsEnabled</paramref> is false.
        /// </summary>
        /// <param name="item">The modified menu item</param>
        /// <param name="isVisible">True if the menu should be visible</param>
        /// <param name="isEnabled">True if the menu should be enabled</param>
        public static void PrepareMenuItem(System.Windows.Controls.MenuItem item, bool isVisible, bool isEnabled)
        {
            item.Visibility = isVisible ? (isEnabled ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden) : System.Windows.Visibility.Collapsed;
        }

        private static void SetNotifyIconText(NotifyIcon ni, string text)
        {
            if (text != null && text.Length >= 128)
                throw new ArgumentOutOfRangeException("Text limited to 127 characters");

            Type t = typeof(NotifyIcon);
            System.Reflection.BindingFlags hidden = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
            t.GetField("text", hidden).SetValue(ni, text);
            if ((bool)t.GetField("added", hidden).GetValue(ni))
                t.GetMethod("UpdateIcon", hidden).Invoke(ni, new object[] { true });
        }

        private static void SetNotifyIcon(NotifyIcon ni, Icon icon)
        {
            Type t = typeof(NotifyIcon);
            System.Reflection.BindingFlags hidden = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
            t.GetField("icon", hidden).SetValue(ni, icon);
            if ((bool)t.GetField("added", hidden).GetValue(ni))
                t.GetMethod("UpdateIcon", hidden).Invoke(ni, new object[] { true });
        }

        private ToolStripMenuItem GetMenuItemFromCommand(ICommand command)
        {
            foreach (KeyValuePair<ToolStripMenuItem, ICommand> Entry in CommandTable)
                if (Entry.Value == command)
                    return Entry.Key;

            throw new InvalidCommandException(command);
        }

        /// <summary>
        /// Event raised before the menu pops up.
        /// </summary>
        public event EventHandler MenuOpening;

        /// <summary>
        /// Event raised when the icon is clicked.
        /// </summary>
        public event EventHandler IconClicked;
        #endregion

        #region Events
        private static void OnClick(object sender, EventArgs e)
        {
            if (e is System.Windows.Forms.MouseEventArgs AsMouseEventArgs)
            {
                foreach (TaskbarIcon Item in ActiveIconList)
                    if (Item.NotifyIcon == sender)
                    {
                        Item.OnClick(AsMouseEventArgs.Button);
                        break;
                    }
            }
        }

        private void OnClick(MouseButtons button)
        {
            switch (button)
            {
                case MouseButtons.Left:
                    IconClicked.Invoke(this, new EventArgs());
                    break;

                case MouseButtons.Right:
                    MenuOpening.Invoke(this, new EventArgs());
                    break;
            }
        }

        private static void OnMenuClicked(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem MenuItem)
                OnMenuClicked(MenuItem);
        }
        #endregion

        #region Menu
        private ContextMenuStrip MenuToMenuStrip(System.Windows.Controls.ContextMenu menu)
        {
            if (menu == null)
                return null;

            ContextMenuStrip Result = new ContextMenuStrip();
            ConvertToolStripMenuItems(menu.Items, Result.Items);

            return Result;
        }

        private void ConvertToolStripMenuItems(System.Windows.Controls.ItemCollection sourceItems, ToolStripItemCollection destinationItems)
        {
            foreach (System.Windows.Controls.Control Item in sourceItems)
                if (Item is System.Windows.Controls.MenuItem AsMenuItem)
                    if (AsMenuItem.Items.Count > 0)
                        AddSubmenuItem(destinationItems, AsMenuItem);
                    else
                        AddMenuItem(destinationItems, AsMenuItem);

                else if (Item is System.Windows.Controls.Separator AsSeparator)
                    AddSeparator(destinationItems);
        }

        private void AddSubmenuItem(ToolStripItemCollection destinationItems, System.Windows.Controls.MenuItem menuItem)
        {
            string MenuHeader = menuItem.Header as string;
            ToolStripMenuItem NewMenuItem = new ToolStripMenuItem(MenuHeader);

            ConvertToolStripMenuItems(menuItem.Items, NewMenuItem.DropDownItems);

            destinationItems.Add(NewMenuItem);
        }

        private void AddMenuItem(ToolStripItemCollection destinationItems, System.Windows.Controls.MenuItem menuItem)
        {
            string MenuHeader = menuItem.Header as string;

            ToolStripMenuItem NewMenuItem;

            if (menuItem.Icon is Bitmap MenuBitmap)
                NewMenuItem = new ToolStripMenuItem(MenuHeader, MenuBitmap);

            else if (menuItem.Icon is Icon MenuIcon)
                NewMenuItem = new ToolStripMenuItem(MenuHeader, MenuIcon.ToBitmap());

            else
                NewMenuItem = new ToolStripMenuItem(MenuHeader);

            NewMenuItem.Click += OnMenuClicked;
            // See PrepareMenuItem for using the visibility to carry Visible/Enabled flags
            NewMenuItem.Visible = (menuItem.Visibility != System.Windows.Visibility.Collapsed);
            NewMenuItem.Enabled = (menuItem.Visibility == System.Windows.Visibility.Visible);
            NewMenuItem.Checked = menuItem.IsChecked;

            destinationItems.Add(NewMenuItem);
            MenuTable.Add(NewMenuItem, this);
            CommandTable.Add(NewMenuItem, menuItem.Command);
        }

        private void AddSeparator(ToolStripItemCollection destinationItems)
        {
            ToolStripSeparator NewSeparator = new ToolStripSeparator();
            destinationItems.Add(NewSeparator);
        }

        private static void OnMenuClicked(ToolStripMenuItem menuItem)
        {
            if (MenuTable.ContainsKey(menuItem) && CommandTable.ContainsKey(menuItem))
            {
                TaskbarIcon TaskbarIcon = MenuTable[menuItem];
                if (CommandTable[menuItem] is RoutedCommand Command && TaskbarIcon.Target != null)
                    Command.Execute(TaskbarIcon, TaskbarIcon.Target);
            }
        }

        private static Dictionary<ToolStripMenuItem, TaskbarIcon> MenuTable = new Dictionary<ToolStripMenuItem, TaskbarIcon>();
        private static Dictionary<ToolStripMenuItem, ICommand> CommandTable = new Dictionary<ToolStripMenuItem, ICommand>();
        #endregion

        #region Implementation of IDisposable
        protected virtual void Dispose(bool isDisposing)
        {
            if (isDisposing)
                DisposeNow();
        }

        private void DisposeNow()
        {
            using (NotifyIcon ToRemove = NotifyIcon)
            {
                ToRemove.Visible = false;
                ToRemove.Click -= OnClick;

                foreach (TaskbarIcon Item in ActiveIconList)
                    if (Item.NotifyIcon == NotifyIcon)
                    {
                        ActiveIconList.Remove(Item);
                        break;
                    }

                NotifyIcon = null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~TaskbarIcon()
        {
            Dispose(false);
        }
        #endregion
    }

    public class IconCreationFailedException : Exception
    {
        public IconCreationFailedException(Exception originalException) { OriginalException = originalException; }
        public Exception OriginalException { get; private set; }
    }

    public class InvalidCommandException : Exception
    {
        public InvalidCommandException(ICommand command) { Command = command; }
        public ICommand Command { get; private set; }
    }
}
