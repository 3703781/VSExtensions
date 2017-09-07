//------------------------------------------------------------------------------
// <copyright file="TWTestCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace _8._4.TWToolbar
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TWTestCommand
    {
        private int currentMCCommand;

        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("6b096892-0707-4ecd-9401-9d4ca8e3ff50");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TWTestCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TWTestCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);

                for (int i = TWTestCommandPackage.cmdidMCItem1; i <= TWTestCommandPackage.cmdidMCItem3; i++)
                {
                    CommandID cmdID = new
                    CommandID(new Guid(TWTestCommandPackage.guidTWTestCommandPackageCmdSet), i);
                    OleMenuCommand mc = new OleMenuCommand(new EventHandler(OnMCItemClicked), cmdID);
                    mc.BeforeQueryStatus += new EventHandler(OnMCItemQueryStatus);
                    commandService.AddCommand(mc);
                    // The first item is, by default, checked.   
                    if (TWTestCommandPackage.cmdidMCItem1 == i)
                    {
                        mc.Checked = true;
                        this.currentMCCommand = i;
                    }
                }
            }
        }

        private void OnMCItemQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand mc = sender as OleMenuCommand;
            if (null != mc)
            {
                mc.Checked = (mc.CommandID.ID == this.currentMCCommand);
            }
        }

        private void OnMCItemClicked(object sender, EventArgs e)
        {
            OleMenuCommand mc = sender as OleMenuCommand;
            if (null != mc)
            {
                string selection;
                switch (mc.CommandID.ID)
                {
                    case TWTestCommandPackage.cmdidMCItem1:
                        selection = "Menu controller Item 1";
                        break;

                    case TWTestCommandPackage.cmdidMCItem2:
                        selection = "Menu controller Item 2";
                        break;

                    case TWTestCommandPackage.cmdidMCItem3:
                        selection = "Menu controller Item 3";
                        break;

                    default:
                        selection = "Unknown command";
                        break;
                }
                this.currentMCCommand = mc.CommandID.ID;

                IVsUIShell uiShell =
                  (IVsUIShell)ServiceProvider.GetService(typeof(SVsUIShell));
                Guid clsid = Guid.Empty;
                int result;
                uiShell.ShowMessageBox(
                       0,
                       ref clsid,
                       "Test Tool Window Toolbar Package",
                       string.Format(CultureInfo.CurrentCulture, "You selected {0}", selection),
                       string.Empty,
                       0,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO,
                       0,
                       out result);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TWTestCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new TWTestCommand(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void MenuItemCallback(object sender, EventArgs e)
        {
            string message = string.Format(CultureInfo.CurrentCulture, "Inside {0}.MenuItemCallback()", this.GetType().FullName);
            string title = "TWTestCommand";

            // Show a message box to prove we were here
            VsShellUtilities.ShowMessageBox(
                this.ServiceProvider,
                message,
                title,
                OLEMSGICON.OLEMSGICON_INFO,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }
    }
}
