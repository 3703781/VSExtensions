//------------------------------------------------------------------------------
// <copyright file="ChangeMenuText.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Security.Permissions;
namespace _7._9.改变命令项的外观
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ChangeMenuText
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid MenuGroup = new Guid("50ef76f6-e3ae-4d6c-830d-ed0e92692984");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ChangeMenuText"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private ChangeMenuText(Package package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                CommandID menuCommandID = new CommandID(MenuGroup, CommandId);
                EventHandler eventhandler = this.ShowMessageBox;
                OleMenuCommand menuItem = new OleMenuCommand(ShowMessageBox, menuCommandID);
                menuItem.BeforeQueryStatus += new EventHandler(OnBeforeQueryStatus);
                commandService.AddCommand(menuItem);
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ChangeMenuText Instance
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
            Instance = new ChangeMenuText(package);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void ShowMessageBox(object sender, EventArgs e)
        {
            var command = sender as OleMenuCommand;
            if (command.Text =="New Text")
            {
                ChangeMyCommand(command.CommandID.ID, false);
            }
        }

        public bool ChangeMyCommand(int cmdID,bool enableCmd)
        {
            bool cmdUpdated = false;
            var mcs = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            var newCmdID = new CommandID(new Guid(ChangeMenuTextPackage.guidChangeMenuTextPackageCmdSet), cmdID);
            MenuCommand mc = mcs.FindCommand(newCmdID);
            if (mc != null)
            {
                mc.Enabled = enableCmd;
                cmdUpdated = true;
            }
            return cmdUpdated;
        }

        private void OnBeforeQueryStatus(object sender,EventArgs e)
        {
            if(sender is OleMenuCommand myCommand)
            {
                myCommand.Text = "New Text";
                //myCommand也有Checked,Enabled,Visible等属性
                //myCommand.Checked = true;
                //myCommand.Enabled = false;
                //myCommand.Visible = true;
            }
        }
    }
}
