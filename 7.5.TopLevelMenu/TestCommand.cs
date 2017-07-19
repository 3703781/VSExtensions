//------------------------------------------------------------------------------
// <copyright file="TestCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace _7._5.TopLevelMenu
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class TestCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("22c36517-7e27-46e2-a004-cdae1dc0b176");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package package;

        /// <summary>
        /// Initializes a new instance of the <see cref="TestCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private TestCommand(Package package)
        {
            this.package = package ?? throw new ArgumentNullException("package");

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);
                InitMRUMenu(commandService);
            }
        }
        private int numMRUItems = 4;
        private int baseMRUID = (int)TestCommandPackage.cmdMRUList;
        private ArrayList mruList;
        /// <summary>
        /// 填充列表
        /// </summary>
        private void InitializeMRUList()
        {
            if (this.mruList == null)
            {
                this.mruList = new ArrayList();
                if (this.mruList != null)
                {
                    for (int i = 0; i < this.numMRUItems; i++)
                    {
                        this.mruList.Add(string.Format(CultureInfo.CurrentCulture, "Item{0}", i + 1));
                    }
                }
            }
        }
        /// <summary>
        /// 初始化每一个MRU菜单命令
        /// </summary>
        /// <param name="mcs"></param>
        private void InitMRUMenu(OleMenuCommandService mcs)
        {
            InitializeMRUList();
            for (int i = 0; i < this.numMRUItems; i++)
            {
                var cmdID = new CommandID(new Guid(TestCommandPackage.guidTestCommandPackageCmdSet), this.baseMRUID + i);
                var mc = new OleMenuCommand(new EventHandler(OnMRUExec), cmdID);
                mc.BeforeQueryStatus += new EventHandler(OnMRUQueryStatus);
                mcs.AddCommand(mc);
            }
        }
        /// <summary>
        /// 让指定菜单命令的Text与列表内容同步
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMRUQueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand menuCommand)
            {
                //if (menuCommand.CommandID.ID == 0x201)
                //    menuCommand.Checked = true;
                int MRUItemIndex = menuCommand.CommandID.ID - this.baseMRUID;
                if (MRUItemIndex >= 0 && MRUItemIndex < this.mruList.Count)
                {
                    menuCommand.Text = this.mruList[MRUItemIndex] as string;
                }
            }
        }
        /// <summary>
        /// 把被点击的放到列表里第一个
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnMRUExec(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand menuCommand)
            {
                int MRUItemIndex = menuCommand.CommandID.ID - this.baseMRUID;
                object selectedItem = new object();
                if (MRUItemIndex >= 0 && MRUItemIndex < this.mruList.Count)
                {
                    selectedItem = this.mruList[MRUItemIndex];
                    this.mruList.RemoveAt(MRUItemIndex);
                    this.mruList.Insert(0, selectedItem);
                }
                VsShellUtilities.ShowMessageBox(
                    this.ServiceProvider,
                    string.Format(CultureInfo.CurrentCulture, "Selected {0}", (string)selectedItem),
                    (string)selectedItem,
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST
                    );
            }
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static TestCommand Instance
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
            Instance = new TestCommand(package);
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
            string title = "TestCommand";

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
