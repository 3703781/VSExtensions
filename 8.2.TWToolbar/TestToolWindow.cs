//------------------------------------------------------------------------------
// <copyright file="TestToolWindow.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

namespace _8._2.TWToolbar
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.VisualStudio.Shell;
    using System.ComponentModel.Design;
    /// <summary>
    /// This class implements the tool window exposed by this package and hosts a user control.
    /// </summary>
    /// <remarks>
    /// In Visual Studio tool windows are composed of a frame (implemented by the shell) and a pane,
    /// usually implemented by the package implementer.
    /// <para>
    /// This class derives from the ToolWindowPane class provided from the MPF in order to use its
    /// implementation of the IVsUIElementPane interface.
    /// </para>
    /// </remarks>
    [Guid("98c10a94-2937-48a9-8da1-a44b521e2ef5")]
    public class TestToolWindow : ToolWindowPane
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="TestToolWindow"/> class.
        /// </summary>
        public TestToolWindow() : base(null)
        {
            this.Caption = "TestToolWindow";

            // This is the user control hosted by the tool window; Note that, even if this class implements IDisposable,
            // we are not calling Dispose on this object. This is because ToolWindowPane calls Dispose on
            // the object returned by the Content property.
            this.Content = new TestToolWindowControl();
            this.ToolBar = new CommandID(new Guid(TWTestCommandPackage.guidTWTestCommandPackageCmdSet), TWTestCommandPackage.TWToolbar);
        }
    }
}
