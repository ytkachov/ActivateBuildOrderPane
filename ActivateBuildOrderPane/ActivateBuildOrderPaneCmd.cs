//------------------------------------------------------------------------------
// <copyright file="ActivateBuildOrderPaneCmd.cs" company="YuriTkachov">
//     Copyright (c) YuriTkachov [yuri.tkachov@gmail.com].  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace ActivateBuildOrderPane
{
  /// <summary>
  /// Command handler
  /// </summary>
  internal sealed class ActivateBuildOrderPaneCmd
  {
    /// <summary>
    /// Command ID.
    /// </summary>
    public const int CommandId = 0x0100;

    /// <summary>
    /// Command menu group (command set GUID).
    /// </summary>
    public static readonly Guid CommandSet = new Guid("34e57e03-6e24-4d58-ae36-bb07bac7e0d4");

    /// <summary>
    /// VS Package that provides this command, not null.
    /// </summary>
    private readonly Package package;
    private static readonly string _regKeyName = "General";
    private static readonly string _regValueName = "ActivateBuildOrderPane";
    /// <summary>
    /// Initializes a new instance of the <see cref="ActivateBuildOrderPaneCmd"/> class.
    /// Adds our command handlers for menu (commands must exist in the command table file)
    /// </summary>
    /// <param name="package">Owner package, not null.</param>
    private ActivateBuildOrderPaneCmd(Package package)
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
        var menuItem = new OleMenuCommand(this.MenuItemCallback, menuCommandID);
        menuItem.BeforeQueryStatus += new EventHandler(OnBeforeQueryStatus);
        commandService.AddCommand(menuItem);
      }
    }

    /// <summary>
    /// Gets the instance of the command.
    /// </summary>
    public static ActivateBuildOrderPaneCmd Instance
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
      Instance = new ActivateBuildOrderPaneCmd(package);

      var urr = package.UserRegistryRoot;
      var general = urr.OpenSubKey(_regKeyName);
      Object res = general.GetValue(_regValueName, 0);
      if (res != null)
        Pressed = Convert.ToInt32(res) == 0 ? false : true;
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
      Pressed = !Pressed;

      var urr = package.UserRegistryRoot;
      var general = urr.OpenSubKey(_regKeyName, true);
      general.SetValue(_regValueName, Pressed ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
    }

    private void OnBeforeQueryStatus(object sender, EventArgs e)
    {
      var command = sender as OleMenuCommand;
      if (command != null)
      {
        command.Checked = Pressed;
      }
    }

    public static bool Pressed
    {
      get;
      private set;
    }
  }
}
