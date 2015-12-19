//------------------------------------------------------------------------------
// <copyright file="UWPTileGeneratorCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Linq;
using System.ComponentModel.Design;
using EnvDTE80;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace UWPTileGenerator
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class UWPTileGeneratorCommand
	{
		/// <summary>
		/// The output window for the VS Window
		/// </summary>
		private IVsOutputWindowPane outputWindow;

		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("b40237da-1c50-4dc7-898d-21c4e08d9b99");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package package;

		/// <summary>
		/// Initializes a new instance of the <see cref="UWPTileGeneratorCommand"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		private UWPTileGeneratorCommand(Package package)
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
			}

			outputWindow = this.ServiceProvider.GetService(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static UWPTileGeneratorCommand Instance
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
			Instance = new UWPTileGeneratorCommand(package);
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
			var dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
			var hierarchy = dte.ToolWindows.SolutionExplorer;
			var solutionRoot = hierarchy.UIHierarchyItems.Item(1);

			for (int i = 1; i <= solutionRoot.UIHierarchyItems.Count; i++)
			{
				var uiHierarchyItems = solutionRoot.UIHierarchyItems.Item(i).UIHierarchyItems;

				foreach (UIHierarchyItem uiHierarchy in uiHierarchyItems)
				{
					if (uiHierarchy.Name.Equals("Package.appxmanifest"))
					{
						var projectItem = uiHierarchy.Object as ProjectItem;
						outputWindow.OutputString($"The package manifest is located at {projectItem.Properties.Item("FullPath").Value.ToString()} \n");
					}
				}
			}

			var selectedItems = (Array)hierarchy.SelectedItems;

			if (selectedItems != null)
			{
				foreach (UIHierarchyItem selectedItem in selectedItems)
				{
					var projectItem = selectedItem.Object as ProjectItem;
					outputWindow.OutputString($"The selected file is located at {projectItem.Properties.Item("FullPath").Value.ToString()} \n");
				}
			}
		}
	}
}