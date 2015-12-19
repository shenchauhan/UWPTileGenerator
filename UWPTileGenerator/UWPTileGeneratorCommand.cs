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
using System.Xml;
using System.Xml.Linq;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Collections.Generic;

namespace UWPTileGenerator
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class UWPTileGeneratorCommand
	{
		private Dictionary<string, Size> tileSizes = new Dictionary<string, Size>();

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

			this.PopulateTileSizes();

			outputWindow = this.ServiceProvider.GetService(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;
		}

		private void PopulateTileSizes()
		{
			this.tileSizes.Clear();

			// Small
			this.tileSizes.Add("Square71x71Logo.scale-100.png", new Size(71, 71));
			this.tileSizes.Add("Square71x71Logo.scale-200.png", new Size(142, 142));
			this.tileSizes.Add("Square71x71Logo.scale-400.png", new Size(284, 284));

			// Medium
			this.tileSizes.Add("Square150x150Logo.scale-100.png", new Size(150, 150));
			this.tileSizes.Add("Square150x150Logo.scale-200.png", new Size(300, 300));
			this.tileSizes.Add("Square150x150Logo.scale-400.png", new Size(600, 600));

			// Wide							 
			this.tileSizes.Add("Wide310x150Logo.scale-100.png", new Size(310, 150));
			this.tileSizes.Add("Wide310x150Logo.scale-200.png", new Size(620, 300));
			this.tileSizes.Add("Wide310x150Logo.scale-400.png", new Size(1240, 600));

			// Large						 
			this.tileSizes.Add("Square310x310Logo.scale-100.png", new Size(310, 310));
			this.tileSizes.Add("Square310x310Logo.scale-200.png", new Size(620, 620));
			this.tileSizes.Add("Square310x310Logo.scale-400.png", new Size(1240, 1240));

			// App list
			this.tileSizes.Add("Square44x44Logo.scale-100.png", new Size(44, 44));
			this.tileSizes.Add("Square44x44Logo.scale-200.png", new Size(88, 88));
			this.tileSizes.Add("Square44x44Logo.scale-400.png", new Size(176, 176));

			// Target size list assets with plate
			this.tileSizes.Add("Square44x44Logo.targetsize-16.png", new Size(16, 16));
			this.tileSizes.Add("Square44x44Logo.targetsize-24.png", new Size(24, 24));
			this.tileSizes.Add("Square44x44Logo.targetsize-32.png", new Size(32, 32));
			this.tileSizes.Add("Square44x44Logo.targetsize-48.png", new Size(48, 48));
			this.tileSizes.Add("Square44x44Logo.targetsize-256.png", new Size(256, 256));
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
			var selectedItems = (Array)hierarchy.SelectedItems;

			if (selectedItems != null)
			{
				foreach (UIHierarchyItem selectedItem in selectedItems)
				{
					var projectItem = selectedItem.Object as ProjectItem;
					var path = projectItem.Properties.Item("FullPath").Value.ToString();
					outputWindow.OutputString($"The selected file is located at {path} \n");
					var project = projectItem.ContainingProject;
					var selectedFileName = Path.GetFileName(path);

					this.tileSizes.Keys.AsParallel().ForAll((i) =>
					{
						if (selectedFileName != i)
						{
							var newImagePath = this.GenerateTiles(path, i);
							project.ProjectItems.AddFromFile(newImagePath);
						}
					});

					project.Save();
				}
			}

			var solutionRoot = hierarchy.UIHierarchyItems.Item(1);

			for (int i = 1; i <= solutionRoot.UIHierarchyItems.Count; i++)
			{
				var uiHierarchyItems = solutionRoot.UIHierarchyItems.Item(i).UIHierarchyItems;

				foreach (UIHierarchyItem uiHierarchy in uiHierarchyItems)
				{
					if (uiHierarchy.Name.Equals("Package.appxmanifest"))
					{
						var projectItem = uiHierarchy.Object as ProjectItem;
						var path = projectItem.Properties.Item("FullPath").Value.ToString();
						projectItem = null;
						outputWindow.OutputString($"The package manifest is located at {path} \n");
						this.ManipulatePackageManifest(path);
					}
				}
			}
		}

		private void ManipulatePackageManifest(string path)
		{
			var xdocument = XDocument.Parse(File.ReadAllText(path));
			var xmlNamespace = "http://schemas.microsoft.com/appx/manifest/uap/windows10";

			var visualElemment = xdocument.Descendants(XName.Get("VisualElements", xmlNamespace)).FirstOrDefault();
			if (visualElemment != null)
			{
				visualElemment.AddAttribute("Square150x150Logo", @"Assets\Square150x150Logo.png");
				visualElemment.AddAttribute("Square44x44Logo", @"Assets\Square44x44Logo.png");
			}

			var defaultTitle = xdocument.Descendants(XName.Get("DefaultTile", xmlNamespace)).FirstOrDefault();
			if (defaultTitle != null)
			{
				defaultTitle.AddAttribute("Wide310x150Logo", @"Assets\Wide310x150Logo.png");
				defaultTitle.AddAttribute("Square310x310Logo", @"Assets\Square310x310Logo.png");
				defaultTitle.AddAttribute("Square71x71Logo", @"Assets\Square71x71Logo.png");
			}

			xdocument.Save(path);
		}

		private string GenerateTiles(string path, string sizeKey)
		{
			var originalImage = Image.FromFile(path);
			var size = this.tileSizes[sizeKey];

			var resizedImage = ResizeImage(originalImage, size);

			var directory = Path.GetDirectoryName(path);
			var fileName = Path.GetFileNameWithoutExtension(path);

			var newImagePath = Path.Combine(directory, sizeKey);

			resizedImage.Save(newImagePath);

			return newImagePath;
		}

		public static Image ResizeImage(Image image, Size size, bool preserveAspectRatio = true)
		{
			int newWidth;
			int newHeight;
			if (preserveAspectRatio)
			{
				int originalWidth = image.Width;
				int originalHeight = image.Height;
				float percentWidth = (float)size.Width / (float)originalWidth;
				float percentHeight = (float)size.Height / (float)originalHeight;
				float percent = percentHeight < percentWidth ? percentHeight : percentWidth;
				newWidth = (int)(originalWidth * percent);
				newHeight = (int)(originalHeight * percent);
			}
			else
			{
				newWidth = size.Width;
				newHeight = size.Height;
			}

			var newImage = new Bitmap(size.Width, size.Height);

			var xPosition = (size.Width - newWidth) / 2;
			var yPosition = (size.Height - newHeight) / 2;

			using (Graphics graphicsHandle = Graphics.FromImage(newImage))
			{
				graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
				graphicsHandle.DrawImage(image, xPosition, yPosition, newWidth, newHeight);
			}

			return newImage;
		}
	}
}