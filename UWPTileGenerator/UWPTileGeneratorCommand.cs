//------------------------------------------------------------------------------
// <copyright file="UWPTileGeneratorCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.Linq;
using System.ComponentModel.Design;
using EnvDTE80;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.Xml.Linq;
using System.IO;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Collections.Generic;
using System.Windows.Forms;
using Svg;
using Svg.Transforms;

namespace UWPTileGenerator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class UWPTileGeneratorCommand
    {
        /// <summary>
        /// The tile sizes based on target size.
        /// </summary>
        private readonly Dictionary<string, Size> _tileSizes = new Dictionary<string, Size>();

        /// <summary>
        /// The splash sizes based on target size.
        /// </summary>
        private readonly Dictionary<string, Size> _splashSizes = new Dictionary<string, Size>();

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly Package _package;

        /// <summary>
        /// The output window for the VS Window
        /// </summary>
        private readonly IVsOutputWindowPane _outputWindow;

        /// <summary>
        /// UWP tile command identifier.
        /// </summary>
        public const int UwpTileCommandId = 0x0100;

        /// <summary>
        /// The uwp splash command identifier.
        /// </summary>
        public const int UwpSplashCommandId = 0x0200;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("b40237da-1c50-4dc7-898d-21c4e08d9b99");

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static void Initialize(Package package)
        {
            Instance = new UWPTileGeneratorCommand(package);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static UWPTileGeneratorCommand Instance { get; private set; }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider => this._package;

        /// <summary>
        /// Initializes a new instance of the <see cref="UWPTileGeneratorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        private UWPTileGeneratorCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            _package = package;

            var commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var splashCommandId = new CommandID(CommandSet, UwpSplashCommandId);
                var splashMenuItem = new MenuCommand(SplashTileCallback, splashCommandId);
                commandService.AddCommand(splashMenuItem);

                var tileCommandId = new CommandID(CommandSet, UwpTileCommandId);
                var tileMenuItem = new MenuCommand(UWPTileCallback, tileCommandId);
                commandService.AddCommand(tileMenuItem);
            }

            PopulateTileSizes();
            PopulateSplashSizes();

            _outputWindow = ServiceProvider.GetService(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;
        }

        /// <summary>
        /// Generates Splashes screen images.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        /// <exception cref="Exception">No file was selected</exception>
        private void SplashTileCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            var hierarchy = dte.ToolWindows.SolutionExplorer;
            var selectedItems = (Array)hierarchy.SelectedItems;
            var selectedItem = selectedItems.Length > 0 ? (UIHierarchyItem)selectedItems.GetValue(0) : null;

            if (selectedItem == null)
            {
                throw new Exception("No file was selected");
            }

            var projectItem = selectedItem.Object as ProjectItem;

            if (projectItem == null)
            {
                return;
            }

            var path = projectItem.Properties.Item("FullPath").Value.ToString();
            _outputWindow.OutputString($"The selected file is located at {path} \n");

            var selectedFileName = Path.GetFileName(path);
            var extension = Path.GetExtension(path);

            if (ValidateImageFormat(extension)) return;

            IsSquare(extension, path);

            Cursor.Current = Cursors.WaitCursor;

            var project = projectItem.ContainingProject;

            _splashSizes.Keys.AsParallel().ForAll(i =>
            {
                // I don't resize the selected file.
                if (selectedFileName == i) return;

                var newImagePath = GenerateTiles(path, i);
                project.ProjectItems.AddFromFile(newImagePath);
                _outputWindow.OutputString($"Added {newImagePath} to the project \n");
            });

            project.Save();
            Cursor.Current = Cursors.Default;

            FindAndUpdatePackageManifest(Path.GetDirectoryName(path), hierarchy, ManipulatePackageManifestForSplash);
            _outputWindow.OutputString("Tile generation complete. \n");
        }

        private bool ValidateImageFormat(string extension)
        {
            if (extension != ".png" && extension != ".svg")
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider, "You need to select a valid png or svg", "",
                    OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return true;
            }

            return false;
        }

        /// <summary>
        /// Finds the package manifest and updates it.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="hierarchy">The hierarchy.</param>
        /// <param name="updatePackageManifestAction">The update package manifest action.</param>
        private void FindAndUpdatePackageManifest(string directory, UIHierarchy hierarchy, Action<string, string> updatePackageManifestAction)
        {
            var solutionRoot = hierarchy.UIHierarchyItems.Item(1);

            for (var i = 1; i <= solutionRoot.UIHierarchyItems.Count; i++)
            {
                var uiHierarchyItems = solutionRoot.UIHierarchyItems.Item(i).UIHierarchyItems;

                foreach (UIHierarchyItem uiHierarchy in uiHierarchyItems)
                {
                    if (!uiHierarchy.Name.ToLower().Equals("package.appxmanifest")) continue;

                    var projectItem = uiHierarchy.Object as ProjectItem;
                    var path = projectItem?.Properties.Item("FullPath").Value.ToString();

                    _outputWindow.OutputString($"The package manifest is located at {path} \n");

                    // this is needed in case the image is in the root of the project.
                    var imageDirectory = directory == Path.GetDirectoryName(path)
                        ? string.Empty
                        : string.Concat(directory.Replace(Path.GetDirectoryName(path) + "\\", ""), "\\");

                    updatePackageManifestAction(path, imageDirectory);

                    _outputWindow.OutputString("The package manifest has been updated \n");
                }
            }
        }

        /// <summary>
        /// Determines whether the specified image is square with a small margin of error.
        /// </summary>
        /// <param name="extension">The extension.</param>
        /// <param name="path">The path.</param>
        /// <returns>a boolean specifying if the path is square</returns>
        private bool IsSquare(string extension, string path)
        {
            var width = 0;
            var height = 0;

            if (extension == ".png")
            {
                using (var selectedImage = Image.FromFile(path))
                {
                    width = selectedImage.Width;
                    height = selectedImage.Height;
                }
            }
            else if (extension == ".svg")
            {
                var selectedImage = SvgDocument.Open(path);
                width = (int)selectedImage.Width.Value;
                height = (int)selectedImage.Height.Value;
            }

            var isSquare = Math.Abs(width - height) < 5;

            if (!isSquare)
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider,
                        "The selected item must be square and ideally with no padding", "", OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

            if (width < 400 || height < 400)
            {
                VsShellUtilities.ShowMessageBox(this.ServiceProvider,
                    "The image you have provided may not scale well due to it's inital size. For better results try a square image larger than 400x400 pixels - bigger the better",
                    "Quality warning", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

            return isSquare;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void UWPTileCallback(object sender, EventArgs e)
        {
            var dte = (DTE2)this.ServiceProvider.GetService(typeof(DTE));
            var hierarchy = dte.ToolWindows.SolutionExplorer;
            var selectedItems = (Array)hierarchy.SelectedItems;
            var directory = string.Empty;

            if (selectedItems != null)
            {
                foreach (UIHierarchyItem selectedItem in selectedItems)
                {
                    var projectItem = selectedItem.Object as ProjectItem;
                    if (projectItem != null)
                    {
                        var path = projectItem.Properties.Item("FullPath").Value.ToString();

                        directory = Path.GetDirectoryName(path);
                        _outputWindow.OutputString($"The selected file is located at {path} \n");
                        var project = projectItem.ContainingProject;
                        var selectedFileName = Path.GetFileName(path);
                        var extension = Path.GetExtension(path);

                        if (extension != ".png" && extension != ".svg")
                        {
                            VsShellUtilities.ShowMessageBox(this.ServiceProvider, "You need to select a valid png or svg", "Invalid file format", OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                            return;
                        }

                        var isSquare = false;
                        if (extension == ".png")
                        {
                            using (var selectedImage = Image.FromFile(path))
                            {
                                isSquare = Math.Abs(selectedImage.Width - selectedImage.Height) > 5;

                                if (!isSquare)
                                {
                                    if (selectedImage.Width < 400 || selectedImage.Height < 400)
                                    {
                                        VsShellUtilities.ShowMessageBox(this.ServiceProvider, "The image you have provided may not scale well due to it's inital size. For better results try a square image larger than 400x400 pixels", "Quality warning", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                                    }
                                }
                            }
                        }
                        else if (extension == ".svg")
                        {
                            var selectedImage = SvgDocument.Open(path);
                            isSquare = Math.Abs(selectedImage.Width - selectedImage.Height) > 5;
                        }

                        if (isSquare)
                        {
                            VsShellUtilities.ShowMessageBox(this.ServiceProvider, "The selected item must be square and ideally with no padding", "", OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                            return;
                        }

                        var imagePaths = new List<string>();

                        Cursor.Current = Cursors.WaitCursor;

                        _tileSizes.Keys.AsParallel().ForAll((i) =>
                        {
                            if (selectedFileName != i)
                            {
                                var newImagePath = GenerateTiles(path, i);
                                imagePaths.Add(newImagePath);
                            }
                        });

                        foreach (var item in imagePaths)
                        {
                            project.ProjectItems.AddFromFile(item);
                            _outputWindow.OutputString($"Added {item} to the project \n");
                        }

                        project.Save();
                    }

                    Cursor.Current = Cursors.Default;
                }
            }

            var solutionRoot = hierarchy.UIHierarchyItems.Item(1);

            for (var i = 1; i <= solutionRoot.UIHierarchyItems.Count; i++)
            {
                var uiHierarchyItems = solutionRoot.UIHierarchyItems.Item(i).UIHierarchyItems;

                foreach (UIHierarchyItem uiHierarchy in uiHierarchyItems)
                {
                    if (uiHierarchy.Name.ToLower().Equals("package.appxmanifest"))
                    {
                        var projectItem = uiHierarchy.Object as ProjectItem;
                        var path = projectItem.Properties.Item("FullPath").Value.ToString();
                        _outputWindow.OutputString($"The package manifest is located at {path} \n");

                        // this is needed in case the image is in the root of the project.
                        if (directory != null)
                        {
                            var imageDirectory = directory == Path.GetDirectoryName(path) ? string.Empty : string.Concat(directory.Replace(Path.GetDirectoryName(path) + "\\", ""), "\\");

                            ManipulatePackageManifestForTiles(path, imageDirectory);
                        }

                        _outputWindow.OutputString($"The package manifest has been updated \n");
                    }
                }
            }

            _outputWindow.OutputString($"Tile generation complete. \n");
        }

        /// <summary>
        /// Populates the tile sizes dictionary.
        /// </summary>
        private void PopulateTileSizes()
        {
            this._tileSizes.Clear();

            // Small
            this._tileSizes.Add("Square71x71Logo.scale-100.png", new Size(71, 71));
            this._tileSizes.Add("Square71x71Logo.scale-125.png", new Size(89, 89));
            this._tileSizes.Add("Square71x71Logo.scale-150.png", new Size(107, 107));
            this._tileSizes.Add("Square71x71Logo.scale-200.png", new Size(142, 142));
            this._tileSizes.Add("Square71x71Logo.scale-400.png", new Size(284, 284));

            // Medium
            this._tileSizes.Add("Square150x150Logo.scale-100.png", new Size(150, 150));
            this._tileSizes.Add("Square150x150Logo.scale-125.png", new Size(188, 188));
            this._tileSizes.Add("Square150x150Logo.scale-150.png", new Size(225, 225));
            this._tileSizes.Add("Square150x150Logo.scale-200.png", new Size(300, 300));
            this._tileSizes.Add("Square150x150Logo.scale-400.png", new Size(600, 600));

            // Wide							 
            this._tileSizes.Add("Wide310x150Logo.scale-100.png", new Size(310, 150));
            this._tileSizes.Add("Wide310x150Logo.scale-125.png", new Size(388, 188));
            this._tileSizes.Add("Wide310x150Logo.scale-150.png", new Size(465, 225));
            this._tileSizes.Add("Wide310x150Logo.scale-200.png", new Size(620, 300));
            this._tileSizes.Add("Wide310x150Logo.scale-400.png", new Size(1240, 600));

            // Large						 
            this._tileSizes.Add("Square310x310Logo.scale-100.png", new Size(310, 310));
            this._tileSizes.Add("Square310x310Logo.scale-125.png", new Size(388, 388));
            this._tileSizes.Add("Square310x310Logo.scale-150.png", new Size(465, 465));
            this._tileSizes.Add("Square310x310Logo.scale-200.png", new Size(620, 620));
            this._tileSizes.Add("Square310x310Logo.scale-400.png", new Size(1240, 1240));

            // App list
            this._tileSizes.Add("Square44x44Logo.scale-100.png", new Size(44, 44));
            this._tileSizes.Add("Square44x44Logo.scale-125.png", new Size(55, 55));
            this._tileSizes.Add("Square44x44Logo.scale-150.png", new Size(66, 66));
            this._tileSizes.Add("Square44x44Logo.scale-200.png", new Size(88, 88));
            this._tileSizes.Add("Square44x44Logo.scale-400.png", new Size(176, 176));

            // Target size list assets with plate
            this._tileSizes.Add("Square44x44Logo.targetsize-16.png", new Size(16, 16));
            this._tileSizes.Add("Square44x44Logo.targetsize-24.png", new Size(24, 24));
            this._tileSizes.Add("Square44x44Logo.targetsize-32.png", new Size(32, 32));
            this._tileSizes.Add("Square44x44Logo.targetsize-48.png", new Size(48, 48));
            this._tileSizes.Add("Square44x44Logo.targetsize-256.png", new Size(256, 256));

            this._tileSizes.Add("Square44x44Logo.targetsize-16_altform-unplated.png", new Size(16, 16));
            this._tileSizes.Add("Square44x44Logo.targetsize-24_altform-unplated.png", new Size(24, 24));
            this._tileSizes.Add("Square44x44Logo.targetsize-32_altform-unplated.png", new Size(32, 32));
            this._tileSizes.Add("Square44x44Logo.targetsize-48_altform-unplated.png", new Size(48, 48));
            this._tileSizes.Add("Square44x44Logo.targetsize-256_altform-unplated.png", new Size(256, 256));

            this._tileSizes.Add("NewStoreLogo.scale-100.png", new Size(50, 50));
            this._tileSizes.Add("NewStoreLogo.scale-125.png", new Size(63, 63));
            this._tileSizes.Add("NewStoreLogo.scale-150.png", new Size(75, 75));
            this._tileSizes.Add("NewStoreLogo.scale-200.png", new Size(100, 100));
            this._tileSizes.Add("NewStoreLogo.scale-400.png", new Size(200, 200));
        }

        /// <summary>
        /// Populates the splash sizes dictionary.
        /// </summary>
        private void PopulateSplashSizes()
        {
            this._splashSizes.Clear();

            this._splashSizes.Add("SplashScreen.scale-400.png", new Size(2480, 1200));
            this._splashSizes.Add("SplashScreen.scale-200.png", new Size(1240, 600));
            this._splashSizes.Add("SplashScreen.scale-150.png", new Size(930, 450));
            this._splashSizes.Add("SplashScreen.scale-125.png", new Size(775, 375));
            this._splashSizes.Add("SplashScreen.scale-100.png", new Size(620, 300));
        }

        /// <summary>
        /// Manipulates the package manifest for splash.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="directory">The directory.</param>
        private static void ManipulatePackageManifestForSplash(string path, string directory)
        {
            var xdocument = XDocument.Parse(File.ReadAllText(path));
            var xmlNamespace = "http://schemas.microsoft.com/appx/manifest/uap/windows10";

            var visualElemment = xdocument.Descendants(XName.Get("VisualElements", xmlNamespace)).FirstOrDefault();
            if (visualElemment != null)
            {

                var splashScreen = xdocument.Descendants(XName.Get("SplashScreen", xmlNamespace)).FirstOrDefault();
                if (splashScreen == null)
                {
                    visualElemment.Add(new XElement(XName.Get("SplashScreen", xmlNamespace)));
                    splashScreen = xdocument.Descendants(XName.Get("SplashScreen", xmlNamespace)).FirstOrDefault();
                }

                splashScreen.AddAttribute("Image", $@"{directory}SplashScreen.png");
                xdocument.Save(path);
            }
        }

        /// <summary>
        /// Manipulates the package manifest for tiles.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="directory">The directory.</param>
        private static void ManipulatePackageManifestForTiles(string path, string directory)
        {
            var xdocument = XDocument.Parse(File.ReadAllText(path));
            var xmlNamespace = "http://schemas.microsoft.com/appx/manifest/uap/windows10";
            var defaultNamespace = "http://schemas.microsoft.com/appx/manifest/foundation/windows10";

            var logo = xdocument.Descendants(XName.Get("Logo", defaultNamespace)).First();
            logo.Value = $@"{directory}NewStoreLogo.png";

            var visualElemment = xdocument.Descendants(XName.Get("VisualElements", xmlNamespace)).FirstOrDefault();
            if (visualElemment != null)
            {
                visualElemment.AddAttribute("Square150x150Logo", $@"{directory}Square150x150Logo.png");
                visualElemment.AddAttribute("Square44x44Logo", $@"{directory}Square44x44Logo.png");
            }

            var defaultTitle = xdocument.Descendants(XName.Get("DefaultTile", xmlNamespace)).FirstOrDefault();
            if (defaultTitle == null)
            {
                visualElemment.Add(new XElement(XName.Get("DefaultTile", xmlNamespace)));
                defaultTitle = xdocument.Descendants(XName.Get("DefaultTile", xmlNamespace)).FirstOrDefault();
            }

            defaultTitle.AddAttribute("Wide310x150Logo", $@"{directory}Wide310x150Logo.png");
            defaultTitle.AddAttribute("Square310x310Logo", $@"{directory}Square310x310Logo.png");
            defaultTitle.AddAttribute("Square71x71Logo", $@"{directory}Square71x71Logo.png");

            xdocument.Save(path);
        }

        /// <summary>
        /// Generates the tiles.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sizeKey">The size key.</param>
        /// <returns></returns>
        private string GenerateTiles(string path, string sizeKey)
        {
            var size = sizeKey.StartsWith("Splash") ? this._splashSizes[sizeKey] : this._tileSizes[sizeKey];
            double xMarginSize = 1;
            double yMarginSize = 1;

            if (sizeKey.StartsWith("Square71x71Logo"))
            {
                xMarginSize = 0.5;
                yMarginSize = 0.5;
            }
            else if (sizeKey.StartsWith("Square150x150Logo") || sizeKey.StartsWith("Wide310x150Logo") || sizeKey.StartsWith("Square310x310Logo"))
            {
                xMarginSize = 0.33;
                yMarginSize = 0.33;
            }
            else if (sizeKey.StartsWith("SplashScreen"))
            {
                xMarginSize = 0.33;
                yMarginSize = 0.33;
            }

            var newImagePath = Path.Combine(Path.GetDirectoryName(path), sizeKey);

            var extension = Path.GetExtension(path);
            if (extension == ".png")
            {
                using (var originalImage = Image.FromFile(path))
                {
                    using (var resizedImage = ResizeImage((Bitmap)originalImage, size, xMargin: xMarginSize, yMargin: yMarginSize))
                    {
                        resizedImage.Save(newImagePath);
                    }
                }
            }
            else if (extension == ".svg")
            {
                using (var resizedImage = ResizeImage(SvgDocument.Open(path), size, xMargin: xMarginSize, yMargin: yMarginSize))
                {
                    resizedImage.Save(newImagePath);
                }
            }

            _outputWindow.OutputString($"Generated image: {newImagePath} \n");

            return newImagePath;
        }

        /// <summary>
        /// Resizes the image.
        /// </summary>
        /// <param name="image">The image.</param>
        /// <param name="size">The size.</param>
        /// <param name="xMargin">The x margin.</param>
        /// <param name="yMargin">The y margin.</param>
        /// <param name="preserveAspectRatio">if set to <c>true</c> [preserve aspect ratio].</param>
        /// <returns></returns>
        public static Image ResizeImage(SvgDocument image, Size size, double xMargin = 1, double yMargin = 1, bool preserveAspectRatio = true)
        {
            int newWidth;
            int newHeight;

            var originalWidth = image.Width.Value;
            var originalHeight = image.Height.Value;

            float percentWidth = (float)size.Width / (float)originalWidth;
            float percentHeight = (float)size.Height / (float)originalHeight;
            float percent = percentHeight < percentWidth ? percentHeight : percentWidth;

            newWidth = (int)((originalWidth * percent) * xMargin);
            newHeight = (int)((originalHeight * percent) * yMargin);

            image.Transforms.Add(new SvgScale(newWidth / originalWidth, newHeight / originalHeight));

            var xPosition = (size.Width - newWidth) / 2;
            var yPosition = (size.Height - newHeight) / 2;

            var newImage = new Bitmap(size.Width, size.Height);

            using (Graphics graphicsHandle = Graphics.FromImage(newImage))
            {
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
                using (var bitmap = image.Draw(newWidth + 10, newHeight + 10))
                {
                    bitmap.MakeTransparent();
                    graphicsHandle.DrawImage(bitmap, xPosition, yPosition);
                }
            }

            return newImage;
        }

        /// <summary>
        /// Resizes the image.
        /// </summary>
        /// <param name="image">The image.</param>
        /// <param name="size">The size.</param>
        /// <param name="xMargin">The x margin.</param>
        /// <param name="yMargin">The y margin.</param>
        /// <param name="preserveAspectRatio">if set to <c>true</c> [preserve aspect ratio].</param>
        /// <returns></returns>
        public static Image ResizeImage(Bitmap image, Size size, double xMargin = 1, double yMargin = 1, bool preserveAspectRatio = true)
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

                newWidth = (int)((originalWidth * percent) * xMargin);
                newHeight = (int)((originalHeight * percent) * yMargin);
            }
            else
            {
                newWidth = size.Width;
                newHeight = size.Height;
            }

            var xPosition = (size.Width - newWidth) / 2;
            var yPosition = (size.Height - newHeight) / 2;

            var newImage = new Bitmap(size.Width, size.Height);

            var firstPixel = image.GetPixel(0, 0);

            var brush = new SolidBrush(firstPixel);

            using (Graphics graphicsHandle = Graphics.FromImage(newImage))
            {
                graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphicsHandle.FillRectangle(brush, new Rectangle(0, 0, xPosition + 10, size.Height + 10));
                graphicsHandle.FillRectangle(brush, new Rectangle(newWidth + xPosition - 10, 0, size.Width - (newWidth + xPosition) + 10, size.Height + 10));
                graphicsHandle.FillRectangle(brush, new Rectangle(0, 0, size.Width + 10, yPosition + 10));
                graphicsHandle.FillRectangle(brush, new Rectangle(0, yPosition + newHeight - 10, size.Width + 10, yPosition + 10));
                graphicsHandle.DrawImage(image, xPosition, yPosition, newWidth, newHeight);
            }

            return newImage;
        }
    }
}