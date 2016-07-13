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
        private IServiceProvider ServiceProvider => _package;

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
                var splashMenuItem = new MenuCommand(GenerateSplashTiles, splashCommandId);
                commandService.AddCommand(splashMenuItem);

                var tileCommandId = new CommandID(CommandSet, UwpTileCommandId);
                var tileMenuItem = new MenuCommand(GenerateTiles, tileCommandId);
                commandService.AddCommand(tileMenuItem);
            }

            PopulateTileSizes();
            PopulateSplashSizes();

            _outputWindow = ServiceProvider.GetService(typeof(SVsGeneralOutputWindowPane)) as IVsOutputWindowPane;
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void GenerateTiles(object sender, EventArgs e)
        {
            TileCallback((path, project) =>
            {
                var selectedFileName = Path.GetFileName(path);

                _tileSizes.Keys.AsParallel().ForAll((i) =>
                {
                    if (selectedFileName != i)
                    {
                        var newImagePath = GenerateImage(path, i);
                        project.ProjectItems.AddFromFile(newImagePath);
                        _outputWindow.OutputString($"Added {newImagePath} to the project \n");
                    }
                });
            }, PackageManifestEditor.ManipulatePackageManifestForTiles);

            _outputWindow.OutputString("Tile generation complete. \n");
        }

        /// <summary>
        /// Generates Splashes screen images.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="e">The <see cref="EventArgs"/> instance containing the event data.</param>
        /// <exception cref="Exception">No file was selected</exception>
        private void GenerateSplashTiles(object sender, EventArgs e)
        {
            TileCallback((path, project) =>
            {
                var selectedFileName = Path.GetFileName(path);

                _splashSizes.Keys.AsParallel().ForAll(i =>
                {
                    // I don't resize the selected file.
                    if (selectedFileName == i) return;

                    var newImagePath = GenerateImage(path, i);
                    project.ProjectItems.AddFromFile(newImagePath);
                    _outputWindow.OutputString($"Added {newImagePath} to the project \n");
                });
            }, PackageManifestEditor.ManipulatePackageManifestForSplash);
        }

        /// <summary>
        /// All the logic to get the selected file and then perform the various steps to generate the tiles or splash.
        /// </summary>
        /// <param name="tileProcessingAction">The tile processing action.</param>
        /// <param name="packageManifestAction">The package manifest action.</param>
        /// <exception cref="Exception">No file was selected</exception>
        private void TileCallback(Action<string, Project> tileProcessingAction, Action<string, string> packageManifestAction)
        {
            var dte = (DTE2)ServiceProvider.GetService(typeof(DTE));
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

            var extension = Path.GetExtension(path);

            if (ValidateImageFormat(extension)) return;

            IsSquare(extension, path);

            Cursor.Current = Cursors.WaitCursor;

            var project = projectItem.ContainingProject;

            tileProcessingAction(path, project);

            project.Save();
            Cursor.Current = Cursors.Default;

            FindAndUpdatePackageManifest(Path.GetDirectoryName(path), hierarchy, packageManifestAction);
            _outputWindow.OutputString("Tile generation complete. \n");
        }

        /// <summary>
        /// Validates the image format to make sure it's a PNG or SVG.
        /// </summary>
        /// <param name="extension">The extension.</param>
        /// <returns>bool to state if the image is a valid format.</returns>
        private bool ValidateImageFormat(string extension)
        {
            if (extension != ".png" && extension != ".svg")
            {
                VsShellUtilities.ShowMessageBox(ServiceProvider, "You need to select a valid png or svg", "",
                    OLEMSGICON.OLEMSGICON_CRITICAL, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                return true;
            }

            return false;
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
                VsShellUtilities.ShowMessageBox(ServiceProvider,
                        "The selected item must be square and ideally with no padding", "", OLEMSGICON.OLEMSGICON_CRITICAL,
                        OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

            if (width < 400 || height < 400)
            {
                VsShellUtilities.ShowMessageBox(ServiceProvider,
                    "The image you have provided may not scale well due to it's inital size. For better results try a square image larger than 400x400 pixels - bigger the better",
                    "Quality warning", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }

            return isSquare;
        }

        /// <summary>
        /// Populates the tile sizes dictionary.
        /// </summary>
        private void PopulateTileSizes()
        {
            _tileSizes.Clear();

            // Small
            _tileSizes.Add("Square71x71Logo.scale-100.png", new Size(71, 71));
            _tileSizes.Add("Square71x71Logo.scale-125.png", new Size(89, 89));
            _tileSizes.Add("Square71x71Logo.scale-150.png", new Size(107, 107));
            _tileSizes.Add("Square71x71Logo.scale-200.png", new Size(142, 142));
            _tileSizes.Add("Square71x71Logo.scale-400.png", new Size(284, 284));

            // Medium
            _tileSizes.Add("Square150x150Logo.scale-100.png", new Size(150, 150));
            _tileSizes.Add("Square150x150Logo.scale-125.png", new Size(188, 188));
            _tileSizes.Add("Square150x150Logo.scale-150.png", new Size(225, 225));
            _tileSizes.Add("Square150x150Logo.scale-200.png", new Size(300, 300));
            _tileSizes.Add("Square150x150Logo.scale-400.png", new Size(600, 600));

            // Wide							 
            _tileSizes.Add("Wide310x150Logo.scale-100.png", new Size(310, 150));
            _tileSizes.Add("Wide310x150Logo.scale-125.png", new Size(388, 188));
            _tileSizes.Add("Wide310x150Logo.scale-150.png", new Size(465, 225));
            _tileSizes.Add("Wide310x150Logo.scale-200.png", new Size(620, 300));
            _tileSizes.Add("Wide310x150Logo.scale-400.png", new Size(1240, 600));

            // Large						 
            _tileSizes.Add("Square310x310Logo.scale-100.png", new Size(310, 310));
            _tileSizes.Add("Square310x310Logo.scale-125.png", new Size(388, 388));
            _tileSizes.Add("Square310x310Logo.scale-150.png", new Size(465, 465));
            _tileSizes.Add("Square310x310Logo.scale-200.png", new Size(620, 620));
            _tileSizes.Add("Square310x310Logo.scale-400.png", new Size(1240, 1240));

            // App list
            _tileSizes.Add("Square44x44Logo.scale-100.png", new Size(44, 44));
            _tileSizes.Add("Square44x44Logo.scale-125.png", new Size(55, 55));
            _tileSizes.Add("Square44x44Logo.scale-150.png", new Size(66, 66));
            _tileSizes.Add("Square44x44Logo.scale-200.png", new Size(88, 88));
            _tileSizes.Add("Square44x44Logo.scale-400.png", new Size(176, 176));

            // Target size list assets with plate
            _tileSizes.Add("Square44x44Logo.targetsize-16.png", new Size(16, 16));
            _tileSizes.Add("Square44x44Logo.targetsize-24.png", new Size(24, 24));
            _tileSizes.Add("Square44x44Logo.targetsize-32.png", new Size(32, 32));
            _tileSizes.Add("Square44x44Logo.targetsize-48.png", new Size(48, 48));
            _tileSizes.Add("Square44x44Logo.targetsize-256.png", new Size(256, 256));

            _tileSizes.Add("Square44x44Logo.targetsize-16_altform-unplated.png", new Size(16, 16));
            _tileSizes.Add("Square44x44Logo.targetsize-24_altform-unplated.png", new Size(24, 24));
            _tileSizes.Add("Square44x44Logo.targetsize-32_altform-unplated.png", new Size(32, 32));
            _tileSizes.Add("Square44x44Logo.targetsize-48_altform-unplated.png", new Size(48, 48));
            _tileSizes.Add("Square44x44Logo.targetsize-256_altform-unplated.png", new Size(256, 256));

            _tileSizes.Add("NewStoreLogo.scale-100.png", new Size(50, 50));
            _tileSizes.Add("NewStoreLogo.scale-125.png", new Size(63, 63));
            _tileSizes.Add("NewStoreLogo.scale-150.png", new Size(75, 75));
            _tileSizes.Add("NewStoreLogo.scale-200.png", new Size(100, 100));
            _tileSizes.Add("NewStoreLogo.scale-400.png", new Size(200, 200));
        }

        /// <summary>
        /// Populates the splash sizes dictionary.
        /// </summary>
        private void PopulateSplashSizes()
        {
            _splashSizes.Clear();

            _splashSizes.Add("SplashScreen.scale-400.png", new Size(2480, 1200));
            _splashSizes.Add("SplashScreen.scale-200.png", new Size(1240, 600));
            _splashSizes.Add("SplashScreen.scale-150.png", new Size(930, 450));
            _splashSizes.Add("SplashScreen.scale-125.png", new Size(775, 375));
            _splashSizes.Add("SplashScreen.scale-100.png", new Size(620, 300));
        }

        /// <summary>
        /// Finds the package manifest and updates it.
        /// </summary>
        /// <param name="directory">The directory.</param>
        /// <param name="hierarchy">The hierarchy.</param>
        /// <param name="updatePackageManifestAction">The update package manifest tileProcessingAction.</param>
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
        /// Generates the tiles.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sizeKey">The size key.</param>
        /// <returns></returns>
        private string GenerateImage(string path, string sizeKey)
        {
            var size = sizeKey.StartsWith("Splash") ? _splashSizes[sizeKey] : _tileSizes[sizeKey];
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
        private static Image ResizeImage(SvgDocument image, Size size, double xMargin = 1, double yMargin = 1, bool preserveAspectRatio = true)
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
        private static Image ResizeImage(Bitmap image, Size size, double xMargin = 1, double yMargin = 1, bool preserveAspectRatio = true)
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