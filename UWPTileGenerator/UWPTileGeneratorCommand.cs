// ******************************************************************
// This code is licensed under the MIT License (MIT).
// THE CODE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
// INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
// IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
// DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
// TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH
// THE CODE OR THE USE OR OTHER DEALINGS IN THE CODE.
// ******************************************************************

using System;
using System.Linq;
using System.ComponentModel.Design;
using EnvDTE80;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using Svg;

namespace UWPTileGenerator
{
    /// <summary>
    /// Command handler for the UWP Tile Generator.
    /// </summary>
    internal sealed class UWPTileGeneratorCommand
    {
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

                ImageGeneration.TileSizes.Keys.AsParallel().ForAll((i) =>
                {
                    if (selectedFileName != i)
                    {
                        var newImagePath = ImageGeneration.GenerateImage(path, i);
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

                ImageGeneration.SplashSizes.Keys.AsParallel().ForAll(i =>
                {
                    // I don't resize the selected file.
                    if (selectedFileName == i) return;

                    var newImagePath = ImageGeneration.GenerateImage(path, i);
                    project.ProjectItems.AddFromFile(newImagePath);
                    _outputWindow.OutputString($"Added {newImagePath} to the project \n");
                });
            }, PackageManifestEditor.ManipulatePackageManifestForSplash);

            _outputWindow.OutputString("Splash generation complete. \n");
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

                if (width < 400 || height < 400)
                {
                    VsShellUtilities.ShowMessageBox(ServiceProvider,
                        "The image you have provided may not scale well due to it's inital size. For better results try a square image larger than 400x400 pixels - bigger the better",
                        "Quality warning", OLEMSGICON.OLEMSGICON_WARNING, OLEMSGBUTTON.OLEMSGBUTTON_OK,
                        OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
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

            return isSquare;
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
    }
}