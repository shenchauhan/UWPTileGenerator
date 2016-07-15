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

using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace UWPTileGenerator
{
    /// <summary>
    /// Where package manifests are manipulated
    /// </summary>
    public static class PackageManifestEditor
    {
        /// <summary>
        /// Manipulates the package manifest for splash.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="directory">The directory.</param>
        public static void ManipulatePackageManifestForSplash(string path, string directory)
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
        public static void ManipulatePackageManifestForTiles(string path, string directory)
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
    }
}
