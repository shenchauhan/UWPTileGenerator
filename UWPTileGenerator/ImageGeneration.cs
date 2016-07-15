// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using Svg;
using System;

namespace UWPTileGenerator
{
    /// <summary>
    /// All the image generation for tiles and splash is done here.
    /// </summary>
    public static class ImageGeneration
    {
        /// <summary>
        /// Gets or sets the tile sizes.
        /// </summary>
        /// <value>
        /// The tile sizes.
        /// </value>
        public static Dictionary<string, Size> TileSizes { get; } = new Dictionary<string, Size>();

        /// <summary>
        /// Gets or sets the splash sizes.
        /// </summary>
        /// <value>
        /// The splash sizes.
        /// </value>
        public static Dictionary<string, Size> SplashSizes { get; } = new Dictionary<string, Size>();

        /// <summary>
        /// Initializes the <see cref="ImageGeneration"/> class.
        /// </summary>
        static ImageGeneration()
        {
            PopulateTileSizes();
            PopulateSplashSizes();
        }

        /// <summary>
        /// Populates the tile sizes dictionary.
        /// </summary>
        private static void PopulateTileSizes()
        {
            TileSizes.Clear();

            // Small
            TileSizes.Add("Square71x71Logo.scale-100.png", new Size(71, 71));
            TileSizes.Add("Square71x71Logo.scale-125.png", new Size(89, 89));
            TileSizes.Add("Square71x71Logo.scale-150.png", new Size(107, 107));
            TileSizes.Add("Square71x71Logo.scale-200.png", new Size(142, 142));
            TileSizes.Add("Square71x71Logo.scale-400.png", new Size(284, 284));

            // Medium
            TileSizes.Add("Square150x150Logo.scale-100.png", new Size(150, 150));
            TileSizes.Add("Square150x150Logo.scale-125.png", new Size(188, 188));
            TileSizes.Add("Square150x150Logo.scale-150.png", new Size(225, 225));
            TileSizes.Add("Square150x150Logo.scale-200.png", new Size(300, 300));
            TileSizes.Add("Square150x150Logo.scale-400.png", new Size(600, 600));

            // Wide							 
            TileSizes.Add("Wide310x150Logo.scale-100.png", new Size(310, 150));
            TileSizes.Add("Wide310x150Logo.scale-125.png", new Size(388, 188));
            TileSizes.Add("Wide310x150Logo.scale-150.png", new Size(465, 225));
            TileSizes.Add("Wide310x150Logo.scale-200.png", new Size(620, 300));
            TileSizes.Add("Wide310x150Logo.scale-400.png", new Size(1240, 600));

            // Large						 
            TileSizes.Add("Square310x310Logo.scale-100.png", new Size(310, 310));
            TileSizes.Add("Square310x310Logo.scale-125.png", new Size(388, 388));
            TileSizes.Add("Square310x310Logo.scale-150.png", new Size(465, 465));
            TileSizes.Add("Square310x310Logo.scale-200.png", new Size(620, 620));
            TileSizes.Add("Square310x310Logo.scale-400.png", new Size(1240, 1240));

            // App list
            TileSizes.Add("Square44x44Logo.scale-100.png", new Size(44, 44));
            TileSizes.Add("Square44x44Logo.scale-125.png", new Size(55, 55));
            TileSizes.Add("Square44x44Logo.scale-150.png", new Size(66, 66));
            TileSizes.Add("Square44x44Logo.scale-200.png", new Size(88, 88));
            TileSizes.Add("Square44x44Logo.scale-400.png", new Size(176, 176));

            // Target size list assets with plate
            TileSizes.Add("Square44x44Logo.targetsize-16.png", new Size(16, 16));
            TileSizes.Add("Square44x44Logo.targetsize-24.png", new Size(24, 24));
            TileSizes.Add("Square44x44Logo.targetsize-32.png", new Size(32, 32));
            TileSizes.Add("Square44x44Logo.targetsize-48.png", new Size(48, 48));
            TileSizes.Add("Square44x44Logo.targetsize-256.png", new Size(256, 256));

            TileSizes.Add("Square44x44Logo.targetsize-16_altform-unplated.png", new Size(16, 16));
            TileSizes.Add("Square44x44Logo.targetsize-24_altform-unplated.png", new Size(24, 24));
            TileSizes.Add("Square44x44Logo.targetsize-32_altform-unplated.png", new Size(32, 32));
            TileSizes.Add("Square44x44Logo.targetsize-48_altform-unplated.png", new Size(48, 48));
            TileSizes.Add("Square44x44Logo.targetsize-256_altform-unplated.png", new Size(256, 256));

            TileSizes.Add("NewStoreLogo.scale-100.png", new Size(50, 50));
            TileSizes.Add("NewStoreLogo.scale-125.png", new Size(63, 63));
            TileSizes.Add("NewStoreLogo.scale-150.png", new Size(75, 75));
            TileSizes.Add("NewStoreLogo.scale-200.png", new Size(100, 100));
            TileSizes.Add("NewStoreLogo.scale-400.png", new Size(200, 200));
        }

        /// <summary>
        /// Populates the splash sizes dictionary.
        /// </summary>
        private static void PopulateSplashSizes()
        {
            SplashSizes.Clear();

            SplashSizes.Add("SplashScreen.scale-400.png", new Size(2480, 1200));
            SplashSizes.Add("SplashScreen.scale-200.png", new Size(1240, 600));
            SplashSizes.Add("SplashScreen.scale-150.png", new Size(930, 450));
            SplashSizes.Add("SplashScreen.scale-125.png", new Size(775, 375));
            SplashSizes.Add("SplashScreen.scale-100.png", new Size(620, 300));
        }

        /// <summary>
        /// Generates the tiles.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="sizeKey">The size key.</param>
        /// <returns></returns>
        public static string GenerateImage(string path, string sizeKey)
        {
            var size = sizeKey.StartsWith("Splash") ? SplashSizes[sizeKey] : TileSizes[sizeKey];
            double xMarginSize = 1;
            double yMarginSize = 1;

            if (sizeKey.StartsWith("Square44x44Logo"))
            {
                xMarginSize = 0.75;
                yMarginSize = 0.75;
            }
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
            var originalImageSize = new Size((int)image.Width.Value, (int)image.Height.Value);
            return ResizeImage((newImage, x, y, width, height) =>
            {
                using (Graphics graphicsHandle = Graphics.FromImage(newImage))
                {
                    graphicsHandle.InterpolationMode = InterpolationMode.Default;
                    using (var bitmap = image.Draw(width, height))
                    {
                        bitmap.MakeTransparent();
                        graphicsHandle.DrawImage(bitmap, x, y);
                    }
                }
            },
            originalImageSize,
            size,
            xMargin,
            yMargin,
            preserveAspectRatio);
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
            var originalImageSize = new Size(image.Width, image.Height);
            return ResizeImage((newImage, x, y, width, height) =>
            {
                var firstPixel = image.GetPixel(0, 0);
                var brush = new SolidBrush(firstPixel);

                using (var graphicsHandle = Graphics.FromImage(newImage))
                {
                    graphicsHandle.InterpolationMode = InterpolationMode.Default;
                    graphicsHandle.FillRectangle(brush, new Rectangle(0, 0, size.Width, size.Height));
                    graphicsHandle.DrawImage(image, x, y, width, height);
                }
            },
            originalImageSize,
            size,
            xMargin,
            yMargin,
            preserveAspectRatio);
        }

        /// <summary>
        /// Resizes the image.
        /// </summary>
        /// <param name="action">The action.</param>
        /// <param name="originalImageSize">Size of the original image.</param>
        /// <param name="requestedSize">Size of the requested.</param>
        /// <param name="xMargin">The x margin.</param>
        /// <param name="yMargin">The y margin.</param>
        /// <param name="preserveAspectRatio">if set to <c>true</c> [preserve aspect ratio].</param>
        /// <returns></returns>
        private static Bitmap ResizeImage(Action<Bitmap, int, int, int, int> action, Size originalImageSize, Size requestedSize, double xMargin = 1, double yMargin = 1, bool preserveAspectRatio = true)
        {
            int newWidth;
            int newHeight;

            if (preserveAspectRatio)
            {
                var originalWidth = originalImageSize.Width;
                var originalHeight = originalImageSize.Height;

                var percentWidth = requestedSize.Width / (float)originalWidth;
                var percentHeight = requestedSize.Height / (float)originalHeight;
                var percent = percentHeight < percentWidth ? percentHeight : percentWidth;

                newWidth = (int)(originalWidth * percent * xMargin);
                newHeight = (int)(originalHeight * percent * yMargin);
            }
            else
            {
                newWidth = requestedSize.Width;
                newHeight = requestedSize.Height;
            }

            var xPosition = (requestedSize.Width - newWidth) / 2;
            var yPosition = (requestedSize.Height - newHeight) / 2;

            var newImage = new Bitmap(requestedSize.Width, requestedSize.Height);

            action(newImage, xPosition, yPosition, newWidth, newHeight);

            return newImage;
        }
    }
}
