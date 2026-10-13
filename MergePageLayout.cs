using System;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PDFProcessor
{
    // Shared by export and preview so paper settings have the same meaning in both.
    internal static class MergePageLayout
    {
        internal static int NormalizeRotation(int rotation) => (rotation % 360 + 360) % 360;

        internal static (double Width, double Height) GetSize(
            string paper, string orientation, double sourceWidth, double sourceHeight)
        {
            var size = paper switch
            {
                "A4" => (595.2756, 841.8898),
                "A3" => (841.8898, 1190.5512),
                "A5" => (419.5276, 595.2756),
                "Letter" => (612.0, 792.0),
                "Legal" => (612.0, 1008.0),
                "원본" => (sourceWidth, sourceHeight),
                _ => throw new ArgumentException("지원하지 않는 용지 크기입니다.", nameof(paper))
            };
            bool landscape = orientation == "가로" ||
                (orientation == "원본" && sourceWidth > sourceHeight);
            return landscape
                ? (Math.Max(size.Item1, size.Item2), Math.Min(size.Item1, size.Item2))
                : (Math.Min(size.Item1, size.Item2), Math.Max(size.Item1, size.Item2));
        }

        internal static void AddPage(PdfDocument output, string path, int pageIndex,
            int rotation, string paper, string orientation)
        {
            rotation = NormalizeRotation(rotation);
            if (System.IO.Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                // Preserve the complete imported page when no paper conversion is requested.
                if (paper == "원본" && orientation == "원본")
                {
                    using var input = PdfReader.Open(path, PdfDocumentOpenMode.Import);
                    var imported = output.AddPage(input.Pages[pageIndex]);
                    imported.Rotate = NormalizeRotation(imported.Rotate + rotation);
                    return;
                }

                using var form = XPdfForm.FromFile(path);
                form.PageNumber = pageIndex + 1;
                var source = form.Page;
                int originalRotation = source.Rotate;
                var originalOrientation = source.Orientation;
                var originalBox = source.MediaBox;
                var visibleBox = source.CropBox;
                if (visibleBox.Width <= 0 || visibleBox.Height <= 0) visibleBox = originalBox;
                double width = visibleBox.Width;
                double height = visibleBox.Height;
                // Apply native PDF rotation and user rotation once, in the drawing transform.
                source.Rotate = 0;
                source.Orientation = PdfSharp.PageOrientation.Portrait;
                source.MediaBox = visibleBox;
                try
                {
                    DrawPage(output, form, width, height,
                        NormalizeRotation(originalRotation + rotation), paper, orientation);
                }
                finally
                {
                    source.Rotate = originalRotation;
                    source.Orientation = originalOrientation;
                    source.MediaBox = originalBox;
                }
            }
            else
            {
                using var image = XImage.FromFile(path);
                DrawPage(output, image, image.PointWidth, image.PointHeight,
                    rotation, paper, orientation);
            }
        }

        private static void DrawPage(PdfDocument output, XImage image, double width,
            double height, int rotation, string paper, string orientation)
        {
            bool sideways = rotation == 90 || rotation == 270;
            double visibleWidth = sideways ? height : width;
            double visibleHeight = sideways ? width : height;
            var size = GetSize(paper, orientation, visibleWidth, visibleHeight);
            var page = output.AddPage();
            page.Width = XUnit.FromPoint(size.Width);
            page.Height = XUnit.FromPoint(size.Height);
            using var graphics = XGraphics.FromPdfPage(page);
            double scale = Math.Min(size.Width / visibleWidth, size.Height / visibleHeight);
            graphics.TranslateTransform(size.Width / 2, size.Height / 2);
            graphics.RotateTransform(rotation);
            graphics.ScaleTransform(scale);
            graphics.DrawImage(image, -width / 2, -height / 2, width, height);
        }
    }
}
