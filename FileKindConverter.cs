using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Data;
using System.Windows.Media;

namespace PDFProcessor
{
    /// <summary>
    /// Maps a file name to an icon glyph (default), an accent brush (parameter "brush") or a short
    /// type label (parameter "label") for the merge file list.
    /// </summary>
    public sealed class FileKindConverter : IValueConverter
    {
        private static readonly string[] Images = { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".gif" };

        private static (string Glyph, Color Color, string Label) Describe(string name)
        {
            string ext = Path.GetExtension(name ?? "").ToLowerInvariant();
            if (Images.Contains(ext)) return ("", Color.FromRgb(124, 58, 237), "이미지");
            if (OfficeConverter.WordExtensions.Contains(ext)) return ("", Color.FromRgb(37, 99, 235), "워드");
            if (OfficeConverter.ExcelExtensions.Contains(ext)) return ("", Color.FromRgb(22, 163, 74), "엑셀");
            if (OfficeConverter.PowerPointExtensions.Contains(ext)) return ("", Color.FromRgb(234, 88, 12), "PPT");
            return ("", Color.FromRgb(220, 38, 38), "PDF");
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            var kind = Describe(value as string);
            return (parameter as string) switch
            {
                "brush" => new SolidColorBrush(kind.Color),
                "soft" => new SolidColorBrush(Color.FromArgb(28, kind.Color.R, kind.Color.G, kind.Color.B)),
                "label" => kind.Label,
                _ => kind.Glyph,
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
            throw new NotSupportedException();
    }
}
