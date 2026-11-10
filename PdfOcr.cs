using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace PDFProcessor
{
    /// <summary>One line (or table cell) of text recognised from the page image.</summary>
    internal sealed class OcrLine
    {
        internal string Id;
        internal string Text;
        /// <summary>Device points, top-left origin of the displayed page.</summary>
        internal RectangleF Bounds;
        internal double FontSize;
        internal Color TextColor = Color.Black;
        internal Color Background = Color.White;
    }

    /// <summary>Text recognition with the OCR engine built into Windows 10/11 (no download required).</summary>
    internal static class PdfOcr
    {
        private static readonly object Gate = new();
        private static OcrEngine engine;

        internal static string LanguageName => GetEngine()?.RecognizerLanguage.DisplayName ?? "";

        /// <summary>Returns null when OCR can run, otherwise a message telling the user how to enable it.</summary>
        internal static string UnavailableReason()
        {
            try
            {
                return GetEngine() == null
                    ? "Windows OCR 언어 팩이 없습니다. 설정 → 시간 및 언어 → 언어 → 한국어 → 언어 옵션에서 '광학 문자 인식'을 설치하세요."
                    : null;
            }
            catch (Exception ex) { return "이 Windows에서는 OCR을 사용할 수 없습니다: " + ex.Message; }
        }

        private static OcrEngine GetEngine()
        {
            lock (Gate)
            {
                if (engine != null) return engine;
                var korean = new Windows.Globalization.Language("ko");
                engine = OcrEngine.IsLanguageSupported(korean)
                    ? OcrEngine.TryCreateFromLanguage(korean)
                    : OcrEngine.TryCreateFromUserProfileLanguages();
                return engine;
            }
        }

        /// <summary>Renders the page at roughly 300 DPI and recognises its text lines.</summary>
        internal static List<OcrLine> Recognize(PdfiumViewer.PdfDocument rendered, int page)
        {
            var ocr = GetEngine() ?? throw new InvalidOperationException(UnavailableReason());
            var size = rendered.PageSizes[page];
            double limit = Math.Min(OcrEngine.MaxImageDimension, 4200);
            double scale = Math.Min(300.0 / 72, limit / Math.Max(size.Width, size.Height));
            int width = Math.Max(1, (int)(size.Width * scale)), height = Math.Max(1, (int)(size.Height * scale));
            using var image = rendered.Render(page, width, height, 96, 96, PdfiumViewer.PdfRenderFlags.Annotations);
            using var bitmap = new Bitmap(image.Width, image.Height, PixelFormat.Format32bppArgb);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.DrawImage(image, 0, 0, image.Width, image.Height);
            }
            double pixelsPerPoint = bitmap.Width / (double)size.Width;
            var result = ocr.RecognizeAsync(ToSoftwareBitmap(bitmap)).AsTask().GetAwaiter().GetResult();
            var lines = new List<OcrLine>();
            foreach (var line in result.Lines)
            {
                foreach (var words in SplitColumns(line.Words.ToList()))
                {
                    var box = words.Select(word => word.BoundingRect)
                        .Select(rect => RectangleF.FromLTRB((float)(rect.X / pixelsPerPoint), (float)(rect.Y / pixelsPerPoint),
                            (float)((rect.X + rect.Width) / pixelsPerPoint), (float)((rect.Y + rect.Height) / pixelsPerPoint)))
                        .Aggregate(RectangleF.Union);
                    string text = JoinWords(words, pixelsPerPoint);
                    if (string.IsNullOrWhiteSpace(text) || box.Width <= 0 || box.Height <= 0) continue;
                    var colors = ColorSampler.Sample(bitmap, box, pixelsPerPoint);
                    lines.Add(new OcrLine
                    {
                        Id = $"{page}:{Math.Round(box.X)}:{Math.Round(box.Y)}:{text}",
                        Text = text,
                        Bounds = box,
                        // Recognised boxes span ascenders and descenders, slightly more than the em height of Hangul.
                        FontSize = Math.Max(4, box.Height / 0.95),
                        Background = colors.Background,
                        TextColor = colors.Foreground,
                    });
                }
            }
            return lines.OrderBy(line => line.Bounds.Top).ThenBy(line => line.Bounds.Left).ToList();
        }

        // Windows OCR joins table cells on the same baseline into one line; split them at wide gaps.
        private static IEnumerable<List<OcrWord>> SplitColumns(List<OcrWord> words)
        {
            var current = new List<OcrWord>();
            foreach (var word in words.OrderBy(word => word.BoundingRect.X))
            {
                if (current.Count > 0)
                {
                    var previous = current[^1].BoundingRect;
                    double gap = word.BoundingRect.X - (previous.X + previous.Width);
                    double height = Math.Max(previous.Height, word.BoundingRect.Height);
                    if (gap > height * 2.5) { yield return current; current = new List<OcrWord>(); }
                }
                current.Add(word);
            }
            if (current.Count > 0) yield return current;
        }

        private static string JoinWords(List<OcrWord> words, double pixelsPerPoint)
        {
            var text = new StringBuilder();
            for (int index = 0; index < words.Count; index++)
            {
                if (index > 0)
                {
                    var previous = words[index - 1].BoundingRect;
                    var current = words[index].BoundingRect;
                    double gap = current.X - (previous.X + previous.Width);
                    // CJK recognisers may split one word into characters; only keep visible gaps as spaces.
                    bool cjk = IsCjk(words[index - 1].Text) && IsCjk(words[index].Text);
                    if (!cjk || gap > Math.Max(previous.Height, current.Height) * 0.18) text.Append(' ');
                }
                text.Append(words[index].Text);
            }
            return text.ToString().Trim();
        }

        private static bool IsCjk(string value) => value.Any(character =>
            character >= 0xAC00 && character <= 0xD7A3 || character >= 0x3130 && character <= 0x318F ||
            character >= 0x4E00 && character <= 0x9FFF || character >= 0x3040 && character <= 0x30FF);

        private static SoftwareBitmap ToSoftwareBitmap(Bitmap bitmap)
        {
            var data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                var bytes = new byte[data.Stride * data.Height];
                Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                if (data.Stride != bitmap.Width * 4)
                {
                    var packed = new byte[bitmap.Width * 4 * bitmap.Height];
                    for (int row = 0; row < bitmap.Height; row++)
                        Buffer.BlockCopy(bytes, row * data.Stride, packed, row * bitmap.Width * 4, bitmap.Width * 4);
                    bytes = packed;
                }
                var buffer = Windows.Security.Cryptography.CryptographicBuffer.CreateFromByteArray(bytes);
                return SoftwareBitmap.CreateCopyFromBuffer(buffer, BitmapPixelFormat.Bgra8, bitmap.Width, bitmap.Height, BitmapAlphaMode.Premultiplied);
            }
            finally { bitmap.UnlockBits(data); }
        }
    }
}
