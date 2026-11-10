using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace PDFProcessor
{
    /// <summary>Replacement text that is drawn on top of a page with an embedded system font.</summary>
    internal sealed class OverlayText
    {
        /// <summary>Line boxes in device points (top-left origin of the displayed page), filled in order.</summary>
        internal List<RectangleF> Lines = new();
        internal string Text = "";
        internal double FontSize = 12;
        internal bool Bold;
        internal Color Color = Color.Black;
        internal string Family = PdfFontResolver.DefaultFamily;
        /// <summary>Reading direction on screen, counter-clockwise degrees (0, 90, 180 or 270).</summary>
        internal int Angle;
        /// <summary>Optional line box index where each paragraph (split by newline) should start.</summary>
        internal List<int> ParagraphAnchors;
    }

    internal static class PdfOverlay
    {
        /// <summary>
        /// Matrix from device points of the displayed page (what Pdfium reports, rotation and crop box applied)
        /// to the coordinate system of an XGraphics created for the page with its rotation neutralised.
        /// </summary>
        internal static XMatrix DeviceMatrix(PdfPage page, int rotate)
        {
            var media = page.MediaBox;
            var crop = page.Elements.ContainsKey("/CropBox") ? page.CropBox : media;
            double x1 = Math.Max(Math.Min(crop.X1, crop.X2), Math.Min(media.X1, media.X2));
            double y1 = Math.Max(Math.Min(crop.Y1, crop.Y2), Math.Min(media.Y1, media.Y2));
            double x2 = Math.Min(Math.Max(crop.X1, crop.X2), Math.Max(media.X1, media.X2));
            double y2 = Math.Min(Math.Max(crop.Y1, crop.Y2), Math.Max(media.Y1, media.Y2));
            double height = Math.Abs(media.Y2 - media.Y1); // PDFsharp maps XGraphics y to (MediaBox height - y).
            int normalized = ((rotate % 360) + 360) % 360;
            return normalized switch
            {
                90 => new XMatrix(0, -1, 1, 0, x1, height - y1),
                180 => new XMatrix(-1, 0, 0, -1, x2, height - y1),
                270 => new XMatrix(0, 1, -1, 0, x2, height - y2),
                _ => new XMatrix(1, 0, 0, 1, x1, height - y2),
            };
        }

        /// <summary>Throws when the font cannot display a character, so nothing is drawn as an empty box.</summary>
        internal static void EnsureRenderable(string text, string family, bool bold)
        {
            PdfFontResolver.EnsureInstalled();
            var info = new PdfFontResolver().ResolveTypeface(family, bold, false)
                ?? throw new InvalidOperationException("대체 글꼴(맑은 고딕)을 찾을 수 없습니다.");
            var typeface = new System.Windows.Media.GlyphTypeface(new Uri(info.FaceName));
            for (int index = 0; index < text.Length; index++)
            {
                int code = char.IsSurrogatePair(text, index) ? char.ConvertToUtf32(text, index) : text[index];
                if (code > 0xFFFF) index++;
                if (char.IsControl((char)Math.Min(code, 0xFFFF)) || code == ' ') continue;
                if (!typeface.CharacterToGlyphMap.ContainsKey(code))
                    throw new InvalidOperationException($"'{char.ConvertFromUtf32(code)}' 글자는 PDF 글꼴로 표시할 수 없어 입력할 수 없습니다.");
            }
        }

        /// <summary>
        /// Fills the given device rectangles and draws the text. Returns true when the text had to be
        /// shrunk or extended beyond the original line boxes.
        /// </summary>
        internal static bool Draw(PdfPage page, IEnumerable<(RectangleF Box, Color Color)> fills, OverlayText text)
        {
            PdfFontResolver.EnsureInstalled();
            int rotate = page.Elements.GetInteger("/Rotate");
            // XGraphics applies its own landscape transform for rotated pages. Draw in unrotated page space instead.
            // Orientation is deliberately left as Portrait afterwards: restoring Landscape makes PDFsharp swap the
            // MediaBox dimensions when saving, which would change the page size.
            page.Elements.SetInteger("/Rotate", 0);
            page.Orientation = PdfSharp.PageOrientation.Portrait;
            bool adjusted = false;
            try
            {
                using var graphics = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                graphics.MultiplyTransform(DeviceMatrix(page, rotate), XMatrixOrder.Prepend);
                foreach (var fill in fills)
                    graphics.DrawRectangle(new XSolidBrush(XColor.FromArgb(255, fill.Color.R, fill.Color.G, fill.Color.B)),
                        fill.Box.X, fill.Box.Y, fill.Box.Width, fill.Box.Height);
                if (!string.IsNullOrWhiteSpace(text.Text) && text.Lines.Count > 0)
                {
                    int angle = ((text.Angle % 360) + 360) % 360;
                    var lines = text.Lines;
                    if (angle != 0)
                    {
                        // Lay the text out in a frame aligned with its reading direction.
                        graphics.RotateTransform(-angle);
                        lines = text.Lines.Select(box => ToTextFrame(box, angle)).ToList();
                    }
                    adjusted = DrawText(graphics, text, lines);
                }
            }
            finally
            {
                if (rotate != 0) page.Elements.SetInteger("/Rotate", rotate);
                else page.Elements.Remove("/Rotate");
            }
            return adjusted;
        }

        private static XFont CreateFont(OverlayText text, double size) =>
            new(text.Family, Math.Max(1, size), text.Bold ? XFontStyleEx.Bold : XFontStyleEx.Regular,
                new XPdfFontOptions(PdfFontEncoding.Unicode));

        // Inverse of RotateTransform(-angle): device point -> text frame point.
        private static RectangleF ToTextFrame(RectangleF box, int angle)
        {
            double radians = angle * Math.PI / 180;
            double cos = Math.Round(Math.Cos(radians)), sin = Math.Round(Math.Sin(radians));
            var corners = new[] { new PointF(box.Left, box.Top), new PointF(box.Right, box.Top), new PointF(box.Left, box.Bottom), new PointF(box.Right, box.Bottom) }
                .Select(point => new PointF((float)(point.X * cos - point.Y * sin), (float)(point.X * sin + point.Y * cos))).ToList();
            return RectangleF.FromLTRB(corners.Min(point => point.X), corners.Min(point => point.Y),
                corners.Max(point => point.X), corners.Max(point => point.Y));
        }

        private static bool DrawText(XGraphics graphics, OverlayText text, List<RectangleF> lineBoxes)
        {
            // Runs of spaces are kept: forms often align labels with them ("납   기").
            string content = text.Text.Replace("\r", "").Replace('\t', ' ').Trim(' ');
            var boxes = lineBoxes.Where(box => box.Width > 0 && box.Height > 0).ToList();
            if (boxes.Count == 0) return false;
            // Try the original size first, then shrink a little before adding lines below the original area.
            foreach (double scale in new[] { 1.0, 0.94, 0.88, 0.82, 0.76 })
            {
                var font = CreateFont(text, text.FontSize * scale);
                var lines = Layout(graphics, font, content, boxes, text.ParagraphAnchors, extend: false);
                if (lines != null)
                {
                    Render(graphics, font, text, lines);
                    return scale < 1;
                }
            }
            var fallback = CreateFont(text, text.FontSize * 0.76);
            Render(graphics, fallback, text, Layout(graphics, fallback, content, boxes, text.ParagraphAnchors, extend: true));
            return true;
        }

        private static List<(string Text, RectangleF Box)> Layout(XGraphics graphics, XFont font, string content,
            List<RectangleF> boxes, List<int> anchors, bool extend)
        {
            var result = new List<(string, RectangleF)>();
            var paragraphs = content.Split('\n');
            int boxIndex = 0;
            double pitch = boxes.Count > 1
                ? Math.Max(boxes[0].Height, (boxes.Last().Top - boxes[0].Top) / (boxes.Count - 1))
                : Math.Max(boxes[0].Height * 1.35, font.Size * 1.3);
            float maxWidth = boxes.Max(box => box.Width);
            RectangleF NextBox()
            {
                if (boxIndex < boxes.Count) return boxes[boxIndex++];
                if (!extend) return RectangleF.Empty;
                var last = result.Count > 0 ? result[^1].Item2 : boxes.Last();
                boxIndex++;
                return new RectangleF(boxes[0].Left, (float)(last.Top + pitch), Math.Max(maxWidth, last.Width), last.Height);
            }
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++)
            {
                string paragraph = paragraphs[paragraphIndex].Trim(' ');
                // Unchanged lines keep their original line box when earlier text became shorter.
                if (anchors != null && paragraphIndex < anchors.Count && anchors[paragraphIndex] > boxIndex &&
                    anchors[paragraphIndex] < boxes.Count) boxIndex = anchors[paragraphIndex];
                var tokens = Regex.Matches(paragraph, @"\s+|[^\s]+").Select(match => match.Value).ToList();
                var box = NextBox();
                if (box.IsEmpty) return null;
                string line = "";
                for (int index = 0; index < tokens.Count; index++)
                {
                    string token = tokens[index];
                    string candidate = line + token;
                    double capacity = box.Width * 1.03 + 0.5;
                    if (graphics.MeasureString(candidate.TrimEnd(), font).Width <= capacity) { line = candidate; continue; }
                    if (string.IsNullOrWhiteSpace(token)) continue; // A space never starts a new line.
                    if (line.Trim().Length > 0)
                    {
                        result.Add((line.TrimEnd(), box));
                        line = "";
                        box = NextBox();
                        if (box.IsEmpty) return null;
                        index--;
                        continue;
                    }
                    // A single word wider than the line (common in Korean without spaces): break by character.
                    string piece = "";
                    foreach (string character in TextElements(token))
                    {
                        if (piece.Length > 0 && graphics.MeasureString(piece + character, font).Width > box.Width * 1.03 + 0.5)
                        {
                            result.Add((piece, box));
                            piece = "";
                            box = NextBox();
                            if (box.IsEmpty) return null;
                        }
                        piece += character;
                    }
                    line = piece;
                }
                if (line.Trim().Length > 0) result.Add((line.TrimEnd(), box));
            }
            return result;
        }

        private static IEnumerable<string> TextElements(string text)
        {
            var enumerator = System.Globalization.StringInfo.GetTextElementEnumerator(text);
            while (enumerator.MoveNext()) yield return enumerator.GetTextElement();
        }

        private static void Render(XGraphics graphics, XFont font, OverlayText text, List<(string Text, RectangleF Box)> lines)
        {
            var brush = new XSolidBrush(XColor.FromArgb(255, text.Color.R, text.Color.G, text.Color.B));
            foreach (var (value, box) in lines)
            {
                // Glyph boxes are tight; centring the em box on them matches the original baseline closely.
                double baseline = box.Top + box.Height / 2 + font.Size * 0.36;
                graphics.DrawString(value, font, brush, box.Left, baseline, XStringFormats.BaseLineLeft);
            }
        }
    }

    /// <summary>Estimates background and ink colours of a page area from a rendered bitmap.</summary>
    internal static class ColorSampler
    {
        /// <param name="scale">Bitmap pixels per device point.</param>
        internal static (Color Background, Color Foreground) Sample(Bitmap bitmap, RectangleF box, double scale)
        {
            var area = Rectangle.FromLTRB(
                Math.Clamp((int)Math.Floor((box.Left - 1.5) * scale), 0, bitmap.Width - 1),
                Math.Clamp((int)Math.Floor((box.Top - 1.5) * scale), 0, bitmap.Height - 1),
                Math.Clamp((int)Math.Ceiling((box.Right + 1.5) * scale), 1, bitmap.Width),
                Math.Clamp((int)Math.Ceiling((box.Bottom + 1.5) * scale), 1, bitmap.Height));
            if (area.Width <= 0 || area.Height <= 0) return (Color.White, Color.Black);
            var border = new List<Color>();
            var inner = new List<Color>();
            int step = Math.Max(1, Math.Min(area.Width, area.Height) / 60);
            for (int y = area.Top; y < area.Bottom; y += step)
            for (int x = area.Left; x < area.Right; x += step)
            {
                bool edge = x - area.Left < 2 * step || area.Right - 1 - x < 2 * step ||
                    y - area.Top < 2 * step || area.Bottom - 1 - y < 2 * step;
                (edge ? border : inner).Add(bitmap.GetPixel(x, y));
            }
            if (border.Count == 0) border.AddRange(inner);
            var background = Median(border);
            if (inner.Count == 0) return (background, Contrast(background));
            // Ink is the most distant tenth of the inner pixels from the background.
            var ink = inner.OrderByDescending(color => Distance(color, background)).Take(Math.Max(1, inner.Count / 10)).ToList();
            var foreground = Median(ink);
            if (Distance(foreground, background) < 60) foreground = Contrast(background);
            return (background, foreground);
        }

        private static Color Median(List<Color> colors)
        {
            int Pick(Func<Color, int> channel) => colors.Select(channel).OrderBy(value => value).ElementAt(colors.Count / 2);
            return Color.FromArgb(Pick(color => color.R), Pick(color => color.G), Pick(color => color.B));
        }

        private static double Distance(Color first, Color second) =>
            Math.Sqrt(Math.Pow(first.R - second.R, 2) + Math.Pow(first.G - second.G, 2) + Math.Pow(first.B - second.B, 2));

        private static Color Contrast(Color background) =>
            background.R * 0.299 + background.G * 0.587 + background.B * 0.114 > 128 ? Color.Black : Color.White;
    }
}
