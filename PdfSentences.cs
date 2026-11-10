using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PdfiumViewer;

namespace PDFProcessor
{
    internal sealed class PdfSentencePart
    {
        internal PdfTextRun Run;
        internal int Start;
        internal int Length;
        internal RectangleF Bounds;
        internal bool SpaceBefore;
        internal string Text => Run.Text.Substring(Start, Length);
    }

    internal sealed class PdfSentence
    {
        internal readonly List<PdfSentencePart> Parts = new();
        /// <summary>Set when the text was recognised from the page image instead of read from the content stream.</summary>
        internal OcrLine Ocr;
        internal bool IsOcr => Ocr != null;
        /// <summary>True when every glyph is drawn invisibly (typically an OCR text layer over a scanned image).</summary>
        internal bool IsInvisible => !IsOcr && Parts.All(part => part.Run.IsInvisible);

        public string Text => IsOcr ? Ocr.Text : string.Concat(Parts.Select(part => (part.SpaceBefore ? " " : "") + part.Text));
        public string FontName => IsOcr ? "OCR" : Parts[0].Run.FontName;
        public double FontSize => IsOcr ? Ocr.FontSize : Parts[0].Run.FontSize;
        /// <summary>Shown in the sentence list.</summary>
        public string Label => IsOcr ? "[OCR] " + Ocr.Text : Text;
        internal RectangleF Bounds => IsOcr ? Ocr.Bounds : Parts.Select(part => part.Bounds).Aggregate(RectangleF.Union);
        internal List<RectangleF> LineBounds => IsOcr ? new List<RectangleF> { Ocr.Bounds } : GetSegments().Select(segment => segment.Bounds).ToList();

        /// <summary>
        /// Visible font size in points. The Tf size is used when it agrees with the rendered glyph height;
        /// otherwise (text matrix scaling such as "1 Tf" with a scaled Tm) the size is estimated from the glyph boxes.
        /// </summary>
        internal double DisplayFontSize
        {
            get
            {
                if (IsOcr) return Ocr.FontSize;
                bool sideways = Angle == 90 || Angle == 270;
                // Line thickness across the reading direction.
                var heights = LineBounds.Select(bounds => (double)(sideways ? bounds.Width : bounds.Height))
                    .Where(height => height > 0).OrderBy(height => height).ToList();
                double height = heights.Count == 0 ? 0 : heights[heights.Count / 2];
                double effective = Parts[0].Run.EffectiveFontSize;
                if (effective > 0.5 && (height == 0 || height / effective >= 0.3 && height / effective <= 2.2)) return effective;
                if (height == 0) return Math.Max(1, FontSize);
                double ratio = height / Math.Max(0.01, FontSize);
                return ratio >= 0.45 && ratio <= 1.6 ? FontSize : Math.Max(4, height / 0.9);
            }
        }

        /// <summary>Reading direction on screen (counter-clockwise degrees, multiple of 90).</summary>
        internal int Angle => IsOcr ? 0 : Parts[0].Run.VisualAngle;

        internal List<PdfSentenceSegment> GetSegments()
        {
            var result = new List<PdfSentenceSegment>();
            if (IsOcr) return result;
            foreach (var part in Parts)
            {
                var last = result.LastOrDefault();
                if (last == null || !PdfSentences.SameLine(last.Bounds, part.Bounds) ||
                    last.Parts[0].Run.FontKey != part.Run.FontKey ||
                    Math.Abs(last.Parts[0].Run.FontSize - part.Run.FontSize) > 0.1)
                {
                    last = new PdfSentenceSegment();
                    result.Add(last);
                }
                last.Parts.Add(part);
            }
            return result;
        }
    }

    internal sealed class PdfSentenceSegment
    {
        internal readonly List<PdfSentencePart> Parts = new();
        internal RectangleF Bounds => Parts.Select(part => part.Bounds).Aggregate(RectangleF.Union);
        internal string Text => string.Concat(Parts.Select((part, index) => (index > 0 && part.SpaceBefore ? " " : "") + part.Text));
    }

    internal static class PdfSentences
    {
        internal static bool SameLine(RectangleF first, RectangleF second) =>
            Math.Abs((first.Top + first.Bottom) / 2 - (second.Top + second.Bottom) / 2) <= Math.Max(first.Height, second.Height) * 0.45;

        internal static List<PdfSentence> Build(PdfDocument rendered, int page, IEnumerable<PdfEditRegion> regions)
        {
            var pieces = new List<PdfSentencePart>();
            foreach (var region in regions)
            {
                string text = region.Run.Text;
                if (region.Span.Length != text.Length)
                {
                    AddPiece(0, text.Length);
                    continue;
                }
                int start = 0;
                for (int index = 0; index < text.Length; index++)
                {
                    if (!IsTerminal(text, index)) continue;
                    int end = index + 1;
                    while (end < text.Length && "!?。！？\"'”’)]".Contains(text[end])) end++;
                    AddPiece(start, end);
                    start = end;
                    index = end - 1;
                }
                AddPiece(start, text.Length);

                void AddPiece(int begin, int end)
                {
                    while (begin < end && char.IsWhiteSpace(text[begin])) begin++;
                    while (end > begin && char.IsWhiteSpace(text[end - 1])) end--;
                    if (begin == end) return;
                    var box = region.Bounds;
                    if (region.Span.Length == text.Length)
                    {
                        box = RectangleF.Empty;
                        foreach (var rectangle in rendered.GetTextBounds(new PdfTextSpan(page, region.Span.Offset + begin, end - begin)))
                        {
                            var device = rendered.RectangleFromPdf(page, rectangle.Bounds);
                            var normalized = RectangleF.FromLTRB(Math.Min(device.Left, device.Right), Math.Min(device.Top, device.Bottom),
                                Math.Max(device.Left, device.Right), Math.Max(device.Top, device.Bottom));
                            box = box.IsEmpty ? normalized : RectangleF.Union(box, normalized);
                        }
                    }
                    if (!box.IsEmpty) pieces.Add(new PdfSentencePart { Run = region.Run, Start = begin, Length = end - begin, Bounds = box });
                }
            }

            var lines = new List<List<PdfSentencePart>>();
            foreach (var piece in pieces.OrderBy(part => part.Bounds.Top).ThenBy(part => part.Bounds.Left))
            {
                var line = lines.LastOrDefault(candidate => SameLine(candidate[0].Bounds, piece.Bounds));
                if (line == null) { line = new List<PdfSentencePart>(); lines.Add(line); }
                line.Add(piece);
            }
            var sentences = new List<PdfSentence>();
            PdfSentence current = null;
            float? cellLeft = null; // Set when the current sentence started after a wide gap (a table cell).
            RectangleF previousLine = RectangleF.Empty;
            var orderedLines = lines.OrderBy(parts => parts.Min(part => part.Bounds.Top)).ToList();
            string previousLineText = "";
            foreach (var line in orderedLines)
            {
                line.Sort((first, second) => first.Bounds.Left.CompareTo(second.Bounds.Left));
                var lineBox = line.Select(part => part.Bounds).Aggregate(RectangleF.Union);
                // Large vertical gaps, column changes and headings start a separate text block.
                if (current != null && (lineBox.Top - previousLine.Bottom > Math.Max(lineBox.Height, previousLine.Height) * 1.2 ||
                    Math.Abs(lineBox.Left - previousLine.Left) > Math.Max(36, lineBox.Height * 2) ||
                    Math.Max(lineBox.Height, previousLine.Height) / Math.Max(1, Math.Min(lineBox.Height, previousLine.Height)) > 1.5))
                    current = null;
                // A new list item ("1. …", "- …") or consecutive form fields ("수 신 : …") never continue the
                // previous line's sentence.
                string lineText = string.Join(" ", line.Select(part => part.Text));
                if (current != null && (StartsListItem(lineText) || IsFormField(lineText) && IsFormField(previousLineText)))
                    current = null;
                previousLineText = lineText;
                foreach (var part in line)
                {
                    var previous = current?.Parts.LastOrDefault();
                    bool cellGap = previous != null && SameLine(previous.Bounds, part.Bounds) &&
                        part.Bounds.Left - previous.Bounds.Right > Math.Max(28, part.Bounds.Height * 2);
                    if (cellGap) current = null;
                    // A table cell only continues on the next line in its own column, never into the next row's first cell.
                    if (previous != null && !SameLine(previous.Bounds, part.Bounds) && cellLeft.HasValue &&
                        Math.Abs(part.Bounds.Left - cellLeft.Value) > Math.Max(36, part.Bounds.Height * 2))
                        current = null;
                    if (current == null)
                    {
                        current = new PdfSentence();
                        sentences.Add(current);
                        cellLeft = cellGap ? part.Bounds.Left : null;
                    }
                    previous = current.Parts.LastOrDefault();
                    if (previous != null)
                        part.SpaceBefore = !SameLine(previous.Bounds, part.Bounds) ||
                            (previous.Run == part.Run && part.Start > previous.Start + previous.Length) ||
                            part.Bounds.Left - previous.Bounds.Right > Math.Max(1.5, part.Bounds.Height * 0.18);
                    current.Parts.Add(part);
                    if (EndsSentence(part.Text)) current = null;
                }
                previousLine = lineBox;
            }
            return sentences;
        }

        private static bool EndsSentence(string text)
        {
            // A lone list number such as "1." or "12." starts an item rather than ending a sentence.
            if (Regex.IsMatch(text.Trim(), @"^\d{1,3}\.$")) return false;
            int end = text.Length - 1;
            while (end >= 0 && "\"'”’)]".Contains(text[end])) end--;
            return end >= 0 && ".!?。！？".Contains(text[end]);
        }
        private static bool IsTerminal(string text, int index)
        {
            char value = text[index];
            if (!".!?。！？".Contains(value)) return false;
            if (value == '.' && index > 0 && index + 1 < text.Length && char.IsDigit(text[index - 1]) && char.IsDigit(text[index + 1])) return false;
            if (value == '.' && IsListNumber(text, index)) return false;
            // A dot within a token is usually a URL, email address or abbreviation.
            return value != '.' || index + 1 == text.Length || char.IsWhiteSpace(text[index + 1]) || "\"'”’)]!?".Contains(text[index + 1]);
        }

        private static bool StartsListItem(string text) =>
            Regex.IsMatch(text, @"^\s*(\d{1,3}[.)](\s|$)|\(\d{1,3}\)|[-•·※○●◎▶►■□◦*]\s|[가-힣][.)]\s)");

        // "수    신 : …", "T E L : …", "견적일자 : …": a short label followed by a colon near the start of the line.
        private static bool IsFormField(string text)
        {
            int colon = text.IndexOfAny(new[] { ':', '：' });
            return colon > 0 && colon <= 14 && text.Substring(0, colon).Trim().Length > 0 && !text.Substring(0, colon).Contains("://");
        }

        // "1. 납기", "2) …" style markers: one to three digits that start the text (or follow only whitespace).
        private static bool IsListNumber(string text, int dot)
        {
            int start = dot;
            while (start > 0 && char.IsDigit(text[start - 1])) start--;
            int digits = dot - start;
            if (digits < 1 || digits > 3) return false;
            for (int index = 0; index < start; index++) if (!char.IsWhiteSpace(text[index])) return false;
            return true;
        }
    }
}
