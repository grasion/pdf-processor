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
        public string Text => string.Concat(Parts.Select(part => (part.SpaceBefore ? " " : "") + part.Text));
        public string FontName => Parts[0].Run.FontName;
        public double FontSize => Parts[0].Run.FontSize;
        internal RectangleF Bounds => Parts.Select(part => part.Bounds).Aggregate(RectangleF.Union);
        internal List<RectangleF> LineBounds => GetSegments().Select(segment => segment.Bounds).ToList();

        internal List<PdfSentenceSegment> GetSegments()
        {
            var result = new List<PdfSentenceSegment>();
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
            RectangleF previousLine = RectangleF.Empty;
            foreach (var line in lines.OrderBy(parts => parts.Min(part => part.Bounds.Top)))
            {
                line.Sort((first, second) => first.Bounds.Left.CompareTo(second.Bounds.Left));
                var lineBox = line.Select(part => part.Bounds).Aggregate(RectangleF.Union);
                // Large vertical gaps, column changes and headings start a separate text block.
                if (current != null && (lineBox.Top - previousLine.Bottom > Math.Max(lineBox.Height, previousLine.Height) * 1.2 ||
                    Math.Abs(lineBox.Left - previousLine.Left) > Math.Max(36, lineBox.Height * 2) ||
                    Math.Max(lineBox.Height, previousLine.Height) / Math.Max(1, Math.Min(lineBox.Height, previousLine.Height)) > 1.5))
                    current = null;
                foreach (var part in line)
                {
                    var previous = current?.Parts.LastOrDefault();
                    if (previous != null && SameLine(previous.Bounds, part.Bounds) &&
                        part.Bounds.Left - previous.Bounds.Right > Math.Max(28, part.Bounds.Height * 2)) current = null;
                    if (current == null) { current = new PdfSentence(); sentences.Add(current); }
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
            int end = text.Length - 1;
            while (end >= 0 && "\"'”’)]".Contains(text[end])) end--;
            return end >= 0 && ".!?。！？".Contains(text[end]);
        }
        private static bool IsTerminal(string text, int index)
        {
            char value = text[index];
            if (!".!?。！？".Contains(value)) return false;
            if (value == '.' && index > 0 && index + 1 < text.Length && char.IsDigit(text[index - 1]) && char.IsDigit(text[index + 1])) return false;
            // A dot within a token is usually a URL, email address or abbreviation.
            return value != '.' || index + 1 == text.Length || char.IsWhiteSpace(text[index + 1]) || "\"'”’)]!?".Contains(text[index + 1]);
        }
    }
}
