using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;
using PdfSharp.Pdf.IO;

namespace PDFProcessor
{
    internal enum OcrMode { None, IfNoText, Force }

    /// <summary>A sentence edit to apply while preparing a page.</summary>
    internal sealed class PdfEditRequest
    {
        internal PdfSentence Sentence;
        /// <summary>Runs of the active document; used to locate the same runs in a private copy.</summary>
        internal IReadOnlyList<PdfTextRun> SourceRuns;
        internal string Replacement;
        /// <summary>Line boxes of the other text on the page (device points); longer text may grow up to them.</summary>
        internal IReadOnlyList<RectangleF> Obstacles;
        /// <summary>Displayed page size in points; bounds how far a line may grow to the right.</summary>
        internal SizeF PageSize;
    }

    internal sealed class PdfEditPreparation : IDisposable
    {
        internal PdfTextEditor Editor;
        internal byte[] Bytes;
        internal int Page;
        internal BitmapSource Image;
        internal float Width, Height;
        internal List<PdfEditRegion> Regions;
        internal List<PdfSentence> Sentences;
        internal string Status;
        /// <summary>Set when OCR ran during this preparation, for the caller to cache per page.</summary>
        internal List<OcrLine> OcrLines;
        /// <summary>Describes how an applied edit was written (original font or substitute font).</summary>
        internal string EditNotice;
        internal int NativeSentenceCount;

        internal static PdfEditPreparation Build(byte[] input, int page, double zoom, PdfEditRequest request = null,
            IReadOnlyList<OcrLine> ocrLines = null, ICollection<string> hiddenOcr = null, OcrMode ocrMode = OcrMode.None)
        {
            var result = new PdfEditPreparation { Page = page, Bytes = input };
            try
            {
                if (request != null)
                {
                    // Work on a private document. A failed edit never changes the active document or undo history.
                    result.Bytes = ApplyEdit(input, page, request, out result.EditNotice);
                }
                result.Editor = new PdfTextEditor(PdfReader.Open(new MemoryStream(result.Bytes), PdfDocumentOpenMode.Modify));
                List<PdfTextRun> runs;
                string analysisError = null;
                try { runs = result.Editor.GetRuns(page); }
                catch (Exception ex)
                {
                    runs = new List<PdfTextRun>();
                    analysisError = $"텍스트 분석 오류: {ex.Message}";
                }
                using var stream = new MemoryStream(result.Bytes);
                using var rendered = PdfiumViewer.PdfDocument.Load(stream);
                var size = rendered.PageSizes[page];
                result.Width = size.Width;
                result.Height = size.Height;
                double scale = 96.0 / 72 * zoom;
                // Bound bitmap memory while keeping the displayed page and hit regions at the requested zoom.
                double renderScale = Math.Min(scale, Math.Min(2200.0 / size.Width, 3000.0 / size.Height));
                using var image = rendered.Render(page, Math.Max(1, (int)(size.Width * renderScale)),
                    Math.Max(1, (int)(size.Height * renderScale)), 96, 96, false);
                result.Image = ToBitmapSource(image);
                result.Regions = PdfEditRegions.Build(rendered, page, runs);
                var sentences = PdfSentences.Build(rendered, page, result.Regions);
                result.NativeSentenceCount = sentences.Count;

                string ocrError = null;
                if (ocrLines == null && (ocrMode == OcrMode.Force || ocrMode == OcrMode.IfNoText && sentences.Count == 0))
                {
                    string reason = PdfOcr.UnavailableReason();
                    if (reason == null) ocrLines = result.OcrLines = PdfOcr.Recognize(rendered, page);
                    else ocrError = reason;
                }
                int ocrCount = 0;
                if (ocrLines != null)
                {
                    foreach (var line in ocrLines)
                    {
                        if (hiddenOcr != null && hiddenOcr.Contains(line.Id)) continue;
                        // Text that exists in the content stream (including edits written over OCR lines) wins.
                        if (sentences.Any(sentence => !sentence.IsOcr && Overlap(sentence.Bounds, line.Bounds) > 0.3)) continue;
                        sentences.Add(new PdfSentence { Ocr = line });
                        ocrCount++;
                    }
                    sentences = sentences.OrderBy(sentence => Math.Round(sentence.Bounds.Top / 4))
                        .ThenBy(sentence => sentence.Bounds.Left).ToList();
                }
                result.Sentences = sentences;
                result.Status = analysisError ?? ocrError ?? (sentences.Count == 0
                    ? "이 페이지에서 수정할 텍스트를 찾지 못했습니다. 스캔 문서라면 'OCR 인식'을 눌러주세요."
                    : ocrCount > 0
                        ? $"OCR로 {ocrCount}줄을 인식했습니다 · 줄을 클릭해 수정 · Enter 완료 · Esc 취소"
                        : "문장을 클릭해 수정 · Enter 완료 · Esc 취소 · Ctrl+S 저장");
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        private static double Overlap(RectangleF first, RectangleF second)
        {
            var intersection = RectangleF.Intersect(first, second);
            if (intersection.IsEmpty) return 0;
            return intersection.Width * intersection.Height / Math.Max(0.01, second.Width * second.Height);
        }

        private static byte[] ApplyEdit(byte[] input, int page, PdfEditRequest request, out string notice)
        {
            string replacement = request.Replacement ?? "";
            var sentence = request.Sentence;
            if (sentence.IsOcr) return ApplyOcrEdit(input, page, sentence.Ocr, replacement, request, out notice);

            if (!sentence.IsInvisible)
            {
                try
                {
                    // Preferred path: re-encode the sentence with the original embedded font, keeping layout and font.
                    using var document = PdfReader.Open(new MemoryStream(input), PdfDocumentOpenMode.Modify);
                    var editor = new PdfTextEditor(document);
                    editor.ReplaceSentence(page, Map(editor, page, sentence, request.SourceRuns), replacement);
                    notice = null;
                    return Save(document);
                }
                // Missing glyphs, longer text or unusual font data: fall back to drawing with a substitute font.
                // ArgumentException means the sentence no longer matches the page and is reported to the user.
                catch (Exception ex) when (ex is not ArgumentException) { }
            }
            return ApplySubstituteEdit(input, page, sentence, request, replacement, out notice);
        }

        private static PdfSentence Map(PdfTextEditor editor, int page, PdfSentence sentence, IReadOnlyList<PdfTextRun> sourceRuns)
        {
            var newRuns = editor.GetRuns(page);
            var originalRuns = sourceRuns.ToList();
            var mapped = new PdfSentence();
            foreach (var part in sentence.Parts)
            {
                int index = originalRuns.IndexOf(part.Run);
                if (index < 0 || index >= newRuns.Count) throw new ArgumentException("문장을 다시 선택하세요.");
                mapped.Parts.Add(new PdfSentencePart { Run = newRuns[index], Start = part.Start,
                    Length = part.Length, Bounds = part.Bounds, SpaceBefore = part.SpaceBefore });
            }
            return mapped;
        }

        /// <summary>
        /// Removes the original glyphs (or covers them when they cannot be removed) and draws the new text with an
        /// installed font at the same position. Used when the embedded font lacks characters or space.
        /// </summary>
        private static byte[] ApplySubstituteEdit(byte[] input, int page, PdfSentence sentence,
            PdfEditRequest request, string replacement, out string notice)
        {
            var sourceRuns = request.SourceRuns;
            var first = sentence.Parts[0].Run;
            bool bold = first.FontName?.Contains("Bold", StringComparison.OrdinalIgnoreCase) == true ||
                first.FontName?.Contains("Black", StringComparison.OrdinalIgnoreCase) == true;
            string family = PdfFontResolver.PickFamily(first.FontName, replacement);
            PdfOverlay.EnsureRenderable(replacement, family, bold);
            var lines = sentence.LineBounds;
            var color = first.FillColor;

            using var document = PdfReader.Open(new MemoryStream(input), PdfDocumentOpenMode.Modify);
            var editor = new PdfTextEditor(document);
            bool removed = false;
            try
            {
                editor.ReplaceSentence(page, Map(editor, page, sentence, sourceRuns), "");
                removed = true;
            }
            catch (Exception ex) when (ex is not ArgumentException) { /* Covered with a background fill below. */ }

            var fills = new List<(RectangleF, Color)>();
            if (!removed || sentence.IsInvisible)
            {
                // The glyphs stay visible (e.g. scanned image under an OCR text layer): paint over them.
                var sampled = SampleColors(input, page, lines);
                for (int index = 0; index < lines.Count; index++)
                    fills.Add((Inflate(lines[index], 0.12f), sampled[index].Background));
                if (sentence.IsInvisible) color = sampled[0].Foreground;
            }
            var (laidOut, anchors) = PreserveLineBreaks(sentence, replacement);
            bool adjusted = PdfOverlay.Draw(document.Pages[page], fills, new OverlayText
            {
                Lines = GrowIntoFreeSpace(lines, request, sentence.Angle), Text = laidOut, FontSize = sentence.DisplayFontSize, Bold = bold, Color = color, Family = family,
                Angle = sentence.Angle, ParagraphAnchors = anchors,
            });
            string fontName = family == PdfFontResolver.DefaultFamily ? "맑은 고딕" : family;
            notice = sentence.IsInvisible
                ? $"스캔 이미지 위 텍스트를 {WithRo(fontName)} 다시 썼습니다"
                : $"원본 글꼴에 없는 글자 또는 길이 때문에 {WithRo(fontName)} 다시 썼습니다";
            if (adjusted) notice += " · 영역에 맞게 크기/줄을 조정";
            return Save(document);
        }

        /// <summary>
        /// Keeps the original line breaks of a multi-line sentence for its unchanged beginning and end, so editing one
        /// item of a list or address does not reflow the other lines. Returns the text with newlines and, for each
        /// resulting paragraph, the index of the line box it starts in.
        /// </summary>
        internal static (string Text, List<int> Anchors) PreserveLineBreaks(PdfSentence sentence, string replacement)
        {
            replacement = replacement.Replace("\r", "");
            var segments = sentence.GetSegments();
            if (segments.Count < 2 || replacement.Contains('\n')) return (replacement, null);
            string original = sentence.Text;
            // Offset in the sentence text where every segment that starts a new line begins.
            var breaks = new List<(int Offset, int Segment)>();
            int position = 0;
            for (int segmentIndex = 0; segmentIndex < segments.Count; segmentIndex++)
            {
                foreach (var part in segments[segmentIndex].Parts)
                {
                    if (part.SpaceBefore) position++;
                    if (part == segments[segmentIndex].Parts[0] && segmentIndex > 0 &&
                        !PdfSentences.SameLine(segments[segmentIndex - 1].Bounds, segments[segmentIndex].Bounds))
                        breaks.Add((position, segmentIndex));
                    position += part.Length;
                }
            }
            int prefix = 0;
            while (prefix < original.Length && prefix < replacement.Length && original[prefix] == replacement[prefix]) prefix++;
            int suffix = 0;
            while (suffix < original.Length - prefix && suffix < replacement.Length - prefix &&
                original[original.Length - 1 - suffix] == replacement[replacement.Length - 1 - suffix]) suffix++;
            var mapped = new List<(int Offset, int Segment)>();
            int last = 0;
            for (int index = 0; index < breaks.Count; index++)
            {
                var (offset, segment) = breaks[index];
                int target = -1;
                if (offset <= prefix) target = offset;
                else if (offset >= original.Length - suffix) target = offset - original.Length + replacement.Length;
                else
                {
                    // Inside the edited region: keep the break when the line's opening text (e.g. "2. 결") is unique.
                    int end = index + 1 < breaks.Count ? breaks[index + 1].Offset : original.Length;
                    string key = original.Substring(offset, Math.Min(6, end - offset)).TrimEnd();
                    int limit = replacement.Length - suffix;
                    if (key.Length >= 2)
                    {
                        int found = replacement.IndexOf(key, last + 1, StringComparison.Ordinal);
                        if (found > 0 && found + key.Length <= limit &&
                            replacement.IndexOf(key, found + 1, StringComparison.Ordinal) is int again && (again < 0 || again + key.Length > limit))
                            target = found;
                    }
                }
                if (target > last && target < replacement.Length)
                {
                    mapped.Add((target, segment));
                    last = target;
                }
            }
            if (mapped.Count == 0) return (replacement, null);
            var text = new System.Text.StringBuilder(replacement);
            foreach (var (offset, _) in mapped.OrderByDescending(item => item.Offset))
            {
                int at = Math.Clamp(offset, 0, text.Length);
                // Replace the joining space (if any) with the line break.
                if (at > 0 && text[at - 1] == ' ') { text.Remove(at - 1, 1); at--; }
                text.Insert(at, '\n');
            }
            var anchors = new List<int> { 0 };
            anchors.AddRange(mapped.OrderBy(item => item.Offset).Select(item => item.Segment));
            return (text.ToString(), anchors);
        }

        private static byte[] ApplyOcrEdit(byte[] input, int page, OcrLine line, string replacement, PdfEditRequest request, out string notice)
        {
            PdfOverlay.EnsureRenderable(replacement, PdfFontResolver.DefaultFamily, false);
            using var document = PdfReader.Open(new MemoryStream(input), PdfDocumentOpenMode.Modify);
            bool adjusted = PdfOverlay.Draw(document.Pages[page],
                new[] { (Inflate(line.Bounds, 0.15f), line.Background) },
                new OverlayText { Lines = GrowIntoFreeSpace(new List<RectangleF> { line.Bounds }, request, 0), Text = replacement, FontSize = line.FontSize, Color = line.TextColor });
            notice = "OCR 텍스트를 맑은 고딕으로 다시 썼습니다" + (adjusted ? " · 영역에 맞게 크기/줄을 조정" : "");
            return Save(document);
        }

        private static List<(Color Background, Color Foreground)> SampleColors(byte[] input, int page, List<RectangleF> boxes)
        {
            using var stream = new MemoryStream(input);
            using var rendered = PdfiumViewer.PdfDocument.Load(stream);
            var size = rendered.PageSizes[page];
            double scale = Math.Min(3, 2600.0 / Math.Max(size.Width, size.Height));
            using var image = rendered.Render(page, Math.Max(1, (int)(size.Width * scale)), Math.Max(1, (int)(size.Height * scale)), 96, 96, false);
            using var bitmap = new Bitmap(image);
            double pixelsPerPoint = bitmap.Width / (double)size.Width;
            return boxes.Select(box => ColorSampler.Sample(bitmap, box, pixelsPerPoint)).ToList();
        }

        /// <summary>
        /// Lets each line of horizontal text extend to the right into empty space (up to the next text on that line
        /// or the page margin), so a longer replacement keeps its size instead of being shrunk.
        /// </summary>
        internal static List<RectangleF> GrowIntoFreeSpace(List<RectangleF> lines, PdfEditRequest request, int angle)
        {
            if (angle != 0 || request?.Obstacles == null || request.PageSize.Width <= 0 || lines.Count == 0) return lines;
            float left = lines.Min(box => box.Left);
            // Mirror the left text margin on the right (at most 8% of the page width).
            float pageLimit = request.PageSize.Width - Math.Max(12, Math.Min(left, request.PageSize.Width * 0.08f));
            return lines.Select(box =>
            {
                float limit = pageLimit;
                foreach (var other in request.Obstacles)
                {
                    bool sameBand = other.Top < box.Bottom - box.Height * 0.2f && other.Bottom > box.Top + box.Height * 0.2f;
                    if (sameBand && other.Left >= box.Right - 1) limit = Math.Min(limit, other.Left - Math.Max(6, box.Height * 0.6f));
                }
                return limit > box.Right ? RectangleF.FromLTRB(box.Left, box.Top, limit, box.Bottom) : box;
            }).ToList();
        }

        // Korean particle "로/으로" after a font name: 으로 when the last syllable has a final consonant (except ㄹ).
        private static string WithRo(string name)
        {
            char last = string.IsNullOrEmpty(name) ? ' ' : name[^1];
            if (last >= 0xAC00 && last <= 0xD7A3)
            {
                int final = (last - 0xAC00) % 28;
                return name + (final != 0 && final != 8 ? "으로" : "로");
            }
            return name + "(으)로";
        }

        private static RectangleF Inflate(RectangleF box, float ratio)
        {
            float padding = Math.Max(1f, box.Height * ratio);
            return RectangleF.FromLTRB(box.Left - padding, box.Top - padding, box.Right + padding, box.Bottom + padding);
        }

        private static byte[] Save(PdfSharp.Pdf.PdfDocument document)
        {
            using var saved = new MemoryStream();
            document.Save(saved, false);
            return saved.ToArray();
        }

        private static BitmapSource ToBitmapSource(System.Drawing.Image image)
        {
            using var png = new MemoryStream();
            image.Save(png, System.Drawing.Imaging.ImageFormat.Png);
            png.Position = 0;
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.StreamSource = png;
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        public void Dispose() => Editor?.Document.Dispose();
    }
}
