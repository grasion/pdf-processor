using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using PdfiumViewer;

namespace PDFProcessor
{
    internal sealed class PdfEditRegion
    {
        internal PdfTextRun Run;
        internal RectangleF Bounds;
        internal PdfTextSpan Span;
    }

    internal static class PdfEditRegions
    {
        internal static List<PdfEditRegion> Build(PdfDocument rendered, int page,
            IEnumerable<PdfTextRun> runs)
        {
            var result = new List<PdfEditRegion>();
            // Character ranges already claimed by earlier runs. A short run such as a single spaced heading
            // character must never be matched inside another run's text, or editing that text would erase it.
            var used = new List<(int Start, int End)>();
            int cursor = 0;
            foreach (var run in runs)
            {
                // Pdfium generates its own spaces, so retry with normalised whitespace before giving up.
                PdfMatch match = null;
                foreach (string query in new[] { run.Text, run.Text.Trim(), System.Text.RegularExpressions.Regex.Replace(run.Text.Trim(), @"\s+", " ") }.Distinct())
                {
                    if (query.Length == 0) continue;
                    var free = rendered.Search(query, true, false, page, page).Items
                        .Where(item => !used.Any(span => item.TextSpan.Offset < span.End && span.Start < item.TextSpan.Offset + item.TextSpan.Length))
                        .OrderBy(item => item.TextSpan.Offset).ToList();
                    // Pdfium extracts text in content order, so prefer the first free match after the previous run.
                    match = free.FirstOrDefault(item => item.TextSpan.Offset >= cursor) ?? free.FirstOrDefault();
                    if (match != null) break;
                }
                if (match == null) continue;
                used.Add((match.TextSpan.Offset, match.TextSpan.Offset + Math.Max(1, match.TextSpan.Length)));
                cursor = match.TextSpan.Offset + match.TextSpan.Length;
                RectangleF bounds = RectangleF.Empty;
                foreach (var rectangle in rendered.GetTextBounds(match.TextSpan))
                {
                    var device = rendered.RectangleFromPdf(page, rectangle.Bounds);
                    var normalized = RectangleF.FromLTRB(Math.Min(device.Left, device.Right),
                        Math.Min(device.Top, device.Bottom), Math.Max(device.Left, device.Right),
                        Math.Max(device.Top, device.Bottom));
                    bounds = bounds.IsEmpty ? normalized : RectangleF.Union(bounds, normalized);
                }
                if (!bounds.IsEmpty) result.Add(new PdfEditRegion { Run = run, Bounds = bounds, Span = match.TextSpan });
            }
            return result;
        }
    }
}
