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
            var used = new HashSet<int>();
            foreach (var run in runs)
            {
                var matches = rendered.Search(run.Text, true, false, page, page).Items;
                var match = matches.FirstOrDefault(item => !used.Contains(item.TextSpan.Offset));
                if (match == null) continue;
                used.Add(match.TextSpan.Offset);
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
