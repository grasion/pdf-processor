using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media.Imaging;
using PdfSharp.Pdf.IO;

namespace PDFProcessor
{
    internal sealed class PdfEditPreparation : IDisposable
    {
        internal PdfTextEditor Editor;
        internal byte[] Bytes;
        internal BitmapSource Image;
        internal float Width, Height;
        internal List<PdfEditRegion> Regions;
        internal List<PdfSentence> Sentences;
        internal string Status;

        internal static PdfEditPreparation Build(byte[] input, int page, double zoom,
            PdfSentence sentence = null, IReadOnlyList<PdfTextRun> sourceRuns = null, string replacement = null)
        {
            var result = new PdfEditPreparation();
            try
            {
                result.Editor = new PdfTextEditor(PdfReader.Open(new MemoryStream(input), PdfDocumentOpenMode.Modify));
                result.Bytes = input;
                if (sentence != null)
                {
                    // Work on a private document. A failed edit never changes the active document or undo history.
                    var newRuns = result.Editor.GetRuns(page);
                    var originalRuns = sourceRuns.ToList();
                    var mapped = new PdfSentence();
                    foreach (var part in sentence.Parts)
                    {
                        int index = originalRuns.IndexOf(part.Run);
                        if (index < 0 || index >= newRuns.Count) throw new InvalidOperationException("문장을 다시 선택하세요.");
                        mapped.Parts.Add(new PdfSentencePart { Run = newRuns[index], Start = part.Start,
                            Length = part.Length, Bounds = part.Bounds, SpaceBefore = part.SpaceBefore });
                    }
                    result.Editor.ReplaceSentence(page, mapped, replacement);
                    using var saved = new MemoryStream();
                    result.Editor.Document.Save(saved, false);
                    result.Bytes = saved.ToArray();
                    result.Editor.Document.Dispose();
                    result.Editor = new PdfTextEditor(PdfReader.Open(new MemoryStream(result.Bytes), PdfDocumentOpenMode.Modify));
                }
                List<PdfTextRun> runs;
                try
                {
                    runs = result.Editor.GetRuns(page);
                    result.Status = runs.Count == 0
                        ? "이 페이지에는 수정 가능한 텍스트가 없습니다."
                        : "문장을 클릭해 전체 수정 · Enter 완료 · Esc 취소 · Ctrl+S 저장";
                }
                catch (Exception ex)
                {
                    runs = new List<PdfTextRun>();
                    result.Status = $"텍스트 분석 오류: {ex.Message}";
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
                using var png = new MemoryStream();
                image.Save(png, System.Drawing.Imaging.ImageFormat.Png);
                png.Position = 0;
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = png;
                bitmap.CacheOption = BitmapCacheOption.OnLoad;
                bitmap.EndInit();
                bitmap.Freeze();
                result.Image = bitmap;
                result.Regions = PdfEditRegions.Build(rendered, page, runs);
                result.Sentences = PdfSentences.Build(rendered, page, result.Regions);
                return result;
            }
            catch { result.Dispose(); throw; }
        }

        public void Dispose() => Editor?.Document.Dispose();
    }
}
