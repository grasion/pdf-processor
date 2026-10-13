using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PDFProcessor;

internal static class EditChecks
{
    internal static void Run()
    {
        using var document = new PdfDocument();
        var page = document.AddPage();
        var font = new PdfDictionary(document);
        font.Elements.SetName("/Type", "/Font");
        font.Elements.SetName("/Subtype", "/Type1");
        font.Elements.SetName("/BaseFont", "/Helvetica");
        font.Elements.SetName("/Encoding", "/WinAnsiEncoding");
        var fonts = new PdfDictionary(document);
        fonts.Elements["/F1"] = font;
        page.Resources.Elements["/Font"] = fonts;
        page.Contents.ReplaceContent(PdfSharp.Pdf.Content.ContentReader.ReadContent(Encoding.ASCII.GetBytes("BT /F1 18 Tf 40 700 Td (Hello World) Tj 0 -30 Td [(Second) -30 ( line)] TJ ET")));
        using var original = new MemoryStream();
        document.Save(original, false);
        original.Position = 0;
        using var editable = PdfReader.Open(original, PdfDocumentOpenMode.Modify);
        var editor = new PdfTextEditor(editable);
        var runs = editor.GetRuns(0);
        Require(runs.Count == 2 && runs[0].Text == "Hello World" && runs[1].Text == "Second line", "Read Tj/TJ");
        Require(ReferenceEquals(runs[0].GlyphWidths, runs[1].GlyphWidths) && ReferenceEquals(runs[0].ReverseGlyphs, runs[1].ReverseGlyphs),
            "Font metrics and character maps are cached across text operators");
        // The UI takes a snapshot immediately before applying each edit.
        using var before = new MemoryStream(original.ToArray());
        editor.Replace(0, runs[0], "Hello PDF");
        editor.Replace(0, runs[1], "Changed line");
        Require(font.Elements.GetName("/BaseFont") == "/Helvetica", "Original font retained");
        using var after = new MemoryStream();
        editable.Save(after, false);
        after.Position = 0;
        using var rendered = PdfiumViewer.PdfDocument.Load(after);
        Require(rendered.GetPdfText(0).Contains("Hello PDF") && rendered.GetPdfText(0).Contains("Changed line"), "Saved edited text");
        bool rejected = false;
        try { editor.Replace(0, runs[0], "한글"); } catch (InvalidOperationException) { rejected = true; }
        Require(rejected && runs[0].Text == "Hello PDF", "Unsupported glyph blocked without mutation");
        before.Position = 0;
        using var restored = PdfReader.Open(before, PdfDocumentOpenMode.Modify);
        Require(new PdfTextEditor(restored).GetRuns(0)[0].Text == "Hello World", "Undo snapshot");

        // Subset-style ToUnicode font: reverse encoding must use original character codes.
        using var mapped = new PdfDocument();
        var mappedPage = mapped.AddPage();
        var mappedFont = new PdfDictionary(mapped);
        mappedFont.Elements.SetName("/Type", "/Font");
        mappedFont.Elements.SetName("/Subtype", "/Type1");
        mappedFont.Elements.SetName("/BaseFont", "/Helvetica");
        var cmap = new PdfDictionary(mapped);
        cmap.CreateStream(Encoding.ASCII.GetBytes("1 beginbfchar <41> <D55C> endbfchar 1 beginbfrange <42> <43> [<AE00> <C218>] endbfrange"));
        mapped.Internals.AddObject(cmap);
        mappedFont.Elements["/ToUnicode"] = cmap.Reference;
        var mappedFonts = new PdfDictionary(mapped);
        mappedFonts.Elements["/F1"] = mappedFont;
        mappedPage.Resources.Elements["/Font"] = mappedFonts;
        mappedPage.Contents.ReplaceContent(PdfSharp.Pdf.Content.ContentReader.ReadContent(Encoding.ASCII.GetBytes("BT /F1 18 Tf 40 700 Td <414243> Tj ET")));
        var mappedEditor = new PdfTextEditor(mapped);
        var mappedRuns = mappedEditor.GetRuns(0);
        Require(mappedRuns[0].Text == "한글수", "ToUnicode decode");
        mappedEditor.Replace(0, mappedRuns[0], "수한글");
        var reRead = new PdfTextEditor(mapped).GetRuns(0);
        Require(reRead[0].Text == "수한글", "ToUnicode encode without font replacement");
        using var positionedBytes = new MemoryStream();
        mapped.Save(positionedBytes, false);
        using var positioned = PdfReader.Open(new MemoryStream(positionedBytes.ToArray()), PdfDocumentOpenMode.Modify);
        var positionedEditor = new PdfTextEditor(positioned);
        using var positionedRendered = PdfiumViewer.PdfDocument.Load(new MemoryStream(positionedBytes.ToArray()));
        var positionedSentences = PdfSentences.Build(positionedRendered, 0, PdfEditRegions.Build(positionedRendered, 0, positionedEditor.GetRuns(0)));
        positionedEditor.ReplaceSentence(0, positionedSentences[0], "한 글");
        using var positionedSaved = new MemoryStream();
        positioned.Save(positionedSaved, false);
        using var positionedReloaded = PdfReader.Open(new MemoryStream(positionedSaved.ToArray()), PdfDocumentOpenMode.Modify);
        Require(new PdfTextEditor(positionedReloaded).GetRuns(0)[0].Text == "한 글", "Spaces without an embedded space glyph survive a save");
        CheckEmbeddedKoreanFont();
        CheckSentences();
        Console.WriteLine("PASS: PDF text editing, Tj/TJ, original font, ToUnicode mapping, unsupported glyph and undo checks.");
    }

    private static void CheckSentences()
    {
        var fixture = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "embedded-korean.pdf"));
        using var source = PdfReader.Open(new MemoryStream(fixture), PdfDocumentOpenMode.Modify);
        var editor = new PdfTextEditor(source);
        using var originalRendered = PdfiumViewer.PdfDocument.Load(new MemoryStream(fixture));
        var originalRegions = PdfEditRegions.Build(originalRendered, 0, editor.GetRuns(0));
        var sentences = PdfSentences.Build(originalRendered, 0, originalRegions);
        Require(sentences.Count == 1, "Korean sentence grouping");
        editor.ReplaceSentence(0, sentences[0], "수정 한글");
        using var saved = new MemoryStream();
        source.Save(saved, false);
        var bytes = saved.ToArray();
        using var reloaded = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Modify);
        var reloadedEditor = new PdfTextEditor(reloaded);
        using var rendered = PdfiumViewer.PdfDocument.Load(new MemoryStream(bytes));
        Require(rendered.GetPdfText(0).Contains("수정 한글"), "Sentence PDF renders edited text");
        var regions = PdfEditRegions.Build(rendered, 0, reloadedEditor.GetRuns(0));
        Require(regions.Count == 1, "Edited sentence hit region remains available");
        using var fontOriginal = PdfReader.Open(new MemoryStream(fixture), PdfDocumentOpenMode.Modify);
        Require(FontFingerprint(reloaded.Pages[0]) == FontFingerprint(fontOriginal.Pages[0]), "Sentence edit preserves embedded font");
        CheckSplitAndWrappedSentences();
    }

    private static void CheckSplitAndWrappedSentences()
    {
        using var document = new PdfDocument();
        var page = document.AddPage();
        var font = new PdfDictionary(document);
        font.Elements.SetName("/Type", "/Font");
        font.Elements.SetName("/Subtype", "/Type1");
        font.Elements.SetName("/BaseFont", "/Helvetica");
        font.Elements.SetName("/Encoding", "/WinAnsiEncoding");
        var fonts = new PdfDictionary(document);
        fonts.Elements["/F1"] = font;
        page.Resources.Elements["/Font"] = fonts;
        string content = "BT /F1 16 Tf 40 700 Td (First ) Tj (sentence.) Tj ( Second sentence.) Tj " +
            "0 -25 Td (This sentence wraps to) Tj 0 -22 Td (the next original line.) Tj " +
            "0 -60 Td (3.14 is a number. Another sentence.) Tj ET";
        page.Contents.ReplaceContent(PdfSharp.Pdf.Content.ContentReader.ReadContent(Encoding.ASCII.GetBytes(content)));
        using var saved = new MemoryStream();
        document.Save(saved, false);
        var fixture = saved.ToArray();
        File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "sentences.pdf"), fixture);
        using var editable = PdfReader.Open(new MemoryStream(fixture), PdfDocumentOpenMode.Modify);
        var editor = new PdfTextEditor(editable);
        using var original = PdfiumViewer.PdfDocument.Load(new MemoryStream(fixture));
        var sentences = PdfSentences.Build(original, 0, PdfEditRegions.Build(original, 0, editor.GetRuns(0)));
        Require(sentences.Count == 5, "Sentence boundaries: " + string.Join(" | ", sentences.Select(sentence => sentence.Text)));
        Require(sentences[0].Text == "First sentence." && sentences[0].Parts.Count == 2, "Split operators join into one sentence");
        Require(sentences[2].Text == "This sentence wraps to the next original line." && sentences[2].GetSegments().Count == 2,
            "Wrapped sentence joins across original lines");
        Require(sentences[3].Text == "3.14 is a number.", "Decimal is not a sentence boundary");
        var oldSecond = original.GetTextBounds(original.Search("Second sentence.", true, false, 0, 0).Items[0].TextSpan)[0].Bounds;
        var oldAnother = original.GetTextBounds(original.Search("Another sentence.", true, false, 0, 0).Items[0].TextSpan)[0].Bounds;
        var before = editable.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue.ToArray();
        bool rejected = false;
        try { editor.ReplaceSentence(0, sentences[0], new string('W', 200)); } catch (InvalidOperationException) { rejected = true; }
        Require(rejected && before.SequenceEqual(editable.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue), "Oversized sentence does not mutate document");
        editor.ReplaceSentence(0, sentences[0], "Short text.");
        editor.ReplaceSentence(0, sentences[2], "This text stays on the original lines.");
        editor.ReplaceSentence(0, sentences[3], "3.14 is fine.");
        using var editedBytes = new MemoryStream();
        editable.Save(editedBytes, false);
        using var rendered = PdfiumViewer.PdfDocument.Load(new MemoryStream(editedBytes.ToArray()));
        string actual = rendered.GetPdfText(0);
        Require(actual.Contains("Short text.") && actual.Contains("Second sentence."), "Sentence edit retains adjacent sentence: " + actual);
        var newSecond = rendered.GetTextBounds(rendered.Search("Second sentence.", true, false, 0, 0).Items[0].TextSpan)[0].Bounds;
        Require(Math.Abs(oldSecond.X - newSecond.X) < 0.1 && Math.Abs(oldSecond.Y - newSecond.Y) < 0.1, "Following sentence position is unchanged");
        var newAnother = rendered.GetTextBounds(rendered.Search("Another sentence.", true, false, 0, 0).Items[0].TextSpan)[0].Bounds;
        Require(actual.Contains("3.14 is fine.") && Math.Abs(oldAnother.X - newAnother.X) < 0.1 && Math.Abs(oldAnother.Y - newAnother.Y) < 0.1,
            "Editing part of one operator preserves its following sentence");
        Require(actual.Replace("\r", "").Replace("\n", " ").Contains("This text stays on the original lines."), "Wrapped sentence writes all text: " + actual);
        using var reopened = PdfReader.Open(new MemoryStream(editedBytes.ToArray()), PdfDocumentOpenMode.Modify);
        var newSentences = PdfSentences.Build(rendered, 0, PdfEditRegions.Build(rendered, 0, new PdfTextEditor(reopened).GetRuns(0)));
        Require(newSentences.Any(sentence => sentence.Text == "This text stays on the original lines."), "Wrapped sentence remains editable after save");
        Console.WriteLine("PASS: sentence grouping, split text, wrapped lines, decimal boundaries, layout preservation and overflow guard.");
    }

    private static void CheckEmbeddedKoreanFont()
    {
        PdfSharp.Fonts.GlobalFontSettings.FontResolver = new WindowsFontResolver();
        using var source = new PdfDocument();
        var page = source.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
            graphics.DrawString("한글 수정 테스트", new XFont("OriginalKorean", 18), XBrushes.Black, new XPoint(40, 100));
        using var bytes = new MemoryStream();
        source.Save(bytes, false);
        File.WriteAllBytes(Path.Combine(AppContext.BaseDirectory, "fixtures", "embedded-korean.pdf"), bytes.ToArray());
        bytes.Position = 0;
        using var editable = PdfReader.Open(bytes, PdfDocumentOpenMode.Modify);
        var editor = new PdfTextEditor(editable);
        var runs = editor.GetRuns(0);
        Require(runs.Count > 0 && runs[0].Text == "한글 수정 테스트", "Embedded Korean text decode");
        string originalFont = FontFingerprint(editable.Pages[0]);
        editor.Replace(0, runs[0], "수정 한글");
        using var saved = new MemoryStream();
        editable.Save(saved, false);
        saved.Position = 0;
        using var loaded = PdfReader.Open(saved, PdfDocumentOpenMode.Modify);
        Require(FontFingerprint(loaded.Pages[0]) == originalFont, "Embedded font bytes preserved");
        saved.Position = 0;
        using var rendered = PdfiumViewer.PdfDocument.Load(saved);
        Require(rendered.GetPdfText(0).Contains("수정 한글"), "Embedded Korean edited text rendering");
        var regions = PdfEditRegions.Build(rendered, 0, new PdfTextEditor(loaded).GetRuns(0));
        Require(regions.Count == 1 && regions[0].Bounds.Width > 20 && regions[0].Bounds.Height > 8,
            "Inline text hit region");
    }

    private static string FontFingerprint(PdfPage page)
    {
        var fonts = page.Resources.Elements.GetDictionary("/Font");
        foreach (var name in fonts.Elements.KeyNames)
        {
            var font = fonts.Elements.GetDictionary(name.Value);
            var descendants = font.Elements.GetArray("/DescendantFonts");
            if (descendants != null)
            {
                var item = descendants.Elements[0];
                font = item is PdfSharp.Pdf.Advanced.PdfReference reference
                    ? (PdfDictionary)reference.Value : (PdfDictionary)item;
            }
            var descriptor = font.Elements.GetDictionary("/FontDescriptor");
            var embedded = descriptor?.Elements.GetDictionary("/FontFile2");
            if (embedded?.Stream != null)
                return Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(embedded.Stream.UnfilteredValue));
        }
        throw new InvalidOperationException("No embedded font in test fixture");
    }

    private sealed class WindowsFontResolver : PdfSharp.Fonts.IFontResolver
    {
        public byte[] GetFont(string faceName) => File.ReadAllBytes(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "malgun.ttf"));
        public PdfSharp.Fonts.FontResolverInfo ResolveTypeface(string familyName, bool bold, bool italic) => new("OriginalKorean");
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
