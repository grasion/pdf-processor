using PDFProcessor;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using Document = PdfSharp.Pdf.PdfDocument;

string fixtures = Path.Combine(AppContext.BaseDirectory, "fixtures");
Directory.CreateDirectory(fixtures);
if (args.Contains("--edit-only")) { EditChecks.Run(); return; }
var papers = new Dictionary<string, (double W, double H)>
{
    ["A4"] = (595.2756, 841.8898), ["A3"] = (841.8898, 1190.5512),
    ["A5"] = (419.5276, 595.2756), ["Letter"] = (612, 792), ["Legal"] = (612, 1008)
};
int checks = 0;
foreach (int native in new[] { 0, 90, 180, 270 })
foreach (bool cropped in new[] { false, true })
{
    string path = Path.Combine(fixtures, $"source-{native}{(cropped ? "-crop" : "")}.pdf");
    using (var source = new Document())
    {
        var page = source.AddPage();
        page.Width = XUnit.FromPoint(cropped ? 300 : 240);
        page.Height = XUnit.FromPoint(cropped ? 420 : 360);
        int inset = cropped ? 30 : 0;
        if (cropped) page.CropBox = new PdfSharp.Pdf.PdfRectangle(new XRect(30, 30, 240, 360));
        using (var graphics = XGraphics.FromPdfPage(page))
        {
            graphics.DrawRectangle(XBrushes.Red, 10 + inset, 10 + inset, 30, 30);
            graphics.DrawRectangle(XBrushes.Blue, 200 + inset, 320 + inset, 30, 30);
        }
        page.Rotate = native;
        source.Save(path);
    }
    foreach (string paper in papers.Keys.Concat(new[] { "원본" }))
    foreach (string orientation in new[] { "세로", "가로", "원본" })
    foreach (int rotation in new[] { 0, 90, 180, 270 })
        Check(path, paper, orientation, native, rotation);
}
string imagePath = Path.Combine(fixtures, "source.png");
using (var image = new System.Drawing.Bitmap(240, 360))
{
    image.SetResolution(72, 72);
    using var graphics = System.Drawing.Graphics.FromImage(image);
    graphics.Clear(System.Drawing.Color.White);
    graphics.FillRectangle(System.Drawing.Brushes.Red, 10, 10, 30, 30);
    graphics.FillRectangle(System.Drawing.Brushes.Blue, 200, 320, 30, 30);
    image.Save(imagePath);
}
foreach (string paper in papers.Keys.Concat(new[] { "원본" }))
foreach (string orientation in new[] { "세로", "가로", "원본" })
foreach (int rotation in new[] { 0, 90, 180, 270 })
    Check(imagePath, paper, orientation, 0, rotation);

// Exercise mixed source order and more than one page in the same output.
using (var mixed = new Document())
{
    MergePageLayout.AddPage(mixed, Path.Combine(fixtures, "source-90.pdf"), 0, 90, "A3", "가로");
    MergePageLayout.AddPage(mixed, imagePath, 0, 270, "A3", "가로");
    using var stream = new MemoryStream();
    mixed.Save(stream, false);
    stream.Position = 0;
    using var rendered = PdfiumViewer.PdfDocument.Load(stream);
    Assert(rendered.PageCount == 2, "Mixed page count");
}
Console.WriteLine($"PASS: {checks} PDF/image paper, orientation and rotation combinations; mixed merge.");
EditChecks.Run();

void Check(string path, string paper, string orientation, int native, int rotation)
{
    int total = (native + rotation) % 360;
    bool sideways = total == 90 || total == 270;
    double visibleW = sideways ? 360 : 240;
    double visibleH = sideways ? 240 : 360;
    var expected = paper == "원본" ? (W: visibleW, H: visibleH) : papers[paper];
    bool landscape = orientation == "가로" || (orientation == "원본" && visibleW > visibleH);
    expected = landscape ? (Math.Max(expected.W, expected.H), Math.Min(expected.W, expected.H))
        : (Math.Min(expected.W, expected.H), Math.Max(expected.W, expected.H));

    using var output = new Document();
    MergePageLayout.AddPage(output, path, 0, rotation, paper, orientation);
    using var stream = new MemoryStream();
    output.Save(stream, false);
    stream.Position = 0;
    using var rendered = PdfiumViewer.PdfDocument.Load(stream);
    var size = rendered.PageSizes[0];
    string label = $"{Path.GetFileName(path)} {paper}/{orientation} +{rotation}";
    Assert(Math.Abs(size.Width - expected.W) < 0.1 && Math.Abs(size.Height - expected.H) < 0.1,
        $"{label}: size {size.Width}x{size.Height}, expected {expected}");
    double renderScale = 400.0 / Math.Max(size.Width, size.Height);
    using var bitmap = new System.Drawing.Bitmap(rendered.Render(0,
        (int)(size.Width * renderScale), (int)(size.Height * renderScale), 72, 72, false));
    double redX = 0, redY = 0, blueX = 0, blueY = 0;
    int redCount = 0, blueCount = 0;
    int minX = 10000, minY = 10000, maxX = 0, maxY = 0;
    for (int y = 0; y < bitmap.Height; y++)
    for (int x = 0; x < bitmap.Width; x++)
    {
        var color = bitmap.GetPixel(x, y);
        if (color.R > 180 && color.G < 80 && color.B < 80)
        {
            redX += x; redY += y; redCount++;
            minX = Math.Min(minX, x); minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x); maxY = Math.Max(maxY, y);
        }
        if (color.B > 180 && color.G < 80 && color.R < 80) { blueX += x; blueY += y; blueCount++; }
    }
    Assert(redCount > 20 && blueCount > 20, $"{label}: clipped/missing content");
    bool redLeft = redX / redCount < blueX / blueCount;
    bool redAbove = redY / redCount < blueY / blueCount;
    Assert(redLeft == (total == 0 || total == 270) && redAbove == (total == 0 || total == 90),
        $"{label}: incorrect content rotation ({redLeft}, {redAbove})");
    Assert(Math.Abs((maxX - minX) - (maxY - minY)) <= 2, $"{label}: distorted aspect ratio");
    checks++;
}

static void Assert(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException(message);
}
