using System.Reflection;
using System.IO;
using System.Runtime.Loader;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using System.Threading;
using System.Threading.Tasks;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        SynchronizationContext.SetSynchronizationContext(new DispatcherSynchronizationContext());
        string appDirectory = Path.GetFullPath(args[0]);
        AssemblyLoadContext.Default.Resolving += (context, name) =>
        {
            string path = Path.Combine(appDirectory, name.Name + ".dll");
            return File.Exists(path) ? context.LoadFromAssemblyPath(path) : null;
        };
        var assembly = Assembly.LoadFrom(Path.Combine(appDirectory, "PDF통합처리기.dll"));
        var appType = assembly.GetType("PDFProcessor.App");
        var app = (Application)Activator.CreateInstance(appType);
        appType.GetMethod("InitializeComponent").Invoke(app, null);
        var type = assembly.GetType("PDFProcessor.MainWindow");
        var window = (Window)Activator.CreateInstance(type);
        window.WindowState = WindowState.Normal;
        window.Width = 1200;
        window.Height = 800;
        window.Left = -10000;
        window.Top = -10000;
        window.ShowInTaskbar = false;
        window.Show();
        var tabs = (TabControl)window.FindName("MainTabControl");
        tabs.SelectedIndex = 1;
        var files = (List<string>)type.GetField("mergeFiles", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
        files.Add(Path.GetFullPath(args[1]));
        var list = (ListBox)window.FindName("MergeFileList");
        list.Items.Add(Path.GetFileName(args[1]));
        type.GetMethod("LoadAllMergePages", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, null);
        window.UpdateLayout();
        var preview = (Image)window.FindName("MergePreviewImage");
        Require(preview.Source != null && preview.ActualHeight > 100, $"Preview size: {preview.ActualWidth} x {preview.ActualHeight}");
        var content = (FrameworkElement)window.Content;
        foreach (var button in Descendants(content).OfType<Button>().Where(button =>
            new[] { "추가", "위로", "아래로", "제거", "전체 삭제" }.Contains(button.Content?.ToString())))
        {
            var point = button.TransformToAncestor(list.Parent as Visual).Transform(new Point(0, 0));
            var listBottom = list.TransformToAncestor(list.Parent as Visual).Transform(new Point(0, list.ActualHeight)).Y;
            Require(point.Y >= listBottom, $"Overlapping button: {button.Content}");
            Require(button.ActualWidth >= 45 && button.ActualHeight >= 30, $"Clipped button: {button.Content}");
        }
        Capture(content, Path.Combine(appDirectory, "merge-ui.png"));
        tabs.SelectedIndex = 3;
        window.UpdateLayout();
        Require(((TabItem)tabs.Items[3]).Header.ToString() == "PDF 수정", "Edit tab");
        if (args.Length > 2)
        {
            var pdfAssembly = Assembly.Load("PdfSharp");
            var modeType = pdfAssembly.GetType("PdfSharp.Pdf.IO.PdfDocumentOpenMode");
            var reader = pdfAssembly.GetType("PdfSharp.Pdf.IO.PdfReader");
            var open = reader.GetMethods().First(method => method.Name == "Open" &&
                method.GetParameters().Length >= 2 && method.GetParameters()[0].ParameterType == typeof(string) &&
                method.GetParameters()[1].ParameterType == modeType);
            var openArgs = open.GetParameters().Select(parameter => parameter.HasDefaultValue ? parameter.DefaultValue : null).ToArray();
            openArgs[0] = Path.GetFullPath(args[2]);
            openArgs[1] = Enum.Parse(modeType, "Modify");
            var document = open.Invoke(null, openArgs);
            var editorType = assembly.GetType("PDFProcessor.PdfTextEditor");
            var editor = Activator.CreateInstance(editorType, BindingFlags.Instance | BindingFlags.NonPublic,
                null, new[] { document }, null);
            type.GetField("pdfEditor", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(window, editor);
            InvokeEdit(type, window, "ShowEditPage");
            var textList = (ListBox)window.FindName("EditTextList");
            Require(textList.Items.Count > 0, "Korean editable text list");
            if (args.Length > 3 && args[3] == "--affected")
            {
                CheckAffectedFile(type, window, textList);
                goto EditChecksComplete;
            }
            bool sentenceFixture = args.Length > 3 && args[3] == "--sentences";
            int sentenceIndex = sentenceFixture ? 2 : 0;
            string initialText = sentenceFixture ? "This sentence wraps to the next original line." : "한글 수정 테스트";
            string editedText = sentenceFixture ? "This text stays on the original lines." : "수정 한글";
            double originalSize = sentenceFixture ? 16 : 18;
            textList.SelectedIndex = sentenceIndex;
            Require(((TextBox)window.FindName("EditReplacement")).Visibility == Visibility.Visible, "Inline editing is visible on the page");
            Require(((Canvas)window.FindName("EditHitLayer")).Children.Count > 0, "Clickable text regions");
            Require(((TextBox)window.FindName("EditReplacement")).Text == initialText, "Entire sentence is selected");
            Require(Math.Abs(((TextBox)window.FindName("EditReplacement")).FontSize - originalSize * 96 / 72) < 0.1, "Original font at 100% zoom");
            ((TextBox)window.FindName("EditReplacement")).Text = editedText;
            InvokeEdit(type, window, "ApplyPdfEdit_Click", new object[] { null, new RoutedEventArgs() });
            Require(((TextBlock)window.FindName("EditStatusLabel")).Text.Contains("수정됨"), "Edit applied in UI");
            Require(((Image)window.FindName("EditPreviewImage")).Source != null, "Edited preview");
            InvokeEdit(type, window, "UndoPdfEdit_Click", new object[] { null, new RoutedEventArgs() });
            textList.SelectedIndex = sentenceIndex;
            Require(((TextBox)window.FindName("EditReplacement")).Text == initialText, "Undo in UI");
            InvokeEdit(type, window, "RedoPdfEdit_Click", new object[] { null, new RoutedEventArgs() });
            textList.SelectedIndex = sentenceIndex;
            Require(((TextBox)window.FindName("EditReplacement")).Text == editedText, "Redo in UI");
            ((TextBox)window.FindName("EditReplacement")).Text = "아직 저장하지 않은 입력";
            var cancelKey = new System.Windows.Input.KeyEventArgs(System.Windows.Input.Keyboard.PrimaryDevice,
                PresentationSource.FromVisual(window), 0, System.Windows.Input.Key.Escape)
                { RoutedEvent = System.Windows.Input.Keyboard.PreviewKeyDownEvent };
            type.GetMethod("EditInline_KeyDown", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, new object[] { null, cancelKey });
            Require(((TextBox)window.FindName("EditReplacement")).Visibility == Visibility.Collapsed, "Escape cancels pending input");
            textList.SelectedIndex = -1;
            textList.SelectedIndex = sentenceIndex;
            Require(((TextBox)window.FindName("EditReplacement")).Text == editedText, "Cancelled input not committed");
            var zoom = (ComboBox)window.FindName("EditZoom");
            zoom.SelectedIndex = 3;
            WaitForEdit(type, window);
            window.UpdateLayout();
            Require(((Canvas)window.FindName("EditPageCanvas")).Width > 1100, "150% zoom enlarges page");
            textList.SelectedIndex = sentenceIndex;
            Require(Math.Abs(((TextBox)window.FindName("EditReplacement")).FontSize - originalSize * 2) < 0.1, "Font size tracks zoom");
            window.UpdateLayout();
            EditChecksComplete: window.UpdateLayout();
        }
        Capture(content, Path.Combine(appDirectory, "edit-ui.png"));
        type.GetMethod("CancelInlineEdit", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, null);
        type.GetField("editDirty", BindingFlags.Instance | BindingFlags.NonPublic).SetValue(window, false);
        window.Close();
        Console.WriteLine("PASS: WPF merge button layout, visible preview and PDF edit tab; screenshots saved.");
    }

    private static void CheckAffectedFile(Type type, Window window, ListBox list)
    {
        int index = Enumerable.Range(0, list.Items.Count).First(value =>
            ((string)list.Items[value].GetType().GetProperty("Text").GetValue(list.Items[value])).Length > 4);
        list.SelectedIndex = index;
        WaitForEdit(type, window);
        var textbox = (TextBox)window.FindName("EditReplacement");
        string original = textbox.Text;
        textbox.Text = original.Substring(1);
        var canvas = (Canvas)window.FindName("EditPageCanvas");
        canvas.Focusable = true;
        canvas.Focus(); // Reproduce LostKeyboardFocus followed by a click on blank page space.
        var click = new System.Windows.Input.MouseButtonEventArgs(System.Windows.Input.Mouse.PrimaryDevice, 0, System.Windows.Input.MouseButton.Left)
            { RoutedEvent = UIElement.MouseLeftButtonDownEvent, Source = canvas };
        int beats = InvokeEdit(type, window, "EditPageCanvas_Click", new object[] { canvas, click });
        Require(((TextBlock)window.FindName("EditStatusLabel")).Text.Contains("수정됨"), "Blank-space click applies the edit: " + ((TextBlock)window.FindName("EditStatusLabel")).Text);
        Require(beats > 0, "UI dispatcher remains responsive while applying an edit");
        var undo = (System.Collections.ICollection)type.GetField("editUndo", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
        Require(undo.Count == 1, "Focus loss and blank click commit only once");
        list.SelectedIndex = index;
        WaitForEdit(type, window);
        textbox.Text += "🧪";
        var before = (byte[])type.GetField("editCurrentBytes", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
        textbox.Focus();
        canvas.Focus();
        InvokeEdit(type, window, "EditPageCanvas_Click", new object[] { canvas, click });
        var after = (byte[])type.GetField("editCurrentBytes", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
        Require(before.SequenceEqual(after) && undo.Count == 1, "Invalid pending input preserves PDF and undo history");
        Require(textbox.Visibility == Visibility.Visible, "Rejected input remains available to correct");
        Require(!(bool)type.GetField("editBusy", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window), "Blank click does not leave the UI busy");
        Console.WriteLine($"PASS: affected PDF blank click, focus loss, single commit, failed-input recovery; {beats} UI heartbeats.");
    }

    private static int InvokeEdit(Type type, Window window, string method, object[] args = null)
    {
        var result = type.GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(window, args);
        int beats = WaitForEdit(type, window);
        if (result is Task task) task.GetAwaiter().GetResult();
        return beats;
    }

    private static int WaitForEdit(Type type, Window window)
    {
        var frame = new DispatcherFrame();
        int heartbeats = 0, idleTicks = 0;
        var watch = System.Diagnostics.Stopwatch.StartNew();
        var timer = new DispatcherTimer(DispatcherPriority.Normal) { Interval = TimeSpan.FromMilliseconds(2) };
        bool timedOut = false;
        timer.Tick += (sender, args) =>
        {
            bool busy = (bool)type.GetField("editBusy", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
            bool committing = (bool)type.GetField("editCommitting", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
            var pending = (Task)type.GetField("editPendingOperation", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(window);
            if (busy) heartbeats++;
            if (!busy && !committing && pending.IsCompleted) idleTicks++; else idleTicks = 0;
            timedOut = watch.Elapsed > TimeSpan.FromSeconds(20);
            if (idleTicks >= 2 || timedOut) frame.Continue = false;
        };
        timer.Start();
        Dispatcher.PushFrame(frame);
        timer.Stop();
        Require(!timedOut, "PDF edit operation did not complete within 20 seconds");
        return heartbeats;
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject parent)
    {
        for (int i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
        {
            var child = VisualTreeHelper.GetChild(parent, i);
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }

    private static void Capture(FrameworkElement content, string path)
    {
        var bitmap = new RenderTargetBitmap((int)content.ActualWidth, (int)content.ActualHeight, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(content);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using var file = File.Create(path);
        encoder.Save(file);
    }

    private static void Require(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }
}
