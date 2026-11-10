using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PDFProcessor
{
    public partial class MainWindow
    {
        /// <summary>Document state restored by undo/redo: bytes plus OCR lines already rewritten.</summary>
        private sealed record EditSnapshot(byte[] Bytes, HashSet<string> HiddenOcr);

        private PdfTextEditor pdfEditor;
        private string editSourcePath;
        private int editPageIndex;
        private byte[] editCurrentBytes;
        private readonly Stack<EditSnapshot> editUndo = new();
        private bool editBusy;
        private Task editPendingOperation = Task.CompletedTask;
        private readonly Dictionary<int, List<OcrLine>> editOcrPages = new();
        private HashSet<string> editOcrHidden = new();

        private double EditZoomValue => double.Parse((EditZoom.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "1", CultureInfo.InvariantCulture);
        private bool EditAutoOcrEnabled => EditAutoOcr?.IsChecked == true;

        private async void OpenEditPdf_Click(object sender, RoutedEventArgs e)
        {
            if (editBusy || !await ConfirmEditDiscard()) return;
            var dialog = new OpenFileDialog { Filter = "PDF 파일|*.pdf", Title = "수정할 PDF 선택" };
            if (dialog.ShowDialog() != true) return;
            await OpenEditDocument(dialog.FileName);
        }

        private async Task<bool> OpenEditDocument(string path)
        {
            double zoom = EditZoomValue;
            var mode = EditAutoOcrEnabled ? OcrMode.IfNoText : OcrMode.None;
            byte[] bytes;
            try { bytes = File.ReadAllBytes(path); }
            catch (Exception ex) { EditStatusLabel.Text = ex.Message; return false; }
            return await PrepareEditAsync(() => PdfEditPreparation.Build(bytes, 0, zoom, ocrMode: mode), () =>
            {
                editSourcePath = path;
                editPageIndex = 0;
                editUndo.Clear();
                editRedo.Clear();
                editOcrPages.Clear();
                editOcrHidden = new HashSet<string>();
                editDirty = false;
                editSavedPath = null;
                EditFileLabel.Text = Path.GetFileName(path);
            }, mode != OcrMode.None);
        }

        private Task<bool> PrepareEditAsync(Func<PdfEditPreparation> work, Action success = null, bool mayRecognize = false)
        {
            var task = PrepareEditCoreAsync(work, success, mayRecognize);
            editPendingOperation = task;
            return task;
        }

        private async Task<bool> PrepareEditCoreAsync(Func<PdfEditPreparation> work, Action success, bool mayRecognize)
        {
            if (editBusy) return false;
            SetEditBusy(true, mayRecognize ? "문서 처리 중… 텍스트가 없는 페이지는 OCR로 인식합니다." : null);
            try
            {
                var prepared = await Task.Run(work);
                success?.Invoke();
                if (prepared.OcrLines != null) editOcrPages[prepared.Page] = prepared.OcrLines;
                editRefreshing = true;
                try
                {
                    CancelInlineEdit();
                    var oldEditor = pdfEditor;
                    pdfEditor = prepared.Editor;
                    editCurrentBytes = prepared.Bytes;
                    oldEditor?.Document.Dispose();
                    editDisplayScale = 96.0 / 72 * EditZoomValue;
                    EditTextList.ItemsSource = null;
                    EditReplacement.Text = "";
                    EditFontLabel.Text = prepared.Sentences.Any(sentence => sentence.IsOcr)
                        ? "스캔 페이지 · OCR 텍스트는 맑은 고딕으로 다시 씀"
                        : "원본 글꼴 유지";
                    EditPreviewImage.Source = prepared.Image;
                    EditPreviewImage.Width = EditPageCanvas.Width = EditHitLayer.Width = prepared.Width * editDisplayScale;
                    EditPreviewImage.Height = EditPageCanvas.Height = EditHitLayer.Height = prepared.Height * editDisplayScale;
                    editPageSize = new System.Drawing.SizeF(prepared.Width, prepared.Height);
                    editRegions = prepared.Regions;
                    editSentences = prepared.Sentences;
                    EditTextList.ItemsSource = editSentences;
                    RenderEditRegions();
                    EditPageLabel.Text = $"{editPageIndex + 1} / {pdfEditor.Document.PageCount}";
                    EditUndoButton.IsEnabled = editUndo.Count > 0;
                    EditRedoButton.IsEnabled = editRedo.Count > 0;
                    EditStatusLabel.Text = prepared.Status;
                    lastEditNotice = prepared.EditNotice;
                    EditEmptyState.Visibility = Visibility.Collapsed;
                }
                finally { editRefreshing = false; }
                return true;
            }
            catch (Exception ex)
            {
                // Keep the active PDF, pending input and history when a background operation fails.
                EditStatusLabel.Text = ex.Message.Replace('\n', ' ');
                return false;
            }
            finally { SetEditBusy(false); }
        }

        private string lastEditNotice;
        private System.Drawing.SizeF editPageSize;

        private void SetEditBusy(bool busy, string message = null)
        {
            editBusy = busy;
            MainTabControl.IsEnabled = !busy;
            EditProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            if (busy) EditStatusLabel.Text = message ?? "문서 처리 중… 잠시 기다려주세요.";
        }

        /// <summary>Builds the current page from the given bytes with the cached OCR state of that page.</summary>
        private Func<PdfEditPreparation> PageBuilder(byte[] bytes, int page, HashSet<string> hidden,
            PdfEditRequest request = null, OcrMode? mode = null)
        {
            double zoom = EditZoomValue;
            editOcrPages.TryGetValue(page, out var cached);
            var ocrMode = mode ?? (cached == null && EditAutoOcrEnabled ? OcrMode.IfNoText : OcrMode.None);
            if (ocrMode == OcrMode.Force) cached = null;
            var hiddenCopy = new HashSet<string>(hidden);
            return () => PdfEditPreparation.Build(bytes, page, zoom, request, cached, hiddenCopy, ocrMode);
        }

        private Task<bool> ShowEditPage()
        {
            if (pdfEditor == null || editBusy) return Task.FromResult(false);
            int page = editPageIndex;
            var bytes = editCurrentBytes;
            if (bytes == null)
            {
                using var stream = new MemoryStream();
                pdfEditor.Document.Save(stream, false);
                bytes = stream.ToArray();
            }
            bool recognize = !editOcrPages.ContainsKey(page) && EditAutoOcrEnabled;
            return PrepareEditAsync(PageBuilder(bytes, page, editOcrHidden), null, recognize);
        }

        private async void EditTextList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (editBusy || editRefreshing || editSelecting || EditTextList.SelectedItem is not PdfSentence sentence) return;
            await BeginInlineEdit(sentence);
        }

        private async void ApplyPdfEdit_Click(object sender, RoutedEventArgs e) => await CommitInlineEdit();

        private async void UndoPdfEdit_Click(object sender, RoutedEventArgs e)
        {
            if (editBusy || !await CommitInlineEdit() || editUndo.Count == 0) return;
            var target = editUndo.Peek();
            var current = new EditSnapshot(editCurrentBytes, editOcrHidden);
            await PrepareEditAsync(PageBuilder(target.Bytes, editPageIndex, target.HiddenOcr, mode: OcrMode.None), () =>
            {
                editUndo.Pop();
                editRedo.Push(current);
                editOcrHidden = target.HiddenOcr;
                editDirty = true;
            });
        }

        private async void RunOcr_Click(object sender, RoutedEventArgs e)
        {
            if (editBusy) return;
            if (pdfEditor == null) { EditStatusLabel.Text = "먼저 PDF를 열어주세요."; return; }
            if (!await CommitInlineEdit()) return;
            string reason = await Task.Run(PdfOcr.UnavailableReason);
            if (reason != null)
            {
                EditStatusLabel.Text = reason;
                MessageBox.Show(this, reason, "OCR 사용 불가", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            int page = editPageIndex;
            await PrepareEditAsync(PageBuilder(editCurrentBytes, page, editOcrHidden, mode: OcrMode.Force), null, true);
        }

        private void CopyEditText_Click(object sender, RoutedEventArgs e)
        {
            if (editSentences.Count == 0) { EditStatusLabel.Text = "복사할 텍스트가 없습니다."; return; }
            try
            {
                Clipboard.SetText(string.Join(Environment.NewLine, editSentences.Select(sentence => sentence.Text)));
                EditStatusLabel.Text = $"페이지 텍스트 {editSentences.Count}줄을 클립보드에 복사했습니다.";
            }
            catch (Exception ex) { EditStatusLabel.Text = "클립보드 복사 실패: " + ex.Message; }
        }

        private async void SaveEditedPdf_Click(object sender, RoutedEventArgs e) => await SaveEditedPdf();

        private async Task<bool> SaveEditedPdf()
        {
            if (pdfEditor == null || editBusy || !await CommitInlineEdit()) return false;
            string target = editSavedPath;
            if (target == null)
            {
                var dialog = new SaveFileDialog { Filter = "PDF 파일|*.pdf", DefaultExt = ".pdf",
                    FileName = Path.GetFileNameWithoutExtension(editSourcePath) + "_수정.pdf" };
                if (dialog.ShowDialog() != true) return false;
                target = dialog.FileName;
            }
            SetEditBusy(true);
            try
            {
                if (string.Equals(Path.GetFullPath(target), Path.GetFullPath(editSourcePath), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("원본 파일을 보존하도록 다른 이름으로 저장하세요.");
                var bytes = editCurrentBytes;
                await Task.Run(() =>
                {
                    string temporary = Path.Combine(Path.GetDirectoryName(target), $".edit-{Guid.NewGuid():N}.tmp");
                    try { File.WriteAllBytes(temporary, bytes); File.Move(temporary, target, true); }
                    finally { if (File.Exists(temporary)) File.Delete(temporary); }
                });
                editSavedPath = target;
                editDirty = false;
                EditFileLabel.Text = Path.GetFileName(target);
                EditStatusLabel.Text = $"저장 완료: {target}";
                return true;
            }
            catch (Exception ex) { EditStatusLabel.Text = ex.Message; return false; }
            finally { SetEditBusy(false); }
        }

        private async void EditPrevPage_Click(object sender, RoutedEventArgs e) => await ChangeEditPage(-1);
        private async void EditNextPage_Click(object sender, RoutedEventArgs e) => await ChangeEditPage(1);

        private async Task<bool> ChangeEditPage(int direction)
        {
            if (editBusy || pdfEditor == null || !await CommitInlineEdit()) return false;
            int target = editPageIndex + direction;
            if (target < 0 || target >= pdfEditor.Document.PageCount) return false;
            bool recognize = !editOcrPages.ContainsKey(target) && EditAutoOcrEnabled;
            return await PrepareEditAsync(PageBuilder(editCurrentBytes, target, editOcrHidden), () => editPageIndex = target, recognize);
        }

        // Number of wheel notches received while already at the top/bottom edge (negative = top).
        private int editWheelEdgeHits;

        /// <summary>
        /// Mouse wheel: scrolls the page, and continues to the next/previous page at the bottom/top edge.
        /// Ctrl+wheel changes the zoom.
        /// </summary>
        private async void EditPreview_MouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (pdfEditor == null) return;
            if (System.Windows.Input.Keyboard.Modifiers == System.Windows.Input.ModifierKeys.Control)
            {
                e.Handled = true;
                if (editBusy) return;
                int zoom = Math.Clamp(EditZoom.SelectedIndex + (e.Delta > 0 ? 1 : -1), 0, EditZoom.Items.Count - 1);
                if (zoom != EditZoom.SelectedIndex) EditZoom.SelectedIndex = zoom;
                return;
            }
            if (editBusy) { e.Handled = true; return; }
            var scroll = EditPreviewScroll;
            bool fits = scroll.ScrollableHeight < 1;
            bool down = e.Delta < 0;
            bool atEdge = down ? scroll.VerticalOffset >= scroll.ScrollableHeight - 1 : scroll.VerticalOffset <= 1;
            bool hasPage = down ? editPageIndex < pdfEditor.Document.PageCount - 1 : editPageIndex > 0;
            if (!atEdge || !hasPage) { editWheelEdgeHits = 0; return; }
            e.Handled = true;
            // A scrolled page needs one extra notch at the edge so a fast flick does not skip pages.
            editWheelEdgeHits = Math.Sign(editWheelEdgeHits) == (down ? 1 : -1) ? editWheelEdgeHits + (down ? 1 : -1) : (down ? 1 : -1);
            if (!fits && Math.Abs(editWheelEdgeHits) < 2) return;
            editWheelEdgeHits = 0;
            if (!await ChangeEditPage(down ? 1 : -1)) return;
            scroll.UpdateLayout();
            if (down) scroll.ScrollToTop(); else scroll.ScrollToBottom();
        }

        private void EditDropZone_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void EditDropZone_Drop(object sender, DragEventArgs e)
        {
            if (editBusy || e.Data.GetData(DataFormats.FileDrop) is not string[] files || files.Length == 0) return;
            if (!files[0].EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                EditStatusLabel.Text = "PDF 파일만 열 수 있습니다.";
                return;
            }
            if (!await ConfirmEditDiscard()) return;
            await OpenEditDocument(files[0]);
        }
    }
}
