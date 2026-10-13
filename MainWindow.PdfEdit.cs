using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace PDFProcessor
{
    public partial class MainWindow
    {
        private PdfTextEditor pdfEditor;
        private string editSourcePath;
        private int editPageIndex;
        private byte[] editCurrentBytes;
        private readonly Stack<byte[]> editUndo = new();
        private bool editBusy;
        private Task editPendingOperation = Task.CompletedTask;

        private double EditZoomValue => double.Parse((EditZoom.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "1", CultureInfo.InvariantCulture);

        private async void OpenEditPdf_Click(object sender, RoutedEventArgs e)
        {
            if (editBusy || !await ConfirmEditDiscard()) return;
            var dialog = new OpenFileDialog { Filter = "PDF 파일|*.pdf", Title = "수정할 PDF 선택" };
            if (dialog.ShowDialog() != true) return;
            double zoom = EditZoomValue;
            await PrepareEditAsync(() => PdfEditPreparation.Build(File.ReadAllBytes(dialog.FileName), 0, zoom), () =>
            {
                editSourcePath = dialog.FileName;
                editPageIndex = 0;
                editUndo.Clear();
                editRedo.Clear();
                editDirty = false;
                editSavedPath = null;
                EditFileLabel.Text = Path.GetFileName(dialog.FileName);
            });
        }

        private Task<bool> PrepareEditAsync(Func<PdfEditPreparation> work, Action success = null)
        {
            var task = PrepareEditCoreAsync(work, success);
            editPendingOperation = task;
            return task;
        }

        private async Task<bool> PrepareEditCoreAsync(Func<PdfEditPreparation> work, Action success)
        {
            if (editBusy) return false;
            SetEditBusy(true);
            try
            {
                var prepared = await Task.Run(work);
                success?.Invoke();
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
                    EditFontLabel.Text = "원본 글꼴 유지";
                    EditPreviewImage.Source = prepared.Image;
                    EditPreviewImage.Width = EditPageCanvas.Width = EditHitLayer.Width = prepared.Width * editDisplayScale;
                    EditPreviewImage.Height = EditPageCanvas.Height = EditHitLayer.Height = prepared.Height * editDisplayScale;
                    editRegions = prepared.Regions;
                    editSentences = prepared.Sentences;
                    EditTextList.ItemsSource = editSentences;
                    RenderEditRegions();
                    EditPageLabel.Text = $"페이지: {editPageIndex + 1} / {pdfEditor.Document.PageCount}";
                    EditUndoButton.IsEnabled = editUndo.Count > 0;
                    EditRedoButton.IsEnabled = editRedo.Count > 0;
                    EditStatusLabel.Text = prepared.Status;
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

        private void SetEditBusy(bool busy)
        {
            editBusy = busy;
            MainTabControl.IsEnabled = !busy;
            EditProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            if (busy) EditStatusLabel.Text = "문서 처리 중… 잠시 기다려주세요.";
        }

        private Task<bool> ShowEditPage()
        {
            if (pdfEditor == null || editBusy) return Task.FromResult(false);
            int page = editPageIndex;
            double zoom = EditZoomValue;
            var bytes = editCurrentBytes;
            var editor = pdfEditor;
            return PrepareEditAsync(() =>
            {
                if (bytes == null)
                {
                    using var stream = new MemoryStream();
                    editor.Document.Save(stream, false);
                    bytes = stream.ToArray();
                }
                return PdfEditPreparation.Build(bytes, page, zoom);
            });
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
            byte[] target = editUndo.Peek(), before = editCurrentBytes;
            int page = editPageIndex;
            double zoom = EditZoomValue;
            await PrepareEditAsync(() => PdfEditPreparation.Build(target, page, zoom), () =>
            {
                editUndo.Pop();
                editRedo.Push(before);
                editDirty = true;
            });
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

        private async Task ChangeEditPage(int direction)
        {
            if (editBusy || pdfEditor == null || !await CommitInlineEdit()) return;
            int target = editPageIndex + direction;
            if (target < 0 || target >= pdfEditor.Document.PageCount) return;
            byte[] bytes = editCurrentBytes;
            double zoom = EditZoomValue;
            await PrepareEditAsync(() => PdfEditPreparation.Build(bytes, target, zoom), () => editPageIndex = target);
        }
    }
}