using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace PDFProcessor
{
    public partial class MainWindow
    {
        private List<PdfEditRegion> editRegions = new();
        private List<PdfSentence> editSentences = new();
        private PdfSentence editActiveSentence;
        private double editDisplayScale = 96.0 / 72;
        private bool editRefreshing;
        private bool editCommitting;
        private bool editSelecting;
        private bool editComposing;
        private bool editDirty;
        private string editSavedPath;
        private readonly Stack<EditSnapshot> editRedo = new();
        private bool editCloseApproved;

        private void InitializeInlineEditing()
        {
            TextCompositionManager.AddPreviewTextInputStartHandler(EditReplacement, (sender, args) => editComposing = true);
            TextCompositionManager.AddPreviewTextInputHandler(EditReplacement, (sender, args) => editComposing = false);
            Closing += EditWindow_Closing;
        }

        private void RenderEditRegions()
        {
            EditHitLayer.Children.Clear();
            foreach (var sentence in editSentences)
            foreach (var bounds in sentence.LineBounds)
            {
                // OCR lines are tinted so users can see which text was recognised from the image.
                var accent = sentence.IsOcr ? Color.FromRgb(16, 185, 129) : Color.FromRgb(37, 99, 235);
                Brush idle = sentence.IsOcr ? new SolidColorBrush(Color.FromArgb(28, accent.R, accent.G, accent.B)) : Brushes.Transparent;
                Brush idleBorder = sentence.IsOcr ? new SolidColorBrush(Color.FromArgb(90, accent.R, accent.G, accent.B)) : Brushes.Transparent;
                var hover = new SolidColorBrush(Color.FromArgb(36, accent.R, accent.G, accent.B));
                var border = new Border
                {
                    Width = Math.Max(8, bounds.Width * editDisplayScale + 6),
                    Height = Math.Max(12, bounds.Height * editDisplayScale + 6),
                    Background = idle,
                    BorderBrush = idleBorder,
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(3),
                    Cursor = Cursors.IBeam,
                    ToolTip = sentence.IsOcr ? "OCR 인식 텍스트 · 클릭하여 수정" : "클릭하여 문장 수정",
                    Tag = sentence
                };
                Canvas.SetLeft(border, bounds.X * editDisplayScale - 3);
                Canvas.SetTop(border, bounds.Y * editDisplayScale - 3);
                border.MouseEnter += (sender, args) => { border.BorderBrush = new SolidColorBrush(accent); border.Background = hover; };
                border.MouseLeave += (sender, args) => { border.BorderBrush = idleBorder; border.Background = idle; };
                border.MouseLeftButtonDown += async (sender, args) =>
                {
                    args.Handled = true;
                    await BeginInlineEdit(sentence, args.GetPosition(EditPageCanvas));
                };
                EditHitLayer.Children.Add(border);
            }
        }

        private async Task BeginInlineEdit(PdfSentence run, Point? clicked = null)
        {
            if (editBusy) return;
            if (editActiveSentence == run && EditReplacement.Visibility == Visibility.Visible) return;
            var anchor = run.Bounds;
            if (!await CommitInlineEdit()) return;
            // Committing rebuilds the page and its text runs; identify the corresponding fresh run.
            if (!editSentences.Contains(run))
                run = editSentences.OrderBy(candidate => Math.Abs(candidate.Bounds.X - anchor.X) +
                    Math.Abs(candidate.Bounds.Y - anchor.Y)).FirstOrDefault();
            if (run == null) return;
            var region = editSentences.FirstOrDefault(item => item == run);
            if (region == null)
            {
                EditStatusLabel.Text = "이 텍스트는 화면 위치를 확인할 수 없어 직접 편집할 수 없습니다.";
                return;
            }
            editActiveSentence = run;
            var bounds = region.Bounds;
            editSelecting = true;
            EditTextList.SelectedItem = run;
            editSelecting = false;
            double left = Math.Max(0, bounds.X * editDisplayScale - 3);
            double top = Math.Max(0, bounds.Y * editDisplayScale - 4);
            Canvas.SetLeft(EditReplacement, left);
            Canvas.SetTop(EditReplacement, top);
            EditReplacement.Width = Math.Min(EditPageCanvas.Width - left,
                Math.Max(160, bounds.Width * editDisplayScale + 70));
            double fontSize = run.DisplayFontSize;
            EditReplacement.Height = Math.Max(48, bounds.Height * editDisplayScale + fontSize * editDisplayScale * 1.6);
            EditReplacement.FontSize = Math.Max(6, fontSize * editDisplayScale);
            string family = run.IsOcr ? "맑은 고딕" : Regex.Replace((run.FontName ?? "").TrimStart('/'), @"^[A-Z]{6}\+", "");
            family = Regex.Replace(family, @"[-,](BoldItalic|Bold|Italic|Regular)$", "", RegexOptions.IgnoreCase);
            var installed = Fonts.SystemFontFamilies.FirstOrDefault(font => font.Source.Equals(family, StringComparison.OrdinalIgnoreCase));
            EditReplacement.FontFamily = installed ?? new FontFamily("맑은 고딕");
            EditReplacement.FontWeight = run.FontName.Contains("Bold", StringComparison.OrdinalIgnoreCase) ? FontWeights.Bold : FontWeights.Normal;
            EditReplacement.FontStyle = run.FontName.Contains("Italic", StringComparison.OrdinalIgnoreCase) ? FontStyles.Italic : FontStyles.Normal;
            EditReplacement.Text = run.Text;
            EditReplacement.IsUndoEnabled = false;
            EditReplacement.IsUndoEnabled = true;
            EditReplacement.Visibility = Visibility.Visible;
            EditReplacement.Focus();
            EditReplacement.CaretIndex = run.Text.Length;
            if (clicked.HasValue)
            {
                EditReplacement.UpdateLayout();
                var point = new Point(clicked.Value.X - left, clicked.Value.Y - top);
                int index = EditReplacement.GetCharacterIndexFromPoint(point, true);
                if (index >= 0) EditReplacement.CaretIndex = index;
            }
            EditFontLabel.Text = run.IsOcr
                ? $"OCR 인식 텍스트 · 약 {fontSize:0.#}pt · 맑은 고딕으로 다시 씀"
                : run.IsInvisible
                    ? $"스캔 이미지의 텍스트 레이어 · 약 {fontSize:0.#}pt"
                    : $"{family} · {fontSize:0.#}pt · 원본 글꼴 유지 (없는 글자는 맑은 고딕)";
            EditStatusLabel.Text = "문장 편집 중 · Enter 완료 · Shift+Enter 줄바꿈 · Esc 취소";
        }

        private async Task<bool> CommitInlineEdit()
        {
            if (editBusy || editCommitting) return false;
            if (editRefreshing || editActiveSentence == null) return true;
            var run = editActiveSentence;
            if (EditReplacement.Text == run.Text)
            {
                CancelInlineEdit();
                return true;
            }
            editCommitting = true;
            try
            {
                var before = new EditSnapshot(editCurrentBytes, editOcrHidden);
                int page = editPageIndex;
                var request = new PdfEditRequest
                {
                    Sentence = run,
                    SourceRuns = run.IsOcr ? Array.Empty<PdfTextRun>() : pdfEditor.GetRuns(page).ToArray(),
                    Replacement = EditReplacement.Text,
                    Obstacles = editSentences.Where(sentence => sentence != run).SelectMany(sentence => sentence.LineBounds).ToList(),
                    PageSize = editPageSize,
                };
                var hidden = new HashSet<string>(editOcrHidden);
                if (run.IsOcr) hidden.Add(run.Ocr.Id);
                bool success = await PrepareEditAsync(PageBuilder(before.Bytes, page, hidden, request, OcrMode.None), () =>
                {
                    editUndo.Push(before);
                    editRedo.Clear();
                    editOcrHidden = hidden;
                    editDirty = true;
                });
                if (success)
                    EditStatusLabel.Text = lastEditNotice == null
                        ? "수정됨 · 다른 문장을 클릭해 계속 편집 · Ctrl+S 저장"
                        : $"수정됨 · {lastEditNotice} · Ctrl+S 저장";
                return success;
            }
            catch (Exception ex)
            {
                // Keep the pending text available for correction instead of discarding user input.
                EditStatusLabel.Text = ex.Message.Replace('\n', ' ');
                return false;
            }
            finally { editCommitting = false; }
        }

        private void CancelInlineEdit()
        {
            editActiveSentence = null;
            editComposing = false;
            EditReplacement.Visibility = Visibility.Collapsed;
        }

        private async void EditInline_KeyDown(object sender, KeyEventArgs e)
        {
            if (editComposing || e.Key == Key.ImeProcessed) return;
            if (e.Key == Key.Escape)
            {
                CancelInlineEdit();
                EditStatusLabel.Text = "문장 수정을 취소했습니다.";
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && Keyboard.Modifiers.HasFlag(ModifierKeys.Shift))
            {
                // Explicit line break inside the replacement text.
                int caret = EditReplacement.SelectionStart;
                EditReplacement.SelectedText = "\n";
                EditReplacement.CaretIndex = caret + 1;
                e.Handled = true;
            }
            else if (e.Key == Key.Enter)
            {
                e.Handled = true;
                await CommitInlineEdit();
            }
        }

        private void EditInline_LostFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (editBusy || editComposing || editRefreshing || editSelecting || editCommitting) return;
            // Toolbar and navigation handlers explicitly commit before acting.
            if (e.NewFocus is Button || e.NewFocus is ComboBox || e.NewFocus is ComboBoxItem || e.NewFocus is ListBoxItem) return;
            var sentence = editActiveSentence;
            // Finish the focus transition first. Never refresh the visual tree or force focus inside LostKeyboardFocus.
            Dispatcher.BeginInvoke(DispatcherPriority.Background, new Action(async () =>
            {
                if (!editBusy && !editCommitting && sentence != null && editActiveSentence == sentence &&
                    !EditReplacement.IsKeyboardFocusWithin) await CommitInlineEdit();
            }));
        }

        private async void EditPageCanvas_Click(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource == EditPageCanvas || e.OriginalSource == EditHitLayer)
            {
                e.Handled = true;
                await CommitInlineEdit();
            }
        }

        private void EditWorkspace_KeyDown(object sender, KeyEventArgs e)
        {
            if (editBusy || editComposing) return;
            if (Keyboard.Modifiers == ModifierKeys.None && (e.Key == Key.PageDown || e.Key == Key.PageUp) &&
                !EditReplacement.IsKeyboardFocusWithin && pdfEditor != null)
            {
                e.Handled = true;
                _ = ChangeEditPage(e.Key == Key.PageDown ? 1 : -1);
                return;
            }
            if (Keyboard.Modifiers != ModifierKeys.Control) return;
            switch (e.Key)
            {
                case Key.S: SaveEditedPdf_Click(sender, e); e.Handled = true; break;
                case Key.O: OpenEditPdf_Click(sender, e); e.Handled = true; break;
                case Key.Z:
                    if (!EditReplacement.IsKeyboardFocusWithin) { UndoPdfEdit_Click(sender, e); e.Handled = true; }
                    break;
                case Key.Y:
                    if (!EditReplacement.IsKeyboardFocusWithin) { RedoPdfEdit_Click(sender, e); e.Handled = true; }
                    break;
            }
        }

        private async void RedoPdfEdit_Click(object sender, RoutedEventArgs e)
        {
            if (editBusy || !await CommitInlineEdit() || editRedo.Count == 0) return;
            var target = editRedo.Peek();
            var before = new EditSnapshot(editCurrentBytes, editOcrHidden);
            await PrepareEditAsync(PageBuilder(target.Bytes, editPageIndex, target.HiddenOcr, mode: OcrMode.None), () =>
            {
                editRedo.Pop();
                editUndo.Push(before);
                editOcrHidden = target.HiddenOcr;
                editDirty = true;
            });
        }

        private async void EditZoom_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsInitialized || pdfEditor == null || editBusy || editRefreshing) return;
            if (!await CommitInlineEdit()) return;
            await ShowEditPage();
        }

        private void EditTextPanel_Changed(object sender, RoutedEventArgs e)
        {
            if (EditTextPanel != null)
                EditTextPanel.Visibility = (sender as CheckBox)?.IsChecked == true ? Visibility.Visible : Visibility.Collapsed;
        }

        private async Task<bool> ConfirmEditDiscard()
        {
            if (editBusy || !await CommitInlineEdit()) return false;
            if (!editDirty) return true;
            var answer = MessageBox.Show(this, "수정한 PDF를 저장하시겠습니까?", "PDF 수정",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
            if (answer == MessageBoxResult.Cancel) return false;
            if (answer == MessageBoxResult.No) return true;
            return await SaveEditedPdf();
        }

        private async void EditWindow_Closing(object sender, CancelEventArgs e)
        {
            if (editCloseApproved) return;
            if (!editBusy && !editDirty && (editActiveSentence == null || EditReplacement.Text == editActiveSentence.Text)) return;
            e.Cancel = true;
            if (editBusy) { EditStatusLabel.Text = "문서 처리 완료 후 닫아주세요."; return; }
            if (await ConfirmEditDiscard()) { editCloseApproved = true; Close(); }
        }
    }
}
