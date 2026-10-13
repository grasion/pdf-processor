using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using PdfiumViewer;
using Velopack;
using Velopack.Sources;

namespace PDFProcessor
{
    public partial class MainWindow : Window
    {
        // 버전 정보
        private const string APP_VERSION = "1.0.8";
        private const string APP_NAME = "PDF 통합 처리기";
        private const string APP_COPYRIGHT = "Copyright © 2024";
        
        // 업데이트 관리자
        private UpdateManager? _updateManager;
        
        // 분할 관련 변수
        private string splitPdfPath = "";
        private string splitOutputFolder = "";
        private PdfSharp.Pdf.PdfDocument splitPdfDoc = null;
        private int splitCurrentPage = 0;
        private int splitTotalPages = 0;

        // 병합 관련 변수
        private List<string> mergeFiles = new List<string>();
        private string mergeOutputFolder = "";
        private Dictionary<int, int> pageRotations = new Dictionary<int, int>(); // 페이지별 회전 각도 (0, 90, 180, 270)
        private List<MergePageInfo> allPages = new List<MergePageInfo>(); // 모든 페이지 정보
        private int currentPreviewPageIndex = 0; // 현재 미리보기 페이지 인덱스
        private HashSet<int> selectedPageIndices = new HashSet<int>(); // 선택된 페이지 인덱스들

        // 변환 관련 변수
        private string convertPdfPath = "";
        private string convertOutputFolder = "";
        private PdfSharp.Pdf.PdfDocument convertPdfDoc = null;
        private int convertCurrentPage = 0;
        private int convertTotalPages = 0;

        public MainWindow()
        {
            InitializeComponent();
            ExtractPdfiumDll();
            SetWindowIcon();
            InitializeDefaults();
            InitializeInlineEditing();
            
            // 업데이트 확인 (비동기)
            _ = CheckForUpdatesAsync();
        }


        private void ExtractPdfiumDll()
        {
            try
            {
                string dllPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "pdfium.dll");
                
                // DLL이 이미 존재하면 건너뛰기
                if (File.Exists(dllPath))
                {
                    return;
                }

                // 임베디드 리소스에서 DLL 추출
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string resourceName = "pdfium.dll";

                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (FileStream fileStream = new FileStream(dllPath, FileMode.Create))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // DLL 추출 실패 시 경고 (프로그램은 계속 실행)
                System.Diagnostics.Debug.WriteLine($"pdfium.dll 추출 실패: {ex.Message}");
            }
        }

        private void SetWindowIcon()
        {
            try
            {
                // 임베디드 리소스에서 아이콘 추출
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                string iconResourceName = assembly.GetManifestResourceNames()
                    .FirstOrDefault(r => r.EndsWith("9040508_filetype_pdf_icon.ico"));
                
                if (iconResourceName != null)
                {
                    using (Stream stream = assembly.GetManifestResourceStream(iconResourceName))
                    {
                        if (stream != null)
                        {
                            // 임시 파일로 저장
                            string tempIconPath = Path.Combine(Path.GetTempPath(), "pdf_icon.ico");
                            using (FileStream fileStream = new FileStream(tempIconPath, FileMode.Create))
                            {
                                stream.CopyTo(fileStream);
                            }
                            
                            this.Icon = new BitmapImage(new Uri(tempIconPath, UriKind.Absolute));
                            System.Diagnostics.Debug.WriteLine("아이콘 로드 완료");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 아이콘 설정 실패 시 무시
                System.Diagnostics.Debug.WriteLine($"아이콘 로드 실패: {ex.Message}");
            }
        }

        private void InitializeDefaults()
        {
            // 병합 기본값 설정
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            MergeSavePath.Text = desktopPath;
            mergeOutputFolder = desktopPath;

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            MergeFileName.Text = $"병합_{timestamp}.pdf";
            
            // 변환 기본값 설정
            ConvertFileNamePattern.Text = "{filename}_page_{page}";
        }

        private void ShowAbout_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 프로그램 정보 창 생성
                Window aboutWindow = new Window
                {
                    Title = "프로그램 정보",
                    Width = 450,
                    Height = 400,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    ResizeMode = ResizeMode.NoResize,
                    Background = System.Windows.Media.Brushes.White
                };

                // 내용 구성
                StackPanel panel = new StackPanel
                {
                    Margin = new Thickness(30)
                };

                // 아이콘 (PDF 아이콘)
                TextBlock icon = new TextBlock
                {
                    Text = "📄",
                    FontSize = 48,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 15)
                };
                panel.Children.Add(icon);

                // 프로그램 이름
                TextBlock appName = new TextBlock
                {
                    Text = APP_NAME,
                    FontSize = 24,
                    FontWeight = FontWeights.Bold,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 5)
                };
                panel.Children.Add(appName);

                // 버전
                TextBlock version = new TextBlock
                {
                    Text = $"버전 {APP_VERSION}",
                    FontSize = 14,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                panel.Children.Add(version);

                // 구분선
                Border separator = new Border
                {
                    Height = 1,
                    Background = System.Windows.Media.Brushes.LightGray,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                panel.Children.Add(separator);

                // 기능 설명
                TextBlock features = new TextBlock
                {
                    Text = "주요 기능",
                    FontSize = 14,
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                panel.Children.Add(features);

                TextBlock featureList = new TextBlock
                {
                    Text = "• PDF 분할 (페이지 범위, 개별, 단위)\n" +
                           "• PDF 병합 (PDF + 이미지)\n" +
                           "• PDF → 이미지 변환 (PNG, JPG, BMP, TIFF)\n" +
                           "• PDF 텍스트 수정 (원본 글꼴 유지)\n" +
                           "• 실시간 미리보기\n" +
                           "• 드래그 앤 드롭 지원",
                    FontSize = 12,
                    Foreground = System.Windows.Media.Brushes.DarkGray,
                    LineHeight = 20,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                panel.Children.Add(featureList);

                // 저작권
                TextBlock copyright = new TextBlock
                {
                    Text = APP_COPYRIGHT,
                    FontSize = 11,
                    Foreground = System.Windows.Media.Brushes.Gray,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                panel.Children.Add(copyright);

aboutWindow.Content = panel;
                aboutWindow.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"정보 창 표시 오류:\n{ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 업데이트 확인 메서드
        private async Task CheckForUpdatesAsync()
        {
            try
            {
                var source = new GithubSource(
                    "https://github.com/grasion/pdf-processor", 
                    null, // 토큰 (public 저장소는 불필요)
                    false // prerelease 포함 여부
                );
                
                _updateManager = new UpdateManager(source);
                
                var updateInfo = await _updateManager.CheckForUpdatesAsync();
                
                if (updateInfo != null)
                {
                    var result = MessageBox.Show(
                        $"새 버전 {updateInfo.TargetFullRelease.Version}이 있습니다.\n업데이트하시겠습니까?",
                        "업데이트 확인",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question
                    );
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        // 진행률 표시 (간단한 메시지)
                        var progressWindow = new Window
                        {
                            Title = "업데이트 중",
                            Width = 300,
                            Height = 150,
                            WindowStartupLocation = WindowStartupLocation.CenterOwner,
                            Owner = this,
                            ResizeMode = ResizeMode.NoResize,
                            WindowStyle = WindowStyle.ToolWindow
                        };
                        
                        var progressPanel = new StackPanel
                        {
                            Margin = new Thickness(20),
                            VerticalAlignment = VerticalAlignment.Center
                        };
                        
                        progressPanel.Children.Add(new TextBlock
                        {
                            Text = "업데이트를 다운로드하는 중...",
                            FontSize = 14,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Margin = new Thickness(0, 0, 0, 20)
                        });
                        
                        var progressBar = new ProgressBar
                        {
                            IsIndeterminate = true,
                            Height = 20
                        };
                        progressPanel.Children.Add(progressBar);
                        
                        progressWindow.Content = progressPanel;
                        progressWindow.Show();
                        
                        await _updateManager.DownloadUpdatesAsync(updateInfo);
                        
                        progressWindow.Close();
                        
                        MessageBox.Show(
                            "업데이트가 다운로드되었습니다.\n프로그램을 재시작합니다.",
                            "업데이트 완료",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                        
                        _updateManager.ApplyUpdatesAndRestart(updateInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"업데이트 확인 실패: {ex.Message}");
                // 업데이트 확인 실패는 조용히 처리 (사용자에게 알리지 않음)
            }
        }

        private async void CheckUpdate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                MessageBox.Show(
                    $"현재 버전: {APP_VERSION}\n저장소: grasion/pdf-processor\n\n업데이트를 확인합니다...",
                    "업데이트 확인",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information
                );
                
                var source = new GithubSource(
                    "https://github.com/grasion/pdf-processor", 
                    null,
                    false
                );
                
                _updateManager = new UpdateManager(source);
                
                var updateInfo = await _updateManager.CheckForUpdatesAsync();
                
                if (updateInfo != null)
                {
                    var result = MessageBox.Show(
                        $"새 버전을 발견했습니다!\n\n현재 버전: {APP_VERSION}\n새 버전: {updateInfo.TargetFullRelease.Version}\n\n업데이트하시겠습니까?",
                        "업데이트 확인",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question
                    );
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        await _updateManager.DownloadUpdatesAsync(updateInfo);
                        
                        MessageBox.Show(
                            "업데이트가 다운로드되었습니다.\n프로그램을 재시작합니다.",
                            "업데이트 완료",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                        
                        _updateManager.ApplyUpdatesAndRestart(updateInfo);
                    }
                }
                else
                {
                    MessageBox.Show(
                        $"현재 최신 버전을 사용 중입니다.\n\n현재 버전: {APP_VERSION}",
                        "업데이트 확인",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"업데이트 확인 중 오류가 발생했습니다:\n\n{ex.Message}\n\n상세 정보:\n{ex.ToString()}",
                    "오류",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error
                );
            }
        }


        #region PDF 분할 기능

        private void SplitDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void SplitDropZone_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void SplitDropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string file = files[0];
                    if (Path.GetExtension(file).ToLower() == ".pdf")
                    {
                        splitPdfPath = file;
                        SplitFileLabel.Text = $"선택된 파일: {Path.GetFileName(splitPdfPath)}";
                        SplitStatusLabel.Text = "파일이 선택되었습니다. 분할 옵션을 설정하세요.";

                        if (string.IsNullOrEmpty(splitOutputFolder))
                        {
                            splitOutputFolder = Path.GetDirectoryName(splitPdfPath);
                            SplitOutputLabel.Text = $"출력 폴더: {splitOutputFolder}";
                        }

                        LoadSplitPdfPreview();
                    }
                    else
                    {
                        MessageBox.Show("PDF 파일만 선택할 수 있습니다.", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        private void SelectSplitFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "PDF 파일 선택"
            };

            if (dialog.ShowDialog() == true)
            {
                splitPdfPath = dialog.FileName;
                SplitFileLabel.Text = $"선택된 파일: {Path.GetFileName(splitPdfPath)}";
                SplitStatusLabel.Text = "파일이 선택되었습니다. 분할 옵션을 설정하세요.";

                if (string.IsNullOrEmpty(splitOutputFolder))
                {
                    splitOutputFolder = Path.GetDirectoryName(splitPdfPath);
                    SplitOutputLabel.Text = $"출력 폴더: {splitOutputFolder}";
                }

                LoadSplitPdfPreview();
            }
        }

        private void SelectSplitOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                splitOutputFolder = dialog.SelectedPath;
                SplitOutputLabel.Text = $"출력 폴더: {splitOutputFolder}";
            }
        }

        private void SplitMode_Changed(object sender, RoutedEventArgs e)
        {
            if (SplitRangePanel == null || SplitChunkPanel == null) return;

            if (SplitModePages.IsChecked == true)
            {
                SplitRangePanel.Visibility = Visibility.Visible;
                SplitChunkPanel.Visibility = Visibility.Collapsed;
            }
            else if (SplitModeChunk.IsChecked == true)
            {
                SplitRangePanel.Visibility = Visibility.Collapsed;
                SplitChunkPanel.Visibility = Visibility.Visible;
            }
            else
            {
                SplitRangePanel.Visibility = Visibility.Collapsed;
                SplitChunkPanel.Visibility = Visibility.Collapsed;
            }
        }

        private void LoadSplitPdfPreview()
        {
            try
            {
                if (splitPdfDoc != null)
                {
                    splitPdfDoc.Dispose();
                }

                splitPdfDoc = PdfReader.Open(splitPdfPath, PdfDocumentOpenMode.Import);
                splitTotalPages = splitPdfDoc.PageCount;
                splitCurrentPage = 0;

                // 썸네일 로드
                LoadSplitThumbnails();
                
                ShowSplitPreviewPage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 미리보기 로드 실패:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadSplitThumbnails()
        {
            SplitPageThumbnailList.Items.Clear();
            
            for (int i = 0; i < splitTotalPages; i++)
            {
                try
                {
                    var thumbnail = RenderPdfPageToImageFromFile(splitPdfPath, i, 72);
                    SplitPageThumbnailList.Items.Add(new SimplePageInfo
                    {
                        PageIndex = i,
                        Thumbnail = thumbnail,
                        PageLabel = $"페이지 {i + 1}"
                    });
                }
                catch { }
            }
        }

        private void SplitPageThumbnail_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SplitPageThumbnailList.SelectedItem is SimplePageInfo pageInfo)
            {
                splitCurrentPage = pageInfo.PageIndex;
                ShowSplitPreviewPage();
            }
        }

        private void SplitPreview_MouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (splitTotalPages == 0) return;
            
            if (e.Delta > 0 && splitCurrentPage > 0)
            {
                splitCurrentPage--;
                ShowSplitPreviewPage();
                e.Handled = true;
            }
            else if (e.Delta < 0 && splitCurrentPage < splitTotalPages - 1)
            {
                splitCurrentPage++;
                ShowSplitPreviewPage();
                e.Handled = true;
            }
        }

        private void ShowSplitPreviewPage()
        {
            if (string.IsNullOrEmpty(splitPdfPath) || splitTotalPages == 0) return;

            try
            {
                // PDF 페이지를 이미지로 렌더링
                var image = RenderPdfPageToImageFromFile(splitPdfPath, splitCurrentPage);
                if (image != null)
                {
                    SplitPreviewImage.Source = image;
                }
                SplitPageLabel.Text = $"페이지: {splitCurrentPage + 1}/{splitTotalPages}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"미리보기 표시 오류:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SplitPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (splitCurrentPage > 0)
            {
                splitCurrentPage--;
                ShowSplitPreviewPage();
            }
        }

        private void SplitNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (splitCurrentPage < splitTotalPages - 1)
            {
                splitCurrentPage++;
                ShowSplitPreviewPage();
            }
        }

        private void ExecuteSplit_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(splitPdfPath))
            {
                MessageBox.Show("PDF 파일을 먼저 선택하세요.", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(splitOutputFolder))
            {
                splitOutputFolder = Path.GetDirectoryName(splitPdfPath);
            }

            try
            {
                SplitProgressBar.Value = 0;
                SplitStatusLabel.Text = "PDF 분할 중...";

                PdfSharp.Pdf.PdfDocument inputDoc = PdfReader.Open(splitPdfPath, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;
                string baseName = Path.GetFileNameWithoutExtension(splitPdfPath);

                if (SplitModeIndividual.IsChecked == true)
                {
                    // 각 페이지를 개별 파일로
                    SplitProgressBar.Maximum = totalPages;
                    for (int i = 0; i < totalPages; i++)
                    {
                        PdfSharp.Pdf.PdfDocument outputDoc = new PdfSharp.Pdf.PdfDocument();
                        outputDoc.AddPage(inputDoc.Pages[i]);
                        
                        string outputPath = Path.Combine(splitOutputFolder, 
                            $"{baseName}_page_{i + 1}.pdf");
                        outputDoc.Save(outputPath);
                        outputDoc.Dispose();

                        SplitProgressBar.Value = i + 1;
                        SplitStatusLabel.Text = $"페이지 {i + 1}/{totalPages} 처리 중...";
                    }

                    SplitStatusLabel.Text = $"완료! {totalPages}개의 파일로 분할되었습니다.";
                }
                else if (SplitModeChunk.IsChecked == true)
                {
                    // 페이지 단위로 분할
                    if (!int.TryParse(SplitChunkSize.Text, out int chunkSize) || chunkSize <= 0)
                    {
                        MessageBox.Show("페이지 단위를 올바르게 입력하세요.", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int totalChunks = (int)Math.Ceiling((double)totalPages / chunkSize);
                    SplitProgressBar.Maximum = totalChunks;
                    int fileCount = 0;

                    for (int start = 0; start < totalPages; start += chunkSize)
                    {
                        int end = Math.Min(start + chunkSize, totalPages);
                        PdfSharp.Pdf.PdfDocument outputDoc = new PdfSharp.Pdf.PdfDocument();

                        for (int i = start; i < end; i++)
                        {
                            outputDoc.AddPage(inputDoc.Pages[i]);
                        }

                        fileCount++;
                        string outputPath = Path.Combine(splitOutputFolder, 
                            $"{baseName}_part_{fileCount}_pages_{start + 1}-{end}.pdf");
                        outputDoc.Save(outputPath);
                        outputDoc.Dispose();

                        SplitProgressBar.Value = fileCount;
                        SplitStatusLabel.Text = $"파일 {fileCount}/{totalChunks} 생성 중...";
                    }

                    SplitStatusLabel.Text = $"완료! {chunkSize}페이지씩 {fileCount}개의 파일로 분할되었습니다.";
                }
                else
                {
                    // 페이지 범위로 분할
                    if (!int.TryParse(SplitStartPage.Text, out int startPage) || 
                        !int.TryParse(SplitEndPage.Text, out int endPage))
                    {
                        MessageBox.Show("페이지 번호를 올바르게 입력하세요.", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    startPage--; // 0-based index
                    if (startPage < 0 || endPage > totalPages || startPage >= endPage)
                    {
                        MessageBox.Show($"잘못된 페이지 범위입니다.\n전체 페이지 수: {totalPages}", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    int pageCount = endPage - startPage;
                    SplitProgressBar.Maximum = pageCount;

                    PdfSharp.Pdf.PdfDocument outputDoc = new PdfSharp.Pdf.PdfDocument();
                    for (int i = startPage; i < endPage; i++)
                    {
                        outputDoc.AddPage(inputDoc.Pages[i]);
                        SplitProgressBar.Value = i - startPage + 1;
                        SplitStatusLabel.Text = $"페이지 {i + 1} 처리 중... ({i - startPage + 1}/{pageCount})";
                    }

                    string outputPath = Path.Combine(splitOutputFolder, 
                        $"{baseName}_pages_{startPage + 1}-{endPage}.pdf");
                    outputDoc.Save(outputPath);
                    outputDoc.Dispose();

                    SplitStatusLabel.Text = $"완료! 페이지 {startPage + 1}-{endPage}가 분할되었습니다.";
                }

                inputDoc.Dispose();
                MessageBox.Show($"PDF 분할이 완료되었습니다.\n출력 폴더: {splitOutputFolder}", "완료", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 분할 중 오류가 발생했습니다:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region PDF 병합 기능

        private void MergeDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void MergeDropZone_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void MergeDropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    int addedCount = 0;
                    foreach (string file in files)
                    {
                        string ext = Path.GetExtension(file).ToLower();
                        // PDF 및 이미지 파일만 허용
                        if (ext == ".pdf" || ext == ".png" || ext == ".jpg" || 
                            ext == ".jpeg" || ext == ".bmp" || ext == ".tiff" || ext == ".gif")
                        {
                            if (!mergeFiles.Contains(file))
                            {
                                mergeFiles.Add(file);
                                MergeFileList.Items.Add(Path.GetFileName(file));
                                addedCount++;
                            }
                        }
                    }

                    if (addedCount > 0)
                    {
                        MergeStatusLabel.Text = $"{mergeFiles.Count}개의 파일이 준비되었습니다.";
                        LoadAllMergePages(); // 모든 페이지 로드
                    }
                    else if (files.Length > 0)
                    {
                        MessageBox.Show("PDF 또는 이미지 파일만 추가할 수 있습니다.\n지원 형식: PDF, PNG, JPG, BMP, TIFF, GIF", 
                            "알림", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
        }

        private void AddMergeFiles_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "지원 파일|*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif|" +
                        "PDF 파일|*.pdf|이미지 파일|*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.gif",
                Title = "병합할 파일 선택 (PDF, 이미지만 지원)",
                Multiselect = true
            };

            if (dialog.ShowDialog() == true)
            {
                foreach (string file in dialog.FileNames)
                {
                    if (!mergeFiles.Contains(file))
                    {
                        mergeFiles.Add(file);
                        MergeFileList.Items.Add(Path.GetFileName(file));
                    }
                }

                MergeStatusLabel.Text = $"{mergeFiles.Count}개의 파일이 준비되었습니다.";
                
                // 모든 페이지 로드 및 미리보기 업데이트
                LoadAllMergePages();
            }
        }

        private void MoveUp_Click(object sender, RoutedEventArgs e)
        {
            int index = MergeFileList.SelectedIndex;
            if (index > 0)
            {
                string temp = mergeFiles[index];
                mergeFiles[index] = mergeFiles[index - 1];
                mergeFiles[index - 1] = temp;

                object tempItem = MergeFileList.Items[index];
                MergeFileList.Items[index] = MergeFileList.Items[index - 1];
                MergeFileList.Items[index - 1] = tempItem;

                LoadAllMergePages();
                MergeFileList.SelectedIndex = index - 1;
            }
        }

        private void MoveDown_Click(object sender, RoutedEventArgs e)
        {
            int index = MergeFileList.SelectedIndex;
            if (index >= 0 && index < mergeFiles.Count - 1)
            {
                string temp = mergeFiles[index];
                mergeFiles[index] = mergeFiles[index + 1];
                mergeFiles[index + 1] = temp;

                object tempItem = MergeFileList.Items[index];
                MergeFileList.Items[index] = MergeFileList.Items[index + 1];
                MergeFileList.Items[index + 1] = tempItem;

                LoadAllMergePages();
                MergeFileList.SelectedIndex = index + 1;
            }
        }

        private void RemoveMergeFile_Click(object sender, RoutedEventArgs e)
        {
            int index = MergeFileList.SelectedIndex;
            if (index >= 0)
            {
                mergeFiles.RemoveAt(index);
                MergeFileList.Items.RemoveAt(index);
                
                MergeStatusLabel.Text = mergeFiles.Count > 0 
                    ? $"{mergeFiles.Count}개의 파일이 준비되었습니다." 
                    : "병합할 PDF 파일들을 추가하세요.";
                
                // 모든 페이지 다시 로드
                if (mergeFiles.Count > 0)
                {
                    LoadAllMergePages();
                }
                else
                {
                    allPages.Clear();
                    pageRotations.Clear();
                    selectedPageIndices.Clear();
                    MergePageThumbnailList.Items.Clear();
                    MergePreviewImage.Source = null;
                    MergePageInfo.Text = "페이지: 0 / 0";
                }
            }
        }

        private void ClearMergeFiles_Click(object sender, RoutedEventArgs e)
        {
            mergeFiles.Clear();
            MergeFileList.Items.Clear();
            pageRotations.Clear();
            allPages.Clear();
            selectedPageIndices.Clear();
            currentPreviewPageIndex = 0;
            
            MergePageThumbnailList.Items.Clear();
            MergePreviewImage.Source = null;
            MergePageInfo.Text = "페이지: 0 / 0";
            
            MergeStatusLabel.Text = "병합할 PDF 파일들을 추가하세요.";
        }

        private void BrowseMergeSavePath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                mergeOutputFolder = dialog.SelectedPath;
                MergeSavePath.Text = mergeOutputFolder;
            }
        }

        private void MergeFileList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 파일 목록 선택 시 해당 파일의 첫 페이지로 이동
            int fileIndex = MergeFileList.SelectedIndex;
            if (fileIndex >= 0 && fileIndex < mergeFiles.Count)
            {
                // 해당 파일의 첫 페이지 찾기
                var firstPageOfFile = allPages.FirstOrDefault(p => p.FileIndex == fileIndex);
                if (firstPageOfFile != null)
                {
                    ShowPreviewPage(firstPageOfFile.PageIndex);
                }
            }
        }

        private void RotateLeft_Click(object sender, RoutedEventArgs e)
        {
            if (allPages.Count == 0) return;
            
            // 선택된 페이지가 있으면 선택된 페이지들 회전, 없으면 현재 페이지만 회전
            var pagesToRotate = selectedPageIndices.Count > 0 
                ? selectedPageIndices.ToList() 
                : new List<int> { currentPreviewPageIndex };
            
            foreach (int pageIndex in pagesToRotate)
            {
                int currentRotation = pageRotations.ContainsKey(pageIndex) ? pageRotations[pageIndex] : 0;
                int newRotation = (currentRotation - 90 + 360) % 360;
                pageRotations[pageIndex] = newRotation;
            }
            
            // 미리보기 및 썸네일 업데이트
            ShowPreviewPage(currentPreviewPageIndex);
            UpdateThumbnailList();
        }

        private void RotateRight_Click(object sender, RoutedEventArgs e)
        {
            if (allPages.Count == 0) return;
            
            var pagesToRotate = selectedPageIndices.Count > 0 
                ? selectedPageIndices.ToList() 
                : new List<int> { currentPreviewPageIndex };
            
            foreach (int pageIndex in pagesToRotate)
            {
                int currentRotation = pageRotations.ContainsKey(pageIndex) ? pageRotations[pageIndex] : 0;
                int newRotation = (currentRotation + 90) % 360;
                pageRotations[pageIndex] = newRotation;
            }
            
            ShowPreviewPage(currentPreviewPageIndex);
            UpdateThumbnailList();
        }

        private void Rotate180_Click(object sender, RoutedEventArgs e)
        {
            if (allPages.Count == 0) return;
            
            var pagesToRotate = selectedPageIndices.Count > 0 
                ? selectedPageIndices.ToList() 
                : new List<int> { currentPreviewPageIndex };
            
            foreach (int pageIndex in pagesToRotate)
            {
                int currentRotation = pageRotations.ContainsKey(pageIndex) ? pageRotations[pageIndex] : 0;
                int newRotation = (currentRotation + 180) % 360;
                pageRotations[pageIndex] = newRotation;
            }
            
            ShowPreviewPage(currentPreviewPageIndex);
            UpdateThumbnailList();
        }

        private void UpdateRotationStatus(int pageIndex)
        {
            if (pageRotations.ContainsKey(pageIndex) && pageRotations[pageIndex] != 0)
            {
                RotationStatusLabel.Text = $"회전: {pageRotations[pageIndex]}°";
            }
            else
            {
                RotationStatusLabel.Text = "";
            }
        }

        private BitmapSource RotateBitmapSource(BitmapSource source, int angle)
        {
            if (angle == 0) return source;
            
            TransformedBitmap transformedBitmap = new TransformedBitmap();
            transformedBitmap.BeginInit();
            transformedBitmap.Source = source;
            
            RotateTransform rotateTransform = new RotateTransform(angle);
            transformedBitmap.Transform = rotateTransform;
            
            transformedBitmap.EndInit();
            transformedBitmap.Freeze();
            
            return transformedBitmap;
        }

        private void LoadAllMergePages()
        {
            // Keep rotations attached to their source page when files move or are removed.
            var rotationsByPage = allPages.ToDictionary(
                page => (page.FilePath, page.PageInFile),
                page => pageRotations.TryGetValue(page.PageIndex, out int angle) ? angle : 0);
            var checkedPages = allPages.Where(page => selectedPageIndices.Contains(page.PageIndex))
                .Select(page => (page.FilePath, page.PageInFile)).ToHashSet();
            allPages.Clear();
            pageRotations.Clear();
            selectedPageIndices.Clear();
            currentPreviewPageIndex = 0;
            
            int globalPageIndex = 0;
            
            for (int fileIndex = 0; fileIndex < mergeFiles.Count; fileIndex++)
            {
                string filePath = mergeFiles[fileIndex];
                string ext = Path.GetExtension(filePath).ToLower();
                
                try
                {
                    if (ext == ".pdf")
                    {
                        // PDF 파일의 모든 페이지 로드
                        using (var pdfDoc = PdfiumViewer.PdfDocument.Load(filePath))
                        {
                            for (int pageInFile = 0; pageInFile < pdfDoc.PageCount; pageInFile++)
                            {
                                var thumbnail = RenderPdfPageToImageFromFile(filePath, pageInFile, 72); // 낮은 DPI로 썸네일
                                
                                allPages.Add(new MergePageInfo
                                {
                                    PageIndex = globalPageIndex,
                                    FileIndex = fileIndex,
                                    PageInFile = pageInFile,
                                    FilePath = filePath,
                                    Thumbnail = thumbnail,
                                    PageLabel = $"페이지 {globalPageIndex + 1}",
                                    RotationInfo = "",
                                    IsChecked = false
                                });
                                
                                globalPageIndex++;
                            }
                        }
                    }
                    else
                    {
                        // 이미지 파일
                        BitmapImage bitmap = new BitmapImage();
                        bitmap.BeginInit();
                        bitmap.UriSource = new Uri(filePath);
                        bitmap.DecodePixelWidth = 150; // 썸네일 크기
                        bitmap.CacheOption = BitmapCacheOption.OnLoad;
                        bitmap.EndInit();
                        
                        allPages.Add(new MergePageInfo
                        {
                            PageIndex = globalPageIndex,
                            FileIndex = fileIndex,
                            PageInFile = 0,
                            FilePath = filePath,
                            Thumbnail = bitmap,
                            PageLabel = $"페이지 {globalPageIndex + 1}",
                            RotationInfo = "",
                            IsChecked = false
                        });
                        
                        globalPageIndex++;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"페이지 로드 오류: {ex.Message}");
                }
            }
            
            foreach (var page in allPages)
            {
                var key = (page.FilePath, page.PageInFile);
                if (rotationsByPage.TryGetValue(key, out int angle) && angle != 0)
                    pageRotations[page.PageIndex] = angle;
                page.IsChecked = checkedPages.Contains(key);
                if (page.IsChecked) selectedPageIndices.Add(page.PageIndex);
            }
            UpdateRotationModeLabel();
            // 썸네일 리스트 업데이트
            UpdateThumbnailList();
            
            // 첫 페이지 미리보기
            if (allPages.Count > 0)
            {
                ShowPreviewPage(0);
            }
        }

        private void UpdateThumbnailList()
        {
            MergePageThumbnailList.Items.Clear();
            
            foreach (var page in allPages)
            {
                // 회전 정보 업데이트
                if (pageRotations.ContainsKey(page.PageIndex) && pageRotations[page.PageIndex] != 0)
                {
                    page.RotationInfo = $"회전: {pageRotations[page.PageIndex]}°";
                }
                else
                {
                    page.RotationInfo = "";
                }
                
                MergePageThumbnailList.Items.Add(page);
            }
        }

        private void ShowPreviewPage(int pageIndex)
        {
            if (pageIndex < 0 || pageIndex >= allPages.Count) return;
            
            currentPreviewPageIndex = pageIndex;
            var pageInfo = allPages[pageIndex];
            
            try
            {
                using var preview = new PdfSharp.Pdf.PdfDocument();
                int rotation = pageRotations.TryGetValue(pageIndex, out int angle) ? angle : 0;
                MergePageLayout.AddPage(preview, pageInfo.FilePath, pageInfo.PageInFile,
                    rotation, SelectedMergePaper, SelectedMergeOrientation);
                using var stream = new MemoryStream();
                preview.Save(stream, false);
                stream.Position = 0;
                using var rendered = PdfiumViewer.PdfDocument.Load(stream);
                var size = rendered.PageSizes[0];
                double scale = Math.Min(800.0 / size.Width, 1000.0 / size.Height);
                using var image = rendered.Render(0, Math.Max(1, (int)(size.Width * scale)),
                    Math.Max(1, (int)(size.Height * scale)), 96, 96, false);
                using var bitmap = new System.Drawing.Bitmap(image);
                MergePreviewImage.Source = ConvertBitmapToBitmapSource(bitmap);
                MergePageInfo.Text = $"페이지: {pageIndex + 1} / {allPages.Count} · {SelectedMergePaper} / {SelectedMergeOrientation}";
                UpdateRotationStatus(pageIndex);
            }
            catch (Exception ex)
            {
                MergePreviewImage.Source = CreateErrorImage($"미리보기를 불러올 수 없습니다.\n{ex.Message}");
                MergeStatusLabel.Text = $"미리보기 오류: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"미리보기 오류: {ex.Message}");
            }
        }

        private void MergePreview_MouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (allPages.Count == 0) return;
            
            if (e.Delta > 0)
            {
                // 휠 위로 - 이전 페이지
                if (currentPreviewPageIndex > 0)
                {
                    ShowPreviewPage(currentPreviewPageIndex - 1);
                }
            }
            else
            {
                // 휠 아래로 - 다음 페이지
                if (currentPreviewPageIndex < allPages.Count - 1)
                {
                    ShowPreviewPage(currentPreviewPageIndex + 1);
                }
            }
            
            e.Handled = true;
        }

        private void MergePageThumbnail_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MergePageThumbnailList.SelectedItem is MergePageInfo pageInfo)
            {
                ShowPreviewPage(pageInfo.PageIndex);
            }
        }

        private void PageCheckBox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox checkBox && checkBox.Tag is int pageIndex)
            {
                if (checkBox.IsChecked == true)
                {
                    selectedPageIndices.Add(pageIndex);
                }
                else
                {
                    selectedPageIndices.Remove(pageIndex);
                }
                
                // 회전 모드 레이블 업데이트
                UpdateRotationModeLabel();
            }
        }

        private void SelectAllPages_Click(object sender, RoutedEventArgs e)
        {
            selectedPageIndices.Clear();
            for (int i = 0; i < allPages.Count; i++)
            {
                selectedPageIndices.Add(i);
            }
            
            // 체크박스 업데이트
            foreach (var item in MergePageThumbnailList.Items)
            {
                if (item is MergePageInfo pageInfo)
                {
                    pageInfo.IsChecked = true;
                }
            }
            
            UpdateThumbnailList();
            UpdateRotationModeLabel();
        }

        private void DeselectAllPages_Click(object sender, RoutedEventArgs e)
        {
            selectedPageIndices.Clear();
            
            // 체크박스 업데이트
            foreach (var item in MergePageThumbnailList.Items)
            {
                if (item is MergePageInfo pageInfo)
                {
                    pageInfo.IsChecked = false;
                }
            }
            
            UpdateThumbnailList();
            UpdateRotationModeLabel();
        }

        private void UpdateRotationModeLabel()
        {
            if (selectedPageIndices.Count > 0)
            {
                RotationModeLabel.Text = $"선택한 {selectedPageIndices.Count}개 페이지 회전";
            }
            else
            {
                RotationModeLabel.Text = "현재 페이지 회전";
            }
        }

        private string SelectedMergePaper =>
            (MergePageSize.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "A4";

        private string SelectedMergeOrientation =>
            (MergePageOrientation.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "세로";

        private void MergeLayout_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (!IsInitialized || MergePreviewImage == null || allPages.Count == 0) return;
            ShowPreviewPage(currentPreviewPageIndex);
        }

        private void ExecuteMerge_Click(object sender, RoutedEventArgs e)
        {
            if (mergeFiles.Count == 0 || allPages.Count == 0)
            {
                MessageBox.Show("병합할 파일을 추가하세요.", "안내",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string temporaryPath = null;
            try
            {
                string filename = MergeFileName.Text.Trim();
                if (string.IsNullOrWhiteSpace(filename) || Path.GetFileName(filename) != filename ||
                    filename.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    throw new InvalidOperationException("올바른 PDF 파일명을 입력하세요.");
                if (!filename.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) filename += ".pdf";
                string outputPath = Path.GetFullPath(Path.Combine(MergeSavePath.Text, filename));
                if (mergeFiles.Any(path => string.Equals(Path.GetFullPath(path), outputPath, StringComparison.OrdinalIgnoreCase)))
                    throw new InvalidOperationException("원본 파일과 다른 파일명 또는 저장 폴더를 선택하세요.");
                if (allPages.Select(page => page.FilePath).Distinct().Count() != mergeFiles.Count)
                    throw new InvalidOperationException("불러오지 못한 파일이 있습니다. 파일 목록을 확인하세요.");
                if (File.Exists(outputPath) && MessageBox.Show("같은 이름의 파일이 있습니다. 덮어쓰시겠습니까?",
                    "파일 저장", MessageBoxButton.YesNo, MessageBoxImage.Question) != MessageBoxResult.Yes) return;

                MergeProgressBar.Maximum = allPages.Count;
                MergeProgressBar.Value = 0;
                MergeStatusLabel.Text = "PDF 병합 중...";
                using var outputDoc = new PdfSharp.Pdf.PdfDocument();
                for (int pageIndex = 0; pageIndex < allPages.Count; pageIndex++)
                {
                    var pageInfo = allPages[pageIndex];
                    int rotation = pageRotations.TryGetValue(pageIndex, out int angle) ? angle : 0;
                    MergePageLayout.AddPage(outputDoc, pageInfo.FilePath, pageInfo.PageInFile,
                        rotation, SelectedMergePaper, SelectedMergeOrientation);
                    MergeProgressBar.Value = pageIndex + 1;
                }

                temporaryPath = Path.Combine(Path.GetDirectoryName(outputPath), $".merge-{Guid.NewGuid():N}.tmp");
                outputDoc.Save(temporaryPath);
                File.Move(temporaryPath, outputPath, true);
                MergeFileName.Text = filename;
                MergeStatusLabel.Text = $"완료 · {allPages.Count}페이지 · {SelectedMergePaper} / {SelectedMergeOrientation}";
                MessageBox.Show($"PDF 병합이 완료되었습니다.\n저장 위치: {outputPath}", "완료",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MergeStatusLabel.Text = "병합 실패 · 설정과 파일을 확인하세요.";
                MessageBox.Show($"PDF 병합 중 오류가 발생했습니다:\n{ex.Message}", "오류",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (temporaryPath != null && File.Exists(temporaryPath)) File.Delete(temporaryPath);
            }
        }
        #endregion

        #region PDF → 이미지 변환 기능

        private void ConvertDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void ConvertDropZone_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            e.Handled = false; // 이벤트 전파 허용
        }

        private void ConvertDropZone_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files != null && files.Length > 0)
                {
                    string file = files[0];
                    if (Path.GetExtension(file).ToLower() == ".pdf")
                    {
                        convertPdfPath = file;
                        ConvertFileLabel.Text = $"선택된 파일: {Path.GetFileName(convertPdfPath)}";
                        ConvertStatusLabel.Text = "파일이 선택되었습니다. 변환 옵션을 설정하세요.";

                        if (string.IsNullOrEmpty(convertOutputFolder))
                        {
                            convertOutputFolder = Path.GetDirectoryName(convertPdfPath);
                            ConvertOutputLabel.Text = $"출력 폴더: {convertOutputFolder}";
                        }

                        LoadConvertPdfPreview();
                    }
                    else
                    {
                        MessageBox.Show("PDF 파일만 선택할 수 있습니다.", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        private void SelectConvertFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog
            {
                Filter = "PDF files (*.pdf)|*.pdf",
                Title = "PDF 파일 선택"
            };

            if (dialog.ShowDialog() == true)
            {
                convertPdfPath = dialog.FileName;
                ConvertFileLabel.Text = $"선택된 파일: {Path.GetFileName(convertPdfPath)}";
                ConvertStatusLabel.Text = "파일이 선택되었습니다. 변환 옵션을 설정하세요.";

                if (string.IsNullOrEmpty(convertOutputFolder))
                {
                    convertOutputFolder = Path.GetDirectoryName(convertPdfPath);
                    ConvertOutputLabel.Text = $"출력 폴더: {convertOutputFolder}";
                }

                LoadConvertPdfPreview();
            }
        }

        private void SelectConvertOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                convertOutputFolder = dialog.SelectedPath;
                ConvertOutputLabel.Text = $"출력 폴더: {convertOutputFolder}";
            }
        }

        private void ConvertMode_Changed(object sender, RoutedEventArgs e)
        {
            if (ConvertRangePanel == null) return;

            ConvertRangePanel.Visibility = ConvertModeRange.IsChecked == true 
                ? Visibility.Visible 
                : Visibility.Collapsed;
        }

        private void LoadConvertPdfPreview()
        {
            try
            {
                if (convertPdfDoc != null)
                {
                    convertPdfDoc.Dispose();
                }

                convertPdfDoc = PdfReader.Open(convertPdfPath, PdfDocumentOpenMode.Import);
                convertTotalPages = convertPdfDoc.PageCount;
                convertCurrentPage = 0;

                // 썸네일 로드
                LoadConvertThumbnails();
                
                ShowConvertPreviewPage();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 미리보기 로드 실패:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadConvertThumbnails()
        {
            ConvertPageThumbnailList.Items.Clear();
            
            for (int i = 0; i < convertTotalPages; i++)
            {
                try
                {
                    var thumbnail = RenderPdfPageToImageFromFile(convertPdfPath, i, 72);
                    ConvertPageThumbnailList.Items.Add(new SimplePageInfo
                    {
                        PageIndex = i,
                        Thumbnail = thumbnail,
                        PageLabel = $"페이지 {i + 1}"
                    });
                }
                catch { }
            }
        }

        private void ConvertPageThumbnail_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ConvertPageThumbnailList.SelectedItem is SimplePageInfo pageInfo)
            {
                convertCurrentPage = pageInfo.PageIndex;
                ShowConvertPreviewPage();
            }
        }

        private void ConvertPreview_MouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            if (convertTotalPages == 0) return;
            
            if (e.Delta > 0 && convertCurrentPage > 0)
            {
                convertCurrentPage--;
                ShowConvertPreviewPage();
                e.Handled = true;
            }
            else if (e.Delta < 0 && convertCurrentPage < convertTotalPages - 1)
            {
                convertCurrentPage++;
                ShowConvertPreviewPage();
                e.Handled = true;
            }
        }

        private void ShowConvertPreviewPage()
        {
            if (string.IsNullOrEmpty(convertPdfPath) || convertTotalPages == 0) return;

            try
            {
                var image = RenderPdfPageToImageFromFile(convertPdfPath, convertCurrentPage);
                if (image != null)
                {
                    ConvertPreviewImage.Source = image;
                }
                ConvertPageLabel.Text = $"페이지: {convertCurrentPage + 1}/{convertTotalPages}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"미리보기 표시 오류:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ConvertPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (convertCurrentPage > 0)
            {
                convertCurrentPage--;
                ShowConvertPreviewPage();
            }
        }

        private void ConvertNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (convertCurrentPage < convertTotalPages - 1)
            {
                convertCurrentPage++;
                ShowConvertPreviewPage();
            }
        }

        private void ExecuteConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(convertPdfPath))
            {
                MessageBox.Show("PDF 파일을 먼저 선택하세요.", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(convertOutputFolder))
            {
                convertOutputFolder = Path.GetDirectoryName(convertPdfPath);
            }

            try
            {
                ConvertProgressBar.Value = 0;
                ConvertStatusLabel.Text = "PDF 변환 중...";

                PdfSharp.Pdf.PdfDocument inputDoc = PdfReader.Open(convertPdfPath, PdfDocumentOpenMode.Import);
                int totalPages = inputDoc.PageCount;

                int startPage = 0;
                int endPage = totalPages;

                if (ConvertModeRange.IsChecked == true)
                {
                    if (!int.TryParse(ConvertStartPage.Text, out startPage) || 
                        !int.TryParse(ConvertEndPage.Text, out endPage))
                    {
                        MessageBox.Show("페이지 번호를 올바르게 입력하세요.", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    startPage--; // 0-based index
                    if (startPage < 0 || endPage > totalPages || startPage >= endPage)
                    {
                        MessageBox.Show($"잘못된 페이지 범위입니다.\n전체 페이지 수: {totalPages}", "오류", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }

                string baseName = Path.GetFileNameWithoutExtension(convertPdfPath);
                string pattern = ConvertFileNamePattern.Text;
                string format = ((ComboBoxItem)ConvertImageFormat.SelectedItem).Content.ToString().ToLower();

                int pageCount = endPage - startPage;
                ConvertProgressBar.Maximum = pageCount;

                // DPI 가져오기
                int dpi = int.Parse(((ComboBoxItem)ConvertDPI.SelectedItem).Content.ToString());

                // PDF를 이미지로 변환
                try
                {
                    using (var pdfDocument = PdfiumViewer.PdfDocument.Load(convertPdfPath))
                    {
                        for (int i = startPage; i < endPage; i++)
                        {
                            try
                            {
                                string filename = pattern
                                    .Replace("{filename}", baseName)
                                    .Replace("{page}", (i + 1).ToString());
                                string outputPath = Path.Combine(convertOutputFolder, $"{filename}.{format}");

                                // 페이지 크기 계산
                                var pageSize = pdfDocument.PageSizes[i];
                                int width = (int)(pageSize.Width * dpi / 72.0);
                                int height = (int)(pageSize.Height * dpi / 72.0);

                                // 크기 제한
                                width = Math.Max(100, Math.Min(width, 4000));
                                height = Math.Max(100, Math.Min(height, 6000));

                                // PDF 페이지를 이미지로 렌더링
                                using (var image = pdfDocument.Render(i, width, height, dpi, dpi, false))
                                {
                                    // 이미지를 Bitmap으로 변환하여 저장
                                    using (var bitmap = new System.Drawing.Bitmap(image))
                                    {
                                        SaveBitmapToFile(bitmap, outputPath, format);
                                    }
                                }

                                ConvertProgressBar.Value = i - startPage + 1;
                                ConvertStatusLabel.Text = $"변환 중... {i - startPage + 1}/{pageCount}";
                            }
                            catch (Exception pageEx)
                            {
                                MessageBox.Show($"페이지 {i + 1} 변환 오류:\n{pageEx.Message}", "경고", 
                                    MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                    }
                }
                catch (Exception pdfEx)
                {
                    throw new Exception($"PDF 로드 오류: {pdfEx.Message}", pdfEx);
                }

                if (inputDoc != null)
                    inputDoc.Dispose();

                ConvertStatusLabel.Text = $"완료! {pageCount}개의 이미지가 생성되었습니다.";
                MessageBox.Show($"PDF를 이미지로 변환했습니다.\n{pageCount}개의 파일이 생성되었습니다.\n\n출력 폴더: {convertOutputFolder}", 
                    "완료", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"PDF 변환 중 오류가 발생했습니다:\n{ex.Message}", "오류", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        #region 유틸리티 메서드

        private BitmapSource RenderPdfPageToImageFromFile(string pdfPath, int pageIndex, int dpi = 150)
        {
            try
            {
                // 파일 존재 확인
                if (!File.Exists(pdfPath))
                {
                    return CreateErrorImage("파일을 찾을 수 없습니다.");
                }

                using (var document = PdfiumViewer.PdfDocument.Load(pdfPath))
                {
                    if (pageIndex >= document.PageCount || pageIndex < 0)
                    {
                        return CreateErrorImage($"잘못된 페이지 번호: {pageIndex + 1}");
                    }

                    // DPI에 따른 크기 계산
                    var page = document.PageSizes[pageIndex];
                    int width = (int)(page.Width * dpi / 72.0);
                    int height = (int)(page.Height * dpi / 72.0);

                    // 미리보기용 크기 제한 (더 작게)
                    int maxWidth = 800;
                    int maxHeight = 1000;
                    
                    if (width > maxWidth || height > maxHeight)
                    {
                        double scale = Math.Min((double)maxWidth / width, (double)maxHeight / height);
                        width = (int)(width * scale);
                        height = (int)(height * scale);
                    }

                    width = Math.Max(100, width);
                    height = Math.Max(100, height);

                    // PDF 페이지를 이미지로 렌더링
                    using (var image = document.Render(pageIndex, width, height, dpi, dpi, false))
                    {
                        using (var bitmap = new System.Drawing.Bitmap(image))
                        {
                            return ConvertBitmapToBitmapSource(bitmap);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return CreateErrorImage($"미리보기 오류:\n{ex.Message}");
            }
        }

        private BitmapSource CreateErrorImage(string message)
        {
            var bitmap = new System.Drawing.Bitmap(400, 600);
            using (var g = System.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(System.Drawing.Color.White);
                var font = new System.Drawing.Font("맑은 고딕", 12);
                g.DrawString(message, font, 
                    System.Drawing.Brushes.Red, 
                    new System.Drawing.RectangleF(20, 20, 360, 560));
            }
            return ConvertBitmapToBitmapSource(bitmap);
        }

        private BitmapSource ConvertBitmapToBitmapSource(System.Drawing.Bitmap bitmap)
        {
            try
            {
                using (var memory = new MemoryStream())
                {
                    bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                    memory.Position = 0;

                    var bitmapImage = new BitmapImage();
                    bitmapImage.BeginInit();
                    bitmapImage.StreamSource = memory;
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                    bitmapImage.EndInit();
                    bitmapImage.Freeze();

                    return bitmapImage;
                }
            }
            catch
            {
                // 오류 발생 시 빈 이미지 반환
                return BitmapSource.Create(1, 1, 96, 96, 
                    System.Windows.Media.PixelFormats.Bgr24, null, 
                    new byte[3], 3);
            }
        }

        private void SaveBitmapSourceToFile(BitmapSource image, string filePath, string format)
        {
            BitmapEncoder encoder;
            
            switch (format.ToLower())
            {
                case "jpg":
                case "jpeg":
                    encoder = new JpegBitmapEncoder { QualityLevel = 95 };
                    break;
                case "bmp":
                    encoder = new BmpBitmapEncoder();
                    break;
                case "tiff":
                    encoder = new TiffBitmapEncoder();
                    break;
                default:
                    encoder = new PngBitmapEncoder();
                    break;
            }

            encoder.Frames.Add(BitmapFrame.Create(image));

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                encoder.Save(stream);
            }
        }

        private void SaveBitmapToFile(System.Drawing.Bitmap bitmap, string filePath, string format)
        {
            switch (format.ToLower())
            {
                case "jpg":
                case "jpeg":
                    bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                    break;
                case "bmp":
                    bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Bmp);
                    break;
                case "tiff":
                    bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Tiff);
                    break;
                default:
                    bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                    break;
            }
        }

        #endregion
    }
}

    // 병합 페이지 정보 클래스
    public class MergePageInfo
    {
        public int PageIndex { get; set; } // 전체 페이지 인덱스
        public int FileIndex { get; set; } // 파일 인덱스
        public int PageInFile { get; set; } // 파일 내 페이지 번호
        public string FilePath { get; set; }
        public BitmapSource Thumbnail { get; set; }
        public string PageLabel { get; set; }
        public string RotationInfo { get; set; }
        public bool IsChecked { get; set; }
    }
    
    // 간단한 페이지 정보 클래스 (분할/변환용)
    public class SimplePageInfo
    {
        public int PageIndex { get; set; }
        public BitmapSource Thumbnail { get; set; }
        public string PageLabel { get; set; }
    }
