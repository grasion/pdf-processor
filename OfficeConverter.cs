using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace PDFProcessor
{
    /// <summary>
    /// Converts Word, Excel and PowerPoint files to PDF with the installed Microsoft Office (COM automation),
    /// or LibreOffice when Office is not available. Results are cached per file version.
    /// </summary>
    internal static class OfficeConverter
    {
        internal static readonly string[] WordExtensions = { ".doc", ".docx", ".docm", ".rtf", ".odt" };
        internal static readonly string[] ExcelExtensions = { ".xls", ".xlsx", ".xlsm", ".xlsb", ".csv", ".ods" };
        internal static readonly string[] PowerPointExtensions = { ".ppt", ".pptx", ".pptm", ".odp" };
        private static readonly TimeSpan Timeout = TimeSpan.FromMinutes(3);
        private static readonly object Gate = new();

        private static string CacheDirectory => Path.Combine(Path.GetTempPath(), "PDF통합처리기", "converted");

        internal static bool IsOfficeFile(string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            return WordExtensions.Contains(extension) || ExcelExtensions.Contains(extension) || PowerPointExtensions.Contains(extension);
        }

        internal static string KindName(string path)
        {
            string extension = Path.GetExtension(path).ToLowerInvariant();
            if (WordExtensions.Contains(extension)) return "워드 문서";
            if (ExcelExtensions.Contains(extension)) return "엑셀 문서";
            return "파워포인트 문서";
        }

        /// <summary>Returns the path of a PDF rendering of the document. Blocking; call off the UI thread.</summary>
        internal static string ConvertToPdf(string path)
        {
            var info = new FileInfo(path);
            if (!info.Exists) throw new FileNotFoundException("파일을 찾을 수 없습니다.", path);
            Directory.CreateDirectory(CacheDirectory);
            string key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(
                $"{info.FullName.ToLowerInvariant()}|{info.Length}|{info.LastWriteTimeUtc.Ticks}")))[..24];
            string output = Path.Combine(CacheDirectory, $"{Path.GetFileNameWithoutExtension(path)}_{key}.pdf");
            if (File.Exists(output) && new FileInfo(output).Length > 0) return output;

            // Office runs one conversion at a time; parallel automation of the same application is unreliable.
            lock (Gate)
            {
                string temporary = Path.Combine(CacheDirectory, $".{Guid.NewGuid():N}.pdf");
                Exception officeError = null;
                try
                {
                    try { RunOnStaThread(() => ConvertWithOffice(info.FullName, temporary)); }
                    catch (Exception ex) { officeError = ex; }
                    if (!File.Exists(temporary) || new FileInfo(temporary).Length == 0)
                    {
                        string soffice = FindLibreOffice();
                        if (soffice != null) ConvertWithLibreOffice(soffice, info.FullName, temporary);
                        else if (officeError is OfficeMissingException)
                            throw new InvalidOperationException(
                                $"{KindName(path)}를 PDF로 바꾸려면 Microsoft Office 또는 LibreOffice가 필요합니다.");
                        else if (officeError != null)
                            throw new InvalidOperationException($"{Path.GetFileName(path)} 변환 실패: {Describe(officeError)}");
                    }
                    if (!File.Exists(temporary) || new FileInfo(temporary).Length == 0)
                        throw new InvalidOperationException($"{Path.GetFileName(path)}을(를) PDF로 변환하지 못했습니다.");
                    File.Move(temporary, output, true);
                    return output;
                }
                finally
                {
                    if (File.Exists(temporary)) File.Delete(temporary);
                }
            }
        }

        private static string Describe(Exception ex)
        {
            while (ex.InnerException != null && ex is System.Reflection.TargetInvocationException) ex = ex.InnerException;
            string message = ex.Message.Replace('\r', ' ').Replace('\n', ' ').Trim();
            if (message.Contains("print", StringComparison.OrdinalIgnoreCase) || message.Contains("인쇄"))
                return "인쇄할 내용이 없습니다(빈 시트).";
            if (message.Contains("password", StringComparison.OrdinalIgnoreCase) || message.Contains("암호"))
                return "암호가 걸린 문서입니다.";
            return message;
        }

        private sealed class OfficeMissingException : Exception { }

        private static void RunOnStaThread(Action action)
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try { action(); }
                catch (Exception ex) { error = ex; }
            }) { IsBackground = true, Name = "Office conversion" };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!thread.Join(Timeout)) throw new TimeoutException("문서 변환 시간이 너무 오래 걸립니다. 문서를 Office에서 직접 열어 확인해 주세요.");
            if (error != null) throw error;
        }

        private static object CreateApplication(string progId)
        {
            var type = Type.GetTypeFromProgID(progId) ?? throw new OfficeMissingException();
            return Activator.CreateInstance(type) ?? throw new OfficeMissingException();
        }

        private static void ConvertWithOffice(string source, string target)
        {
            string extension = Path.GetExtension(source).ToLowerInvariant();
            // A non-empty dummy password makes protected files fail instead of showing a hidden password prompt.
            const string noPassword = "\u0001";
            object missing = System.Reflection.Missing.Value;
            if (WordExtensions.Contains(extension))
            {
                dynamic word = CreateApplication("Word.Application");
                dynamic document = null;
                // Word/PowerPoint may hand back an instance the user already has open: never hide or quit that one.
                bool owned = word.Documents.Count == 0 && !word.Visible;
                try
                {
                    if (owned)
                    {
                        word.Visible = false;
                        word.DisplayAlerts = 0; // wdAlertsNone
                    }
                    // FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert
                    document = word.Documents.Open(source, false, true, false, noPassword, noPassword, true);
                    document.ExportAsFixedFormat(target, 17); // wdExportFormatPDF
                }
                finally
                {
                    try { document?.Close(0); } catch { }
                    try { if (owned && word.Documents.Count == 0) word.Quit(0); } catch { }
                    Release(document);
                    Release(word);
                }
            }
            else if (ExcelExtensions.Contains(extension))
            {
                dynamic excel = CreateApplication("Excel.Application");
                dynamic workbook = null;
                bool owned = excel.Workbooks.Count == 0 && !excel.Visible;
                try
                {
                    if (owned)
                    {
                        excel.Visible = false;
                        excel.DisplayAlerts = false;
                        excel.AskToUpdateLinks = false;
                    }
                    // Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended
                    workbook = excel.Workbooks.Open(source, 0, true, missing, noPassword, noPassword, true);
                    workbook.ExportAsFixedFormat(0, target); // xlTypePDF, whole workbook
                }
                finally
                {
                    try { workbook?.Close(false); } catch { }
                    try { if (owned && excel.Workbooks.Count == 0) excel.Quit(); } catch { }
                    Release(workbook);
                    Release(excel);
                }
            }
            else
            {
                dynamic powerPoint = CreateApplication("PowerPoint.Application");
                dynamic presentation = null;
                bool owned = powerPoint.Presentations.Count == 0;
                try
                {
                    // FileName, ReadOnly (msoTrue), Untitled (msoFalse), WithWindow (msoFalse)
                    presentation = powerPoint.Presentations.Open(source, -1, 0, 0);
                    presentation.SaveAs(target, 32); // ppSaveAsPDF
                }
                finally
                {
                    try { presentation?.Close(); } catch { }
                    try { if (owned && powerPoint.Presentations.Count == 0) powerPoint.Quit(); } catch { }
                    Release(presentation);
                    Release(powerPoint);
                }
            }
        }

        private static void Release(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                try { Marshal.FinalReleaseComObject(value); } catch { }
            }
        }

        private static string FindLibreOffice()
        {
            foreach (var root in new[] { Environment.SpecialFolder.ProgramFiles, Environment.SpecialFolder.ProgramFilesX86 })
            {
                string path = Path.Combine(Environment.GetFolderPath(root), "LibreOffice", "program", "soffice.exe");
                if (File.Exists(path)) return path;
            }
            return null;
        }

        private static void ConvertWithLibreOffice(string soffice, string source, string target)
        {
            string directory = Path.Combine(CacheDirectory, Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                var start = new ProcessStartInfo(soffice)
                {
                    UseShellExecute = false, CreateNoWindow = true,
                    RedirectStandardOutput = true, RedirectStandardError = true,
                };
                foreach (string argument in new[] { "--headless", "--norestore", "--convert-to", "pdf", "--outdir", directory, source })
                    start.ArgumentList.Add(argument);
                using var process = Process.Start(start) ?? throw new InvalidOperationException("LibreOffice를 실행할 수 없습니다.");
                if (!process.WaitForExit((int)Timeout.TotalMilliseconds))
                {
                    try { process.Kill(true); } catch { }
                    throw new TimeoutException("LibreOffice 변환 시간이 초과되었습니다.");
                }
                string produced = Directory.GetFiles(directory, "*.pdf").FirstOrDefault()
                    ?? throw new InvalidOperationException("LibreOffice가 PDF를 만들지 못했습니다.");
                File.Move(produced, target, true);
            }
            finally
            {
                try { Directory.Delete(directory, true); } catch { }
            }
        }
    }
}
