using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace PDFProcessor
{
    internal sealed class PdfTextRun
    {
        public string Text { get; internal set; }
        public string FontName { get; internal set; }
        public double FontSize { get; internal set; }
        public string Label => $"{Text}  ·  {FontName}";
        internal COperator Operator;
        internal Dictionary<string, string> Glyphs;
        internal string OriginalText;
        internal string FontKey;
        internal Dictionary<string, double> GlyphWidths;
        internal double CharacterSpacing;
        internal double WordSpacing;
        internal Dictionary<string, string> ReverseGlyphs;
        internal int[] UnicodeLengths;
        internal int[] CodeLengths;
    }

    internal sealed partial class PdfTextEditor
    {
        internal readonly PdfDocument Document;
        private readonly Dictionary<int, CSequence> contents = new();
        private readonly Dictionary<int, List<PdfTextRun>> runs = new();
        private readonly Dictionary<PdfDictionary, FontData> fontCache = new();

        private sealed class FontData
        {
            internal Dictionary<string, string> Glyphs, Reverse;
            internal Dictionary<string, double> Widths;
            internal int[] UnicodeLengths, CodeLengths;
        }

        private FontData GetFontData(PdfDictionary font)
        {
            if (fontCache.TryGetValue(font, out var cached)) return cached;
            var glyphs = ReadGlyphs(font);
            var reverse = glyphs.GroupBy(pair => pair.Value).ToDictionary(group => group.Key, group => group.First().Key);
            var data = new FontData { Glyphs = glyphs, Reverse = reverse, Widths = ReadWidths(font, glyphs.Keys),
                UnicodeLengths = reverse.Keys.Select(key => key.Length).Where(length => length > 0).Distinct().OrderByDescending(length => length).ToArray(),
                CodeLengths = glyphs.Keys.Select(key => key.Length).Distinct().OrderByDescending(length => length).ToArray() };
            fontCache.Add(font, data);
            return data;
        }

        internal PdfTextEditor(PdfDocument document) { Document = document; }

        internal List<PdfTextRun> GetRuns(int pageIndex)
        {
            if (runs.TryGetValue(pageIndex, out var cached)) return cached;
            var page = Document.Pages[pageIndex];
            var sequence = ContentReader.ReadContent(page);
            PrepareStrings(sequence);
            var result = new List<PdfTextRun>();
            var fonts = page.Resources.Elements.GetDictionary("/Font");
            string currentFont = null;
            double currentSize = 12;
            double currentCharacterSpacing = 0, currentWordSpacing = 0;
            var stack = new Stack<(string Font, double Size, double Character, double Word)>();
            foreach (var item in sequence)
            {
                if (item is not COperator op) continue;
                if (op.Name == "q") stack.Push((currentFont, currentSize, currentCharacterSpacing, currentWordSpacing));
                else if (op.Name == "Q" && stack.Count > 0)
                {
                    var state = stack.Pop();
                    currentFont = state.Font;
                    currentSize = state.Size;
                    currentCharacterSpacing = state.Character;
                    currentWordSpacing = state.Word;
                }
                else if (op.Name == "Tf" && op.Operands.Count > 0 && op.Operands[0] is CName name)
                {
                    currentFont = name.Name;
                    if (op.Operands.Count > 1 && double.TryParse(op.Operands[1].ToString(),
                        System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double size))
                        currentSize = Math.Abs(size);
                }
                else if ((op.Name == "Tc" || op.Name == "Tw") && op.Operands.Count > 0)
                {
                    double spacing = ContentNumber(op.Operands[0]);
                    if (op.Name == "Tc") currentCharacterSpacing = spacing;
                    else currentWordSpacing = spacing;
                }
                else if ((op.Name == "Tj" || op.Name == "TJ" || op.Name == "'" || op.Name == "\"") &&
                    currentFont != null && fonts != null)
                {
                    var font = fonts.Elements.GetDictionary(currentFont);
                    if (font == null) continue;
                    var fontData = GetFontData(font);
                    var glyphs = fontData.Glyphs;
                    if (glyphs.Count == 0) continue;
                    string decoded = DecodeOperator(op, glyphs, fontData.CodeLengths);
                    if (string.IsNullOrWhiteSpace(decoded)) continue;
                    if (op.Name == "\"" && op.Operands.Count >= 3)
                    {
                        currentWordSpacing = ContentNumber(op.Operands[0]);
                        currentCharacterSpacing = ContentNumber(op.Operands[1]);
                    }
                    result.Add(new PdfTextRun { Text = decoded, OriginalText = decoded,
                        FontName = font.Elements.GetName("/BaseFont"), FontSize = currentSize, Operator = op, Glyphs = glyphs,
                        FontKey = currentFont, GlyphWidths = fontData.Widths,
                        CharacterSpacing = currentCharacterSpacing, WordSpacing = currentWordSpacing,
                        ReverseGlyphs = fontData.Reverse, UnicodeLengths = fontData.UnicodeLengths,
                        CodeLengths = fontData.CodeLengths });
                }
            }
            contents[pageIndex] = sequence;
            runs[pageIndex] = result;
            return result;
        }

        internal static double MeasureText(PdfTextRun run, string text)
        {
            string encoded = Encode(run, text);
            return RawAdvance(run, encoded) * run.FontSize / 1000;
        }

        private static string Encode(PdfTextRun run, string text)
        {
            var encoded = new StringBuilder();
            for (int index = 0; index < text.Length;)
            {
                int length = run.UnicodeLengths.FirstOrDefault(size => index + size <= text.Length && run.ReverseGlyphs.ContainsKey(text.Substring(index, size)));
                string match = length == 0 ? null : text.Substring(index, length);
                if (match == null) throw new InvalidOperationException($"원본 글꼴에 '{text[index]}' 글자가 없습니다. 다른 글꼴로 바꾸지 않고는 입력할 수 없습니다.");
                encoded.Append(run.ReverseGlyphs[match]);
                index += match.Length;
            }
            return encoded.ToString();
        }

        private static double RawAdvance(PdfTextRun run, string raw)
        {
            double result = 0;
            for (int index = 0; index < raw.Length;)
            {
                int length = run.CodeLengths.First(size => index + size <= raw.Length && run.Glyphs.ContainsKey(raw.Substring(index, size)));
                string code = raw.Substring(index, length);
                result += run.GlyphWidths[code] + run.CharacterSpacing * 1000 / Math.Max(1, run.FontSize);
                if (code == " ") result += run.WordSpacing * 1000 / Math.Max(1, run.FontSize);
                index += code.Length;
            }
            return result;
        }

        private static double ContentNumber(CObject value) => value is CInteger integer ? integer.Value : value is CReal real ? real.Value : 0;
        private static double PdfNumber(PdfItem value) => value is PdfInteger integer ? integer.Value : value is PdfReal real ? real.Value : 0;

        private static Dictionary<string, double> ReadWidths(PdfDictionary font, IEnumerable<string> codes)
        {
            var descendantArray = font.Elements.GetArray("/DescendantFonts");
            bool composite = descendantArray != null;
            if (composite)
            {
                var item = descendantArray.Elements[0];
                font = item is PdfSharp.Pdf.Advanced.PdfReference reference ? (PdfDictionary)reference.Value : (PdfDictionary)item;
            }
            var result = new Dictionary<string, double>();
            var widths = font.Elements.GetArray(composite ? "/W" : "/Widths");
            double defaultWidth = composite ? (font.Elements.ContainsKey("/DW") ? font.Elements.GetInteger("/DW") : 1000) : 500;
            foreach (string code in codes)
            {
                int cid = 0;
                foreach (char value in code) cid = (cid << 8) | value;
                double width = defaultWidth;
                if (widths != null && !composite)
                {
                    int offset = cid - font.Elements.GetInteger("/FirstChar");
                    if (offset >= 0 && offset < widths.Elements.Count) width = PdfNumber(widths.Elements[offset]);
                }
                else if (widths != null)
                {
                    for (int index = 0; index + 1 < widths.Elements.Count;)
                    {
                        int first = (int)PdfNumber(widths.Elements[index++]);
                        if (widths.Elements[index] is PdfArray array)
                        {
                            index++;
                            if (cid >= first && cid < first + array.Elements.Count) { width = PdfNumber(array.Elements[cid - first]); break; }
                        }
                        else
                        {
                            int last = (int)PdfNumber(widths.Elements[index++]);
                            if (index >= widths.Elements.Count) break;
                            double shared = PdfNumber(widths.Elements[index++]);
                            if (cid >= first && cid <= last) { width = shared; break; }
                        }
                    }
                }
                // Standard fonts without a Widths array: use their defined metrics where available.
                if (!composite && widths == null)
                {
                    width = StandardFontWidths.Get(font.Elements.GetName("/BaseFont"), cid, font.Elements.GetName("/Encoding"));
                }
                result[code] = width;
            }
            return result;
        }

        internal void Replace(int pageIndex, PdfTextRun run, string replacement)
        {
            if (!GetRuns(pageIndex).Contains(run)) throw new InvalidOperationException("현재 페이지의 텍스트를 선택하세요.");
            if (replacement.Contains('\n') || replacement.Contains('\r'))
                throw new InvalidOperationException("선택한 텍스트 구간은 한 줄로 수정하세요.");
            // Encode through the existing font's character codes. Never add or substitute a font.
            var reverse = run.Glyphs.GroupBy(pair => pair.Value)
                .ToDictionary(group => group.Key, group => group.First().Key);
            var encoded = new StringBuilder();
            for (int index = 0; index < replacement.Length;)
            {
                string match = reverse.Keys.Where(text => text.Length > 0 &&
                        index + text.Length <= replacement.Length &&
                        string.CompareOrdinal(replacement, index, text, 0, text.Length) == 0)
                    .OrderByDescending(text => text.Length).FirstOrDefault();
                if (match == null)
                    throw new InvalidOperationException($"원본 글꼴에 '{replacement[index]}' 글자가 포함되어 있지 않습니다.\n글꼴을 유지하기 위해 이 변경은 적용하지 않았습니다.");
                encoded.Append(reverse[match]);
                index += match.Length;
            }
            var value = new SerializableString { Value = encoded.ToString(), CStringType = CStringType.String };
            var op = run.Operator;
            if (op.Name == "TJ")
            {
                var array = new CArray();
                array.Add(value);
                op.Operands[0] = array;
            }
            else op.Operands[op.Operands.Count - 1] = value;
            Document.Pages[pageIndex].Contents.ReplaceContent(contents[pageIndex]);
            run.Text = replacement;
        }

        private static IEnumerable<CString> GetStrings(COperator op)
        {
            foreach (var operand in op.Operands)
            {
                if (operand is CString text) yield return text;
                else if (operand is CArray array)
                    foreach (var child in array) if (child is CString part) yield return part;
            }
        }

        private static string Decode(string raw, Dictionary<string, string> glyphs, int[] lengths)
        {
            var text = new StringBuilder();
            for (int index = 0; index < raw.Length;)
            {
                int length = lengths.FirstOrDefault(size => index + size <= raw.Length &&
                    glyphs.ContainsKey(raw.Substring(index, size)));
                if (length == 0) return null;
                text.Append(glyphs[raw.Substring(index, length)]);
                index += length;
            }
            return text.ToString();
        }

        private static string DecodeOperator(COperator op, Dictionary<string, string> glyphs, int[] lengths)
        {
            if (op.Name != "TJ")
                return Decode(((CString)op.Operands[op.Operands.Count - 1]).Value, glyphs, lengths);
            var array = (CArray)op.Operands[0];
            var result = new StringBuilder();
            for (int index = 0; index < array.Count; index++)
            {
                if (array[index] is CString raw)
                {
                    string text = Decode(raw.Value, glyphs, lengths);
                    if (text == null) return null;
                    result.Append(text);
                }
                else if (IsPositionedSpace(array, index, result.ToString(), glyphs, lengths)) result.Append(' ');
            }
            return result.ToString();
        }

        private static bool IsPositionedSpace(CArray array, int index, string previous,
            Dictionary<string, string> glyphs, int[] lengths)
        {
            if (ContentNumber(array[index]) > -250 || previous.Length == 0 || char.IsWhiteSpace(previous[previous.Length - 1])) return false;
            for (int next = index + 1; next < array.Count; next++)
            {
                if (array[next] is not CString raw || raw.Value.Length == 0) continue;
                string text = Decode(raw.Value, glyphs, lengths);
                return !string.IsNullOrEmpty(text) && !char.IsWhiteSpace(text[0]);
            }
            return false; // A trailing adjustment preserves positions; it is not visible text.
        }

        private static string RawCode(string hex) => new string(Convert.FromHexString(hex).Select(value => (char)value).ToArray());
        private static string Unicode(string hex) => Encoding.BigEndianUnicode.GetString(Convert.FromHexString(hex));

        // PDFsharp 6.1 does not serialize parsed hexadecimal strings. Preserve their bytes explicitly.
        private sealed class SerializableString : CString
        {
            public override string ToString()
            {
                byte[] bytes = CStringType == CStringType.UnicodeString || CStringType == CStringType.UnicodeHexString
                    ? new byte[] { 0xFE, 0xFF }.Concat(Encoding.BigEndianUnicode.GetBytes(Value)).ToArray()
                    : Value.Select(character => checked((byte)character)).ToArray();
                return "<" + Convert.ToHexString(bytes) + ">";
            }
        }

        private static void PrepareStrings(CSequence sequence)
        {
            for (int index = 0; index < sequence.Count; index++)
            {
                if (sequence[index] is CString value)
                    sequence[index] = new SerializableString { Value = value.Value, CStringType = value.CStringType };
                else if (sequence[index] is COperator op) PrepareStrings(op.Operands);
                else if (sequence[index] is CArray array) PrepareStrings(array);
            }
        }

        private static Dictionary<string, string> ReadGlyphs(PdfDictionary font)
        {
            var result = new Dictionary<string, string>();
            var map = font.Elements.GetDictionary("/ToUnicode");
            if (map?.Stream != null)
            {
                string cmap = Encoding.ASCII.GetString(map.Stream.UnfilteredValue);
                foreach (Match block in Regex.Matches(cmap, @"beginbfchar(.*?)endbfchar", RegexOptions.Singleline))
                    foreach (Match pair in Regex.Matches(block.Groups[1].Value, @"<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>"))
                        result[RawCode(pair.Groups[1].Value)] = Unicode(pair.Groups[2].Value);
                foreach (Match block in Regex.Matches(cmap, @"beginbfrange(.*?)endbfrange", RegexOptions.Singleline))
                foreach (Match range in Regex.Matches(block.Groups[1].Value,
                    @"<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<([0-9a-fA-F]+)>|\[(.*?)\])", RegexOptions.Singleline))
                {
                    int start = Convert.ToInt32(range.Groups[1].Value, 16);
                    int end = Convert.ToInt32(range.Groups[2].Value, 16);
                    if (end - start > 65535 || end < start) continue;
                    var targets = Regex.Matches(range.Groups[5].Value, @"<([0-9a-fA-F]+)>");
                    for (int code = start; code <= end; code++)
                    {
                        string source = code.ToString("X" + range.Groups[1].Value.Length);
                        string target;
                        if (range.Groups[4].Success)
                        {
                            var bytes = Convert.FromHexString(range.Groups[4].Value);
                            int carry = code - start;
                            for (int i = bytes.Length - 1; i >= 0 && carry > 0; i--)
                            { int sum = bytes[i] + carry; bytes[i] = (byte)sum; carry = sum >> 8; }
                            target = Encoding.BigEndianUnicode.GetString(bytes);
                        }
                        else
                        {
                            if (code - start >= targets.Count) break;
                            target = Unicode(targets[code - start].Groups[1].Value);
                        }
                        result[RawCode(source)] = target;
                    }
                }
            }
            else
            {
                string name = font.Elements.GetName("/BaseFont").TrimStart('/');
                string encoding = font.Elements.GetName("/Encoding");
                // The standard Latin fonts have defined ASCII glyphs. Unknown encodings stay read-only.
                if ((name.StartsWith("Helvetica") || name.StartsWith("Times-") || name.StartsWith("Courier")) &&
                    (encoding == "/WinAnsiEncoding" || encoding == "/StandardEncoding" || string.IsNullOrEmpty(encoding)))
                    for (int code = 32; code < 127; code++) result[((char)code).ToString()] = ((char)code).ToString();
            }
            return result;
        }
    }
}
