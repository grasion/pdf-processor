using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;

namespace PDFProcessor
{
    internal sealed partial class PdfTextEditor
    {
        private sealed class SentenceEdit
        {
            internal PdfTextRun Run;
            internal int Start, Length;
            internal string Text;
        }

        private sealed class TextAtom
        {
            internal int Start, Length;
            internal CObject Value;
            internal double Advance;
        }

        internal void ReplaceSentence(int pageIndex, PdfSentence sentence, string text)
        {
            var pageRuns = GetRuns(pageIndex);
            if (sentence.Parts.Any(part => !pageRuns.Contains(part.Run)))
                throw new InvalidOperationException("현재 페이지의 문장을 다시 선택하세요.");
            text = System.Text.RegularExpressions.Regex.Replace(text.Trim(), @"\s+", " ");
            var segments = sentence.GetSegments();
            var available = sentence.Parts.SelectMany(part => part.Run.ReverseGlyphs.Keys).ToHashSet();
            var availableLengths = available.Select(value => value.Length).Where(value => value > 0).Distinct().OrderByDescending(value => value).ToArray();
            for (int offset = 0; offset < text.Length;)
            {
                int length = availableLengths.FirstOrDefault(value => offset + value <= text.Length && available.Contains(text.Substring(offset, value)));
                if (text[offset] == ' ') { offset++; continue; }
                if (length == 0) throw new InvalidOperationException($"원본 문장의 글꼴에 '{text[offset]}' 글자가 없습니다.");
                offset += length;
            }
            var edits = new List<SentenceEdit>();
            string remaining = text;
            for (int index = 0; index < segments.Count; index++)
            {
                var segment = segments[index];
                var first = segment.Parts[0];
                // Keep every line within the width of its original sentence segment.
                double capacity = SentenceAdvance(first.Run, segment.Text) * first.Run.FontSize / 1000;
                int count = 0;
                while (count < remaining.Length)
                {
                    int next = count + (char.IsHighSurrogate(remaining[count]) && count + 1 < remaining.Length ? 2 : 1);
                    double width;
                    try { width = SentenceAdvance(first.Run, remaining.Substring(0, next)) * first.Run.FontSize / 1000; }
                    catch (InvalidOperationException) { break; } // Continue in the next original font/style segment.
                    if (width > capacity) break;
                    count = next;
                }
                if (count < remaining.Length && index + 1 < segments.Count &&
                    !PdfSentences.SameLine(segment.Bounds, segments[index + 1].Bounds))
                {
                    // Prefer a word boundary when the sentence wraps to its next original line.
                    int space = remaining.LastIndexOf(' ', Math.Max(0, count - 1), count);
                    if (space > 0) count = space + 1;
                }
                string replacement = remaining.Substring(0, count).TrimEnd();
                remaining = remaining.Substring(count).TrimStart();
                edits.Add(new SentenceEdit { Run = first.Run, Start = first.Start, Length = first.Length, Text = replacement });
                foreach (var part in segment.Parts.Skip(1))
                    edits.Add(new SentenceEdit { Run = part.Run, Start = part.Start, Length = part.Length, Text = "" });
            }
            if (remaining.Length > 0)
                throw new InvalidOperationException("수정한 문장이 원본 문장 영역보다 깁니다. 원본 배치를 유지하려면 문장을 조금 줄여주세요.");

            // Validate every font and build every replacement before changing any page content.
            var sequence = contents[pageIndex];
            var plans = edits.GroupBy(edit => edit.Run).Select(group =>
            {
                var runEdits = group.OrderBy(edit => edit.Start).ToList();
                return (Run: group.Key, Edits: runEdits, Index: sequence.IndexOf(group.Key.Operator),
                    Operators: BuildSentenceOperators(group.Key, runEdits));
            }).ToList();
            foreach (var plan in plans.OrderByDescending(plan => plan.Index))
            {
                sequence.RemoveAt(plan.Index);
                for (int index = 0; index < plan.Operators.Count; index++) sequence.Insert(plan.Index + index, plan.Operators[index]);
                plan.Run.Operator = plan.Operators.Last();
                foreach (var edit in plan.Edits.OrderByDescending(edit => edit.Start))
                    plan.Run.Text = plan.Run.Text.Remove(edit.Start, edit.Length).Insert(edit.Start, edit.Text);
            }
            Document.Pages[pageIndex].Contents.ReplaceContent(sequence);
        }

        private static List<COperator> BuildSentenceOperators(PdfTextRun run, List<SentenceEdit> edits)
        {
            var atoms = GetAtoms(run);
            var array = new CArray();
            int cursor = 0;
            foreach (var edit in edits)
            {
                int end = edit.Start + edit.Length;
                while (cursor < atoms.Count && (atoms[cursor].Start < edit.Start ||
                    (atoms[cursor].Length == 0 && atoms[cursor].Start == edit.Start))) array.Add(atoms[cursor++].Value);
                if (cursor < atoms.Count && atoms[cursor].Start != edit.Start)
                    throw new InvalidOperationException("이 글꼴의 연결된 문자는 문장 경계에서 나눌 수 없습니다.");
                double originalAdvance = 0;
                while (cursor < atoms.Count && atoms[cursor].Start < end)
                {
                    if (atoms[cursor].Start + atoms[cursor].Length > end)
                        throw new InvalidOperationException("이 글꼴의 연결된 문자는 문장 경계에서 나눌 수 없습니다.");
                    originalAdvance += atoms[cursor++].Advance;
                }
                foreach (var value in EncodeSentence(run, edit.Text)) array.Add(value);
                // TJ adjustments keep the following sentence and existing text operators at their old positions.
                double adjustment = SentenceAdvance(run, edit.Text) - originalAdvance;
                if (Math.Abs(adjustment) > 0.001) array.Add(new CReal { Value = adjustment });
            }
            while (cursor < atoms.Count) array.Add(atoms[cursor++].Value);
            var op = (COperator)ContentReader.ReadContent(Encoding.ASCII.GetBytes("[] TJ"))[0];
            op.Operands.Clear();
            op.Operands.Add((CObject)array);
            var result = new List<COperator>();
            if (run.Operator.Name == "'" || run.Operator.Name == "\"")
            {
                string prefix = run.Operator.Name == "\""
                    ? run.WordSpacing.ToString(CultureInfo.InvariantCulture) + " Tw " + run.CharacterSpacing.ToString(CultureInfo.InvariantCulture) + " Tc T*"
                    : "T*";
                foreach (var item in ContentReader.ReadContent(Encoding.ASCII.GetBytes(prefix))) result.Add((COperator)item);
            }
            result.Add(op);
            return result;
        }

        private static List<TextAtom> GetAtoms(PdfTextRun run)
        {
            var values = run.Operator.Name == "TJ" ? (CArray)run.Operator.Operands[0]
                : new CArray();
            if (run.Operator.Name != "TJ") values.Add(run.Operator.Operands[run.Operator.Operands.Count - 1]);
            var result = new List<TextAtom>();
            int offset = 0;
            var decoded = new StringBuilder();
            for (int valueIndex = 0; valueIndex < values.Count; valueIndex++)
            {
                var value = values[valueIndex];
                if (value is not CString raw)
                {
                    bool space = IsPositionedSpace(values, valueIndex, decoded.ToString(), run.Glyphs, run.CodeLengths);
                    result.Add(new TextAtom { Start = offset, Length = space ? 1 : 0, Value = value, Advance = -ContentNumber(value) });
                    if (space) { offset++; decoded.Append(' '); }
                    continue;
                }
                for (int index = 0; index < raw.Value.Length;)
                {
                    int length = run.CodeLengths.First(size => index + size <= raw.Value.Length && run.Glyphs.ContainsKey(raw.Value.Substring(index, size)));
                    string code = raw.Value.Substring(index, length);
                    string text = run.Glyphs[code];
                    result.Add(new TextAtom { Start = offset, Length = text.Length,
                        Value = new SerializableString { Value = code, CStringType = CStringType.String }, Advance = RawAdvance(run, code) });
                    offset += text.Length;
                    decoded.Append(text);
                    index += code.Length;
                }
            }
            return result;
        }

        private static double PositionedSpaceAdvance(PdfTextRun run) =>
            1000.0 / 3 + (run.CharacterSpacing + run.WordSpacing) * 1000 / Math.Max(1, run.FontSize);

        private static double SentenceAdvance(PdfTextRun run, string text)
        {
            if (run.ReverseGlyphs.ContainsKey(" ")) return RawAdvance(run, Encode(run, text));
            var parts = text.Split(' ');
            return parts.Sum(part => RawAdvance(run, Encode(run, part))) + (parts.Length - 1) * PositionedSpaceAdvance(run);
        }

        private static IEnumerable<CObject> EncodeSentence(PdfTextRun run, string text)
        {
            if (run.ReverseGlyphs.ContainsKey(" "))
            {
                yield return new SerializableString { Value = Encode(run, text), CStringType = CStringType.String };
                yield break;
            }
            // PDFs may position words without embedding a space glyph. Express spaces as TJ advances using the same font.
            var parts = text.Split(' ');
            for (int index = 0; index < parts.Length; index++)
            {
                if (index > 0) yield return new CReal { Value = -PositionedSpaceAdvance(run) };
                yield return new SerializableString { Value = Encode(run, parts[index]), CStringType = CStringType.String };
            }
        }
    }
}
