using System;
using System.Collections.Generic;
using System.IO;
using PdfSharp.Fonts;

namespace PDFProcessor
{
    /// <summary>
    /// Supplies Windows system fonts to PDFsharp so replacement text (Korean included) can be embedded.
    /// </summary>
    internal sealed class PdfFontResolver : IFontResolver
    {
        internal const string DefaultFamily = "Malgun Gothic";
        private static readonly object Gate = new();
        private static bool installed;

        private static readonly string FontDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);

        // family -> (regular, bold) file candidates, in order of preference.
        private static readonly Dictionary<string, (string[] Regular, string[] Bold)> Families =
            new(StringComparer.OrdinalIgnoreCase)
            {
                [DefaultFamily] = (new[] { "malgun.ttf", "gulim.ttc", "arial.ttf" }, new[] { "malgunbd.ttf", "malgun.ttf", "arialbd.ttf" }),
                ["Arial"] = (new[] { "arial.ttf" }, new[] { "arialbd.ttf", "arial.ttf" }),
                ["Times New Roman"] = (new[] { "times.ttf" }, new[] { "timesbd.ttf", "times.ttf" }),
                ["Courier New"] = (new[] { "cour.ttf" }, new[] { "courbd.ttf", "cour.ttf" }),
            };

        internal static void EnsureInstalled()
        {
            lock (Gate)
            {
                if (installed) return;
                if (GlobalFontSettings.FontResolver == null) GlobalFontSettings.FontResolver = new PdfFontResolver();
                installed = true;
            }
        }

        /// <summary>Maps an original PDF font name to the closest installed family that covers Korean when needed.</summary>
        internal static string PickFamily(string pdfFontName, string text)
        {
            string name = (pdfFontName ?? "").ToLowerInvariant();
            bool asciiOnly = true;
            foreach (char value in text ?? "") if (value > 0x7E) { asciiOnly = false; break; }
            if (asciiOnly)
            {
                if (name.Contains("times") || name.Contains("serif") && !name.Contains("sans")) return Exists("Times New Roman") ? "Times New Roman" : DefaultFamily;
                if (name.Contains("courier") || name.Contains("mono")) return Exists("Courier New") ? "Courier New" : DefaultFamily;
                if (name.Contains("arial") || name.Contains("helvetica")) return Exists("Arial") ? "Arial" : DefaultFamily;
            }
            return DefaultFamily;
        }

        private static bool Exists(string family) =>
            Families.TryGetValue(family, out var files) && Find(files.Regular) != null;

        private static string Find(IEnumerable<string> candidates)
        {
            foreach (string file in candidates)
            {
                string path = Path.Combine(FontDirectory, file);
                if (File.Exists(path)) return path;
            }
            return null;
        }

        public FontResolverInfo ResolveTypeface(string familyName, bool bold, bool italic)
        {
            if (!Families.TryGetValue(familyName ?? "", out var files)) files = Families[DefaultFamily];
            string path = Find(bold ? files.Bold : files.Regular) ?? Find(Families[DefaultFamily].Regular);
            if (path == null) return null;
            bool simulateBold = bold && !Path.GetFileNameWithoutExtension(path).EndsWith("bd", StringComparison.OrdinalIgnoreCase);
            return new FontResolverInfo(path, simulateBold, italic);
        }

        public byte[] GetFont(string faceName)
        {
            byte[] bytes = File.ReadAllBytes(faceName);
            if (!faceName.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase)) return bytes;
            return ExtractFirstFace(bytes);
        }

        // PDFsharp expects a single TrueType face. Rebuild the first face of a collection as a standalone font.
        private static byte[] ExtractFirstFace(byte[] collection)
        {
            uint ReadU32(byte[] data, int offset) => (uint)(data[offset] << 24 | data[offset + 1] << 16 | data[offset + 2] << 8 | data[offset + 3]);
            ushort ReadU16(byte[] data, int offset) => (ushort)(data[offset] << 8 | data[offset + 1]);
            int faceOffset = (int)ReadU32(collection, 12);
            int tableCount = ReadU16(collection, faceOffset + 4);
            int headerLength = 12 + tableCount * 16;
            var tables = new List<(int Record, int Offset, int Length)>();
            for (int index = 0; index < tableCount; index++)
            {
                int record = faceOffset + 12 + index * 16;
                tables.Add((record, (int)ReadU32(collection, record + 8), (int)ReadU32(collection, record + 12)));
            }
            using var output = new MemoryStream();
            output.Write(collection, faceOffset, headerLength);
            var header = output.ToArray();
            int position = headerLength;
            var body = new MemoryStream();
            for (int index = 0; index < tables.Count; index++)
            {
                var table = tables[index];
                int recordInHeader = 12 + index * 16;
                header[recordInHeader + 8] = (byte)(position >> 24);
                header[recordInHeader + 9] = (byte)(position >> 16);
                header[recordInHeader + 10] = (byte)(position >> 8);
                header[recordInHeader + 11] = (byte)position;
                body.Write(collection, table.Offset, table.Length);
                position += table.Length;
                while (position % 4 != 0) { body.WriteByte(0); position++; }
            }
            var result = new MemoryStream();
            result.Write(header, 0, header.Length);
            body.Position = 0;
            body.CopyTo(result);
            return result.ToArray();
        }
    }
}
