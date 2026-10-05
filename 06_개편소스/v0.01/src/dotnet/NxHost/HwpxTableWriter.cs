using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml;

namespace LH.NxHost {
    // Pure managed port of the product's embedded table generator. No Excel objects cross this boundary.
    internal sealed class HwpxTableWriter {
        internal const int MaxPayloadChars = 16 * 1024 * 1024;
        internal const string TemplateHash = "9a8446816cc377c8f848a0666f1e7e8e38320311e248c4c2594c40ca896cf44b";
        private const string HP = "http://www.hancom.co.kr/hwpml/2011/paragraph";
        private const string HH = "http://www.hancom.co.kr/hwpml/2011/head";
        private const string HC = "http://www.hancom.co.kr/hwpml/2011/core";
        private const string HS = "http://www.hancom.co.kr/hwpml/2011/section";
        private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
        private readonly CancellationToken cancellation;
        private int rows, cols, fontSize = 13;
        private bool title = true, simple;
        private string font = "중고딕";
        private int[] widths, heights;
        private string[,] values;
        private readonly Dictionary<int, int[]> merges = new Dictionary<int, int[]>();
        private readonly HashSet<int> covered = new HashSet<int>();

        private HwpxTableWriter(CancellationToken token) { cancellation = token; }
        internal static void Write(string payloadJson, string templatePath, string outputPath, CancellationToken token) {
            if (HwpxPictureWriter.IsPicturePayload(payloadJson)) {
                HwpxPictureWriter.Write(payloadJson, templatePath, outputPath, token);
                return;
            }
            var writer = new HwpxTableWriter(token);
            writer.Parse(payloadJson);
            token.ThrowIfCancellationRequested();
            // Read once, hash those exact bytes, then parse; replacement after validation cannot race the reader.
            if (new FileInfo(templatePath).Length > 4 * 1024 * 1024) throw new InvalidDataException("TEMPLATE_SIZE");
            byte[] template = File.ReadAllBytes(templatePath);
            using (var sha = SHA256.Create()) {
                if (BitConverter.ToString(sha.ComputeHash(template)).Replace("-", "").ToLowerInvariant() != TemplateHash)
                    throw new InvalidDataException("TEMPLATE_HASH");
            }
            var entries = new List<KeyValuePair<string, byte[]>>();
            using (var memory = new MemoryStream(template)) using (var zip = new ZipArchive(memory, ZipArchiveMode.Read)) {
                entries.Add(new KeyValuePair<string, byte[]>("mimetype", Read(zip.GetEntry("mimetype"))));
                foreach (var entry in zip.Entries) {
                    token.ThrowIfCancellationRequested();
                    if (entry.FullName == "mimetype") continue;
                    byte[] bytes = Read(entry);
                    if (entry.FullName == "Contents/section0.xml") bytes = writer.Section(bytes);
                    else if (entry.FullName == "Contents/header.xml") bytes = writer.Header(bytes);
                    else if (entry.FullName == "Preview/PrvText.txt") bytes = writer.Preview();
                    entries.Add(new KeyValuePair<string, byte[]>(entry.FullName, bytes));
                }
            }
            bool created = false;
            try {
                using (var output = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    created = true;
                    writer.StoredZip(entries, output);
                    output.Flush(true);
                }
                token.ThrowIfCancellationRequested();
            } catch { if (created) try { File.Delete(outputPath); } catch { } throw; }
        }
        private static Dictionary<string, object> Map(object value) {
            var result = value as Dictionary<string, object>;
            if (result == null) throw new InvalidDataException("INPUT_OBJECT");
            return result;
        }
        private static object Get(Dictionary<string, object> map, params string[] keys) {
            foreach (string key in keys) { object value; if (map.TryGetValue(key, out value)) return value; }
            return null;
        }
        private static int Integer(object value) {
            if (value == null || value is bool || value is string) throw new InvalidDataException("INPUT_INTEGER");
            double number = Convert.ToDouble(value, Invariant);
            if (number != Math.Truncate(number) || number < 0 || number > 16384) throw new InvalidDataException("INPUT_INTEGER");
            return (int)number;
        }
        private static IList Items(Dictionary<string, object> map, params string[] keys) {
            object value = Get(map, keys);
            if (value == null) return new object[0];
            var list = value as IList;
            if (list == null) throw new InvalidDataException("INPUT_ARRAY");
            return list;
        }
        private static void Only(Dictionary<string, object> map, string allowed) {
            foreach (string key in map.Keys) if (Array.IndexOf(allowed.Split('|'), key) < 0) throw new InvalidDataException("INPUT_FIELD: " + key);
        }
        private void Parse(string json) {
            if (String.IsNullOrEmpty(json) || json.Length > MaxPayloadChars) throw new InvalidDataException("INPUT_SIZE");
            var serializer = new JavaScriptSerializer { MaxJsonLength = MaxPayloadChars, RecursionLimit = 32 };
            var raw = Map(serializer.DeserializeObject(json.TrimStart('\uFEFF')));
            Only(raw, "schema_version|rows|rowCount|columns|cols|columnCount|column_widths|row_heights|rowHeights|rowsMeta|cells|merges|settings|source_mode|title_row");
            if (Integer(Get(raw, "schema_version")) != 1) throw new InvalidDataException("INPUT_SCHEMA");
            rows = Integer(Get(raw, "rows", "rowCount"));
            object columns = Get(raw, "columns");
            cols = Integer(columns != null && !(columns is IList) ? columns : Get(raw, "cols", "columnCount"));
            if (rows < 1 || cols < 1 || (long)rows * cols > 16384) throw new InvalidDataException("INPUT_DIMENSIONS");
            object settings = Get(raw, "settings");
            if (settings != null) {
                var s = Map(settings);
                Only(s, "font_name|font_type|font_size_pt|title_row|include_hidden|table_style");
                if (s.ContainsKey("font_name")) font = Convert.ToString(s["font_name"], Invariant);
                if (Array.IndexOf(new [] { "중고딕", "함초롬돋움", "맑은 고딕", "굴림" }, font) < 0) throw new InvalidDataException("SETTINGS_FONT");
                if (s.ContainsKey("font_type") && (string)s["font_type"] != "HFT") throw new InvalidDataException("SETTINGS_FONT_TYPE");
                if (s.ContainsKey("font_size_pt")) fontSize = Integer(s["font_size_pt"]);
                if (fontSize < 9 || fontSize > 24) throw new InvalidDataException("SETTINGS_FONT_SIZE");
                if (s.ContainsKey("title_row")) title = (bool)s["title_row"];
                if (s.ContainsKey("include_hidden") && !(s["include_hidden"] is bool)) throw new InvalidDataException("SETTINGS_HIDDEN");
                if (s.ContainsKey("table_style")) {
                    string style = (string)s["table_style"];
                    if (style != "공문형" && style != "간결형") throw new InvalidDataException("SETTINGS_STYLE");
                    simple = style == "간결형";
                }
            }
            values = new string[rows, cols];
            var occupiedCells = new HashSet<int>();
            foreach (object item in Items(raw, "cells")) {
                cancellation.ThrowIfCancellationRequested();
                var cell = Map(item);
                int r = Integer(Get(cell, "row", "r")) - 1, c = Integer(Get(cell, "column", "c")) - 1;
                if (r < 0 || r >= rows || c < 0 || c >= cols || !occupiedCells.Add(r * cols + c)) throw new InvalidDataException("INPUT_CELL");
                string text = Convert.ToString(Get(cell, "text", "v"), Invariant) ?? "";
                if (text.Length > 32767) throw new InvalidDataException("INPUT_CELL_TEXT");
                XmlConvert.VerifyXmlChars(text);
                // Remove Excel display-format padding, preserving internal spacing/newlines.
                values[r, c] = text.Trim(' ', '\t', '\u00a0', '\u3000');
            }
            var occupiedMerges = new HashSet<int>();
            foreach (object item in Items(raw, "merges")) {
                cancellation.ThrowIfCancellationRequested();
                var merge = Map(item);
                int r = Integer(Get(merge, "row")) - 1, c = Integer(Get(merge, "column")) - 1;
                int rs = Integer(Get(merge, "row_span")), cs = Integer(Get(merge, "column_span"));
                if (r < 0 || c < 0 || rs < 1 || cs < 1 || r + rs > rows || c + cs > cols) throw new InvalidDataException("INPUT_MERGE");
                for (int y = r; y < r + rs; y++) for (int x = c; x < c + cs; x++) {
                    int key = y * cols + x;
                    if (!occupiedMerges.Add(key)) throw new InvalidDataException("INPUT_MERGE_OVERLAP");
                    if (y != r || x != c) covered.Add(key);
                }
                merges.Add(r * cols + c, new [] { rs, cs });
            }
            IList widthItems = Get(raw, "column_widths") != null ? Items(raw, "column_widths") : columns as IList ?? new object[0];
            var shares = new double[cols]; double sum = 0;
            for (int c = 0; c < cols; c++) { shares[c] = c < widthItems.Count ? Dimension(widthItems[c], "width", .1) : 1; sum += shares[c]; }
            double min = Math.Min(.08, .7 / cols), max = Math.Max(.18, Math.Min(.42, 1 - (cols - 1) * min)), adjusted = 0;
            for (int c = 0; c < cols; c++) { shares[c] = Math.Min(Math.Max(shares[c] / sum, min), max); adjusted += shares[c]; }
            widths = new int[cols]; int totalWidth = 0;
            for (int c = 0; c < cols; c++) { widths[c] = Math.Max(1, (int)Math.Round(47341 * shares[c] / adjusted)); totalWidth += widths[c]; }
            widths[cols - 1] += 47341 - totalWidth;
            if (widths[cols - 1] <= 0) throw new InvalidDataException("INPUT_WIDTH");
            IList heightItems = Items(raw, "row_heights", "rowHeights", "rowsMeta");
            heights = new int[rows];
            for (int r = 0; r < rows; r++) {
                int headerHeight = simple ? 2000 : 2200;
                heights[r] = r == 0 && title ? headerHeight : 1865;
                if (r < heightItems.Count) {
                    int minimum = r == 0 && (title || heightItems[r] is Dictionary<string, object>) ? headerHeight : 1500;
                    heights[r] = Math.Max(minimum, (int)Math.Round(Dimension(heightItems[r], "height", 8) * 100));
                }
            }
        }
        private static double Dimension(object item, string name, double minimum) {
            var map = item as Dictionary<string, object>;
            object value = map == null ? item : Get(map, name);
            double result;
            if (value == null || !Double.TryParse(Convert.ToString(value, Invariant), NumberStyles.Float, Invariant, out result)) return 1;
            if (Double.IsNaN(result) || Double.IsInfinity(result) || Math.Abs(result) > 100000) throw new InvalidDataException("INPUT_DIMENSION_VALUE");
            return Math.Max(minimum, result);
        }
        private static byte[] Read(ZipArchiveEntry entry) {
            if (entry == null || entry.Length > 16 * 1024 * 1024) throw new InvalidDataException("TEMPLATE_ENTRY");
            using (var input = entry.Open()) using (var output = new MemoryStream()) { input.CopyTo(output); return output.ToArray(); }
        }
        private static XmlDocument Load(byte[] bytes) {
            var doc = new XmlDocument { PreserveWhitespace = true, XmlResolver = null };
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64 * 1024 * 1024 };
            using (var stream = new MemoryStream(bytes)) using (var reader = XmlReader.Create(stream, settings)) doc.Load(reader);
            return doc;
        }
        private static XmlNamespaceManager Ns(XmlDocument doc) {
            var ns = new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("hp", HP); ns.AddNamespace("hh", HH); ns.AddNamespace("hc", HC); ns.AddNamespace("hs", HS); return ns;
        }
        private static XmlElement Need(XmlNode parent, string path) {
            var node = parent.SelectSingleNode(path, Ns(parent as XmlDocument ?? parent.OwnerDocument)) as XmlElement;
            if (node == null) throw new InvalidDataException("TEMPLATE_NODE: " + path);
            return node;
        }
        private static XmlElement Child(XmlNode parent, string prefix, string name) {
            var doc = parent.OwnerDocument;
            var existing = parent.SelectSingleNode(prefix + ":" + name, Ns(doc)) as XmlElement;
            if (existing != null) return existing;
            var node = doc.CreateElement(prefix, name, prefix == "hp" ? HP : prefix == "hh" ? HH : HC); parent.AppendChild(node); return node;
        }
        private static void Attr(XmlElement element, params object[] values) {
            for (int i = 0; i < values.Length; i += 2) element.SetAttribute((string)values[i], Convert.ToString(values[i + 1], Invariant));
        }
        private static void ClearChildren(XmlNode node) { while (node.HasChildNodes) node.RemoveChild(node.FirstChild); }
        private static void Remove(XmlNode node, string xpath) { foreach (XmlNode child in node.SelectNodes(xpath, Ns(node.OwnerDocument))) node.RemoveChild(child); }
        private static byte[] Bytes(XmlDocument doc) {
            using (var output = new MemoryStream()) {
                using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = false })) doc.Save(writer);
                return output.ToArray();
            }
        }
        private XmlElement Paragraph(XmlDocument doc, int id, string para, string ch, string text) {
            var p = doc.CreateElement("hp", "p", HP); Attr(p, "id", id, "paraPrIDRef", para, "styleIDRef", 0, "pageBreak", 0, "columnBreak", 0, "merged", 0);
            var run = Child(p, "hp", "run"); Attr(run, "charPrIDRef", ch); Child(run, "hp", "t").AppendChild(doc.CreateTextNode(text)); return p;
        }
        private static bool Numeric(string text) {
            text = text.Trim().Replace(",", "").Replace(" ", "");
            if (text.StartsWith("(") && text.EndsWith(")")) text = "-" + text.Substring(1, text.Length - 2);
            foreach (string prefix in new [] { "₩", "\\" }) if (text.StartsWith(prefix, StringComparison.Ordinal)) text = text.Substring(prefix.Length);
            foreach (string suffix in new [] { "원", "%" }) if (text.EndsWith(suffix, StringComparison.Ordinal)) text = text.Substring(0, text.Length - suffix.Length);
            double n; return Double.TryParse(text, NumberStyles.Float, Invariant, out n);
        }
        private int Border(int r, int c) {
            if (cols == 1) return title && r == 0 ? 17 : r == rows - 1 ? 19 : 18;
            if (title && r == 0) return c == 0 ? 7 : c == cols - 1 ? 9 : c == 1 ? 8 : 10;
            if (r == rows - 1) return c == 0 ? 5 : c == cols - 1 ? 15 : c == 1 ? 14 : 16;
            return c == 0 ? 6 : c == cols - 1 ? 12 : c == 1 ? 11 : 13;
        }
        private byte[] Section(byte[] bytes) {
            var original = Load(bytes); var doc = Load(bytes); var section = doc.DocumentElement; ClearChildren(section);
            var secPara = doc.ImportNode(Need(original, "/hs:sec/hp:p[.//hp:secPr]"), true);
            foreach (XmlNode text in secPara.SelectNodes(".//hp:t", Ns(doc))) text.InnerText = text.InnerText.Replace("12pt", "13pt");
            section.AppendChild(secPara);
            var sourceTable = Need(original, "//hp:tbl"); var sourceRow = Need(sourceTable, "hp:tr"); var sourceCell = Need(sourceRow, "hp:tc");
            var p = (XmlElement)doc.ImportNode(sourceTable.ParentNode.ParentNode, true); ClearChildren(p);
            Attr(p, "id", 1000000002, "paraPrIDRef", 0, "styleIDRef", 0, "pageBreak", 0, "columnBreak", 0, "merged", 0);
            var run = (XmlElement)doc.ImportNode(sourceTable.ParentNode, true); ClearChildren(run); Attr(run, "charPrIDRef", 0);
            var table = (XmlElement)doc.ImportNode(sourceTable, true); Remove(table, "hp:tr");
            Attr(table, "id", 1000000999, "zOrder", 6, "numberingType", "TABLE", "textWrap", "TOP_AND_BOTTOM", "textFlow", "BOTH_SIDES", "lock", 0, "dropcapstyle", "None", "pageBreak", "CELL", "repeatHeader", 1, "rowCnt", rows, "colCnt", cols, "cellSpacing", 0, "borderFillIDRef", 3, "noAdjust", 0);
            int totalHeight = 0; foreach (int h in heights) totalHeight = checked(totalHeight + h);
            Attr(Child(table, "hp", "sz"), "width", 47341, "widthRelTo", "ABSOLUTE", "height", totalHeight, "heightRelTo", "ABSOLUTE", "protect", 0);
            Attr(Child(table, "hp", "pos"), "treatAsChar", 1, "affectLSpacing", 0, "flowWithText", 1, "allowOverlap", 0, "holdAnchorAndSO", 0, "vertRelTo", "PARA", "horzRelTo", "COLUMN", "vertAlign", "TOP", "horzAlign", "LEFT", "vertOffset", 0, "horzOffset", 0);
            Attr(Child(table, "hp", "outMargin"), "left", 0, "right", 0, "top", 0, "bottom", 0);
            Attr(Child(table, "hp", "inMargin"), "left", 340, "right", 340, "top", 120, "bottom", 120);
            int pid = 1000000100;
            for (int r = 0; r < rows; r++) {
                var tr = (XmlElement)doc.ImportNode(sourceRow, true); ClearChildren(tr);
                for (int c = 0; c < cols; c++) {
                    cancellation.ThrowIfCancellationRequested();
                    if (covered.Contains(r * cols + c)) continue;
                    int[] span; if (!merges.TryGetValue(r * cols + c, out span)) span = new [] { 1, 1 };
                    int w = 0, h = 0; for (int x = c; x < c + span[1]; x++) w += widths[x]; for (int y = r; y < r + span[0]; y++) h = checked(h + heights[y]);
                    var tc = (XmlElement)doc.ImportNode(sourceCell, true);
                    Attr(tc, "name", "", "header", 0, "hasMargin", 0, "protect", 0, "editable", 0, "dirty", 0, "borderFillIDRef", Border(r, c));
                    var sub = Child(tc, "hp", "subList");
                    Attr(sub, "id", "", "textDirection", "HORIZONTAL", "lineWrap", "BREAK", "vertAlign", "CENTER", "linkListIDRef", 0, "linkListNextIDRef", 0, "textWidth", 0, "textHeight", 0, "hasTextRef", 0, "hasNumRef", 0);
                    Remove(sub, "hp:p"); string text = values[r, c] ?? "";
                    sub.AppendChild(Paragraph(doc, pid++, title && r == 0 ? "11" : Numeric(text) ? "15" : "11", title && r == 0 ? "10" : "2", text));
                    Attr(Child(tc, "hp", "cellAddr"), "colAddr", c, "rowAddr", r); Attr(Child(tc, "hp", "cellSpan"), "colSpan", span[1], "rowSpan", span[0]);
                    Attr(Child(tc, "hp", "cellSz"), "width", w, "height", h); Attr(Child(tc, "hp", "cellMargin"), "left", 340, "right", 340, "top", 120, "bottom", 120);
                    tr.AppendChild(tc);
                }
                table.AppendChild(tr);
            }
            run.AppendChild(table); p.AppendChild(run); section.AppendChild(p); section.AppendChild(Paragraph(doc, 1000000003, "14", "0", "")); return Bytes(doc);
        }
        private byte[] Header(byte[] bytes) {
            var doc = Load(bytes); var ns = Ns(doc);
            string[] languages = { "HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER" }; int[] fontIds = { 3, 2, 3, 3, 2, 3, 2 };
            foreach (XmlElement face in doc.SelectNodes("//hh:fontface", ns)) {
                int index = Array.IndexOf(languages, face.GetAttribute("lang")); if (index < 0) continue;
                var item = face.SelectSingleNode("hh:font[@id='" + fontIds[index] + "']", ns) as XmlElement;
                if (item == null) { item = doc.CreateElement("hh", "font", HH); face.AppendChild(item); Attr(item, "id", fontIds[index]); }
                Attr(item, "face", font, "type", "HFT", "isEmbedded", 0);
                Attr(Child(item, "hh", "typeInfo"), "familyType", "FCAT_GOTHIC", "weight", 0, "proportion", 0, "contrast", 0, "strokeVariation", 0, "armStyle", 0, "letterform", 0, "midline", 0, "xHeight", 0);
            }
            foreach (XmlElement item in doc.SelectNodes("//hh:font[@face='한양중고딕']", ns)) Attr(item, "face", font, "type", "HFT");
            foreach (string id in new [] { "0", "2", "3", "10" }) {
                var ch = Need(doc, "//hh:charPr[@id='" + id + "']"); Attr(ch, "height", fontSize * 100);
                var reference = Need(ch, "hh:fontRef"); for (int i = 0; i < languages.Length; i++) Attr(reference, languages[i].ToLowerInvariant(), fontIds[i]);
            }
            foreach (string id in new [] { "11", "15", "17" }) {
                var margin = Child(Need(doc, "//hh:paraPr[@id='" + id + "']"), "hh", "margin");
                foreach (string name in new [] { "intent", "left", "right", "prev", "next" }) Attr(Child(margin, "hc", name), "value", 0, "unit", "HWPUNIT");
            }
            string inner = simple ? "#C0C0C0" : "#808080", outer = simple ? "#808080" : "#000000";
            var fills = Need(doc, "//hh:borderFills");
            for (int id = 17; id <= 19; id++) {
                if (fills.SelectSingleNode("hh:borderFill[@id='" + id + "']", ns) != null) continue;
                var fill = doc.CreateElement("hh", "borderFill", HH); fills.AppendChild(fill);
                Attr(fill, "id", id, "threeD", 0, "shadow", 0, "centerLine", "NONE", "breakCellSeparateLine", 0);
                foreach (string slash in new [] { "slash", "backSlash" }) Attr(Child(fill, "hh", slash), "type", "NONE", "Crooked", 0, "isCounter", 0);
                foreach (string side in new [] { "leftBorder", "rightBorder", "topBorder", "bottomBorder", "diagonal" }) Attr(Child(fill, "hh", side), "type", "NONE", "width", "0.1 mm", "color", outer);
                if (id == 17 && !simple) Attr(Child(Child(fill, "hc", "fillBrush"), "hc", "winBrush"), "faceColor", "#EEF3F8", "hatchColor", "#000000", "alpha", 0);
            }
            if (simple) Remove(Need(fills, "hh:borderFill[@id='17']"), "hc:fillBrush");
            Attr(fills, "itemCnt", fills.ChildNodes.Count);
            for (int id = 5; id <= 19; id++) {
                bool header = id == 7 || id == 8 || id == 9 || id == 10 || id == 17;
                bool bottom = id == 5 || id == 14 || id == 15 || id == 16 || id == 19;
                bool left = id == 5 || id == 6 || id == 7 || id >= 17;
                bool right = id == 9 || id == 12 || id == 15 || id >= 17;
                var fill = Need(fills, "hh:borderFill[@id='" + id + "']");
                Attr(Child(fill, "hh", "leftBorder"), "type", left ? "NONE" : "SOLID", "width", "0.1 mm", "color", inner);
                Attr(Child(fill, "hh", "rightBorder"), "type", right ? "NONE" : "SOLID", "width", "0.1 mm", "color", inner);
                Attr(Child(fill, "hh", "topBorder"), "type", "SOLID", "width", header && !simple ? "0.3 mm" : "0.1 mm", "color", header ? outer : inner);
                Attr(Child(fill, "hh", "bottomBorder"), "type", header ? "DOUBLE_SLIM" : "SOLID", "width", header ? "0.5 mm" : simple ? "0.1 mm" : bottom ? "0.3 mm" : "0.1 mm", "color", bottom ? outer : inner);
                Attr(Child(fill, "hh", "diagonal"), "type", "NONE", "width", "0.1 mm", "color", inner);
            }
            return Bytes(doc);
        }
        private byte[] Preview() {
            var output = new StringBuilder();
            for (int r = 0; r < rows; r++) { for (int c = 0; c < cols; c++) { if (c != 0) output.Append('\t'); output.Append(values[r, c] ?? ""); } output.Append('\n'); }
            return Encoding.UTF8.GetBytes(output.ToString());
        }
        private void StoredZip(List<KeyValuePair<string, byte[]>> entries, Stream output) {
            // Framework ZipArchive NoCompression still uses method 8; HWPX mimetype requires method 0.
            using (var central = new MemoryStream()) using (var directory = new BinaryWriter(central, Encoding.UTF8, true)) using (var zip = new BinaryWriter(output, Encoding.UTF8, true)) {
                foreach (var entry in entries) {
                    cancellation.ThrowIfCancellationRequested();
                    byte[] name = Encoding.UTF8.GetBytes(entry.Key), data = entry.Value; uint offset = checked((uint)output.Position), size = checked((uint)data.Length), crc = 0xffffffff;
                    for (int i = 0; i < data.Length; i++) { if ((i & 4095) == 0) cancellation.ThrowIfCancellationRequested(); crc ^= data[i]; for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xedb88320 : crc >> 1; }
                    crc ^= 0xffffffff;
                    zip.Write(0x04034b50u); zip.Write((ushort)20); zip.Write((ushort)0x800); zip.Write((ushort)0); zip.Write((ushort)0); zip.Write((ushort)33); zip.Write(crc); zip.Write(size); zip.Write(size); zip.Write((ushort)name.Length); zip.Write((ushort)0); zip.Write(name); zip.Write(data);
                    directory.Write(0x02014b50u); directory.Write((ushort)20); directory.Write((ushort)20); directory.Write((ushort)0x800); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)33); directory.Write(crc); directory.Write(size); directory.Write(size); directory.Write((ushort)name.Length); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write(0u); directory.Write(offset); directory.Write(name);
                }
                uint start = checked((uint)output.Position); directory.Flush(); byte[] bytes = central.ToArray(); zip.Write(bytes);
                zip.Write(0x06054b50u); zip.Write((ushort)0); zip.Write((ushort)0); zip.Write((ushort)entries.Count); zip.Write((ushort)entries.Count); zip.Write((uint)bytes.Length); zip.Write(start); zip.Write((ushort)0);
            }
        }
    }
}
