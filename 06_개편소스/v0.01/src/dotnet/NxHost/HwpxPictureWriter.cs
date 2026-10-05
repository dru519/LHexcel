using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Xml;

namespace LH.NxHost {
    internal static class HwpxPictureWriter {
        private const string NamespaceRoot = "http" + "://www.hancom.co.kr/hwpml/2011/";
        private const string HP = NamespaceRoot + "paragraph";
        private const string HH = NamespaceRoot + "head";
        private const string HC = NamespaceRoot + "core";
        private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;

        private sealed class Picture {
            internal string Title, Path;
            internal DateTime Taken;
            internal byte[] Bytes;
        }
        private sealed class Layout {
            internal int Width, Height, Columns, Rows, Factor, Pages;
            internal List<Picture> Pictures;
            internal bool TitleFileNames;
        }

        internal static bool IsPicturePayload(string json) {
            if (String.IsNullOrEmpty(json) || json.Length > HwpxTableWriter.MaxPayloadChars) return false;
            try {
                var serializer = new JavaScriptSerializer { MaxJsonLength = HwpxTableWriter.MaxPayloadChars, RecursionLimit = 32 };
                var root = serializer.DeserializeObject(json.TrimStart('\uFEFF')) as Dictionary<string, object>;
                return root != null && root.ContainsKey("picture_layout");
            } catch { return false; }
        }

        internal static void Write(string json, string templatePath, string outputPath, CancellationToken token) {
            Layout layout = Parse(json, token);
            string temporary = outputPath + ".picture-" + Guid.NewGuid().ToString("N") + ".tmp";
            bool temporaryCreated = false, outputCreated = false;
            try {
                HwpxTableWriter.Write(SyntheticPayload(layout), templatePath, temporary, token);
                temporaryCreated = true;
                var entries = Transform(temporary, layout, token);
                using (var output = new FileStream(outputPath, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                    outputCreated = true;
                    WriteStoredZip(entries, output, token);
                    output.Flush(true);
                }
                token.ThrowIfCancellationRequested();
            } catch {
                if (outputCreated) try { File.Delete(outputPath); } catch { }
                throw;
            } finally {
                if (temporaryCreated) try { File.Delete(temporary); } catch { }
            }
        }

        private static Layout Parse(string json, CancellationToken token) {
            var serializer = new JavaScriptSerializer { MaxJsonLength = HwpxTableWriter.MaxPayloadChars, RecursionLimit = 32 };
            var root = Map(serializer.DeserializeObject(json.TrimStart('\uFEFF')));
            Only(root, "schema_version|picture_layout");
            if (Integer(root["schema_version"]) != 1) throw new InvalidDataException("INPUT_SCHEMA");
            var source = Map(root["picture_layout"]);
            Only(source, "width_cm|height_cm|columns|rows|titles|title_filenames|sort|files");
            foreach (string field in new [] { "width_cm", "height_cm", "columns", "rows", "titles", "sort", "files" })
                if (!source.ContainsKey(field)) throw new InvalidDataException("PICTURE_FIELD: " + field);
            double width = Number(source["width_cm"]), height = Number(source["height_cm"]);
            int columns = Integer(source["columns"]), rows = Integer(source["rows"]);
            if (width < .5 || width > 16.7 || height < .5 || height > 24) throw new InvalidDataException("PICTURE_SIZE");
            if (columns < 1 || columns > 8 || rows < 1 || rows > 12) throw new InvalidDataException("PICTURE_LAYOUT");
            if (!(source["titles"] is bool)) throw new InvalidDataException("PICTURE_OPTIONS");
            bool titles = (bool)source["titles"];
            if (source.ContainsKey("title_filenames") && !(source["title_filenames"] is bool)) throw new InvalidDataException("PICTURE_OPTIONS");
            bool titleFileNames = !source.ContainsKey("title_filenames") || (bool)source["title_filenames"];
            string sort = source["sort"] as string;
            if (sort != "title" && sort != "taken") throw new InvalidDataException("PICTURE_OPTIONS");
            if (width * columns > 16.7 || (height + (titles ? .7 : 0)) * rows > 24) throw new InvalidDataException("PICTURE_PAGE_SIZE");
            var files = source["files"] as IList;
            if (files == null || files.Count < 1 || files.Count > 200) throw new InvalidDataException("PICTURE_COUNT");
            var pictures = new List<Picture>(); var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase); long total = 0;
            foreach (object item in files) {
                token.ThrowIfCancellationRequested();
                string path = item as string;
                if (path == null || path.Length > 1024 || !Path.IsPathRooted(path) || path.StartsWith(@"\\?\") || path.StartsWith(@"\.\") || path.IndexOf(':', 2) >= 0) throw new InvalidDataException("PICTURE_PATH");
                string full = Path.GetFullPath(path);
                if (!seen.Add(full)) continue;
                string extension = Path.GetExtension(full).ToLowerInvariant();
                if (extension != ".png" && extension != ".jpg" && extension != ".jpeg" && extension != ".bmp") throw new InvalidDataException("PICTURE_TYPE");
                using (var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                    if (stream.Length < 4 || stream.Length > 32L * 1024 * 1024) throw new InvalidDataException("PICTURE_FILE_SIZE");
                    using (var image = Image.FromStream(stream, false, true)) {
                        if ((long)image.Width * image.Height > 16000000) throw new InvalidDataException("PICTURE_PIXEL_LIMIT");
                        DateTime taken = ReadTaken(image); ApplyOrientation(image);
                        using (var normalized = new Bitmap(image.Width, image.Height, PixelFormat.Format32bppArgb)) using (var graphics = Graphics.FromImage(normalized)) using (var buffer = new MemoryStream()) {
                            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy; graphics.DrawImageUnscaled(image, 0, 0); normalized.Save(buffer, ImageFormat.Png); byte[] bytes = buffer.ToArray(); total += bytes.Length;
                            if (total > 128L * 1024 * 1024) throw new InvalidDataException("PICTURE_TOTAL_SIZE");
                            pictures.Add(new Picture { Title = Path.GetFileNameWithoutExtension(full), Path = full, Taken = taken, Bytes = bytes });
                        }
                    }
                }
            }
            pictures.Sort(delegate(Picture left, Picture right) {
                int result = sort == "taken" ? DateTime.Compare(left.Taken, right.Taken) : 0;
                if (result == 0) result = StringComparer.OrdinalIgnoreCase.Compare(left.Title, right.Title);
                if (result == 0) result = StringComparer.Ordinal.Compare(left.Path, right.Path);
                return result;
            });
            int factor = titles ? 2 : 1;
            return new Layout { Width = (int)Math.Round(width * 7200 / 2.54), Height = (int)Math.Round(height * 7200 / 2.54), Columns = columns, Rows = rows, Factor = factor, Pages = (int)Math.Ceiling((double)pictures.Count / (columns * rows)), Pictures = pictures, TitleFileNames = titleFileNames };
        }

        private static DateTime ReadTaken(Image image) {
            try {
                var property = image.GetPropertyItem(36867);
                string value = Encoding.ASCII.GetString(property.Value).Trim('\0'); DateTime parsed;
                if (DateTime.TryParseExact(value, "yyyy:MM:dd HH:mm:ss", Invariant, DateTimeStyles.None, out parsed)) return parsed;
            } catch (ArgumentException) { }
            return DateTime.MaxValue;
        }
        private static void ApplyOrientation(Image image) {
            try {
                int orientation = BitConverter.ToUInt16(image.GetPropertyItem(274).Value, 0);
                var rotations = new [] { RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipNone, RotateFlipType.RotateNoneFlipX, RotateFlipType.Rotate180FlipNone, RotateFlipType.Rotate180FlipX, RotateFlipType.Rotate90FlipX, RotateFlipType.Rotate90FlipNone, RotateFlipType.Rotate270FlipX, RotateFlipType.Rotate270FlipNone };
                if (orientation >= 2 && orientation <= 8) image.RotateFlip(rotations[orientation]);
            } catch (ArgumentException) { }
        }
        private static string SyntheticPayload(Layout layout) {
            var heights = new List<int>(); for (int i = 0; i < layout.Rows * layout.Factor * layout.Pages; i++) heights.Add(20);
            var root = new Dictionary<string, object> { { "schema_version", 1 }, { "rows", heights.Count }, { "columns", layout.Columns }, { "row_heights", heights }, { "cells", new object[0] }, { "merges", new object[0] }, { "settings", new Dictionary<string, object> { { "font_name", "중고딕" }, { "font_type", "HFT" }, { "font_size_pt", 9 }, { "title_row", false }, { "table_style", "간결형" } } } };
            return new JavaScriptSerializer().Serialize(root);
        }

        private static List<KeyValuePair<string, byte[]>> Transform(string path, Layout layout, CancellationToken token) {
            var entries = new List<KeyValuePair<string, byte[]>>(); byte[] sectionBytes = null, headerBytes = null, manifestBytes = null;
            using (var stream = File.OpenRead(path)) using (var zip = new ZipArchive(stream, ZipArchiveMode.Read)) {
                foreach (var entry in zip.Entries) {
                    token.ThrowIfCancellationRequested(); byte[] bytes = Read(entry);
                    if (entry.FullName == "Contents/section0.xml") sectionBytes = bytes;
                    else if (entry.FullName == "Contents/header.xml") headerBytes = bytes;
                    else if (entry.FullName == "Contents/content.hpf") manifestBytes = bytes;
                    entries.Add(new KeyValuePair<string, byte[]>(entry.FullName, bytes));
                }
            }
            if (sectionBytes == null || headerBytes == null || manifestBytes == null) throw new InvalidDataException("PICTURE_TEMPLATE");
            byte[] section, header, manifest; TransformXml(sectionBytes, headerBytes, manifestBytes, layout, token, out section, out header, out manifest);
            for (int i = 0; i < entries.Count; i++) {
                if (entries[i].Key == "Contents/section0.xml") entries[i] = new KeyValuePair<string, byte[]>(entries[i].Key, section);
                else if (entries[i].Key == "Contents/header.xml") entries[i] = new KeyValuePair<string, byte[]>(entries[i].Key, header);
                else if (entries[i].Key == "Contents/content.hpf") entries[i] = new KeyValuePair<string, byte[]>(entries[i].Key, manifest);
                else if (entries[i].Key == "Preview/PrvText.txt") entries[i] = new KeyValuePair<string, byte[]>(entries[i].Key, Encoding.UTF8.GetBytes(String.Join("\n", layout.Pictures.ConvertAll(p => p.Title).ToArray())));
            }
            for (int i = 0; i < layout.Pictures.Count; i++) entries.Add(new KeyValuePair<string, byte[]>("BinData/nxphoto" + (i + 1) + ".png", layout.Pictures[i].Bytes));
            return entries;
        }

        private static void TransformXml(byte[] sectionBytes, byte[] headerBytes, byte[] manifestBytes, Layout layout, CancellationToken token, out byte[] sectionResult, out byte[] headerResult, out byte[] manifestResult) {
            XmlDocument section = Load(sectionBytes), header = Load(headerBytes), manifest = Load(manifestBytes); XmlNamespaceManager sns = Ns(section), hns = Ns(header);
            foreach (XmlNode text in section.SelectNodes("/hs:sec/hp:p[.//hp:secPr]//hp:t", sns)) text.InnerText = "";
            XmlElement fills = Need(header, "//hh:borderFills", hns); int nextFill = 1;
            foreach (XmlNode node in fills.ChildNodes) { var element = node as XmlElement; int id; if (element != null && Int32.TryParse(element.GetAttribute("id"), out id)) nextFill = Math.Max(nextFill, id + 1); }
            int emptyFill = nextFill; XmlElement manifestNode = manifest.SelectSingleNode("//*[local-name()='manifest']") as XmlElement;
            if (manifestNode == null) throw new InvalidDataException("PICTURE_MANIFEST");
            for (int index = -1; index < layout.Pictures.Count; index++) {
                token.ThrowIfCancellationRequested(); XmlElement fill = header.CreateElement("hh", "borderFill", HH); Attr(fill, "id", emptyFill + index + 1, "threeD", 0, "shadow", 0, "centerLine", "NONE", "breakCellSeparateLine", 0);
                foreach (string slash in new [] { "slash", "backSlash" }) Attr(Append(header, fill, "hh", slash, HH), "type", "NONE", "Crooked", 0, "isCounter", 0);
                foreach (string side in new [] { "leftBorder", "rightBorder", "topBorder", "bottomBorder", "diagonal" }) Attr(Append(header, fill, "hh", side, HH), "type", side == "diagonal" ? "NONE" : "SOLID", "width", "0.1 mm", "color", "#808080");
                if (index >= 0) {
                    string id = "nxphoto" + (index + 1), entry = "BinData/" + id + ".png";
                    XmlElement brush = Append(header, fill, "hc", "fillBrush", HC), imageBrush = Append(header, brush, "hc", "imgBrush", HC); Attr(imageBrush, "mode", "TOTAL");
                    Attr(Append(header, imageBrush, "hc", "img", HC), "binaryItemIDRef", id, "bright", 0, "contrast", 0, "effect", "REAL_PIC", "alpha", 0);
                    XmlElement item = manifest.CreateElement("opf", "item", manifestNode.NamespaceURI); Attr(item, "id", id, "href", entry, "media-type", "image/png", "isEmbeded", 1); manifestNode.AppendChild(item);
                }
                fills.AppendChild(fill);
            }
            Attr(fills, "itemCnt", fills.ChildNodes.Count);
            XmlElement table = Need(section, "//hp:tbl", sns), paragraph = (XmlElement)table.ParentNode.ParentNode; var allRows = new List<XmlNode>();
            foreach (XmlNode row in table.SelectNodes("hp:tr", sns)) allRows.Add(row); foreach (XmlNode row in allRows) table.RemoveChild(row);
            int groupRows = layout.Rows * layout.Factor, groupHeight = layout.Rows * (layout.Height + (layout.Factor == 2 ? 1984 : 0));
            for (int page = 0; page < layout.Pages; page++) {
                var p = (XmlElement)paragraph.CloneNode(true); Attr(p, "id", 1100000000 + page, "pageBreak", page == 0 ? 0 : 1); XmlElement t = Need(p, ".//hp:tbl", sns);
                Attr(t, "id", 1200000000 + page, "rowCnt", groupRows, "repeatHeader", 0); Attr(Need(t, "hp:sz", sns), "width", layout.Width * layout.Columns, "height", groupHeight);
                for (int r = 0; r < groupRows; r++) {
                    XmlNode row = allRows[page * groupRows + r]; bool caption = layout.Factor == 2 && r % 2 == 1;
                    foreach (XmlNode cellNode in row.SelectNodes("hp:tc", sns)) {
                        var cell = (XmlElement)cellNode; XmlElement address = Need(cell, "hp:cellAddr", sns); int col = Int32.Parse(address.GetAttribute("colAddr"), Invariant); Attr(address, "rowAddr", r);
                        int index = page * layout.Rows * layout.Columns + (r / layout.Factor) * layout.Columns + col;
                        Attr(cell, "borderFillIDRef", !caption && index < layout.Pictures.Count ? emptyFill + 1 + index : emptyFill, "hasMargin", 1);
                        Attr(Need(cell, "hp:cellSz", sns), "width", layout.Width, "height", caption ? 1984 : layout.Height); Attr(Need(cell, "hp:cellMargin", sns), "left", 0, "right", 0, "top", 0, "bottom", 0);
                        Need(cell, ".//hp:t", sns).InnerText = caption && layout.TitleFileNames && index < layout.Pictures.Count ? layout.Pictures[index].Title : "";
                    }
                    t.AppendChild(row);
                }
                paragraph.ParentNode.InsertBefore(p, paragraph);
            }
            paragraph.ParentNode.RemoveChild(paragraph); sectionResult = Bytes(section); headerResult = Bytes(header); manifestResult = Bytes(manifest);
        }

        private static Dictionary<string, object> Map(object value) { var map = value as Dictionary<string, object>; if (map == null) throw new InvalidDataException("INPUT_OBJECT"); return map; }
        private static void Only(Dictionary<string, object> map, string allowed) { foreach (string key in map.Keys) if (Array.IndexOf(allowed.Split('|'), key) < 0) throw new InvalidDataException("INPUT_FIELD: " + key); }
        private static double Number(object value) { if (value == null || value is bool || value is string) throw new InvalidDataException("INPUT_NUMBER"); double number = Convert.ToDouble(value, Invariant); if (Double.IsNaN(number) || Double.IsInfinity(number)) throw new InvalidDataException("INPUT_NUMBER"); return number; }
        private static int Integer(object value) { double number = Number(value); if (number != Math.Truncate(number) || number < 0 || number > 16384) throw new InvalidDataException("INPUT_INTEGER"); return (int)number; }
        private static XmlDocument Load(byte[] bytes) { var doc = new XmlDocument { PreserveWhitespace = true, XmlResolver = null }; var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 64 * 1024 * 1024 }; using (var stream = new MemoryStream(bytes)) using (var reader = XmlReader.Create(stream, settings)) doc.Load(reader); return doc; }
        private static XmlNamespaceManager Ns(XmlDocument doc) { var ns = new XmlNamespaceManager(doc.NameTable); ns.AddNamespace("hp", HP); ns.AddNamespace("hh", HH); ns.AddNamespace("hc", HC); ns.AddNamespace("hs", NamespaceRoot + "section"); return ns; }
        private static XmlElement Need(XmlNode node, string xpath, XmlNamespaceManager ns) { var result = node.SelectSingleNode(xpath, ns) as XmlElement; if (result == null) throw new InvalidDataException("TEMPLATE_NODE: " + xpath); return result; }
        private static XmlElement Append(XmlDocument doc, XmlNode parent, string prefix, string name, string ns) { var element = doc.CreateElement(prefix, name, ns); parent.AppendChild(element); return element; }
        private static void Attr(XmlElement element, params object[] values) { for (int i = 0; i < values.Length; i += 2) element.SetAttribute((string)values[i], Convert.ToString(values[i + 1], Invariant)); }
        private static byte[] Bytes(XmlDocument doc) { using (var output = new MemoryStream()) { using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false, CloseOutput = false })) doc.Save(writer); return output.ToArray(); } }
        private static byte[] Read(ZipArchiveEntry entry) { if (entry == null || entry.Length > 160L * 1024 * 1024) throw new InvalidDataException("PICTURE_ENTRY"); using (var input = entry.Open()) using (var output = new MemoryStream()) { input.CopyTo(output); return output.ToArray(); } }
        private static void WriteStoredZip(List<KeyValuePair<string, byte[]>> entries, Stream output, CancellationToken token) {
            using (var central = new MemoryStream()) using (var directory = new BinaryWriter(central, Encoding.UTF8, true)) using (var zip = new BinaryWriter(output, Encoding.UTF8, true)) {
                foreach (var entry in entries) {
                    token.ThrowIfCancellationRequested(); byte[] name = Encoding.UTF8.GetBytes(entry.Key), data = entry.Value; uint offset = checked((uint)output.Position), size = checked((uint)data.Length), crc = 0xffffffff;
                    for (int i = 0; i < data.Length; i++) { if ((i & 4095) == 0) token.ThrowIfCancellationRequested(); crc ^= data[i]; for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xedb88320 : crc >> 1; } crc ^= 0xffffffff;
                    zip.Write(0x04034b50u); zip.Write((ushort)20); zip.Write((ushort)0x800); zip.Write((ushort)0); zip.Write((ushort)0); zip.Write((ushort)33); zip.Write(crc); zip.Write(size); zip.Write(size); zip.Write((ushort)name.Length); zip.Write((ushort)0); zip.Write(name); zip.Write(data);
                    directory.Write(0x02014b50u); directory.Write((ushort)20); directory.Write((ushort)20); directory.Write((ushort)0x800); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)33); directory.Write(crc); directory.Write(size); directory.Write(size); directory.Write((ushort)name.Length); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write((ushort)0); directory.Write(0u); directory.Write(offset); directory.Write(name);
                }
                uint start = checked((uint)output.Position); directory.Flush(); byte[] bytes = central.ToArray(); zip.Write(bytes); zip.Write(0x06054b50u); zip.Write((ushort)0); zip.Write((ushort)0); zip.Write((ushort)entries.Count); zip.Write((ushort)entries.Count); zip.Write((uint)bytes.Length); zip.Write(start); zip.Write((ushort)0);
            }
        }
    }
}
