using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Text.RegularExpressions;

namespace LH.NxHost {
    // Immutable, bounded samples only. This renderer has no Excel/COM or filesystem access.
    internal static class WorksheetPreviewRenderer {
        private static Bitmap Canvas(int width, int height) {
            if (width < 24 || height < 16 || width > 1800 || height > 1200) throw new ArgumentException("PREVIEW_SIZE");
            return new Bitmap(width * 2, height * 2);
        }
        internal static Bitmap Summary(string text, int width, int height) {
            if (text == null || text.Length > 100000) throw new ArgumentException("PREVIEW_TEXT");
            var rows = new List<string[]>();
            bool changes = false;
            bool comparison = text.IndexOf("비교 표본", StringComparison.Ordinal) >= 0;
            foreach (string line in text.Replace("\r", "").Split('\n')) {
                var match = Regex.Match(line, @"^\s*(\$?[A-Z]{1,3}\$?\d+)\s*:\s*(.*?)\s+→\s+(.*)$");
                if (match.Success) { rows.Add(new[] {match.Groups[1].Value, match.Groups[2].Value, match.Groups[3].Value}); changes = true; }
            }
            if (!changes) foreach (string line in text.Replace("\r", "").Split('\n')) {
                if (String.IsNullOrWhiteSpace(line)) continue;
                rows.Add(line.Split(new[] {" | ", "\t"}, StringSplitOptions.None));
                if (rows.Count >= 20) break;
            }
            int columns = changes ? 3 : 1;
            foreach (var row in rows) columns = Math.Min(6, Math.Max(columns, row.Length));
            var bitmap = Canvas(width, height);
            using (var g = Graphics.FromImage(bitmap)) using (var font = new Font("맑은 고딕", 9, FontStyle.Regular, GraphicsUnit.Point))
            using (var bold = new Font(font, FontStyle.Bold)) using (var pen = new Pen(Color.FromArgb(215,221,225))) {
                g.ScaleTransform(2,2); g.Clear(Color.White); g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                const float rowHeight = 22;
                float columnWidth = (width - 2f) / columns;
                int visible = Math.Max(0, Math.Min(rows.Count, (int)(height / rowHeight) - 1));
                for (int r=0; r<=visible; r++) for (int c=0; c<columns; c++) {
                    var rect = new RectangleF(1+c*columnWidth, 1+r*rowHeight, columnWidth, rowHeight);
                    string value = r==0 ? (changes ? (comparison ? new[] {"셀", "기준값", "비교값"} : new[] {"셀", "변경 전", "변경 후"})[c] : columns==1 ? "미리보기" : ((char)('A'+c)).ToString()) : c<rows[r-1].Length ? rows[r-1][c] : "";
                    bool changed = changes && r>0 && rows[r-1][1]!=rows[r-1][2];
                    var fill = r==0 ? Color.FromArgb(239,243,245) : changed && c>0 ? (comparison ? Color.FromArgb(255,199,206) : Color.FromArgb(226,239,218)) : Color.White;
                    using (var brush = new SolidBrush(fill)) g.FillRectangle(brush,rect);
                    g.DrawRectangle(pen,rect.X,rect.Y,rect.Width,rect.Height);
                    Text(g,value,r==0?bold:font,Color.FromArgb(35,45,50),new RectangleF(rect.X+4,rect.Y+2,rect.Width-8,rect.Height-3),0);
                }
            }
            return bitmap;
        }
        internal static Bitmap Cells(Array cells, int width, int height) {
            if (cells == null || cells.Rank != 2 || cells.GetLength(1)!=10 || cells.GetLength(0)>300) throw new ArgumentException("PREVIEW_CELLS");
            var bitmap = Canvas(width,height);
            try {
                using (var g = Graphics.FromImage(bitmap)) {
                    g.ScaleTransform(2,2); g.Clear(Color.White); g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                    for(int i=0;i<cells.GetLength(0);i++) {
                        int r=i+cells.GetLowerBound(0), c=cells.GetLowerBound(1);
                        float x=Number(cells.GetValue(r,c)), y=Number(cells.GetValue(r,c+1));
                        float w=Number(cells.GetValue(r,c+2)), h=Number(cells.GetValue(r,c+3));
                        float size=Number(cells.GetValue(r,c+7));
                        if (w<=0 || h<=0 || size<1 || size>72 || Math.Abs(x)>2000 || Math.Abs(y)>2000 || w>2000 || h>2000) throw new ArgumentException("PREVIEW_CELL_SIZE");
                        var rect=new RectangleF(x,y,w,h);
                        using(var fill=new SolidBrush(ColorTranslator.FromOle(Convert.ToInt32(cells.GetValue(r,c+5))))) g.FillRectangle(fill,rect);
                        using(var font=new Font("맑은 고딕",size,Convert.ToBoolean(cells.GetValue(r,c+8))?FontStyle.Bold:FontStyle.Regular,GraphicsUnit.Point))
                            Text(g,Convert.ToString(cells.GetValue(r,c+4)),font,ColorTranslator.FromOle(Convert.ToInt32(cells.GetValue(r,c+6))),rect,Convert.ToInt32(cells.GetValue(r,c+9)));
                    }
                }
                return bitmap;
            } catch {bitmap.Dispose();throw;}
        }
        private static float Number(object value) {
            float number=Convert.ToSingle(value,CultureInfo.InvariantCulture);
            if(float.IsNaN(number)||float.IsInfinity(number))throw new ArgumentException("PREVIEW_NUMBER");
            return number;
        }
        private static void Text(Graphics g,string value,Font font,Color color,RectangleF rect,int align) {
            if(value.Length>10000)value=value.Substring(0,10000);
            using(var brush=new SolidBrush(color)) using(var format=new StringFormat()) {
                format.Alignment=align==2?StringAlignment.Center:align==3?StringAlignment.Far:StringAlignment.Near;
                format.LineAlignment=StringAlignment.Center;format.Trimming=StringTrimming.EllipsisCharacter;
                g.DrawString(value,font,brush,rect,format);
            }
        }
    }
}
