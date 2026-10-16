using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace FakturaLite
{
    /// <summary>One printed sheet: either original + copy together (short invoice), or one full-page piece of one copy.</summary>
    public class PageSpec
    {
        public bool Two;            // original and copy on the same sheet
        public string Tag = "";     // ОРИГИНАЛ / "" (single-copy pages)
        public int Start, Count;    // item range
        public bool Last;           // shows totals and signatures
        public string Label = "";   // "стр. 1 от 2"
    }

    public static class InvoicePrint
    {
        const int HALF = 7, MID = 34, LAST = 24, CPL = 50;   // same limits as the full app

        static int Weight(Item it) { int n = (it.Name ?? "").Length; return Math.Max(1, (n + CPL - 1) / CPL); }

        public static List<PageSpec> Plan(Doc d)
        {
            List<Item> items = Store.UsedItems(d);
            List<PageSpec> pages = new List<PageSpec>();
            int w = 0; foreach (Item it in items) w += Weight(it);
            if (w <= HALF) { PageSpec p = new PageSpec(); p.Two = true; p.Tag = "ОРИГИНАЛ"; p.Start = 0; p.Count = items.Count; p.Last = true; pages.Add(p); return pages; }

            // split one copy over several pages (the totals need room on the last one)
            List<int[]> chunks = new List<int[]>();           // {start, count}
            int cur = 0, cw = 0, from = 0;
            for (int i = 0; i < items.Count; i++)
            {
                int x = Weight(items[i]);
                if (cur > 0 && cw + x > MID) { chunks.Add(new int[] { from, cur }); from = i; cur = 0; cw = 0; }
                cur++; cw += x;
            }
            chunks.Add(new int[] { from, cur });
            int[] lastC = chunks[chunks.Count - 1]; int lw = 0;
            for (int i = lastC[0]; i < lastC[0] + lastC[1]; i++) lw += Weight(items[i]);
            if (lw > LAST)
            {
                int k = 1;
                while (k < lastC[1])
                {
                    int rest = 0; for (int i = lastC[0] + k; i < lastC[0] + lastC[1]; i++) rest += Weight(items[i]);
                    if (rest <= LAST) break; k++;
                }
                chunks.RemoveAt(chunks.Count - 1);
                chunks.Add(new int[] { lastC[0], k });
                chunks.Add(new int[] { lastC[0] + k, lastC[1] - k });
            }
            for (int copy = 0; copy < 2; copy++)
                for (int c = 0; c < chunks.Count; c++)
                {
                    PageSpec p = new PageSpec();
                    p.Tag = copy == 0 ? "ОРИГИНАЛ" : ""; p.Start = chunks[c][0]; p.Count = chunks[c][1]; p.Last = c == chunks.Count - 1;
                    p.Label = chunks.Count > 1 ? "стр. " + (c + 1) + " от " + chunks.Count : "";
                    pages.Add(p);
                }
            return pages;
        }

        /* ------------------------------ resources ------------------------------ */
        static Image LoadImage(string name)
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (s == null) return null;
                using (Image im = Image.FromStream(s)) return new Bitmap(im);
            }
        }

        /// <summary>The company logo from data.json (a data: URL written by the main app). Data prepared by an older version has no such key: the logo that used to be built in is used then.</summary>
        static Image Logo(Store st)
        {
            if (!st.ShowLogo) return null;
            object raw = st.Root["logoImage"];
            if (raw is string)
            {
                string s = (string)raw; int i = s.IndexOf("base64,");
                if (i < 0) return null;
                try { using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(s.Substring(i + 7)))) using (Image im = Image.FromStream(ms)) return new Bitmap(im); }
                catch { return null; }
            }
            if (raw != null) return null;
            Image m = LoadImage("logo-mark.png"), t = LoadImage("logo-text.png");
            if (m == null || t == null) { if (m != null) m.Dispose(); if (t != null) t.Dispose(); return null; }
            float hm = 320f, wm = hm * m.Width / m.Height, ht = hm * 0.55f, wt = ht * t.Width / t.Height, gap = hm * 0.18f;
            Bitmap b = new Bitmap((int)Math.Round(wm + gap + wt), (int)hm);
            using (Graphics gg = Graphics.FromImage(b)) { gg.InterpolationMode = InterpolationMode.HighQualityBicubic; gg.DrawImage(m, 0f, 0f, wm, hm); gg.DrawImage(t, wm + gap, (hm - ht) / 2f, wt, ht); }
            m.Dispose(); t.Dispose();
            return b;
        }

        static readonly Color Ink = Color.Black, Gray = Color.Black, Navy = Color.Black;
        static readonly Color Head = Color.FromArgb(214, 218, 225), Box = Color.Black, Grid = Color.Black;

        static string Dmy(string iso)
        {
            if (iso != null && iso.Length >= 10 && iso[4] == '-' && iso[7] == '-') return iso.Substring(8, 2) + "." + iso.Substring(5, 2) + "." + iso.Substring(0, 4);
            return iso ?? "";
        }

        static string Eur(decimal v) { return v.ToString("#,##0.00", CultureInfo.GetCultureInfo("bg-BG")).Replace(' ', ' ') + " €"; }
        static string Qty(string q) { decimal d = Store.Dec(q); return d.ToString("0.####", CultureInfo.InvariantCulture); }

        /* ------------------------------ drawing ------------------------------ */
        /// <param name="shiftX">Printer's unprintable left margin in MILLIMETRES (the drawing origin of a printer graphics sits at the printable area, not at the paper corner).</param>
        public static void DrawPage(Graphics g, Store st, Doc d, PageSpec p, float shiftX = 0f, float shiftY = 0f)
        {
            g.PageUnit = GraphicsUnit.Millimeter;
            g.TranslateTransform(-shiftX, -shiftY);   // must come AFTER the unit switch: the world transform is measured in page units
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            List<Item> items = Store.UsedItems(d);
            List<Item> slice = items.GetRange(p.Start, p.Count);
            if (p.Two)
            {
                DrawCopy(g, st, d, 0f, 146f, false, "ОРИГИНАЛ", 0, slice, true, "", HALF);
                using (Pen dash = new Pen(Color.Black, 0.25f)) { dash.DashStyle = DashStyle.Dash; g.DrawLine(dash, 0f, 145.9f, 210f, 145.9f); }
                DrawCopy(g, st, d, 143f, 149f, false, "", 0, slice, true, "", HALF);      // as in the main app: the second copy starts 3 mm below the dotted line
            }
            else DrawCopy(g, st, d, 0f, 296f, true, p.Tag, p.Start, slice, p.Last, p.Label, 0);
        }

        static void Text(Graphics g, string s, Font f, Color c, float x, float y)
        {
            using (SolidBrush b = new SolidBrush(c)) g.DrawString(s, f, b, x, y, StringFormat.GenericTypographic);
        }

        static void TextR(Graphics g, string s, Font f, Color c, float xRight, float y)
        {
            using (SolidBrush b = new SolidBrush(c)) using (StringFormat sf = new StringFormat(StringFormat.GenericTypographic)) { sf.Alignment = StringAlignment.Far; g.DrawString(s, f, b, new RectangleF(xRight - 150f, y, 150f, 8f), sf); }
        }

        static void TextC(Graphics g, string s, Font f, Color c, float x, float y, float w)
        {
            using (SolidBrush b = new SolidBrush(c)) using (StringFormat sf = new StringFormat(StringFormat.GenericTypographic)) { sf.Alignment = StringAlignment.Center; g.DrawString(s, f, b, new RectangleF(x, y, w, 8f), sf); }
        }

        static void Clipped(Graphics g, string s, Font f, Color c, float x, float y, float w)
        {
            using (SolidBrush b = new SolidBrush(c)) using (StringFormat sf = new StringFormat()) { sf.FormatFlags = StringFormatFlags.NoWrap; sf.Trimming = StringTrimming.EllipsisCharacter; g.DrawString(s, f, b, new RectangleF(x, y, w, 5f), sf); }
        }

        static void RoundRect(Graphics g, Pen pen, float x, float y, float w, float h, float r)
        {
            using (GraphicsPath gp = new GraphicsPath())
            {
                gp.AddArc(x, y, 2 * r, 2 * r, 180, 90); gp.AddArc(x + w - 2 * r, y, 2 * r, 2 * r, 270, 90);
                gp.AddArc(x + w - 2 * r, y + h - 2 * r, 2 * r, 2 * r, 0, 90); gp.AddArc(x, y + h - 2 * r, 2 * r, 2 * r, 90, 90);
                gp.CloseFigure(); g.DrawPath(pen, gp);
            }
        }

        /* Every position below was measured on the main app's sheet (millimetres from the top of a copy; the same Arial sizes), so the portable
           version prints the same layout: band, parties, table, totals, bank details and signatures. */
        const float Em = 25.4f / 72f;                               // 1 pt in millimetres
        static readonly Color Dim = Color.Black, Dim2 = Color.Black, SigLine = Color.Black;

        static float SpaceW(Font f) { return 0.278f * f.Size * Em; }

        static float Width(Graphics g, string s, Font f)
        {
            if (s.Length == 0) return 0f;
            float w = g.MeasureString(s, f, 2000, StringFormat.GenericTypographic).Width;
            foreach (char ch in s) if (ch == ' ') w += 0f;           // spaces inside a string are measured by GDI+ itself
            return w;
        }

        /// <summary>Text with letter spacing (GDI+ has none): every character is drawn on its own. Returns the x after the last character (the spacing after it included, like CSS).</summary>
        static float TextLS(Graphics g, string s, Font f, Color c, float x, float y, float ls)
        {
            using (SolidBrush b = new SolidBrush(c))
                foreach (char ch in s)
                {
                    string t = ch.ToString();
                    float w = ch == ' ' ? SpaceW(f) : g.MeasureString(t, f, 2000, StringFormat.GenericTypographic).Width;
                    if (ch != ' ') g.DrawString(t, f, b, x, y, StringFormat.GenericTypographic);
                    x += w + ls;
                }
            return x;
        }

        /// <summary>Grey label + dark value on one line, right aligned (e.g. "дата на издаване: 13.10.2026").</summary>
        static void TextRR(Graphics g, string label, Font fl, Color cl, string value, Font fv, Color cv, float xRight, float y)
        {
            float wv = Width(g, value, fv), wl = Width(g, label, fl), x = xRight - wv - SpaceW(fl) - wl;
            Text(g, label, fl, cl, x, y); Text(g, value, fv, cv, xRight - wv, y);
        }

        static void Party(Graphics g, Store st, float x, float y, float w, float h, string label, string name, bool logo, string addr, string vat, string eik, string mol)
        {
            using (Pen pen = new Pen(Navy, 0.4f)) RoundRect(g, pen, x + 0.2f, y + 0.2f, w - 0.4f, h - 0.4f, 2f);
            using (Font fl = new Font("Arial", 7.125f, FontStyle.Bold)) TextLS(g, label, fl, Gray, x + 3.21f, y + 2.40f, 0.53f);
            using (Font fn = new Font("Arial", 9f, FontStyle.Bold)) Clipped(g, name, fn, Ink, x + 3.21f, y + 6.71f, w - 6.4f);
            string[] lab = { "Адрес:", "№ по ЗДДС:", "ЕИК/ЕГН:", "МОЛ:" }, val = { addr, vat, eik, mol };
            using (Font f = new Font("Arial", 7.875f))
                for (int i = 0; i < 4; i++)
                {
                    float yy = y + 11.81f + i * 3.61f;
                    Text(g, lab[i], f, Gray, x + 3.21f, yy);
                    Clipped(g, val[i], f, Ink, x + 20.59f, yy, w - 23.8f);   // label column 15.79 mm + the 1.59 mm column gap, as in the main app
                }
        }

        static void DrawCopy(Graphics g, Store st, Doc d, float oy, float h, bool full, string tag, int start, List<Item> rows, bool last, string pgLabel, int padWeight)
        {
            JObj co = st.Company;
            float X0 = full ? 13f : 10f, W = 210f - 2f * X0, y = oy + (full ? 12f : (oy < 1f ? 9f : 6f));

            // ---- title band: the title on the left, the company logo in the middle of the sheet, the number and dates on the right ----
            using (Image logo = Logo(st))
                if (logo != null && logo.Width > 0 && logo.Height > 0)
                {
                    float k = Math.Min(8.5f / logo.Height, 42f / logo.Width), lw = logo.Width * k, lh = logo.Height * k;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.DrawImage(logo, 105f - lw / 2f, y + 14.28f - lh, lw, lh);
                }
            float xEnd;
            using (Font ft = new Font("Arial", 18f, FontStyle.Bold)) xEnd = TextLS(g, "ФАКТУРА", ft, Ink, X0, y + 7.51f, 1.32f);
            if (tag.Length > 0)
                using (Font fg = new Font("Arial", 7.5f, FontStyle.Bold))
                {
                    float rx = xEnd + 2.65f, tw = 0f;
                    foreach (char ch in tag) tw += g.MeasureString(ch.ToString(), fg, 2000, StringFormat.GenericTypographic).Width + 0.4f;
                    float rw = tw + 5.04f;
                    using (Pen pen = new Pen(Navy, 0.4f)) g.DrawRectangle(pen, rx + 0.2f, y + 10.52f + 0.2f, rw - 0.4f, 3.92f - 0.4f);
                    TextLS(g, tag, fg, Ink, rx + 2.52f, y + 10.52f + 0.63f, 0.4f);
                }
            using (Font fb = new Font("Arial", 11.25f, FontStyle.Bold))
                TextRR(g, "№", fb, Dim, Store.Pad(d.No), fb, Ink, X0 + W, y + 0.64f);
            using (Font fs = new Font("Arial", 7.875f))
            {
                TextRR(g, "дата на издаване:", fs, Dim, Dmy(d.Date), fs, Ink, X0 + W, y + 6.38f);
                TextRR(g, "дата на данъчно събитие:", fs, Dim, Dmy(d.Tax.Length > 0 ? d.Tax : d.Date), fs, Ink, X0 + W, y + 10.55f);
            }
            if (pgLabel.Length > 0) using (Font fp = new Font("Arial", 7.5f)) TextR(g, pgLabel, fp, Gray, 200f, oy + 2.5f);
            using (Pen rule = new Pen(Navy, 0.53f)) g.DrawLine(rule, X0, y + 16.45f, X0 + W, y + 16.45f);

            // ---- parties ----
            float by = y + 19.71f, bw = (W - 5f) / 2f, bh = 28.26f;
            Party(g, st, X0, by, bw, bh, "П О Л У Ч А Т Е Л", d.Client, false, d.CAddr, d.CVat, d.CEik, d.CMol);
            Party(g, st, X0 + bw + 5f, by, bw, bh, "Д О С Т А В Ч И К", co.Str("name"), false, co.Str("addr"), co.Str("vat"), co.Str("eik"), co.Str("mol"));

            // ---- table ----
            float ty = by + bh + 3f, hh = 4.82f;
            float[] cw = { 8f, W - 8f - 15f - 20f - 24f - 26f, 15f, 20f, 24f, 26f };
            float[] cx = new float[7]; cx[0] = X0; for (int i = 0; i < 6; i++) cx[i + 1] = cx[i] + cw[i];
            string[] heads = { "№", "наименование на стоката и услугата", "код", "количество", "ед. цена", "стойност" };
            using (SolidBrush hb = new SolidBrush(Head)) g.FillRectangle(hb, X0, ty, W, hh);
            using (Pen bp = new Pen(Box, 0.3f)) g.DrawRectangle(bp, X0, ty, W, hh);
            using (Pen vh = new Pen(Color.FromArgb(154, 161, 173), 0.21f)) for (int i = 1; i < 6; i++) g.DrawLine(vh, cx[i], ty, cx[i], ty + hh);      // thin grey column lines, as in the main app
            using (Font fh = new Font("Arial", 7.5f, FontStyle.Bold)) for (int i = 0; i < 6; i++) TextC(g, heads[i], fh, Ink, cx[i], ty + 1.0f, cw[i]);

            float ry = ty + hh; int weightDrawn = 0;
            using (Font fr = new Font("Arial", 7.875f))
            using (Pen lp = new Pen(Grid, 0.25f))
            using (StringFormat wrap = new StringFormat())
            {
                wrap.Trimming = StringTrimming.Word;
                for (int r = 0; r < rows.Count; r++)
                {
                    Item it = rows[r];
                    SizeF sz = g.MeasureString(it.Name, fr, new SizeF(cw[1] - 2.64f, 500f), wrap);
                    float rh = Math.Max(4.9f, sz.Height + 1.1f);
                    TextC(g, (start + r + 1).ToString(CultureInfo.InvariantCulture), fr, Ink, cx[0], ry + 0.86f, cw[0]);
                    using (SolidBrush b = new SolidBrush(Ink)) g.DrawString(it.Name, fr, b, new RectangleF(cx[1] + 1.32f, ry + 0.86f, cw[1] - 2.64f, rh), wrap);
                    TextC(g, it.Code, fr, Ink, cx[2], ry + 0.86f, cw[2]);
                    TextC(g, Qty(it.Qty), fr, Ink, cx[3], ry + 0.86f, cw[3]);
                    TextR(g, Eur(Store.Dec(it.Price)), fr, Ink, cx[5] - 1.32f, ry + 0.86f);
                    TextR(g, Eur(it.Val), fr, Ink, cx[6] - 1.32f, ry + 0.86f);
                    g.DrawLine(lp, X0, ry + rh, X0 + W, ry + rh);
                    using (Pen vp = new Pen(Color.FromArgb(201, 205, 213), 0.21f)) for (int i = 1; i < 6; i++) g.DrawLine(vp, cx[i], ry, cx[i], ry + rh);
                    ry += rh; weightDrawn += Weight(it);
                }
                for (int r = weightDrawn; r < padWeight; r++)   // empty rows keep the form look
                {
                    g.DrawLine(lp, X0, ry + 4.9f, X0 + W, ry + 4.9f);
                    using (Pen vp = new Pen(Color.FromArgb(201, 205, 213), 0.21f)) for (int i = 1; i < 6; i++) g.DrawLine(vp, cx[i], ry, cx[i], ry + 4.9f);
                    ry += 4.9f;
                }
            }
            using (Pen bp = new Pen(Box, 0.3f)) { g.DrawLine(bp, X0, ty + hh, X0, ry); g.DrawLine(bp, X0 + W, ty + hh, X0 + W, ry); g.DrawLine(bp, X0, ry, X0 + W, ry); }

            if (!last)
            {
                using (Font fi = new Font("Arial", 8f, FontStyle.Italic)) TextR(g, "продължава на следващата страница…", fi, Gray, X0 + W, ry + 3f);
                return;
            }

            // ---- words and totals ----
            float sy = ry + 3.09f;
            using (Font fi = new Font("Arial", 7.875f, FontStyle.Italic)) using (Font fv = new Font("Arial", 7.875f, FontStyle.Bold))
            using (StringFormat wrap = new StringFormat()) using (SolidBrush b = new SolidBrush(Ink))
            {
                string l1 = "ДДС словом: ", l2 = "Сума за плащане словом: ";
                float w1 = Width(g, l1, fi) + SpaceW(fi), w2 = Width(g, l2, fi) + SpaceW(fi);
                Text(g, l1, fi, Ink, X0, sy + 1.56f); g.DrawString(Words.Slov(d.Vat), fv, b, new RectangleF(X0 + w1, sy + 1.56f, W - 69f - w1, 9f), wrap);
                Text(g, l2, fi, Ink, X0, sy + 5.87f); g.DrawString(Words.Slov(d.Total), fv, b, new RectangleF(X0 + w2, sy + 5.87f, W - 69f - w2, 9f), wrap);
                if (d.NoVat)
                {
                    string l3 = "Основание за неначисляване на ДДС: "; float w3 = Width(g, l3, fi) + SpaceW(fi);
                    Text(g, l3, fi, Ink, X0, sy + 10.18f); g.DrawString(d.NoVatReason, fv, b, new RectangleF(X0 + w3, sy + 10.18f, W - 69f - w3, 9f), wrap);
                }
            }
            float tx = X0 + W - 64f;
            using (Font f = new Font("Arial", 7.875f))
            {
                Text(g, "дан. основа:", f, Ink, tx + 1.32f, sy + 0.47f); TextR(g, Eur(d.Base), f, Ink, X0 + W - 1.32f, sy + 0.47f);
                Text(g, "ДДС ставка:", f, Ink, tx + 1.32f, sy + 4.61f); TextR(g, d.Rate.ToString("0.##", CultureInfo.InvariantCulture) + "%", f, Ink, X0 + W - 1.32f, sy + 4.61f);
                Text(g, "ДДС:", f, Ink, tx + 1.32f, sy + 8.76f); TextR(g, Eur(d.Vat), f, Ink, X0 + W - 1.32f, sy + 8.76f);
            }
            using (SolidBrush hb = new SolidBrush(Head)) g.FillRectangle(hb, tx, sy + 12.53f, 64f, 5.4f);
            using (Pen np = new Pen(Navy, 0.4f)) { g.DrawLine(np, tx, sy + 12.53f, tx + 64f, sy + 12.53f); g.DrawLine(np, tx, sy + 12.53f + 5.4f, tx + 64f, sy + 12.53f + 5.4f); }
            using (Font fbold = new Font("Arial", 9f, FontStyle.Bold))
            {
                Text(g, "сума за плащане:", fbold, Ink, tx + 1.32f, sy + 13.51f); TextR(g, Eur(d.Total), fbold, Ink, X0 + W - 1.32f, sy + 13.51f);
            }

            // ---- bottom of the copy: bank details and signatures (anchored to the bottom like in the main app) ----
            float ft0 = oy + h - (full ? 9.2f : 8.2f) - 16.93f;
            using (Font f = new Font("Arial", 7.875f)) using (Font fb = new Font("Arial", 7.875f, FontStyle.Bold))
            {
                string[] bl = { "плащане:", "банка:", "BIC код:", "IBAN:" }, bv = { d.Pay, co.Str("bank"), co.Str("bic"), co.Str("iban") };
                for (int i = 0; i < 4; i++)
                {
                    float by2 = ft0 + 0.75f + i * 4.17f, vx = X0 + Width(g, bl[i], f) + SpaceW(f);
                    Text(g, bl[i], f, Gray, X0, by2); Text(g, bv[i], i == 0 ? fb : f, Ink, vx, by2);
                }
            }
            float colW = (W - 12f) / 3.25f, s1 = X0 + colW * 1.25f + 6f, s2 = s1 + colW + 6f;
            using (Pen sp = new Pen(SigLine, 0.3f)) { g.DrawLine(sp, s1, ft0 + 7.99f, s1 + colW, ft0 + 7.99f); g.DrawLine(sp, s2, ft0 + 7.99f, s2 + colW, ft0 + 7.99f); }
            using (Font fc = new Font("Arial", 7.5f)) using (Font fn = new Font("Arial", 8.25f, FontStyle.Bold))
            {
                TextC(g, "подпис на получателя", fc, Dim2, s1, ft0 + 9.5f, colW);
                TextC(g, "съставител", fc, Dim2, s2, ft0 + 9.5f, colW);
                TextC(g, d.Comp, fn, Ink, s2, ft0 + 13.36f, colW);
            }
        }

        /* ------------------------------ printing / preview helpers ------------------------------ */
        [DllImport("gdi32.dll")] static extern int GetDeviceCaps(IntPtr hdc, int index);
        public static int Caps(IntPtr hdc, int index) { return GetDeviceCaps(hdc, index); }

        public static PrintDocument Create(Store st, Doc d)
        {
            PrintDocument pd = new PrintDocument();
            pd.DocumentName = "Фактура " + Store.Pad(d.No);
            pd.DefaultPageSettings.Margins = new Margins(0, 0, 0, 0);
            pd.OriginAtMargins = false;
            List<PageSpec> pages = Plan(d); int idx = 0;
            pd.BeginPrint += delegate { idx = 0; };
            pd.PrintPage += delegate (object s, PrintPageEventArgs e)
            {
                e.Graphics.ResetTransform();
                // The page is laid out exactly like the main app's sheet (A4, from the paper corner). A printer's graphics starts at the
                // printable area instead, so the offset is read from the printer itself (HardMargin only as a fallback) and the sheet is
                // centred if the driver's paper is not exactly 210 mm wide.
                float sx = e.PageSettings.HardMarginX * 25.4f / 100f, sy = e.PageSettings.HardMarginY * 25.4f / 100f, cx = 0f;
                IntPtr hdc = IntPtr.Zero;
                try
                {
                    hdc = e.Graphics.GetHdc();
                    int dx = GetDeviceCaps(hdc, 88), dy = GetDeviceCaps(hdc, 90);                       // LOGPIXELSX / Y
                    int ox = GetDeviceCaps(hdc, 112), oy = GetDeviceCaps(hdc, 113), pw = GetDeviceCaps(hdc, 110);   // PHYSICALOFFSETX / Y, PHYSICALWIDTH
                    if (dx > 0 && dy > 0)
                    {
                        if (ox > 0 || oy > 0) { sx = ox * 25.4f / dx; sy = oy * 25.4f / dy; }
                        float wmm = pw * 25.4f / dx;
                        if (pw > 0 && wmm > 100f && Math.Abs(wmm - 210f) > 1f) cx = (wmm - 210f) / 2f;
                    }
                }
                catch { }
                finally { if (hdc != IntPtr.Zero) e.Graphics.ReleaseHdc(hdc); }
                DrawPage(e.Graphics, st, d, pages[idx], sx - cx, sy);
                idx++; e.HasMorePages = idx < pages.Count;
            };
            return pd;
        }

        /// <summary>For tests / screenshots: draws page <c>index</c> into a bitmap (A4 at the given dpi).</summary>
        /// <summary>Simulates a printer whose printable area starts <paramref name="hardX"/>/<paramref name="hardY"/> (1/100 inch) inside the paper: the page is drawn
        /// the way PrintPage does and then placed at that offset, so the result shows where the ink really lands on the paper.</summary>
        public static Bitmap RenderOnPrinter(Store st, Doc d, int index, int dpi, float hardX, float hardY)
        {
            List<PageSpec> pages = Plan(d);
            Bitmap paper = new Bitmap((int)(8.27 * dpi), (int)(11.69 * dpi)); paper.SetResolution(dpi, dpi);
            float ox = hardX / 100f * dpi, oy = hardY / 100f * dpi;
            using (Bitmap area = new Bitmap(paper.Width - (int)Math.Round(2 * ox), paper.Height - (int)Math.Round(2 * oy)))
            {
                area.SetResolution(dpi, dpi);
                using (Graphics g = Graphics.FromImage(area)) { g.Clear(Color.White); DrawPage(g, st, d, pages[index], hardX * 25.4f / 100f, hardY * 25.4f / 100f); }
                using (Graphics gp = Graphics.FromImage(paper)) { gp.Clear(Color.White); gp.DrawImageUnscaled(area, (int)Math.Round(ox), (int)Math.Round(oy)); }
            }
            return paper;
        }

        public static Bitmap Render(Store st, Doc d, int index, int dpi)
        {
            List<PageSpec> pages = Plan(d);
            Bitmap bmp = new Bitmap((int)(8.27 * dpi), (int)(11.69 * dpi));
            bmp.SetResolution(dpi, dpi);
            using (Graphics g = Graphics.FromImage(bmp)) { g.Clear(Color.White); DrawPage(g, st, d, pages[index]); }
            return bmp;
        }
    }
}
