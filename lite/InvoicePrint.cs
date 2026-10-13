using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Reflection;

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
        const int HALF = 8, MID = 34, LAST = 24, CPL = 50;   // same limits as the full app

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

        static readonly Color Ink = Color.FromArgb(17, 17, 17), Gray = Color.FromArgb(102, 102, 102), Navy = Color.FromArgb(29, 35, 48);
        static readonly Color Head = Color.FromArgb(233, 236, 241), Box = Color.FromArgb(154, 161, 173), Grid = Color.FromArgb(201, 205, 213);

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
            // safe area: shrink the sheet to 93% around its centre so nothing sits near the paper edge (printers clip 5-8 mm)
            const float SAFE = 0.93f;
            g.TranslateTransform(210f * (1f - SAFE) / 2f, 297f * (1f - SAFE) / 2f);
            g.ScaleTransform(SAFE, SAFE);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.AntiAlias;
            List<Item> items = Store.UsedItems(d);
            List<Item> slice = items.GetRange(p.Start, p.Count);
            if (p.Two)
            {
                DrawCopy(g, st, d, 0f, 146f, false, "ОРИГИНАЛ", 0, slice, true, "", HALF);
                using (Pen dash = new Pen(Color.FromArgb(170, 170, 170), 0.15f)) { dash.DashStyle = DashStyle.Dash; g.DrawLine(dash, 5f, 147.2f, 205f, 147.2f); }
                DrawCopy(g, st, d, 148.5f, 146f, false, "", 0, slice, true, "", HALF);
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

        static void Party(Graphics g, Store st, float x, float y, float w, float h, string label, string name, bool logo, string addr, string vat, string eik, string mol)
        {
            using (Pen pen = new Pen(Box, 0.2f)) RoundRect(g, pen, x, y, w, h, 1.6f);
            using (Font fl = new Font("Arial", 6f, FontStyle.Bold)) Text(g, label, fl, Gray, x + 2.5f, y + 1.8f);
            if (logo)
            {
                using (Image m = LoadImage("logo-mark.png")) using (Image t = LoadImage("logo-text.png"))
                {
                    float hm = 6.6f, wm = m != null ? hm * m.Width / m.Height : 0f, ht = 4.3f, wt = t != null ? ht * t.Width / t.Height : 0f;
                    if (m != null) g.DrawImage(m, x + 2.5f, y + 4.8f, wm, hm);
                    if (t != null) g.DrawImage(t, x + 2.5f + wm + 2f, y + 4.8f + (hm - ht) / 2f, wt, ht);
                    if (m == null && t == null) using (Font fn = new Font("Arial", 9f, FontStyle.Bold)) Text(g, name, fn, Ink, x + 2.5f, y + 5f);
                }
            }
            else using (Font fn = new Font("Arial", 9f, FontStyle.Bold)) Clipped(g, name, fn, Ink, x + 2.5f, y + 5f, w - 5f);
            string[] lab = { "Адрес:", "№ по ЗДДС:", "ЕИК/ЕГН:", "МОЛ:" }, val = { addr, vat, eik, mol };
            using (Font f = new Font("Arial", 7f))
                for (int i = 0; i < 4; i++)
                {
                    float yy = y + 11.4f + i * 3.05f;
                    Text(g, lab[i], f, Gray, x + 2.5f, yy);
                    Clipped(g, val[i], f, Ink, x + 20f, yy, w - 22f);
                }
        }

        static void DrawCopy(Graphics g, Store st, Doc d, float oy, float h, bool full, string tag, int start, List<Item> rows, bool last, string pgLabel, int padWeight)
        {
            JObj co = st.Company;
            float X0 = full ? 13f : 10f, W = 210f - 2f * X0, y = oy + (full ? 12f : 6f);

            // ---- title band ----
            float titleW;
            using (Font ft = new Font("Arial", 19f, FontStyle.Bold))
            {
                Text(g, "ФАКТУРА", ft, Ink, X0, y);
                titleW = g.MeasureString("ФАКТУРА", ft, 200, StringFormat.GenericTypographic).Width;
            }
            if (tag.Length > 0)
                using (Font fg = new Font("Arial", 7f, FontStyle.Bold))
                {
                    float tw = g.MeasureString(tag, fg, 200, StringFormat.GenericTypographic).Width;
                    using (Pen pen = new Pen(Navy, 0.35f)) g.DrawRectangle(pen, X0 + titleW + 4f, y + 2.6f, tw + 4f, 4.4f);
                    Text(g, tag, fg, Ink, X0 + titleW + 6f, y + 3.6f);
                }
            using (Font fb = new Font("Arial", 11f, FontStyle.Bold)) TextR(g, "№ " + Store.Pad(d.No), fb, Ink, X0 + W, y - 0.5f);
            using (Font fs = new Font("Arial", 7.5f))
            {
                TextR(g, "дата на издаване: " + Dmy(d.Date), fs, Ink, X0 + W, y + 5.2f);
                TextR(g, "дата на данъчно събитие: " + Dmy(d.Tax.Length > 0 ? d.Tax : d.Date), fs, Ink, X0 + W, y + 8.7f);
            }
            if (pgLabel.Length > 0) using (Font fp = new Font("Arial", 7f)) TextR(g, pgLabel, fp, Gray, X0 + W, oy + 2f);
            using (Pen rule = new Pen(Navy, 0.5f)) g.DrawLine(rule, X0, y + 12.4f, X0 + W, y + 12.4f);

            // ---- parties ----
            float by = y + 15f, bw = (W - 5f) / 2f, bh = full ? 26f : 25f;
            Party(g, st, X0, by, bw, bh, "П О Л У Ч А Т Е Л", d.Client, false, d.CAddr, d.CVat, d.CEik, d.CMol);
            Party(g, st, X0 + bw + 5f, by, bw, bh, "Д О С Т А В Ч И К", co.Str("name"), st.ShowLogo, co.Str("addr"), co.Str("vat"), co.Str("eik"), co.Str("mol"));

            // ---- table ----
            float ty = by + bh + 3f;
            float[] cw = { 8f, W - 8f - 15f - 20f - 24f - 26f, 15f, 20f, 24f, 26f };
            float[] cx = new float[7]; cx[0] = X0; for (int i = 0; i < 6; i++) cx[i + 1] = cx[i] + cw[i];
            string[] heads = { "№", "наименование на стоката и услугата", "код", "количество", "ед. цена", "стойност" };
            using (SolidBrush hb = new SolidBrush(Head)) g.FillRectangle(hb, X0, ty, W, 5.2f);
            using (Pen bp = new Pen(Box, 0.2f)) g.DrawRectangle(bp, X0, ty, W, 5.2f);
            using (Pen gp = new Pen(Box, 0.15f)) for (int i = 1; i < 6; i++) g.DrawLine(gp, cx[i], ty, cx[i], ty + 5.2f);
            using (Font fh = new Font("Arial", 7f, FontStyle.Bold)) for (int i = 0; i < 6; i++) TextC(g, heads[i], fh, Ink, cx[i], ty + 1.4f, cw[i]);

            float ry = ty + 5.2f; int weightDrawn = 0;
            using (Font fr = new Font("Arial", 8f))
            using (Pen lp = new Pen(Grid, 0.15f))
            using (StringFormat wrap = new StringFormat())
            {
                wrap.Trimming = StringTrimming.Word;
                for (int r = 0; r < rows.Count; r++)
                {
                    Item it = rows[r];
                    SizeF sz = g.MeasureString(it.Name, fr, new SizeF(cw[1] - 2f, 500f), wrap);
                    float rh = Math.Max(4.9f, sz.Height + 1.1f);
                    TextC(g, (start + r + 1).ToString(CultureInfo.InvariantCulture), fr, Ink, cx[0], ry + 0.9f, cw[0]);
                    using (SolidBrush b = new SolidBrush(Ink)) g.DrawString(it.Name, fr, b, new RectangleF(cx[1] + 1f, ry + 0.9f, cw[1] - 2f, rh), wrap);
                    TextC(g, it.Code, fr, Ink, cx[2], ry + 0.9f, cw[2]);
                    TextC(g, Qty(it.Qty), fr, Ink, cx[3], ry + 0.9f, cw[3]);
                    TextR(g, Eur(Store.Dec(it.Price)), fr, Ink, cx[5] - 1.2f, ry + 0.9f);
                    TextR(g, Eur(it.Val), fr, Ink, cx[6] - 1.2f, ry + 0.9f);
                    g.DrawLine(lp, X0, ry + rh, X0 + W, ry + rh);
                    for (int i = 0; i <= 6; i++) g.DrawLine(lp, cx[i], ry, cx[i], ry + rh);
                    ry += rh; weightDrawn += Weight(it);
                }
                for (int r = weightDrawn; r < padWeight; r++)   // empty rows keep the form look
                {
                    g.DrawLine(lp, X0, ry + 4.9f, X0 + W, ry + 4.9f);
                    for (int i = 0; i <= 6; i++) g.DrawLine(lp, cx[i], ry, cx[i], ry + 4.9f);
                    ry += 4.9f;
                }
            }
            using (Pen bp = new Pen(Box, 0.2f)) { g.DrawLine(bp, X0, ty + 5.2f, X0, ry); g.DrawLine(bp, X0 + W, ty + 5.2f, X0 + W, ry); g.DrawLine(bp, X0, ry, X0 + W, ry); }

            if (!last)
            {
                using (Font fi = new Font("Arial", 8f, FontStyle.Italic)) TextR(g, "продължава на следващата страница…", fi, Gray, X0 + W, ry + 3f);
                return;
            }

            // ---- totals + words ----
            float sy = ry + 3.5f;
            using (Font fi = new Font("Arial", 7.8f, FontStyle.Italic))
            using (StringFormat wrap = new StringFormat())
            using (SolidBrush b = new SolidBrush(Ink))
            {
                g.DrawString("ДДС словом: " + Words.Slov(d.Vat), fi, b, new RectangleF(X0, sy + 0.5f, W - 70f, 9f), wrap);
                g.DrawString("Сума за плащане словом: " + Words.Slov(d.Total), fi, b, new RectangleF(X0, sy + 8.5f, W - 70f, 9f), wrap);
            }
            float tx = X0 + W - 64f;
            using (Font f = new Font("Arial", 8f))
            {
                Text(g, "дан. основа:", f, Ink, tx + 1.5f, sy + 0.5f); TextR(g, Eur(d.Base), f, Ink, X0 + W - 1.5f, sy + 0.5f);
                Text(g, "ДДС ставка:", f, Ink, tx + 1.5f, sy + 4.3f); TextR(g, d.Rate.ToString("0.##", CultureInfo.InvariantCulture) + "%", f, Ink, X0 + W - 1.5f, sy + 4.3f);
                Text(g, "ДДС:", f, Ink, tx + 1.5f, sy + 8.1f); TextR(g, Eur(d.Vat), f, Ink, X0 + W - 1.5f, sy + 8.1f);
            }
            using (SolidBrush hb = new SolidBrush(Head)) g.FillRectangle(hb, tx, sy + 12.2f, 64f, 5.6f);
            using (Pen np = new Pen(Navy, 0.35f)) g.DrawLine(np, tx, sy + 12.2f, tx + 64f, sy + 12.2f);
            using (Font fbold = new Font("Arial", 9.5f, FontStyle.Bold))
            {
                Text(g, "сума за плащане:", fbold, Ink, tx + 1.5f, sy + 13.2f); TextR(g, Eur(d.Total), fbold, Ink, X0 + W - 1.5f, sy + 13.2f);
            }

            // ---- bottom: bank details + signatures ----
            float fy = oy + h - (full ? 9f : 4f) - 20f;
            using (Font f = new Font("Arial", 7.3f)) using (Font fb = new Font("Arial", 7.3f, FontStyle.Bold))
            {
                Text(g, "плащане:", f, Gray, X0, fy); Text(g, d.Pay, fb, Ink, X0 + 13f, fy);
                Text(g, "банка:", f, Gray, X0, fy + 3.6f); Text(g, co.Str("bank"), f, Ink, X0 + 13f, fy + 3.6f);
                Text(g, "BIC код:", f, Gray, X0, fy + 7.2f); Text(g, co.Str("bic"), f, Ink, X0 + 13f, fy + 7.2f);
                Text(g, "IBAN:", f, Gray, X0, fy + 10.8f); Text(g, co.Str("iban"), f, Ink, X0 + 13f, fy + 10.8f);
            }
            float colW = (W - 12f) / 3.25f, s1 = X0 + colW * 1.25f + 6f, s2 = s1 + colW + 6f;
            using (Pen sp = new Pen(Ink, 0.2f)) { g.DrawLine(sp, s1, fy + 12f, s1 + colW, fy + 12f); g.DrawLine(sp, s2, fy + 12f, s2 + colW, fy + 12f); }
            using (Font fc = new Font("Arial", 6.5f)) using (Font fn = new Font("Arial", 7.5f, FontStyle.Bold))
            {
                TextC(g, "подпис на получателя", fc, Gray, s1, fy + 12.8f, colW);
                TextC(g, "съставител", fc, Gray, s2, fy + 12.8f, colW);
                TextC(g, d.Comp, fn, Ink, s2, fy + 16f, colW);
            }
        }

        /* ------------------------------ printing / preview helpers ------------------------------ */
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
                // work in millimetres from the physical paper corner: HardMargin is reported in 1/100 inch
                DrawPage(e.Graphics, st, d, pages[idx], e.PageSettings.HardMarginX * 25.4f / 100f, e.PageSettings.HardMarginY * 25.4f / 100f);
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
