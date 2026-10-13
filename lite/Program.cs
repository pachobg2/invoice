using System;
using System.IO;
using System.Windows.Forms;

namespace FakturaLite
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // data lives next to the program: <folder>\invoice-data\data.json (written by the main app when it prepared the flash drive);
            // drives prepared by older versions used the folder name Faktura-data
            string exeDir = Path.GetDirectoryName(Application.ExecutablePath);
            string dataDir = Path.Combine(exeDir, "invoice-data");
            if (!Directory.Exists(dataDir) && Directory.Exists(Path.Combine(exeDir, "Faktura-data"))) dataDir = Path.Combine(exeDir, "Faktura-data");
            string testDir = null;
            foreach (string a in args)
            {
                if (a.StartsWith("/data:")) dataDir = a.Substring(6);       // checks only
                if (a.StartsWith("/snap:")) testDir = a.Substring(6);       // checks only
            }
            Store st = new Store(dataDir);
            if (!File.Exists(st.FilePath))
            {
                MessageBox.Show("Не е намерен файлът:\n" + st.FilePath + "\n\nТази програма работи с данните от папка „invoice-data“, която трябва да е до нея.\nПодгответе флашката от основното приложение: Настройки → Преносима версия.", "Фактури (Lite)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }
            try { st.Load(); }
            catch (Exception ex)
            {
                MessageBox.Show("Файлът с данните не може да се прочете:\n" + ex.Message, "Фактури (Lite)", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }

            foreach (string a in args)
                if (a == "/eik-test")   // checks only: runs the ЕИК rule over every client in the data file
                {
                    int bad = 0, total = 0; string sample = "";
                    foreach (object co in st.Clients)
                    {
                        JObj c = co as JObj; if (c == null) continue; string e = c.Str("eik"); if (e.Length == 0) continue; total++;
                        if (!Validators.EikValid(e)) { bad++; if (bad <= 5) sample += e + ","; }
                    }
                    File.WriteAllText(Path.Combine(st.Dir, "eik-test.txt"), "clients_with_eik=" + total + " invalid=" + bad + " sample=" + sample + " company123456789=" + Validators.EikValid("123456789"));
                    return 0;
                }
            foreach (string a in args)
                if (a.StartsWith("/print-test:"))   // checks only: real printer geometry (no paper is printed) and a print to the PDF printer
                {
                    string dir = a.Substring(12); Directory.CreateDirectory(dir);
                    Doc d = new Doc(); d.No = "3358"; d.Date = "2026-10-13"; d.Tax = d.Date; d.Client = "ТЕСТ КЛИЕНТ"; d.Pay = "В БРОЙ"; d.Comp = "инж. ТЕСТ"; d.Rate = 20;
                    Item it = new Item(); it.Name = "АРТИКУЛ"; it.Code = "БР."; it.Qty = "1"; it.Price = "10"; d.Items.Add(it); Store.Calc(d);
                    string report = "";
                    foreach (string pn in (Environment.GetEnvironmentVariable("PT_PRINTERS") ?? "Microsoft Print to PDF").Split(';'))
                    {
                        System.Drawing.Printing.PrintDocument q = new System.Drawing.Printing.PrintDocument(); q.PrinterSettings.PrinterName = pn;
                        try
                        {
                            System.Drawing.Printing.PageSettings ps = q.DefaultPageSettings;
                            using (System.Drawing.Graphics g = q.PrinterSettings.CreateMeasurementGraphics())
                            {
                                IntPtr h = g.GetHdc();
                                report += pn + ": paper=" + ps.PaperSize.PaperName + " " + ps.PaperSize.Width + "x" + ps.PaperSize.Height + " hard=" + ps.HardMarginX + "," + ps.HardMarginY
                                    + " caps offX=" + InvoicePrint.Caps(h, 112) + " offY=" + InvoicePrint.Caps(h, 113) + " physW=" + InvoicePrint.Caps(h, 110) + " physH=" + InvoicePrint.Caps(h, 111) + " dpi=" + InvoicePrint.Caps(h, 88) + "\r\n";
                                g.ReleaseHdc(h);
                            }
                        }
                        catch (Exception ex) { report += pn + ": " + ex.Message + "\r\n"; }
                    }
                    try
                    {
                        System.Drawing.Printing.PrintDocument pd = InvoicePrint.Create(st, d);
                        pd.PrinterSettings.PrinterName = "Microsoft Print to PDF"; pd.PrinterSettings.PrintToFile = true; pd.PrinterSettings.PrintFileName = Path.Combine(dir, "out.pdf");
                        pd.Print(); report += "pdf printed\r\n";
                    }
                    catch (Exception ex) { report += "pdf failed: " + ex.Message + "\r\n"; }
                    File.WriteAllText(Path.Combine(dir, "print-test.txt"), report);
                    return 0;
                }
            foreach (string a in args)
                if (a.StartsWith("/margin-test:"))   // checks only: where does the ink land for printers with different unprintable margins?
                {
                    string dir = a.Substring(13); Directory.CreateDirectory(dir);
                    Doc d = new Doc(); d.No = "3358"; d.Date = "2026-10-13"; d.Tax = d.Date; d.Client = "ТЕСТ КЛИЕНТ"; d.Pay = "В БРОЙ"; d.Comp = "инж. ТЕСТ"; d.Rate = 20;
                    Item it = new Item(); it.Name = "АРТИКУЛ"; it.Code = "БР."; it.Qty = "1"; it.Price = "10"; d.Items.Add(it); Store.Calc(d);
                    string report = "";
                    foreach (float hm in new float[] { 0f, 16f, 25f })
                        using (System.Drawing.Bitmap b = InvoicePrint.RenderOnPrinter(st, d, 0, 100, hm, hm))
                        {
                            int minX = b.Width, minY = b.Height, maxX = 0;
                            for (int y = 0; y < b.Height; y++) for (int x = 0; x < b.Width; x++) { System.Drawing.Color c = b.GetPixel(x, y); if (c.R < 200 || c.G < 200 || c.B < 200) { if (x < minX) minX = x; if (x > maxX) maxX = x; if (y < minY) minY = y; } }
                            report += "hardMargin=" + hm + " (=" + (hm * 25.4f / 100f).ToString("0.0") + "mm): ink from x=" + (minX * 25.4 / 100).ToString("0.0") + "mm to " + (maxX * 25.4 / 100).ToString("0.0") + "mm, top y=" + (minY * 25.4 / 100).ToString("0.0") + "mm\r\n";
                            b.Save(Path.Combine(dir, "hard_" + (int)hm + ".png"), System.Drawing.Imaging.ImageFormat.Png);
                        }
                    File.WriteAllText(Path.Combine(dir, "margin-test.txt"), report);
                    return 0;
                }
            bool issueTest = false;
            foreach (string a in args) if (a == "/issue-test") issueTest = true;
            if (issueTest)   // checks only: issue one invoice through the same code the window uses, no UI
            {
                Doc d = new Doc(); d.No = st.NextNo.ToString(); d.Date = "2026-10-12"; d.Tax = d.Date; d.Client = "ПЪТЕН КЛИЕНТ ООД"; d.CVat = "BG555666777"; d.CEik = "555666777"; d.Pay = "В БРОЙ"; d.Comp = "инж. ТЕСТ"; d.Rate = st.Rate;
                if (st.FindClient(d.Client) == null) st.AddClient(d.Client, "ВАРНА, УЛ. МОРСКА 1", "555666777", "BG555666777", "ГЕОРГИ ПЪТНИКОВ");
                Item it = new Item(); it.Name = "ПЪТНА УСЛУГА"; it.Code = "ЧАСА"; it.Qty = "2"; it.Price = "50"; d.Items.Add(it);
                Store.Calc(d); st.AddInvoice(d); st.Save();
                return 0;
            }

            MainForm f = new MainForm(st);
            if (testDir != null) { f.Snapshot(testDir); return 0; }
            Application.Run(f);
            return 0;
        }
    }
}
