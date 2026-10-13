using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace FakturaLite
{
    public class Item
    {
        public string Name = "", Code = "", Qty = "", Price = "";
        public decimal Val;     // qty * price, rounded to cents
        public bool Used;       // has a name or an amount
    }

    /// <summary>One invoice as it is shown / printed.</summary>
    public class Doc
    {
        public string No = "", Date = "", Tax = "", Client = "", Pay = "", Comp = "";
        public string CAddr = "", CVat = "", CEik = "", CMol = "";
        public decimal Rate = 20;
        public List<Item> Items = new List<Item>();
        public decimal Base, Vat, Total;
    }

    /// <summary>
    /// Reads/writes invoice-data\data.json — the same file the full app uses — keeping every field it doesn't touch.
    /// </summary>
    public class Store
    {
        public string Dir, FilePath;
        public JObj Root;

        public static string Pad(string no)
        {
            string t = (no ?? "").Trim();
            if (t.Length == 0) return t;
            for (int i = 0; i < t.Length; i++) if (t[i] < '0' || t[i] > '9') return t;
            return t.PadLeft(10, '0');
        }

        public static decimal Dec(string s)
        {
            decimal d;
            if (s != null && decimal.TryParse(s.Trim().Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            return 0m;
        }

        public static decimal R2(decimal d) { return Math.Round(d, 2, MidpointRounding.AwayFromZero); }

        public Store(string dir)
        {
            Dir = dir; FilePath = Path.Combine(dir, "data.json");
        }

        public void Load()
        {
            Root = Json.Parse(File.ReadAllText(FilePath, Encoding.UTF8)) as JObj;
            if (Root == null) throw new FormatException("data.json не съдържа данни на приложението.");
            Root.Arr("clients"); Root.Arr("articles"); Root.Arr("invoices");
        }

        public void Save()
        {
            string tmp = FilePath + ".tmp";
            File.WriteAllText(tmp, Json.Write(Root), new UTF8Encoding(false));
            if (File.Exists(FilePath)) File.Copy(FilePath, FilePath + ".bak", true);
            if (File.Exists(FilePath)) File.Delete(FilePath);
            File.Move(tmp, FilePath);
        }

        /* ------------------ simple accessors ------------------ */
        public JObj Company { get { return Root.Obj("company"); } }
        public List<object> Clients { get { return Root.Arr("clients"); } }
        public List<object> Articles { get { return Root.Arr("articles"); } }
        public List<object> Invoices { get { return Root.Arr("invoices"); } }
        public decimal Rate { get { double r = Root.Num("rate"); return r > 0 ? (decimal)r : 20m; } }
        public bool ShowLogo { get { object v = Root["logo"]; return !(v is bool) || (bool)v; } }

        public long NextNo
        {
            get { double d = Root.Num("nextNo"); return d > 0 ? (long)d : 1; }
            set { Root["nextNo"] = Json.Num(value); }
        }

        public List<string> StringList(string key, string fallback)
        {
            List<string> r = new List<string>();
            foreach (object o in Root.Arr(key)) { string s = o as string; if (!string.IsNullOrEmpty(s)) r.Add(s); }
            if (r.Count == 0 && fallback != null) r.Add(fallback);
            return r;
        }

        public JObj FindClient(string name)
        {
            foreach (object o in Clients) { JObj c = o as JObj; if (c != null && c.Str("name") == name) return c; }
            return null;
        }

        public JObj FindArticle(string name)
        {
            foreach (object o in Articles) { JObj a = o as JObj; if (a != null && a.Str("name") == name) return a; }
            return null;
        }

        public List<string> ClientNames()
        {
            List<string> r = new List<string>();
            foreach (object o in Clients) { JObj c = o as JObj; if (c != null && c.Str("name").Length > 0) r.Add(c.Str("name")); }
            return r;
        }

        public List<string> ArticleNames()
        {
            List<string> r = new List<string>();
            foreach (object o in Articles) { JObj a = o as JObj; if (a != null && a.Str("name").Length > 0) r.Add(a.Str("name")); }
            return r;
        }

        public bool InvoiceNoExists(string no)
        {
            foreach (object o in Invoices) { JObj i = o as JObj; if (i != null && i.Str("no") == no) return true; }
            return false;
        }

        /// <summary>The price a client was last invoiced for an item (from invoices that carry their lines).</summary>
        public JObj LastDeal(string client, string item)
        {
            JObj best = null; string bestDate = "";
            foreach (object o in Invoices)
            {
                JObj inv = o as JObj; if (inv == null || inv.Str("client") != client) continue;
                foreach (object li in inv.Arr("items"))
                {
                    JObj it = li as JObj;
                    if (it != null && it.Str("name") == item && Dec(it.Str("price")) != 0 && string.CompareOrdinal(inv.Str("date"), bestDate) >= 0) { best = it; bestDate = inv.Str("date"); }
                }
            }
            return best;
        }

        /* ------------------ invoice maths ------------------ */
        public static void Calc(Doc d)
        {
            decimal b = 0m;
            foreach (Item it in d.Items)
            {
                it.Val = R2(Dec(it.Qty) * Dec(it.Price));
                it.Used = it.Name.Trim().Length > 0 || it.Val != 0m;
                b += it.Val;
            }
            d.Base = R2(b); d.Vat = R2(d.Base * d.Rate / 100m); d.Total = R2(d.Base + d.Vat);
        }

        public static List<Item> UsedItems(Doc d)
        {
            List<Item> r = new List<Item>();
            foreach (Item it in d.Items) if (it.Used) r.Add(it);
            return r;
        }

        /// <summary>Turns a saved invoice record into a printable Doc (needs the lines to be stored).</summary>
        public Doc FromRecord(JObj rec)
        {
            Doc d = new Doc();
            d.No = rec.Str("no"); d.Date = rec.Str("date"); d.Tax = rec.Str("tax").Length > 0 ? rec.Str("tax") : d.Date;
            d.Client = rec.Str("client"); d.Pay = rec.Str("pay"); d.Comp = rec.Str("comp");
            d.Rate = rec.Num("rate") > 0 ? (decimal)rec.Num("rate") : Rate;
            JObj c = FindClient(d.Client);
            if (c != null) { d.CAddr = c.Str("addr"); d.CVat = c.Str("vat"); d.CEik = c.Str("eik"); d.CMol = c.Str("mol"); }
            else { d.CVat = rec.Str("vat"); d.CEik = rec.Str("eik"); }
            foreach (object o in rec.Arr("items"))
            {
                JObj li = o as JObj; if (li == null) continue;
                Item it = new Item(); it.Name = li.Str("name"); it.Code = li.Str("code"); it.Qty = li.Str("qty"); it.Price = li.Str("price");
                d.Items.Add(it);
            }
            Calc(d);
            return d;
        }

        /// <summary>Appends the invoice to the register (same field layout as the full app) and moves the next number on.</summary>
        public JObj AddInvoice(Doc d)
        {
            JObj rec = new JObj();
            rec["date"] = d.Date; rec["tax"] = d.Tax;
            long no; rec["no"] = long.TryParse(d.No, out no) ? (object)Json.Num(no) : d.No;
            rec["client"] = d.Client; rec["vat"] = d.CVat; rec["eik"] = d.CEik;
            List<object> lines = new List<object>(); List<object> items = new List<object>();
            foreach (Item it in UsedItems(d))
            {
                lines.Add(it.Name);
                JObj li = new JObj(); li["name"] = it.Name; li["code"] = it.Code; li["qty"] = it.Qty; li["price"] = it.Price; items.Add(li);
            }
            rec["lines"] = lines;
            rec["base"] = Json.Num((double)d.Base); rec["dds"] = Json.Num((double)d.Vat); rec["total"] = Json.Num((double)d.Total);
            rec["pay"] = d.Pay; rec["comp"] = d.Comp; rec["rate"] = Json.Num((double)d.Rate);
            rec["items"] = items;
            Invoices.Add(rec);
            if (long.TryParse(d.No, out no) && no >= NextNo) NextNo = no + 1;
            return rec;
        }

        public JObj AddClient(string name, string addr, string eik, string vat, string mol)
        {
            JObj c = new JObj();
            c["name"] = name; c["addr"] = addr; c["vat"] = vat; c["eik"] = eik; c["mol"] = mol;
            Clients.Add(c);
            return c;
        }
    }
}
