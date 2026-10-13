using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace FakturaLite
{
    public class MainForm : Form
    {
        readonly Store st;
        List<string> units = new List<string>();

        TabControl tabs; TabPage tpInv, tpReg;
        ComboBox cbClient, cbPay, cbComp; Label lblInfo, lblRate, lblBase, lblVat, lblTot, lblNoWarn;
        Panel bottomPnl; Label hintLbl;
        Button btnNewClient, btnClearClient, btnPrint, btnPreview, btnSave, btnClear;
        TextBox txtNo; DateTimePicker dtDate, dtTax; CheckBox chkTax;
        SuggestGrid grid;
        SuggestPopup clientPop, artPop; TextBox artEdit; bool artShowing, artPicking;
        ListView lv; Button btnReprint, btnRePreview, btnCopy;
        bool loading;

        public MainForm(Store store)
        {
            st = store;
            Text = "Фактури — преносима версия (Lite " + Application.ProductVersion + ")";
            Font = SystemFonts.MessageBoxFont;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(1040, 700);
            MinimumSize = new Size(900, 600);
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
            BuildUi();
            LoadLists();
            ClearForm();
            RefreshRegister();
            Shown += delegate { LayoutInv(); };
        }

        /* =========================================================== UI */
        Label L(string text, int x, int y, int w) { Label l = new Label(); l.Text = text; l.Location = new Point(x, y); l.Size = new Size(w, 16); l.ForeColor = Color.DimGray; return l; }

        void BuildUi()
        {
            Panel banner = new Panel(); banner.Dock = DockStyle.Top; banner.Height = 24; banner.BackColor = Color.FromArgb(230, 168, 23);
            Label bl = new Label(); bl.Dock = DockStyle.Fill; bl.TextAlign = ContentAlignment.MiddleLeft; bl.Padding = new Padding(8, 0, 0, 0); bl.ForeColor = Color.FromArgb(28, 34, 48);
            bl.Font = new Font(Font, FontStyle.Bold);
            bl.Text = "ПРЕНОСИМА ВЕРСИЯ (Lite) — данните са в папка invoice-data до програмата. При връщане в офиса ги синхронизирайте с основното приложение.";
            banner.Controls.Add(bl);

            tabs = new TabControl(); tabs.Dock = DockStyle.Fill;
            tpInv = new TabPage("Фактура"); tpReg = new TabPage("Издадени");
            tabs.TabPages.Add(tpInv); tabs.TabPages.Add(tpReg);
            Controls.Add(tabs); Controls.Add(banner);
            banner.BringToFront(); tabs.BringToFront();

            // ---- invoice tab ----
            tpInv.Controls.Add(L("Получател (клиент)", 12, 10, 200));
            cbClient = new ComboBox(); cbClient.Location = new Point(12, 28); cbClient.Size = new Size(430, 22);
            cbClient.DropDownStyle = ComboBoxStyle.DropDown; cbClient.MaxDropDownItems = 30;
            cbClient.AutoCompleteMode = AutoCompleteMode.None;
            // own search list: shows every client when the field is clicked, narrows while typing (any part of the name, EIK, address...)
            clientPop = new SuggestPopup(ClientMatches, delegate (string n) { cbClient.Text = n; cbClient.SelectionStart = n.Length; FillPricesForClient(); });
            cbClient.MouseUp += delegate { ShowClientPop(true); };
            cbClient.DropDown += delegate { BeginInvoke(new MethodInvoker(delegate { cbClient.DroppedDown = false; ShowClientPop(true); })); };
            cbClient.TextUpdate += delegate { ShowClientPop(false); };
            cbClient.KeyDown += delegate (object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Down && !clientPop.Visible) { ShowClientPop(true); e.Handled = true; e.SuppressKeyPress = true; }
                else if (clientPop.HandleKey(e.KeyCode)) { e.Handled = true; e.SuppressKeyPress = true; }
            };
            cbClient.Leave += delegate { clientPop.Hide(); };
            Deactivate += delegate { clientPop.Hide(); artPop.Hide(); };
            cbClient.TextChanged += delegate { ShowClientInfo(); };
            cbClient.SelectedIndexChanged += delegate { ShowClientInfo(); FillPricesForClient(); };
            btnNewClient = new Button(); btnNewClient.Text = "+ нов клиент"; btnNewClient.Location = new Point(450, 27); btnNewClient.Size = new Size(96, 24); btnNewClient.Click += delegate { NewClient(); };
            btnClearClient = new Button(); btnClearClient.Text = "✕"; btnClearClient.Location = new Point(552, 27); btnClearClient.Size = new Size(28, 24); btnClearClient.Click += delegate { cbClient.Text = ""; cbClient.Focus(); };
            lblInfo = new Label(); lblInfo.Location = new Point(12, 54); lblInfo.Size = new Size(570, 32); lblInfo.ForeColor = Color.DimGray;
            tpInv.Controls.AddRange(new Control[] { cbClient, btnNewClient, btnClearClient, lblInfo });

            tpInv.Controls.Add(L("Фактура №", 610, 10, 100));
            txtNo = new TextBox(); txtNo.Location = new Point(610, 28); txtNo.Size = new Size(110, 22); txtNo.Leave += delegate { txtNo.Text = Store.Pad(Digits(txtNo.Text)); };
            tpInv.Controls.Add(txtNo);
            tpInv.Controls.Add(L("Дата на издаване", 730, 10, 130));
            dtDate = new DateTimePicker(); dtDate.Format = DateTimePickerFormat.Custom; dtDate.CustomFormat = "dd.MM.yyyy"; dtDate.Location = new Point(730, 28); dtDate.Size = new Size(120, 22);
            dtDate.ValueChanged += delegate { if (chkTax.Checked) dtTax.Value = dtDate.Value; };
            tpInv.Controls.Add(dtDate);
            tpInv.Controls.Add(L("Дата на данъчно събитие", 862, 10, 160));
            dtTax = new DateTimePicker(); dtTax.Format = DateTimePickerFormat.Custom; dtTax.CustomFormat = "dd.MM.yyyy"; dtTax.Location = new Point(862, 28); dtTax.Size = new Size(120, 22);
            dtTax.ValueChanged += delegate { if (dtTax.Value.Date != dtDate.Value.Date) chkTax.Checked = false; };
            chkTax = new CheckBox(); chkTax.Text = "= дата на издаване"; chkTax.Location = new Point(862, 54); chkTax.Size = new Size(150, 20); chkTax.Checked = true;
            chkTax.CheckedChanged += delegate { if (chkTax.Checked) dtTax.Value = dtDate.Value; };
            tpInv.Controls.AddRange(new Control[] { dtTax, chkTax });

            tpInv.Controls.Add(L("Плащане", 12, 90, 100));
            cbPay = new ComboBox(); cbPay.Location = new Point(12, 108); cbPay.Size = new Size(200, 22); cbPay.DropDownStyle = ComboBoxStyle.DropDownList;
            tpInv.Controls.Add(cbPay);
            tpInv.Controls.Add(L("Съставител", 230, 90, 100));
            cbComp = new ComboBox(); cbComp.Location = new Point(230, 108); cbComp.Size = new Size(260, 22); cbComp.DropDownStyle = ComboBoxStyle.DropDownList;
            tpInv.Controls.Add(cbComp);
            lblRate = new Label(); lblRate.Location = new Point(510, 111); lblRate.Size = new Size(300, 18); lblRate.ForeColor = Color.DimGray;
            tpInv.Controls.Add(lblRate);

            grid = new SuggestGrid(); grid.KeyHook = ArtKey;
            grid.Location = new Point(12, 142); grid.Size = new Size(1000, 380);
            grid.ScrollBars = ScrollBars.Vertical; grid.RowHeadersVisible = false; grid.AllowUserToAddRows = true; grid.AllowUserToDeleteRows = true; grid.AllowUserToResizeRows = false;
            grid.BackgroundColor = Color.White; grid.SelectionMode = DataGridViewSelectionMode.CellSelect; grid.EditMode = DataGridViewEditMode.EditOnEnter;
            grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None; grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            DataGridViewTextBoxColumn cNo = new DataGridViewTextBoxColumn(); cNo.HeaderText = "№"; cNo.Width = 34; cNo.ReadOnly = true; cNo.SortMode = DataGridViewColumnSortMode.NotSortable;
            DataGridViewTextBoxColumn cName = new DataGridViewTextBoxColumn(); cName.HeaderText = "Наименование на стоката / услугата"; cName.Width = 400; cName.SortMode = DataGridViewColumnSortMode.NotSortable;
            DataGridViewComboBoxColumn cCode = new DataGridViewComboBoxColumn(); cCode.HeaderText = "Код"; cCode.Width = 90; cCode.FlatStyle = FlatStyle.Standard; cCode.SortMode = DataGridViewColumnSortMode.NotSortable;
            DataGridViewTextBoxColumn cQty = new DataGridViewTextBoxColumn(); cQty.HeaderText = "Количество"; cQty.Width = 90; cQty.SortMode = DataGridViewColumnSortMode.NotSortable; cQty.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewTextBoxColumn cPrice = new DataGridViewTextBoxColumn(); cPrice.HeaderText = "Ед. цена"; cPrice.Width = 100; cPrice.SortMode = DataGridViewColumnSortMode.NotSortable; cPrice.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            DataGridViewTextBoxColumn cVal = new DataGridViewTextBoxColumn(); cVal.HeaderText = "Стойност"; cVal.Width = 110; cVal.ReadOnly = true; cVal.SortMode = DataGridViewColumnSortMode.NotSortable; cVal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cVal.DefaultCellStyle.BackColor = Color.FromArgb(245, 246, 248);
            grid.Columns.AddRange(new DataGridViewColumn[] { cNo, cName, cCode, cQty, cPrice, cVal });
            grid.Resize += delegate { FitGrid(); };
            grid.EditingControlShowing += GridEditing;
            grid.CellEndEdit += GridCellEnd;
            grid.RowsRemoved += delegate { if (!loading) { Renumber(); Recalc(); } };
            grid.DataError += delegate (object s, DataGridViewDataErrorEventArgs e) { e.ThrowException = false; };
            tpInv.Controls.Add(grid);
            Label hint = L("Съвет: щракнете върху клиента или артикула, за да видите списъка, или пишете част от името, за да търсите. Избирането на клиент и артикул попълва последната цена за този клиент. Delete изтрива избран ред.", 12, 526, 1000);
            hintLbl = hint; tpInv.Controls.Add(hint);

            Panel bottom = new Panel(); bottom.Dock = DockStyle.Bottom; bottom.Height = 70; tpInv.Controls.Add(bottom);
            btnPrint = new Button(); btnPrint.Text = "Печат и запис"; btnPrint.Location = new Point(12, 18); btnPrint.Size = new Size(130, 34); btnPrint.Font = new Font(Font, FontStyle.Bold); btnPrint.Click += delegate { DoPrint(true); };
            btnPreview = new Button(); btnPreview.Text = "Преглед"; btnPreview.Location = new Point(150, 18); btnPreview.Size = new Size(90, 34); btnPreview.Click += delegate { DoPreview(); };
            btnSave = new Button(); btnSave.Text = "Само запис"; btnSave.Location = new Point(248, 18); btnSave.Size = new Size(100, 34); btnSave.Click += delegate { DoSaveOnly(); };
            btnClear = new Button(); btnClear.Text = "Изчисти"; btnClear.Location = new Point(356, 18); btnClear.Size = new Size(90, 34); btnClear.Click += delegate { ClearForm(); };
            lblBase = new Label(); lblVat = new Label(); lblTot = new Label();
            foreach (Label tl in new Label[] { lblBase, lblVat, lblTot }) { tl.Size = new Size(150, 40); tl.TextAlign = ContentAlignment.TopRight; }
            bottomPnl = bottom; tpInv.Resize += delegate { LayoutInv(); }; bottom.Resize += delegate { LayoutInv(); };
            lblTot.Font = new Font(Font.FontFamily, Font.Size + 2, FontStyle.Bold);
            bottom.Controls.AddRange(new Control[] { btnPrint, btnPreview, btnSave, btnClear, lblBase, lblVat, lblTot });

            // ---- issued tab ----
            lv = new ListView(); lv.View = View.Details; lv.FullRowSelect = true; lv.GridLines = true; lv.HideSelection = false; lv.MultiSelect = false;
            lv.Columns.Add("Дата", 90); lv.Columns.Add("№", 100); lv.Columns.Add("Клиент", 330); lv.Columns.Add("Артикули", 360); lv.Columns.Add("Общо", 110, HorizontalAlignment.Right); lv.Columns.Add("Плащане", 100);
            lv.Dock = DockStyle.Fill;
            Panel rb = new Panel(); rb.Dock = DockStyle.Bottom; rb.Height = 52;
            btnReprint = new Button(); btnReprint.Text = "Печат отново"; btnReprint.Location = new Point(12, 10); btnReprint.Size = new Size(120, 32); btnReprint.Click += delegate { ReprintSelected(false); };
            btnRePreview = new Button(); btnRePreview.Text = "Преглед"; btnRePreview.Location = new Point(140, 10); btnRePreview.Size = new Size(90, 32); btnRePreview.Click += delegate { ReprintSelected(true); };
            btnCopy = new Button(); btnCopy.Text = "Копирай в нова фактура"; btnCopy.Location = new Point(238, 10); btnCopy.Size = new Size(170, 32); btnCopy.Click += delegate { CopySelected(); };
            lblNoWarn = new Label(); lblNoWarn.Location = new Point(420, 17); lblNoWarn.Size = new Size(580, 18); lblNoWarn.ForeColor = Color.DimGray;
            lblNoWarn.Text = "Показани са фактурите във флашката. Стари записи без редове могат да се копират само като наименования.";
            rb.Controls.AddRange(new Control[] { btnReprint, btnRePreview, btnCopy, lblNoWarn });
            tpReg.Controls.Add(lv); tpReg.Controls.Add(rb);
            lv.DoubleClick += delegate { ReprintSelected(true); };
        }

        /// <summary>Makes the drop-down wide enough that the longest entry is fully visible.</summary>
        string LongestClient() { int n = 0; foreach (object it in cbClient.Items) n = Math.Max(n, Convert.ToString(it).Length); return n.ToString(); }

        static void FitDropDown(ComboBox cb)
        {
            int w = cb.Width;
            foreach (object it in cb.Items) w = Math.Max(w, TextRenderer.MeasureText(Convert.ToString(it), cb.Font).Width + SystemInformation.VerticalScrollBarWidth + 12);
            cb.DropDownWidth = Math.Min(w, Screen.PrimaryScreen.WorkingArea.Width - 40);
        }

        void LayoutInv()
        {
            if (bottomPnl == null) return;
            int w = tpInv.ClientSize.Width, bt = bottomPnl.Top;
            grid.Size = new Size(Math.Max(300, w - 24), Math.Max(100, bt - 142 - 26));
            hintLbl.Location = new Point(12, bt - 22); hintLbl.Size = new Size(Math.Max(300, w - 24), 18);
            int r = bottomPnl.ClientSize.Width - 12;
            lblTot.Location = new Point(r - 150, 6); lblVat.Location = new Point(r - 310, 6); lblBase.Location = new Point(r - 470, 6);
            FitGrid();
        }

        void FitGrid()
        {
            int others = 0; for (int i = 0; i < grid.Columns.Count; i++) if (i != 1) others += grid.Columns[i].Width;
            int w = grid.ClientSize.Width - others - 2;
            if (w > 120) grid.Columns[1].Width = w;
        }

        static string Digits(string s) { string r = ""; foreach (char c in (s ?? "")) if (c >= '0' && c <= '9') r += c; return r.TrimStart('0').Length == 0 ? (r.Length > 0 ? "0" : "") : r.TrimStart('0'); }

        /* =========================================================== lists & state */
        void LoadLists()
        {
            loading = true;
            cbClient.Items.Clear();
            foreach (string n in st.ClientNames()) cbClient.Items.Add(n);
            units = st.StringList("units", "БР.");
            DataGridViewComboBoxColumn cc = (DataGridViewComboBoxColumn)grid.Columns[2];
            cc.Items.Clear(); cc.Items.Add(""); foreach (string u in units) cc.Items.Add(u);
            cbPay.Items.Clear(); foreach (string p in st.StringList("payments", "В БРОЙ")) cbPay.Items.Add(p);
            cbComp.Items.Clear(); foreach (string c in st.StringList("compilers", "")) cbComp.Items.Add(c);
            if (cbPay.Items.Count > 0 && cbPay.SelectedIndex < 0) cbPay.SelectedIndex = 0;
            if (cbComp.Items.Count > 0 && cbComp.SelectedIndex < 0) cbComp.SelectedIndex = 0;
            FitDropDown(cbClient); FitDropDown(cbPay); FitDropDown(cbComp);
            lblRate.Text = "ДДС ставка: " + st.Rate.ToString("0.##", CultureInfo.InvariantCulture) + "%  (променя се от основното приложение)";
            loading = false;
        }

        public void ClearForm()
        {
            loading = true;
            cbClient.Text = ""; txtNo.Text = Store.Pad(st.NextNo.ToString(CultureInfo.InvariantCulture));
            dtDate.Value = DateTime.Today; chkTax.Checked = true; dtTax.Value = DateTime.Today;
            grid.Rows.Clear();
            for (int i = 0; i < 0; i++) grid.Rows.Add();
            loading = false;
            Renumber(); Recalc(); ShowClientInfo();
        }

        void Renumber()
        {
            loading = true;
            for (int i = 0; i < grid.Rows.Count; i++) if (!grid.Rows[i].IsNewRow) grid.Rows[i].Cells[0].Value = (i + 1).ToString();
            loading = false;
        }

        string Cell(DataGridViewRow r, int i) { object v = r.Cells[i].Value; return v == null ? "" : Convert.ToString(v, CultureInfo.InvariantCulture); }

        public Doc BuildDoc()
        {
            Doc d = new Doc();
            d.No = Digits(txtNo.Text); d.Date = dtDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            d.Tax = chkTax.Checked ? d.Date : dtTax.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            d.Client = cbClient.Text.Trim(); d.Pay = cbPay.SelectedItem as string ?? ""; d.Comp = cbComp.SelectedItem as string ?? ""; d.Rate = st.Rate;
            JObj c = st.FindClient(d.Client);
            if (c != null) { d.CAddr = c.Str("addr"); d.CVat = c.Str("vat"); d.CEik = c.Str("eik"); d.CMol = c.Str("mol"); }
            foreach (DataGridViewRow r in grid.Rows)
            {
                if (r.IsNewRow) continue;
                Item it = new Item(); it.Name = Cell(r, 1).Trim(); it.Code = Cell(r, 2); it.Qty = Cell(r, 3); it.Price = Cell(r, 4);
                d.Items.Add(it);
            }
            Store.Calc(d);
            return d;
        }

        void Recalc()
        {
            if (loading) return;
            Doc d = BuildDoc();
            int i = 0;
            loading = true;
            foreach (DataGridViewRow r in grid.Rows)
            {
                if (r.IsNewRow) continue;
                Item it = d.Items[i++];
                r.Cells[5].Value = it.Used ? Eur(it.Val) : "";
            }
            loading = false;
            lblBase.Text = "Данъчна основа\n" + Eur(d.Base); lblVat.Text = "ДДС " + d.Rate.ToString("0.##", CultureInfo.InvariantCulture) + "%\n" + Eur(d.Vat); lblTot.Text = "За плащане\n" + Eur(d.Total);
        }

        static string Eur(decimal v) { return v.ToString("#,##0.00", CultureInfo.GetCultureInfo("bg-BG")) + " €"; }

        void ShowClientInfo()
        {
            if (loading) return;
            string n = cbClient.Text.Trim();
            JObj c = st.FindClient(n);
            if (n.Length == 0) lblInfo.Text = "";
            else if (c == null) lblInfo.Text = "Клиентът не е в списъка — използвайте „+ нов клиент“.";
            else lblInfo.Text = c.Str("addr") + (c.Str("vat").Length > 0 ? "  ·  ЗДДС: " + c.Str("vat") : "") + (c.Str("eik").Length > 0 ? "  ·  ЕИК: " + c.Str("eik") : "") + (c.Str("mol").Length > 0 ? "  ·  МОЛ: " + c.Str("mol") : "");
        }

        /* =========================================================== grid behaviour */
        void GridEditing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox tb = e.Control as TextBox;
            if (tb == null) return;
            tb.AutoCompleteMode = AutoCompleteMode.None; tb.AutoCompleteSource = AutoCompleteSource.None;
            if (artPop == null) artPop = new SuggestPopup(ArticleMatches, ArtPicked);
            if (grid.CurrentCell == null || grid.CurrentCell.ColumnIndex != 1) { artPop.Hide(); return; }
            // the same edit box is reused for every cell: hook it once
            if (!object.ReferenceEquals(tb, artEdit))
            {
                artEdit = tb;
                tb.MouseUp += delegate { ShowArtPop(true); };
                tb.TextChanged += delegate { if (!artShowing && !artPicking && tb.Focused) ShowArtPop(false); };
                tb.Leave += delegate { artPop.Hide(); };
            }
            artShowing = true;      // the text is being loaded into the box, not typed
            BeginInvoke(new MethodInvoker(delegate { artShowing = false; ShowArtPop(true); }));
        }

        /* ---- search lists (client field and article cells) ---- */
        List<string> ClientMatches(string q)
        {
            List<string> r = new List<string>(); string[] tok = SuggestPopup.Tokens(q);
            foreach (string n in st.ClientNames())
            {
                JObj c = st.FindClient(n);
                string hay = n + " " + (c == null ? "" : c.Str("eik") + " " + c.Str("vat") + " " + c.Str("addr") + " " + c.Str("mol"));
                if (SuggestPopup.Matches(tok, hay)) r.Add(n);
            }
            return r;
        }

        List<string> ArticleMatches(string q)
        {
            List<string> r = new List<string>(); string[] tok = SuggestPopup.Tokens(q);
            foreach (string n in st.ArticleNames()) if (SuggestPopup.Matches(tok, n)) r.Add(n);
            return r;
        }

        void ShowClientPop(bool all)
        {
            if (loading || clientPop == null) return;
            string q = cbClient.Text;
            if (all && (q.Trim().Length == 0 || st.FindClient(q.Trim()) != null)) q = "";   // a chosen client: offer the whole list again
            clientPop.Show(cbClient, cbClient.RectangleToScreen(cbClient.ClientRectangle), q);
        }

        void ShowArtPop(bool all)
        {
            if (loading || artPop == null || artEdit == null || grid.CurrentCell == null || grid.CurrentCell.ColumnIndex != 1 || grid.EditingControl == null) return;
            string q = artEdit.Text;
            if (all && (q.Trim().Length == 0 || st.FindArticle(q.Trim()) != null)) q = "";
            Rectangle cell = grid.GetCellDisplayRectangle(1, grid.CurrentCell.RowIndex, false);
            artPop.Show(grid, grid.RectangleToScreen(cell), q);
        }

        void ArtPicked(string name)
        {
            if (artEdit == null || grid.CurrentCell == null) return;
            artPicking = true; artEdit.Text = name; artEdit.SelectionStart = name.Length; artPicking = false;
            int row = grid.CurrentCell.RowIndex;
            // leave the cell (fills unit and price) and go on to the quantity
            BeginInvoke(new MethodInvoker(delegate { if (row < grid.Rows.Count) grid.CurrentCell = grid.Rows[row].Cells[3]; }));
        }

        /// <summary>Keys for the article list while a name cell is being edited (the grid would otherwise use Up/Down/Enter itself).</summary>
        bool ArtKey(Keys key)
        {
            if (artPop == null || grid.CurrentCell == null || grid.CurrentCell.ColumnIndex != 1 || grid.EditingControl == null) return false;
            if (key == Keys.Down && !artPop.Visible) { ShowArtPop(true); return true; }
            return artPop.HandleKey(key);
        }

        void GridCellEnd(object sender, DataGridViewCellEventArgs e)
        {
            if (artPop != null) artPop.Hide();
            if (loading || e.RowIndex < 0) return;
            DataGridViewRow r = grid.Rows[e.RowIndex];
            if (e.ColumnIndex == 1 && !r.IsNewRow) FillRow(r);
            if (e.ColumnIndex == 2 && r.Cells[2].Value != null) EnsureUnit(Convert.ToString(r.Cells[2].Value));
            Renumber(); Recalc();
        }

        void EnsureUnit(string u)
        {
            DataGridViewComboBoxColumn cc = (DataGridViewComboBoxColumn)grid.Columns[2];
            if (u.Length > 0 && !cc.Items.Contains(u)) cc.Items.Add(u);
        }

        /// <summary>Unit and price from the article list, or the client's last price for it.</summary>
        void FillRow(DataGridViewRow r)
        {
            string name = Cell(r, 1).Trim(); if (name.Length == 0) return;
            JObj art = st.FindArticle(name); JObj deal = cbClient.Text.Trim().Length > 0 ? st.LastDeal(cbClient.Text.Trim(), name) : null;
            string unit = deal != null && deal.Str("code").Length > 0 ? deal.Str("code") : (art != null ? art.Str("unit") : "");
            if (unit.Length == 0 && Cell(r, 2).Length == 0 && units.Count > 0) unit = units[0];
            if (Cell(r, 2).Length == 0 && unit.Length > 0) { EnsureUnit(unit); r.Cells[2].Value = unit; }
            if (Store.Dec(Cell(r, 4)) == 0m)
            {
                if (deal != null) { r.Cells[4].Value = deal.Str("price"); if (Cell(r, 3).Length == 0) r.Cells[3].Value = deal.Str("qty").Length > 0 ? deal.Str("qty") : "1"; }
                else if (art != null && art.Str("price").Length > 0) r.Cells[4].Value = art.Str("price");
            }
        }

        void FillPricesForClient()
        {
            if (loading) return;
            foreach (DataGridViewRow r in grid.Rows) if (!r.IsNewRow) FillRow(r);
            Recalc();
        }

        /* =========================================================== actions */
        string Validate(Doc d)
        {
            if (d.Client.Length == 0) return "Няма избран клиент!";
            if (st.FindClient(d.Client) == null) return "Клиентът не е в списъка. Добавете го с „+ нов клиент“.";
            if (d.No.Length == 0) return "Няма номер на фактурата!";
            if (Store.UsedItems(d).Count == 0) return "Няма добавени стоки!";
            return null;
        }

        void NewClient()
        {
            using (ClientForm f = new ClientForm(cbClient.Text.Trim()))
            {
                if (f.ShowDialog(this) != DialogResult.OK) return;
                if (st.FindClient(f.CName) != null) { MessageBox.Show(this, "Клиент с това име вече съществува.", Text); return; }
                string eik = ""; foreach (char c in f.CEik) if (c >= '0' && c <= '9') eik += c;
                string vat = f.CVat.Length > 0 ? f.CVat : (f.IsVat && eik.Length > 0 ? "BG" + eik : "");
                st.AddClient(f.CName, f.CAddr, eik.Length > 0 ? eik : f.CEik, vat, f.CMol);
                try { st.Save(); } catch (Exception ex) { MessageBox.Show(this, "Не може да се запише: " + ex.Message, Text); }
                LoadLists(); cbClient.Text = f.CName; ShowClientInfo();
            }
        }

        bool Save(Doc d)
        {
            if (st.InvoiceNoExists(d.No) && MessageBox.Show(this, "Фактура № " + d.No + " вече съществува в данните. Да продължа ли?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return false;
            foreach (Item it in Store.UsedItems(d))
                if (it.Name.Length > 0 && st.FindArticle(it.Name) == null)
                {
                    JObj a = new JObj(); a["name"] = it.Name; a["unit"] = it.Code; a["price"] = ""; a["added"] = DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture); st.Articles.Add(a);
                }
            st.AddInvoice(d);
            try { st.Save(); }
            catch (Exception ex) { MessageBox.Show(this, "Не може да се запише във флашката: " + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
            return true;
        }

        void PrintDoc(Doc d)
        {
            PrintDocument pd = InvoicePrint.Create(st, d);
            using (PrintDialog dlg = new PrintDialog())
            {
                dlg.Document = pd; dlg.AllowSomePages = false; dlg.AllowSelection = false;
                if (dlg.ShowDialog(this) == DialogResult.OK) pd.Print();
            }
        }

        void PreviewDoc(Doc d)
        {
            using (PrintPreviewDialog pv = new PrintPreviewDialog())
            {
                pv.Document = InvoicePrint.Create(st, d); pv.WindowState = FormWindowState.Maximized; pv.UseAntiAlias = true;
                pv.ShowDialog(this);
            }
        }

        void DoPrint(bool save)
        {
            Doc d = BuildDoc(); string err = Validate(d);
            if (err != null) { MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (save && !Save(d)) return;
            PrintDoc(d);
            LoadLists(); ClearForm(); RefreshRegister();
        }

        void DoPreview()
        {
            Doc d = BuildDoc(); string err = Validate(d);
            if (err != null) { MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            PreviewDoc(d);
        }

        void DoSaveOnly()
        {
            Doc d = BuildDoc(); string err = Validate(d);
            if (err != null) { MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (!Save(d)) return;
            LoadLists(); ClearForm(); RefreshRegister();
            MessageBox.Show(this, "Фактурата е записана във флашката.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /* =========================================================== issued invoices */
        void RefreshRegister()
        {
            lv.BeginUpdate(); lv.Items.Clear();
            List<JObj> list = new List<JObj>();
            foreach (object o in st.Invoices) { JObj i = o as JObj; if (i != null) list.Add(i); }
            list.Sort(delegate (JObj a, JObj b) { int c = string.CompareOrdinal(b.Str("date"), a.Str("date")); return c != 0 ? c : b.Num("no").CompareTo(a.Num("no")); });
            foreach (JObj i in list)
            {
                string dt = i.Str("date"); if (dt.Length == 10) dt = dt.Substring(8, 2) + "." + dt.Substring(5, 2) + "." + dt.Substring(0, 4);
                List<string> names = new List<string>(); foreach (object l in i.Arr("lines")) { string s = l as string; if (!string.IsNullOrEmpty(s)) names.Add(s); }
                ListViewItem it = new ListViewItem(dt);
                it.SubItems.Add(Store.Pad(i.Str("no"))); it.SubItems.Add(i.Str("client")); it.SubItems.Add(string.Join("; ", names.ToArray()));
                it.SubItems.Add(Eur((decimal)i.Num("total"))); it.SubItems.Add(i.Str("pay"));
                it.Tag = i; lv.Items.Add(it);
            }
            lv.EndUpdate();
        }

        JObj Selected()
        {
            if (lv.SelectedItems.Count == 0) { MessageBox.Show(this, "Изберете фактура от списъка.", Text); return null; }
            return lv.SelectedItems[0].Tag as JObj;
        }

        void ReprintSelected(bool preview)
        {
            JObj rec = Selected(); if (rec == null) return;
            if (rec.Arr("items").Count == 0) { MessageBox.Show(this, "Този запис няма запазени редове (стар запис от Excel), затова не може да се отпечата отново от тук.", Text); return; }
            Doc d = st.FromRecord(rec);
            if (preview) PreviewDoc(d); else PrintDoc(d);
        }

        void CopySelected()
        {
            JObj rec = Selected(); if (rec == null) return;
            ClearForm();
            loading = true;
            cbClient.Text = rec.Str("client");
            if (rec.Arr("items").Count > 0)
                foreach (object o in rec.Arr("items"))
                {
                    JObj li = o as JObj; if (li == null) continue;
                    EnsureUnit(li.Str("code"));
                    grid.Rows.Add(new object[] { "", li.Str("name"), li.Str("code"), li.Str("qty"), li.Str("price"), "" });
                }
            else foreach (object l in rec.Arr("lines")) { string s = l as string; if (!string.IsNullOrEmpty(s)) grid.Rows.Add(new object[] { "", s, units.Count > 0 ? units[0] : "", "1", "", "" }); }
            loading = false;
            Renumber(); Recalc(); ShowClientInfo(); tabs.SelectedTab = tpInv;
        }

        /* =========================================================== test hook */
        /// <summary>Used by the build checks: fills a sample invoice and writes screenshots to a folder.</summary>
        public void Snapshot(string dir)
        {
            Directory.CreateDirectory(dir);
            Show(); Application.DoEvents();
            string first = cbClient.Items.Count > 0 ? (string)cbClient.Items[0] : "";
            cbClient.Text = first;
            loading = true;
            EnsureUnit("БР.");
            grid.Rows.Add(new object[] { "", "ПРИМЕРЕН АПАРАТ МОДЕЛ X100", "БР.", "1", "1250", "" });
            grid.Rows.Add(new object[] { "", "ПОЗИЦИОНЕН КЛЮЧ", "БР.", "3", "12.5", "" });
            grid.Rows.Add(new object[] { "", "ПРИМЕРНА УСЛУГА С МНОГО ДЪЛГО НАИМЕНОВАНИЕ, КОЕТО ПРЕМИНАВА НА НЯКОЛКО РЕДА ВЪВ ФАКТУРАТА", "БР.", "1", "591.67", "" });
            loading = false; Renumber(); Recalc(); ShowClientInfo(); Application.DoEvents();
            using (Bitmap ui = new Bitmap(Width, Height)) { DrawToBitmap(ui, new Rectangle(0, 0, Width, Height)); ui.Save(Path.Combine(dir, "ui.png"), System.Drawing.Imaging.ImageFormat.Png); }
            Doc d = BuildDoc();
            int n = InvoicePrint.Plan(d).Count;
            for (int i = 0; i < n; i++) using (Bitmap b = InvoicePrint.Render(st, d, i, 100)) b.Save(Path.Combine(dir, "print_" + (i + 1) + ".png"), System.Drawing.Imaging.ImageFormat.Png);
            // a long invoice for the pagination check
            Doc big = BuildDoc(); big.Items.Clear();
            for (int i = 1; i <= 30; i++) { Item it = new Item(); it.Name = "АРТИКУЛ ПО ДОГОВОР " + i; it.Code = "БР."; it.Qty = "2"; it.Price = "10.5"; big.Items.Add(it); }
            Store.Calc(big);
            int m = InvoicePrint.Plan(big).Count;
            string lay = "clientDD=" + cbClient.DropDownWidth + "/" + cbClient.Width + " maxItems=" + cbClient.MaxDropDownItems + " longest=" + LongestClient() + " form=" + ClientSize + " tab=" + tpInv.ClientSize + " grid=" + grid.Bounds + " cols=";
            foreach (DataGridViewColumn c in grid.Columns) lay += c.Width + (c.Visible ? "" : "(hidden)") + ",";
            lay += " bottom=" + btnPrint.Parent.Bounds + " btns=" + btnPrint.Bounds + btnClear.Bounds + " totals=" + lblBase.Bounds + lblVat.Bounds + lblTot.Bounds + " texts=[" + lblBase.Text.Replace('\n', ' ') + "|" + lblVat.Text.Replace('\n', ' ') + "|" + lblTot.Text.Replace('\n', ' ') + "] rows=" + grid.Rows.Count;
            File.WriteAllText(Path.Combine(dir, "layout.txt"), lay);
            File.WriteAllText(Path.Combine(dir, "pages.txt"), "short=" + n + " long30=" + m + " words=" + Words.Slov(120.45m) + " | " + Words.Slov(1234.56m) + " | " + Words.Slov(21000m));
            for (int i = 0; i < m; i++) using (Bitmap b = InvoicePrint.Render(st, big, i, 60)) b.Save(Path.Combine(dir, "long_" + (i + 1) + ".png"), System.Drawing.Imaging.ImageFormat.Png);
        }
    }

    /// <summary>Small dialog for a new client.</summary>
    public class ClientForm : Form
    {
        TextBox tName, tAddr, tEik, tVat, tMol; CheckBox cVat;
        public string CName { get { return tName.Text.Trim(); } }
        public string CAddr { get { return tAddr.Text.Trim(); } }
        public string CEik { get { return tEik.Text.Trim(); } }
        public string CVat { get { return tVat.Text.Trim(); } }
        public string CMol { get { return tMol.Text.Trim(); } }
        public bool IsVat { get { return cVat.Checked; } }

        TextBox Field(string label, int y)
        {
            Label l = new Label(); l.Text = label; l.Location = new Point(12, y); l.Size = new Size(120, 18); Controls.Add(l);
            TextBox t = new TextBox(); t.Location = new Point(140, y - 3); t.Size = new Size(300, 22); Controls.Add(t); return t;
        }

        public ClientForm(string name)
        {
            Text = "Нов клиент"; Font = SystemFonts.MessageBoxFont; FormBorderStyle = FormBorderStyle.FixedDialog; MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent; ClientSize = new Size(460, 250);
            tName = Field("Име (ГЛАВНИ БУКВИ)", 16); tAddr = Field("Адрес", 46); tEik = Field("ЕИК / БУЛСТАТ", 76); tVat = Field("Номер по ЗДДС", 106); tMol = Field("МОЛ", 166);
            cVat = new CheckBox(); cVat.Text = "Регистриран по ЗДДС (№ = BG + ЕИК)"; cVat.Location = new Point(140, 134); cVat.Size = new Size(300, 20); Controls.Add(cVat);
            tName.Text = name; tName.MaxLength = 51;
            Button ok = new Button(); ok.Text = "Запази"; ok.Location = new Point(250, 208); ok.Size = new Size(90, 28); ok.DialogResult = DialogResult.OK;
            Button cancel = new Button(); cancel.Text = "Отказ"; cancel.Location = new Point(348, 208); cancel.Size = new Size(90, 28); cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(ok); Controls.Add(cancel); AcceptButton = ok; CancelButton = cancel;
            ok.Click += delegate (object s, EventArgs e)
            {
                if (tName.Text.Trim().Length == 0) { MessageBox.Show(this, "Въведете име.", Text); DialogResult = DialogResult.None; return; }
                string eik = ""; foreach (char c in tEik.Text) if (c >= '0' && c <= '9') eik += c;
                if (tEik.Text.Trim().Length > 0 && eik.Length == 0) eik = tEik.Text.Trim();
                string vat = tVat.Text.Trim().Length > 0 ? tVat.Text.Trim() : (cVat.Checked && eik.Length > 0 ? "BG" + eik : "");
                List<string> w = Validators.PartyWarn(eik, vat);
                if (w.Count > 0 && MessageBox.Show(this, "Проверка: " + string.Join("; ", w.ToArray()) + ".\n\nДа запазя ли клиента въпреки това?", Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) != DialogResult.Yes)
                    DialogResult = DialogResult.None;
            };
        }
    }
}
