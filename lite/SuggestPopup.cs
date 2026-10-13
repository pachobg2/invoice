using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FakturaLite
{
    public delegate List<string> MatchSource(string query);
    public delegate bool KeyHandler(Keys key);

    /// <summary>A search drop-down under a text field: shows every item when the field is clicked, narrows the list (any part of the text, any word order)
    /// while typing, and is driven with the mouse or Up/Down/Enter/Esc. It never takes the keyboard focus from the field.</summary>
    /// <summary>A list that never takes the keyboard focus, so clicking it does not make the text field lose focus (and close the list).</summary>
    class NoFocusList : ListBox
    {
        public NoFocusList() { SetStyle(ControlStyles.Selectable, false); TabStop = false; }
    }

    public class SuggestPopup
    {
        readonly ListBox list = new NoFocusList();
        readonly ToolStripDropDown drop = new ToolStripDropDown();
        readonly MatchSource source;
        readonly Action<string> pick;
        const int MaxRows = 14;

        public SuggestPopup(MatchSource source, Action<string> pick)
        {
            this.source = source; this.pick = pick;
            list.IntegralHeight = false; list.BorderStyle = BorderStyle.FixedSingle; list.HorizontalScrollbar = true;
            list.MouseDown += delegate (object s, MouseEventArgs e)
            {
                int i = list.IndexFromPoint(e.Location);
                if (i >= 0 && i < list.Items.Count) Choose(Convert.ToString(list.Items[i]));
            };
            list.MouseMove += delegate (object s, MouseEventArgs e)
            {
                int i = list.IndexFromPoint(e.Location);
                if (i >= 0 && i != list.SelectedIndex) list.SelectedIndex = i;
            };
            ToolStripControlHost host = new ToolStripControlHost(list);
            host.Margin = Padding.Empty; host.Padding = Padding.Empty; host.AutoSize = false;
            drop.Items.Add(host);
            drop.AutoClose = false; drop.Padding = Padding.Empty; drop.Margin = Padding.Empty; drop.DropShadowEnabled = true;
        }

        public bool Visible { get { return drop.Visible; } }

        [System.Runtime.InteropServices.DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        /// <summary>True while this list itself is the active window (some Windows versions make it active, which deactivates the main form).</summary>
        public bool OwnsForeground { get { return drop.Visible && GetForegroundWindow() == drop.Handle; } }

        public static string[] Tokens(string q)
        {
            return (q ?? "").Trim().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }

        public static bool Matches(string[] tokens, string haystack)
        {
            foreach (string t in tokens) if (haystack.IndexOf(t, StringComparison.OrdinalIgnoreCase) < 0) return false;
            return true;
        }

        void Choose(string item) { Hide(); pick(item); }

        /// <summary>Shows the matches for <paramref name="query"/> under the given screen rectangle (the field / cell).</summary>
        public void Show(Control owner, Rectangle fieldOnScreen, string query)
        {
            List<string> items = source(query);
            if (items.Count == 0) { Hide(); return; }
            list.BeginUpdate(); list.Items.Clear();
            int w = fieldOnScreen.Width;
            using (Graphics g = list.CreateGraphics())
                foreach (string s in items) { list.Items.Add(s); w = Math.Max(w, (int)g.MeasureString(s, owner.Font).Width + 28); }
            list.EndUpdate();
            Rectangle wa = Screen.FromControl(owner).WorkingArea;
            w = Math.Min(w, Math.Max(fieldOnScreen.Width, wa.Width - 40));
            int rowH = list.ItemHeight, rows = Math.Min(items.Count, MaxRows);
            int h = rows * rowH + 4;
            int x = Math.Min(fieldOnScreen.Left, wa.Right - w - 4), y = fieldOnScreen.Bottom;
            if (y + h > wa.Bottom) y = Math.Max(wa.Top, fieldOnScreen.Top - h);       // not enough room below: open above
            list.Font = owner.Font; list.Size = new Size(w, h);
            drop.Size = new Size(w, h);
            list.SelectedIndex = -1;
            if (drop.Visible) drop.Location = new Point(x, y); else drop.Show(new Point(x, y));
        }

        public void Hide() { if (drop.Visible) drop.Close(); }

        /// <summary>Navigation keys of the field. Returns true when the key was used by the list.</summary>
        public bool HandleKey(Keys key)
        {
            if (!drop.Visible) return false;
            int n = list.Items.Count; if (n == 0) return false;
            switch (key)
            {
                case Keys.Down: list.SelectedIndex = Math.Min(n - 1, list.SelectedIndex + 1); return true;
                case Keys.Up: list.SelectedIndex = list.SelectedIndex <= 0 ? 0 : list.SelectedIndex - 1; return true;
                case Keys.PageDown: list.SelectedIndex = Math.Min(n - 1, Math.Max(0, list.SelectedIndex) + MaxRows - 1); return true;
                case Keys.PageUp: list.SelectedIndex = Math.Max(0, list.SelectedIndex - MaxRows + 1); return true;
                case Keys.Escape: Hide(); return true;
                case Keys.Enter:
                    if (list.SelectedIndex >= 0) { Choose(Convert.ToString(list.SelectedItem)); return true; }
                    Hide(); return false;
            }
            return false;
        }
    }
}
