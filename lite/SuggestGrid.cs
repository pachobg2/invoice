using System;
using System.Windows.Forms;

namespace FakturaLite
{
    /// <summary>DataGridView that lets the form see Up/Down/Enter/Esc before the grid uses them (for the article search list).</summary>
    public class SuggestGrid : DataGridView
    {
        public KeyHandler KeyHook;

        protected override bool ProcessDataGridViewKey(KeyEventArgs e)
        {
            if (KeyHook != null && KeyHook(e.KeyCode)) return true;
            return base.ProcessDataGridViewKey(e);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (KeyHook != null && (keyData & Keys.Modifiers) == Keys.None && KeyHook(keyData & Keys.KeyCode)) return true;
            return base.ProcessDialogKey(keyData);
        }
    }
}
