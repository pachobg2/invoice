using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace FakturaLite
{
    /// <summary>Print preview: the pages exactly as they lie on the A4 paper (same layout as the main app), with a Print button.</summary>
    public class PreviewForm : Form
    {
        readonly List<Bitmap> bitmaps = new List<Bitmap>();
        readonly List<PictureBox> boxes = new List<PictureBox>();
        readonly Panel scroll;

        public PreviewForm(Store st, Doc d, MethodInvoker print)
        {
            Text = "Преглед — фактура " + Store.Pad(d.No);
            Font = SystemFonts.MessageBoxFont; StartPosition = FormStartPosition.CenterParent; WindowState = FormWindowState.Maximized; ShowInTaskbar = false;
            Panel top = new Panel(); top.Dock = DockStyle.Top; top.Height = 44; top.BackColor = SystemColors.Control;
            Button bPrint = new Button(); bPrint.Text = "Печат…"; bPrint.Location = new Point(10, 8); bPrint.Size = new Size(110, 28);
            bPrint.Click += delegate { if (print != null) print(); };
            Button bClose = new Button(); bClose.Text = "Затвори"; bClose.Location = new Point(130, 8); bClose.Size = new Size(110, 28);
            bClose.Click += delegate { Close(); };
            Label hint = new Label(); hint.Text = "Така ще изглежда листът (A4)."; hint.AutoSize = true; hint.Location = new Point(255, 14); hint.ForeColor = Color.DimGray;
            top.Controls.AddRange(new Control[] { bPrint, bClose, hint });
            scroll = new Panel(); scroll.Dock = DockStyle.Fill; scroll.AutoScroll = true; scroll.BackColor = Color.FromArgb(150, 150, 150);
            Controls.Add(scroll); Controls.Add(top);
            top.BringToFront();
            int n = InvoicePrint.Plan(d).Count;
            for (int i = 0; i < n; i++)
            {
                Bitmap b = InvoicePrint.Render(st, d, i, 110); bitmaps.Add(b);
                PictureBox pb = new PictureBox(); pb.Image = b; pb.SizeMode = PictureBoxSizeMode.AutoSize; pb.BorderStyle = BorderStyle.FixedSingle; boxes.Add(pb); scroll.Controls.Add(pb);
            }
            scroll.Resize += delegate { Place(); };
            Shown += delegate { Place(); };
            KeyPreview = true; KeyDown += delegate (object s, KeyEventArgs e) { if (e.KeyCode == Keys.Escape) Close(); };
        }

        void Place()
        {
            int y = 12 + scroll.AutoScrollPosition.Y;
            foreach (PictureBox pb in boxes)
            {
                pb.Location = new Point(Math.Max(12, (scroll.ClientSize.Width - pb.Width) / 2) + scroll.AutoScrollPosition.X, y);
                y += pb.Height + 12;
            }
            scroll.AutoScrollMinSize = new Size(boxes.Count > 0 ? boxes[0].Width + 24 : 0, y - scroll.AutoScrollPosition.Y);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            foreach (Bitmap b in bitmaps) b.Dispose();
        }
    }
}
