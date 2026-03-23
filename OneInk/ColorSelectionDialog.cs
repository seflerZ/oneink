/*
 *  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 *  Color selection dialog for deleting ink by color.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace OneInk
{
    public partial class ColorSelectionDialog : Form
    {
        public string SelectedColor { get; private set; }

        public ColorSelectionDialog(List<string> colors)
        {
            InitializeComponent();

            colorImageList.ColorDepth = ColorDepth.Depth32Bit;

            for (int i = 0; i < colors.Count; i++)
            {
                string colorHex = colors[i];
                Color c;
                try { c = ColorTranslator.FromHtml(colorHex); }
                catch { c = Color.Black; }

                var bmp = new Bitmap(48, 48);
                using (var g = Graphics.FromImage(bmp))
                {
                    using (var brush = new SolidBrush(c))
                        g.FillRectangle(brush, 0, 0, 48, 48);
                    using (var pen = new Pen(Color.FromArgb(128, 128, 128)))
                        g.DrawRectangle(pen, 0, 0, 47, 47);
                }

                colorImageList.Images.Add(bmp);
                var item = new ListViewItem(colorHex, i) { Tag = colorHex };
                colorListView.Items.Add(item);
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            if (colorListView.SelectedItems.Count == 0)
            {
                MessageBox.Show(Strings.NoSelection, Strings.NoSelectionTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            SelectedColor = colorListView.SelectedItems[0].Tag as string;
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void colorListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (colorListView.SelectedItems.Count > 0)
                SelectedColor = colorListView.SelectedItems[0].Tag as string;
        }

        private void colorListView_DoubleClick(object sender, EventArgs e)
        {
            if (colorListView.SelectedItems.Count > 0)
            {
                SelectedColor = colorListView.SelectedItems[0].Tag as string;
                DialogResult = DialogResult.OK;
                Close();
            }
        }
    }
}
