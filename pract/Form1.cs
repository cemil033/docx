using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pract
{
    public partial class btn_save : Form
    {
        public btn_save()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void btn_save_Load(object sender, EventArgs e)
        {
            foreach (var item in FontFamily.Families)
            {
                cb_fonts.Items.Add(item.Name);
            }
            cb_sizes.Items.Add(8);
            cb_sizes.Items.Add(9);
            cb_sizes.Items.Add(10);
            cb_sizes.Items.Add(11);
            cb_sizes.Items.Add(12);
            cb_sizes.Items.Add(14);
            cb_sizes.Items.Add(16);
            cb_sizes.Items.Add(18);
            cb_sizes.Items.Add(20);
            cb_sizes.Items.Add(22);
            cb_sizes.Items.Add(24);
            cb_sizes.Items.Add(26);
            cb_sizes.Items.Add(28);
            cb_sizes.Items.Add(36);
            cb_sizes.Items.Add(48);
            cb_sizes.Items.Add(72);
            foreach (Color color in new ColorConverter().GetStandardValues())
            {
                cb_color.Items.Add(color.Name);
            }
            cb_bold.Appearance = System.Windows.Forms.Appearance.Button;
            cb_i.Appearance = System.Windows.Forms.Appearance.Button;
            cb_und.Appearance = System.Windows.Forms.Appearance.Button;
            rb_cnt.Appearance = System.Windows.Forms.Appearance.Button;
            rb_lft.Appearance = System.Windows.Forms.Appearance.Button;
            rb_r.Appearance = System.Windows.Forms.Appearance.Button;
            cb_sizes.SelectedIndex = 0;
        }

        private void cb_fonts_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rtb_text.SelectedText == null)
            {
                Font ExistingFont = rtb_text.SelectionFont;
                float t = float.Parse(cb_sizes.SelectedItem.ToString());
                rtb_text.Font = new Font(cb_fonts.SelectedItem.ToString(), t,ExistingFont.Style);
            }
            else
            {
                Font ExistingFont = rtb_text.SelectionFont;
                float t = float.Parse(cb_sizes.SelectedItem.ToString());
                rtb_text.SelectionFont = new Font(cb_fonts.SelectedItem.ToString(), t,ExistingFont.Style);
            }
        }
        int trr=0;
        private void cb_sizes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rtb_text.SelectedText == null)
            {
                Font ExistingFont = rtb_text.SelectionFont;
                float t = float.Parse(cb_sizes.SelectedItem.ToString());
                rtb_text.Font = new Font(cb_fonts.SelectedItem.ToString(),t,ExistingFont.Style);

            }
            else
            {
                if (trr != 0)
                {
                    Font ExistingFont = rtb_text.SelectionFont;
                    float t = float.Parse(cb_sizes.SelectedItem.ToString());
                    rtb_text.SelectionFont = new Font(cb_fonts.SelectedItem.ToString(), t,ExistingFont.Style);
                }
                trr=1;
            }
        }
        
        private void cb_bold_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_bold.Checked)
            {
                if (rtb_text.SelectedText == null)
                {
                    Font ExistingFont = rtb_text.Font;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.Font = new Font(ExistingFont, NewFontStyle);
                }
                else
                {
                    Font ExistingFont = rtb_text.SelectionFont;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.SelectionFont = new Font(ExistingFont, NewFontStyle);
                }
            }
        }

        private void cb_und_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_und.Checked)
            {
                if (rtb_text.SelectedText == null)
                {
                    Font ExistingFont = rtb_text.Font;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.Font = new Font(ExistingFont, NewFontStyle);
                }
                else
                {
                    Font ExistingFont = rtb_text.SelectionFont;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.SelectionFont = new Font(ExistingFont, NewFontStyle);
                }
            }
        }

        private void cb_i_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_i.Checked)
            {
                if (rtb_text.SelectedText == null)
                {
                    Font ExistingFont = rtb_text.Font;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.Font = new Font(ExistingFont, NewFontStyle);
                }
                else
                {
                    Font ExistingFont = rtb_text.SelectionFont;
                    FontStyle NewFontStyle = new FontStyle();
                    if (cb_bold.Checked)
                        NewFontStyle = FontStyle.Bold;

                    if (cb_und.Checked)
                        NewFontStyle |= FontStyle.Underline;

                    if (cb_i.Checked)
                        NewFontStyle |= FontStyle.Italic;
                    rtb_text.SelectionFont = new Font(ExistingFont, NewFontStyle);
                }
            }
        }

        private void rb_lft_CheckedChanged(object sender, EventArgs e)
        {
            if(rtb_text.Text!=null)
            {
                rtb_text.SelectAll();
                rtb_text.SelectionAlignment = HorizontalAlignment.Left;
                rtb_text.DeselectAll();
            }
        }

        private void rb_cnt_CheckedChanged(object sender, EventArgs e)
        {
            if (rtb_text.Text != null)
            {
                rtb_text.SelectAll();
                rtb_text.SelectionAlignment = HorizontalAlignment.Center;
                rtb_text.DeselectAll();
            }
        }

        private void rb_r_CheckedChanged(object sender, EventArgs e)
        {
            if (rtb_text.Text != null)
            {
                rtb_text.SelectAll();
                rtb_text.SelectionAlignment = HorizontalAlignment.Right;
                rtb_text.DeselectAll();
            }
        }

        private void cb_color_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rtb_text.SelectedText == null)
            {
                foreach (Color color in new ColorConverter().GetStandardValues())
                {
                    if(color.Name == cb_color.SelectedItem.ToString())
                    {
                        rtb_text.ForeColor = color;
                    }
                }
            }
            else
            {
                foreach (Color color in new ColorConverter().GetStandardValues())
                {
                    if (color.Name == cb_color.SelectedItem.ToString())
                    {
                        rtb_text.SelectionColor = color;
                    }
                }
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            rtb_text.SaveFile(fb_savef.Text + ".docx");
        }

        private void btn_load_Click(object sender, EventArgs e)
        {
            rtb_text.LoadFile(tb_loadfil.Text + ".docx");
        }
    }
}
