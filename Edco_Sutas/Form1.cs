using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ADODB;

namespace Edco_Sutas
{
    public partial class Form1 : MetroFramework.Forms.MetroForm
    {
        String SQLquery;
        String SQLquery2;
        String SQLquery3;
        String SQLquery4;

        ADODB.Connection connprop;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Konstanta.NMFILEACCESS = Application.StartupPath + "\\adods.dll";
            for (int i = 1; i <= 893; i++)
            {
                String text = "txt" + (i);
                this.Controls.Find(text, true)[0].KeyDown += TextBox_KeyDown;
                this.Controls.Find(text, true)[0].KeyPress += TextBox_KeyPress;
                this.Controls.Find(text, true)[0].Enter += TextBox_Enter;
                this.Controls.Find(text, true)[0].Click += TextBox_Click;
                this.Controls.Find(text, true)[0].Leave += TextBox_Leave;
            }
            txtjam.Focus();
        }

        private void txtjam_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txtmenit.Focus();
            }
        }

        private void txtmenit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                txt1.Focus();
            }
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            Boolean cekkode;

            TextBox txtnya = sender as TextBox;
            Int32 txtsekarang = Convert.ToInt32(txtnya.Name.Substring(3, txtnya.Name.Length - 3));
            Int32 txtselanjutnya = txtsekarang + 1;
            String text = "txt" + (txtselanjutnya);
            String textsekarang = "txt" + (txtsekarang);

            if (e.KeyCode == Keys.Enter)
            {
                if (txtsekarang == 81)
                {
                    tab1.SelectedTab = tp2;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 188)
                {
                    tab1.SelectedTab = tp3;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 284)
                {
                    tab1.SelectedTab = tp4;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 382)
                {
                    tab1.SelectedTab = tp5;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 549)
                {
                    tab1.SelectedTab = tp6;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 658)
                {
                    tab1.SelectedTab = tp7;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 790)
                {
                    tab1.SelectedTab = tp8;
                    this.Controls.Find(text, true)[0].Focus();
                }
                else if (txtsekarang == 893)
                {
                    MessageBox.Show("Selesai");
                }
                else
                {
                    this.Controls.Find(text, true)[0].Focus();
                }
            }
        }

        private void TextBox_Enter(object sender, EventArgs e)
        {
            TextBox txtnya = sender as TextBox;
            String isitext = txtnya.Text;
            txtnya.SelectionStart = 0;
            txtnya.SelectionLength = isitext.Length;
            txtnya.Text = txtnya.Text.Replace(",", "");
        }

        private void TextBox_Click(object sender, EventArgs e)
        {
            TextBox txtnya = sender as TextBox;
            String isitext = txtnya.Text;
            txtnya.SelectionStart = 0;
            txtnya.SelectionLength = isitext.Length;
            txtnya.Text = txtnya.Text.Replace(",", "");
        }

        private void TextBox_Leave(object sender, EventArgs e)
        {
            TextBox txtnya = sender as TextBox;
            Int32 nmr = Convert.ToInt32(txtnya.Name.Substring(3, txtnya.Name.Length - 3));
            String text = "txt" + (nmr);
            if (nmr >= 129)
            {
                String isitxt = txtnya.Text;
                //txtnya.Text = formatribuan2(isitxt);
            }
        }

        private void TextBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            TextBox txtnya = sender as TextBox;
            Int32 nmr = Convert.ToInt32(txtnya.Name.Substring(3, txtnya.Name.Length - 3));
            String text = "txt" + (nmr.ToString());

            if (nmr == 5 || nmr == 6 || nmr == 7 || nmr == 8 || nmr == 10 || nmr == 13 || nmr == 15)
            {
                this.Controls.Find(text, true)[0].Text.ToUpper();
            }
            else
            {
                hanyaangka(e);
            }
        }

        void hanyaangka(KeyPressEventArgs f)
        {
            if (char.IsDigit(f.KeyChar) || (int)f.KeyChar == 8)
            {
                f.Handled = false;
            }

            else
            {
                f.Handled = true;
            }
        }

        void angkatertentu(KeyPressEventArgs f, Int32 awal, Int32 akhir)
        {
            f.Handled = true;
            for (int i = awal; i <= akhir; i++)
            {
                if (f.KeyChar == Convert.ToChar(Convert.ToString(i)) || (int)f.KeyChar == 8)
                {
                    f.Handled = false;
                }
            }
        }

        void formatribuan(TextBox txtY)
        {
            if (txtY.Text.Length > 0)
            {
                txtY.Text = string.Format("{0:#,##0}", double.Parse(txtY.Text));
            }
        }

        public String formatribuan2(String txtY)
        {
            String hasil = "";
            if (txtY.Length > 0)
            {
                hasil = string.Format("{0:#,##0}", double.Parse(txtY));
            }

            return hasil;
        }

        private void cmd1_Click(object sender, EventArgs e)
        {
            if (txt1.Text.Length > 0 && txt2.Text.Length > 0 && txt3.Text.Length > 0 && txt4.Text.Length > 0 && txt5.Text.Length > 0 && txt9.Text.Length > 0)
            {
                cek_lk_editing();
                MessageBox.Show("SELESAI");
            }
            else
            {
                MessageBox.Show("Identitas wilayah harus isi");
            }
        }

        void cek_lk_editing()
        {
            //cek_101();
            //cek_102();
            //cek_103();
            //cek_104();
            //cek_105();
            cek_306();

            cek_403();
        }

        void cek_101()
        {
            bool adasalah = false;

            for (int i = 1; i <= 9; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length == 0)
                {
                    adasalah = true;
                }
            }

            if (adasalah == true)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.01", 1);
            }
        }

        void cek_102()
        {
            bool adasalah = false;

            if (txtjam.Text.Length == 0 || txtmenit.Text.Length == 0)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.02", 2);
            }
        }

        void cek_103()
        {
            bool adasalah = false;

            for (int i = 16; i <= 20; i++)
            {
                String text = "txt" + (i);
                String text1 = "txt" + (i+5);
                String text2 = "txt" + (i+10);
                String text3 = "txt" + (i+15);

                Double jml = this.Controls.Find(text, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text, true)[0].Text) : 0;
                Double jml1 = this.Controls.Find(text1, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text1, true)[0].Text) : 0;
                Double jml2 = this.Controls.Find(text2, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text2, true)[0].Text) : 0;
                Double jml3 = this.Controls.Find(text3, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text3, true)[0].Text) : 0;

                if (jml != (jml1 + jml2 + jml3))
                {
                    adasalah = true;
                }

            }

            if (adasalah == true)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.03", 3);
            }
        }

        void cek_104()
        {
            bool adasalah = false;

            if ((txt19.Text.Length == 0 && txt39.Text.Length == 0) || (txt19.Text == "0" && txt39.Text == "0") || (Convert.ToDouble(txt19.Text)<8 && Convert.ToDouble(txt39.Text)<8))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.04", 4);
            }
        }

        void cek_105()
        {
            bool adasalah = false;

            for (int i = 20; i <= 50; i++)
            {
                if (i % 5 == 0)
                {
                    String text = "txt" + (i);
                    String text1 = "txt" + (i - 4);
                    String text2 = "txt" + (i - 3);
                    String text3 = "txt" + (i - 2);
                    String text4 = "txt" + (i - 1);

                    Double jml = this.Controls.Find(text, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text, true)[0].Text) : 0;
                    Double jml1 = this.Controls.Find(text1, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text1, true)[0].Text) : 0;
                    Double jml2 = this.Controls.Find(text2, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text2, true)[0].Text) : 0;
                    Double jml3 = this.Controls.Find(text3, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text3, true)[0].Text) : 0;
                    Double jml4 = this.Controls.Find(text4, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text4, true)[0].Text) : 0;

                    if (jml != (jml1 + jml2 + jml3 + jml4))
                    {
                        adasalah = true;
                    }
                }
            }

            if (adasalah == true)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.05", 5);
            }
        }

        void cek_106()
        {
            bool adasalah = false;

            for (int i = 46; i <= 50; i++)
            {
                String text = "txt" + (i);
                String text1 = "txt" + (i - 30);
                String text2 = "txt" + (i - 10);
                String text3 = "txt" + (i - 5);
                
                Double jml = this.Controls.Find(text, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text, true)[0].Text) : 0;
                Double jml1 = this.Controls.Find(text1, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text1, true)[0].Text) : 0;
                Double jml2 = this.Controls.Find(text2, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text2, true)[0].Text) : 0;
                Double jml3 = this.Controls.Find(text3, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text3, true)[0].Text) : 0;
                
                if (jml != (jml1 + jml2 + jml3))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.06", 6);
                }
            }

            if ((txt46.Text.Length == 0 && txt47.Text.Length == 0 && txt48.Text.Length == 0) || (txt46.Text == "0" && txt47.Text == "0" && txt48.Text == "0"))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.06", 7);
            }

            if ((txt49.Text.Length == 0 && txt50.Text.Length == 0) || (txt49.Text == "0" && txt50.Text == "0") || (Convert.ToDouble(txt49.Text) < 8 && Convert.ToDouble(txt50.Text) < 8))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.06", 8);
            }

        }

        void cek_107()
        {
            bool adasalah = false;

            for (int i = 79; i <= 81; i++)
            {
                String text = "txt" + (i);
                String text1 = "txt" + (i - 33);
                
                Double jml = this.Controls.Find(text, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text, true)[0].Text) : 0;
                Double jml1 = this.Controls.Find(text1, true)[0].Text.Length > 0 ? Convert.ToDouble(this.Controls.Find(text1, true)[0].Text) : 0;
                
                if (jml != jml1)
                {
                    adasalah = true;
                }
            }

            if (adasalah == true)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "1.07", 9);
            }
        }

        void cek_301()
        {
            bool adasalah = false;

            if (txt131.Text.Length > 0)
            {
                SQLquery = "select*from m_kom where KodeKom='" + txt131.Text + "' and FlagSubsektor='32'";
                if (new KoneksiAccess().cekDataSQL(SQLquery) == false)
                {
                    adasalah = true;
                }
            }

            if (txt148.Text.Length > 0)
            {
                SQLquery2 = "select*from m_kom where KodeKom='" + txt148.Text + "' and FlagSubsektor='32'";
                if (new KoneksiAccess().cekDataSQL(SQLquery2) == false)
                {
                    adasalah = true;
                }
            }

            if (txt165.Text.Length > 0)
            {
                SQLquery3 = "select*from m_kom where KodeKom='" + txt165.Text + "' and FlagSubsektor='32'";
                if (new KoneksiAccess().cekDataSQL(SQLquery3) == false)
                {
                    adasalah = true;
                }
            }

            if (adasalah == true)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.01", 10);
            }
        }

        void cek_302()
        {
            bool adasalah = false;
            bool adaisian = false;

            for (int i = 82; i <= 90; i++)
            {
                String text = "txt" + (i);
                if (i == 83 || i == 86 || i == 89)
                {
                    String textpanen = "txt" + (i);
                    String textproduksi = "txt" + (i + 1);
                    if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                    }
                }
            }

            for (int i = 98; i <= 106; i++)
            {
                String text = "txt" + (i);
                if (i == 99 || i == 102 || i == 105)
                {
                    String textpanen = "txt" + (i);
                    String textproduksi = "txt" + (i + 1);
                    if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                    }
                }
            }

            for (int i = 114; i <= 122; i++)
            {
                String text = "txt" + (i);
                if (i == 115 || i == 118 || i == 121)
                {
                    String textpanen = "txt" + (i);
                    String textproduksi = "txt" + (i + 1);
                    if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                    }
                }
            }


            if (txt131.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 132; i <= 140; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (i == 133 || i == 136 || i == 139)
                    {
                        String textpanen = "txt" + (i);
                        String textproduksi = "txt" + (i+1);
                        if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                        }
                    }
                }

                if (adaisian == false)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 11);
                }
            }

            if (txt148.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 149; i <= 157; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (i == 150 || i == 153 || i == 156)
                    {
                        String textpanen = "txt" + (i);
                        String textproduksi = "txt" + (i + 1);
                        if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                        }
                    }
                }

                if (adaisian == false)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 11);
                }
            }

            if (txt165.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 166; i <= 174; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (i == 167 || i == 170 || i == 173)
                    {
                        String textpanen = "txt" + (i);
                        String textproduksi = "txt" + (i + 1);
                        if ((this.Controls.Find(textpanen, true)[0].Text.Length > 0 && this.Controls.Find(textproduksi, true)[0].Text.Length == 0) || (this.Controls.Find(textpanen, true)[0].Text.Length == 0 && this.Controls.Find(textproduksi, true)[0].Text.Length > 0) || (this.Controls.Find(textpanen, true)[0].Text != "0" && this.Controls.Find(textproduksi, true)[0].Text == "0") || (this.Controls.Find(textpanen, true)[0].Text == "0" && this.Controls.Find(textproduksi, true)[0].Text != "0"))
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 12);
                        }
                    }
                }

                if (adaisian == false)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.02", 11);
                }
            }
        }

        void cek_303()
        {
            bool adasalah = false;
            bool adaisian = false;

            for (int i = 82; i <= 90; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    if(txt91.Text=="2" || txt91.Text=="3")
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 14);
                    }

                    if (txt93.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                    }
                    else if (txt93.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt93.Text) < 1 && Convert.ToInt32(txt93.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                        }
                    }

                    if (txt94.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                    }
                    else if (txt94.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt94.Text) < 1 && Convert.ToInt32(txt94.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                        }
                    }
                }
            }

            for (int i = 98; i <= 106; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    if (txt107.Text == "2" || txt107.Text == "3")
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 14);
                    }

                    if (txt109.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                    }
                    else if (txt109.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt109.Text) < 1 && Convert.ToInt32(txt109.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                        }
                    }

                    if (txt110.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                    }
                    else if (txt110.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt110.Text) < 1 && Convert.ToInt32(txt110.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                        }
                    }
                }
            }

            for (int i = 114; i <= 122; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    if (txt123.Text == "2" || txt123.Text == "3")
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 14);
                    }

                    if (txt125.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                    }
                    else if (txt125.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt125.Text) < 1 && Convert.ToInt32(txt125.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 15);
                        }
                    }

                    if (txt126.Text.Length == 0)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                    }
                    else if (txt126.Text.Length > 0)
                    {
                        if (Convert.ToInt32(txt126.Text) < 1 && Convert.ToInt32(txt126.Text) > 8)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 16);
                        }
                    }
                }
            }

            if (txt131.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 132; i <= 140; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (adaisian == true && (txt141.Text.Length == 0 || txt143.Text.Length == 0 || txt144.Text.Length == 0))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 13);
                    }
                }              
            }

            if (txt148.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 149; i <= 157; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (adaisian == true && (txt158.Text.Length == 0 || txt160.Text.Length == 0 || txt161.Text.Length == 0))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 13);
                    }
                }
            }

            if (txt165.Text.Length > 0)
            {
                adaisian = false;
                for (int i = 166; i <= 174; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = true;
                    }

                    if (adaisian == true && (txt175.Text.Length == 0 || txt177.Text.Length == 0 || txt178.Text.Length == 0))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.03", 13);
                    }
                }
            }
        }

        void cek_304()
        {
            bool adasalah = false;
            bool adaisian = false;

            for (int i = 82; i <= 181; i++)
            {
                if (i == 91 || i == 107 || i == 123 || i == 141 || i == 158 || i == 175)
                {
                    String text = "txt" + (i);
                    String text2 = "txt" + (i+1);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        if ((this.Controls.Find(text, true)[0].Text != "6" && this.Controls.Find(text, true)[0].Text != "7") && this.Controls.Find(text2, true)[0].Text.Length == 0)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.04", 17);
                        }
                        else if ((this.Controls.Find(text, true)[0].Text == "6" || this.Controls.Find(text, true)[0].Text == "7") && this.Controls.Find(text2, true)[0].Text.Length > 0)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.04", 18);
                        }
                    }
                }
            }
        }

        void cek_305()
        {
            bool adasalah = false;
            bool adaisian = false;

            for (int i = 82; i <= 181; i++)
            {
                if (i == 94 || i == 110 || i == 126 || i == 144 || i == 161 || i == 178)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text != "1")
                    {
                        adaisian = true;
                        for (int k = i + 1; k <= i + 3; k++)
                        {
                            String text2 = "txt" + (k);
                            if (this.Controls.Find(text2, true)[0].Text.Length == 0)
                            {
                                adaisian = false;
                            }

                            if (k == i + 1)
                            {
                                String textkab = "txt" + (k);
                                if (this.Controls.Find(textkab, true)[0].Text == String.Concat(txt1.Text, txt2.Text))
                                {
                                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.05", 20);
                                }
                            }

                            if (k == i + 3)
                            {
                                String textkablain = "txt" + (k);
                                String textkabutama = "txt" + (k-1);
                                if (this.Controls.Find(textkablain, true)[0].Text.Length > 0 && this.Controls.Find(textkabutama, true)[0].Text.Length == 0)
                                {
                                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.05", 21);
                                }
                            }

                        }
                        if (adaisian == false)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.05", 19);
                        }

                        if (this.Controls.Find(text, true)[0].Text == "3" || this.Controls.Find(text, true)[0].Text == "5" || this.Controls.Find(text, true)[0].Text == "7")
                        {
                            String textkabutama = "txt" + (i+2);
                            String textkablain = "txt" + (i+3);

                            if (Convert.ToInt32(this.Controls.Find(textkabutama, true)[0].Text) + Convert.ToInt32(this.Controls.Find(textkablain, true)[0].Text) == 100)
                            {
                                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.05", 22);
                            }
                        }

                        if (this.Controls.Find(text, true)[0].Text == "2" || this.Controls.Find(text, true)[0].Text == "4" || this.Controls.Find(text, true)[0].Text == "6")
                        {
                            String textkabutama = "txt" + (i + 2);
                            String textkablain = "txt" + (i + 3);

                            if (Convert.ToInt32(this.Controls.Find(textkabutama, true)[0].Text) + Convert.ToInt32(this.Controls.Find(textkablain, true)[0].Text) != 100)
                            {
                                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.05", 23);
                            }
                        }
                    }
                }
            }
        }

        void cek_306()
        {
            bool adasalah = false;
            bool adaisian = false;

            List<string> padi = new List<string>();

            for (int i = 82; i <= 97; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    padi.Add("1103");
                }
            }

            for (int i = 98; i <= 113; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    padi.Add("1104");
                }
            }

            for (int i = 114; i <= 129; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    padi.Add("1102");
                }
            }

            if (padi.Count > 0)
            {
                if (txt130.Text.Length == 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.06", 25);
                }
                else
                {
                    foreach (string item in padi)
                    {
                        if (item.Contains(txt130.Text))
                        {
                            adaisian = true;
                        }

                    }

                    if (adaisian == false)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.06", 25);
                    }
                }
            }
            else if (padi.Count == 0)
            {
                if (txt130.Text.Length > 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.06", 25);
                }
            }
        }

        void cek_307()
        {
            bool adasalah = false;
            bool adaisian = false;

            List<string> palawija = new List<string>();

            for (int i = 131; i <= 147; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    palawija.Add(txt131.Text);
                }
            }

            for (int i = 148; i <= 164; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    palawija.Add(txt148.Text);
                }
            }

            for (int i = 165; i <= 181; i++)
            {
                String text = "txt" + (i);
                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    palawija.Add(txt165.Text);
                }
            }

            if (palawija.Count > 0)
            {
                if (txt182.Text.Length == 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.07", 27);
                }
                else
                {
                    foreach (string item in palawija)
                    {
                        if (item.Contains(txt182.Text))
                        {
                            adaisian = true;
                        }

                    }

                    if (adaisian == false)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.07", 27);
                    }
                }
            }
            else if (palawija.Count == 0)
            {
                if (txt182.Text.Length > 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.07", 27);
                }
            }
        }

        void cek_308()
        {
            bool adasalah = false;
            bool adaisian = false;

            if (txt130.Text.Length > 0 || txt182.Text.Length > 0)
            {
                if (txt183.Text.Length == 0 && txt184.Text.Length == 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.08", 28);
                }

                if (txt183.Text == "0" && txt184.Text == "0")
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.08", 28);
                }
            }
        }

        void cek_309()
        {
            bool adasalah = false;
            bool adaisian = false;

            if (txt130.Text.Length > 0 || txt182.Text.Length > 0)
            {
                if (txt185.Text.Length == 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.09", 29);
                }
            }
        }

        void cek_310()
        {
            bool adasalah = false;
            bool adaisianpadihibrida = false;
            bool adaisianjagunghibrida = false;

            if (txt130.Text.Length > 0 || txt182.Text.Length > 0)
            {
                if (txt186.Text.Length == 0)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.10", 30);
                }

                for (int i = 82; i <= 97; i++)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisianpadihibrida = true;
                    }
                }

                if (adaisianpadihibrida == true && txt186.Text == "2")
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.10", 30);
                }

                if ((txt131.Text == "1213" || txt148.Text == "1213" || txt165.Text == "1213") && txt186.Text == "2")
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.10", 30);
                }

            }
        }

        void cek_311()
        {
            bool adasalah = false;
            bool adaisianpadihibrida = false;
            bool adaisianjagunghibrida = false;

            if (txt130.Text.Length > 0 && txt187.Text.Length == 0)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.11", 31);
            }

            if (txt182.Text.Length > 0 && txt188.Text.Length == 0)
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "3.11", 31);
            }
        }

        void cek_401()
        {
            bool adasalah = false;

            for (int i = 189; i <= 224; i++)
            {
                if (i == 189 || i == 198 || i == 207 || i == 216)
                {
                    String text = "txt" + (i);
                    SQLquery = "select*from m_kom where KodeKom='" + this.Controls.Find(text, true)[0].Text + "' and FlagSubsektor='41'";
                    if (new KoneksiAccess().cekDataSQL(SQLquery) == false)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.01", 32);
                    }
                }
                
            }
        }

        void cek_402()
        {
            bool adasalah = false;
            bool adaisian = false;

            for (int i = 189; i <= 224; i++)
            {
                if (i == 189 || i == 198 || i == 207 || i == 216)
                {
                    String text = "txt" + (i);
                    if (this.Controls.Find(text, true)[0].Text.Length > 0)
                    {
                        adaisian = false;
                        for (int k = i+1; k <= i+3; k++)
                        {
                            String text2 = "txt" + (k);
                            if (this.Controls.Find(text2, true)[0].Text.Length > 0)
                            {
                                adaisian = true;
                            }
                        }

                        if (adaisian == false)
                        {
                            tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.02", 33);
                        }
                    }
                }

            }
        }

        //no urut 34
        void cek_403()
        {

            for (int i = 189; i <= 224; i += 9)
            {
                if (!jika_terisi_maka_harus_terisi(i, i + 8))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.03", 34);
                }
            }
        }

        //no urut 35
        void cek_404()
        {

            for (int i = 197; i <= 224; i += 9)
            {
                if (!harus_terisi_berikut(i, new[] { "1", "2", "3", "4", "5", "6", "7" }))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.04", 35);
                }
            }
        }

        //no urut 36 & 37
        void cek_405()
        {
            for (int i = 225; i <= 278; i += 18)
            {
                if (i == 225 || i == 243 || i == 261)
                {
                    String text = "txt" + (i);
                    SQLquery = "select*from m_kom where KodeKom='" + this.Controls.Find(text, true)[0].Text + "' and FlagSubsektor='42'";
                    if (new KoneksiAccess().cekDataSQL(SQLquery) == false)
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.05", 36);
                    }
                }

            }
        }

        //no urut 38
        void cek_406()
        {
        }

        //no urut 39, 40 & 41
        void cek_407()
        {
            for (int i = 225; i <= 278; i += 18)
            {
                int jarak_kolom = 4;

                for (int k = 0; k < 4; ++k)
                {
                    if (!jika_terisi_maka_harus_terisi((i + 2 + (jarak_kolom * k)), (i + 2 + (jarak_kolom * k))))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.07", 39);
                    }
                }
            }
        }

        //no urut 42 & 43
        void cek_408()
        {
            for (int i = 225; i <= 278; i += 18)
            {
                if (!jika_terisi_maka_harus_terisi(i, i + 17))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.08", 42);
                }

                if (!harus_terisi_berikut((i + 17), new[] { "1", "2", "3", "4", "5", "6", "7" }))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.08", 43);
                }
            }
        }

        //no urut 44, 45, 46
        void cek_409()
        {
            if (!pengecekan_nilai_produksi_utama(279, new[]{189, 198, 207, 216, 225, 243, 261}))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.09", 44);
            }
        }

        //no urut 47
        void cek_410()
        {
            if (!pengecekan_tenaga_kerja(280, 281, new[] { 189, 198, 207, 216, 225, 243, 261 }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.10", 47);
            }
        }

        //no urut 48, 49
        void cek_411()
        {
            if (!harus_terisi_berikut(282, new[] { "1", "2" }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.11", 48);
            }

            if (!jika_salah_satu_terisi_maka_harus_terisi(282, new[] { 189, 198, 207, 216, 225, 243, 261 }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.11", 49);
            }
        }
        
        
        //no urut 50, 51
        void cek_412()
        {
            if (!harus_terisi_berikut(283, new[] { "1", "2" }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.12", 50);
            }

            if (!jika_salah_satu_terisi_maka_harus_terisi(283, new[] { 189, 198, 207, 216, 225, 243, 261 }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.12", 51);
            }
        }


        //no urut 52, 53
        void cek_413()
        {
            if (!harus_terisi_berikut(284, new[] { "1", "2" }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.13", 52);
            }

            if (!jika_salah_satu_terisi_maka_harus_terisi(284, new[] { 189, 198, 207, 216, 225, 243, 261 }))
            {
                tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "4.13", 53);
            }
        }

        //54
        void cek_501()
        {
            for (int i = 303; i <= 323; i += 10)
            {
                String text = "txt" + (i);
                SQLquery = "select*from m_kom where KodeKom='" + this.Controls.Find(text, true)[0].Text + "' and FlagSubsektor='51'";
                if (new KoneksiAccess().cekDataSQL(SQLquery) == false)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.01", 54);
                }
            }
        }

        //55, 56
        void cek_502()
        {
            for (int i = 288; i <= 330; i++)
            {
                if (i == 288 || i == 297 || i == 307 || i == 317 || i == 327)
                {
                    if(jika_salah_satu_terisi_maka_harus_terisi(i - 2, new[] {i, i+1, i+2, i+3})){
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.02", 55);
                    }
                }
                
            }
        }

        //57, 58
        void cek_503()
        {
            for (int i = 285; i <= 324; i++)
            {
                if (i == 285 || i == 294 || i == 304 || i == 314 || i == 324)
                {
                    if (jika_salah_satu_terisi_maka_harus_terisi(i + 7, new[] { i, i + 1, i + 2}))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.03", 57);
                    }

                    if (jika_salah_satu_terisi_maka_harus_terisi(i + 8, new[] { i, i + 1, i + 2 }))
                    {
                        tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.03", 57);
                    }
                }

                if (harus_terisi_berikut(i + 7, new[] { "1", "2", "3", "4", "5", "6", "7" }))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.03", 58);
                }


                if (harus_terisi_berikut(i + 8, new[] { "1", "2", "3", "4", "5", "6", "7" }))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.03", 58);
                }
            }
        }

        //59
        void cek_504()
        {
            for (int i = 347; i <= 362; i += 15)
            {
                String text = "txt" + (i);
                SQLquery = "select*from m_kom where KodeKom='" + this.Controls.Find(text, true)[0].Text + "' and FlagSubsektor='52'";
                if (new KoneksiAccess().cekDataSQL(SQLquery) == false)
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.04", 59);
                }
            }
        }

        //60
        void cek_505()
        {
            for (int i = 334; i <= 364; i += 15)
            {
                if (jika_salah_satu_terisi_maka_harus_terisi(i - 2, new[] { i, i + 1, i + 2, i + 3 }))
                {
                    tambah_hasil_edit(txt1.Text, txt2.Text, txt3.Text, txt4.Text, txt5.Text, txt9.Text, "5.02", 55);
                }
            }
        }


        //digunakan untuk mengecek kolom komoditas yang memiliki nilai produksi paling besar selama setahun
        //pengecekan dilakukan antara lain
        //harus terisi jika ada isian minimal satu komoditas pada kategori tersebut
        //isian harus terdapat pada daftar komoditas yang dikelola
        bool pengecekan_nilai_produksi_utama(int text_komoditas, int[] text_list_komoditas)
        {
            bool hasil = true;

            bool is_ada_daftar = false;
            List<String> daftar_komoditas = new List<string>();

            foreach (int komoditas in text_list_komoditas)
            {
                String text = "txt" + (komoditas);

                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    daftar_komoditas.Add(this.Controls.Find(text, true)[0].Text);
                    is_ada_daftar = true;
                }
            }

            if (is_ada_daftar && this.Controls.Find("txt" + (text_komoditas), true)[0].Text.Length == 0)
            {
                return false;
            }

            if (daftar_komoditas.Contains(this.Controls.Find("txt" + (text_komoditas), true)[0].Text))
            {
                return false;
            }

            return hasil;
        }

        //mengecek konsistensi jumlah tenaga kerja dan komoditas
        //jika komoditas terisi maka salah satu isian tenaga kerja laki/perempuan harus terisi
        bool pengecekan_tenaga_kerja(int text_laki, int text_perempuan, int[] text_list_komoditas)
        {
            bool hasil = true;

            bool is_ada_daftar = false;
            List<String> daftar_komoditas = new List<string>();

            foreach (int komoditas in text_list_komoditas)
            {
                String text = "txt" + (komoditas);

                if (this.Controls.Find(text, true)[0].Text.Length > 0)
                {
                    is_ada_daftar = true;
                    break;
                }
            }

            if (is_ada_daftar)
            {
                if(this.Controls.Find("txt" + (text_laki), true)[0].Text.Length == 0 && this.Controls.Find("txt" + (text_perempuan), true)[0].Text.Length == 0)
                    return false;
            }

            return hasil;
        }

        //mengecek nilai suatu text harus berada dalam range parameter "option"
        //saat memanggil fungsi ini, pastikan nilai pada text yang diisi tidak kosong
        bool harus_terisi_berikut(int text, string[] option)
        {
            bool hasil = true;

            String text_str = "txt" + (text);

            if(!option.Contains(text_str))
            {
                return false;
            }

            int nilai = Int32.Parse(this.Controls.Find(text_str, true)[0].Text;

            return hasil;
        }

        //mengecek jika nilai text asal berisi, maka text tujuan wajib terisi. Begitu juga sebaliknya
        bool jika_terisi_maka_harus_terisi(int asal, int tujuan)
        {
            bool hasil = true;
            String text_asal = "txt" + (asal);
            String text_tujuan = "txt" + (tujuan);

            if (this.Controls.Find(text_asal, true)[0].Text.Length > 0)
            {
                if (this.Controls.Find(text_tujuan, true)[0].Text.Length == 0)
                {
                    return false;
                }
            }
            else
            {
                if (this.Controls.Find(text_tujuan, true)[0].Text.Length > 0)
                {
                    return false;
                }
            }


            if (this.Controls.Find(text_tujuan, true)[0].Text.Length > 0)
            {
                if (this.Controls.Find(text_asal, true)[0].Text.Length == 0)
                {
                    return false;
                }
            }
            else
            {
                if (this.Controls.Find(text_asal, true)[0].Text.Length > 0)
                {
                    return false;
                }
            }

            return hasil;
        }

        bool jika_salah_satu_terisi_maka_harus_terisi(int text, int[] text_list_komoditas)
        {
            bool hasil = true;

            bool is_ada_daftar = false;
            List<String> daftar_komoditas = new List<string>();

            foreach (int komoditas in text_list_komoditas)
            {
                String text_kom = "txt" + (komoditas);

                if (this.Controls.Find(text_kom, true)[0].Text.Length > 0)
                {
                    is_ada_daftar = true;
                    break;
                }
            }

            if (is_ada_daftar && this.Controls.Find("txt" + (text), true)[0].Text.Length == 0)
            {
                return false;
            }

            return hasil;
        }


        void tambah_hasil_edit(String kdprov, String kdkab, String kdkec, String kddesa, String kdnbs, String nurtp, String nmitem, Int32 nmurut)
        {
            SQLquery2 = "select*from hasil_edit where kode_prov='"+ kdprov +"' and kode_kab='"+ kdkab +"' and kode_kec='"+ kdkec +"' and kode_desa='"+ kddesa +"' and nbs='"+ kdnbs +"' and nu="+ Convert.ToInt32(nurtp) +" and nmritem='"+ nmitem +"' and nmrurut="+ nmurut +"";
            if (new KoneksiAccess().cekDataSQL(SQLquery2) == false)
            {
                String konf_prop = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + Konstanta.NMFILEACCESS + ";Jet OLEDB:Database Password=edcosutas18";
                connprop = new ADODB.Connection();
                connprop.Open(konf_prop, "", "", -1);
                ADODB.Recordset rsprop = new ADODB.Recordset();
                rsprop.Open(SQLquery2, connprop, ADODB.CursorTypeEnum.adOpenStatic, ADODB.LockTypeEnum.adLockBatchOptimistic, 0);
                rsprop.AddNew();
                rsprop.Fields["kode_prov"].Value = kdprov;
                rsprop.Fields["kode_kab"].Value = kdkab;
                rsprop.Fields["kode_kec"].Value = kdkec;
                rsprop.Fields["kode_desa"].Value = kddesa;
                rsprop.Fields["nbs"].Value = kdnbs;
                rsprop.Fields["nu"].Value = Convert.ToInt32(nurtp);
                rsprop.Fields["nmritem"].Value = nmitem;
                rsprop.Fields["nmrurut"].Value = nmurut;
                rsprop.UpdateBatch();
                rsprop.Close();
                connprop.Close();
            }
        }
        
    }
}
