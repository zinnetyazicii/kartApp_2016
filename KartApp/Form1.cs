using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace KartApp
{
    public partial class frmKartApp : Form
    {
        int kategori = 0;
        //OleDbConnection olCon = new OleDbConnection(Yardimci.Baglanti());
        public frmKartApp()
        {
            InitializeComponent();
            grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());

            cmbAktif.Items.Add("Pasif");
            cmbAktif.Items.Add("Aktif");
            cmbAktif.SelectedItem = "Aktif";
            cmbKategori.DataSource = Yardimci.Tablo(Yardimci.KategoriGetir());
            cmbKategori.DisplayMember = "KatAdi";
            cmbKategori.ValueMember = "KatID";




        }
        //------------Kaydet----------------------
        private void btnKaydet_Click(object sender, EventArgs e)
        {
            btnKayıt.Visible = true;

            int al = 0;
            if (btnKaydet.Text=="Kaydet")
            {
                try
                {
                    DataTable dt = new DataTable();
                    dt = Yardimci.Tablo(Yardimci.TumVeriGetir());
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (txtAd.Text.ToUpper() == dt.Rows[i][1].ToString().ToUpper() && txtSoyad.Text.ToUpper() == dt.Rows[i][2].ToString().ToUpper() && txtGsm.Text==dt.Rows[i][6].ToString())
                        {
                            lblSon.Text = "Bu Kayıt Mevcuttur!";
                            lblSon.ForeColor = Color.DarkRed;
                            al = 1;
                            break;
                        }
                    }
                    if (al==0)
                    {
                        Yardimci.Tablo(Yardimci.VeriKaydet(txtAd.Text, txtSoyad.Text, txtUnvan.Text, txtTarih.Text, txtTel.Text, txtGsm.Text, txtFax.Text, txtMail.Text, txtAdres.Text, txtSAdi.Text, txtWeb.Text,cmbKategori.SelectedValue.ToString()));
                        grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());
                        lblSon.Text = "KAYIT BAŞARILI";
                        lblSon.ForeColor = Color.Gold;
                    }
                }
                catch (Exception)
                {
                    lblSon.Text = "Boş Kayıt Girilemez!";
                    lblSon.ForeColor = Color.DarkRed;

                }

                
            }
            else if (btnKaydet.Text == "Güncelle")
            {
                Yardimci.Tablo(Yardimci.VeriGüncelle(txtAd.Text, txtSoyad.Text, txtUnvan.Text, txtTarih.Text, txtTel.Text,txtGsm.Text, txtFax.Text, txtMail.Text, txtAdres.Text, txtSAdi.Text, txtWeb.Text,cmbAktif.SelectedIndex.ToString(),cmbKategori.SelectedValue.ToString(),satir));
                grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());

                lblSon.Text = "GÜNCELLEME İŞLEMİ BAŞARILI..";
                lblSon.ForeColor = Color.Green;

                
            }
            else if (btnKaydet.Text=="Geri Al")
            {
                Yardimci.Tablo(Yardimci.VeriGeriAl(satir));
                grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());

                checkBox1.Checked = false;
                lblSon.Text = "Kayıt Geri Alınmıştır.";
                lblSon.ForeColor = Color.Green;
                
            }
        }
        //--------Yazdırma işlemi-------

            
        private void frmKartApp_Load(object sender, EventArgs e)
        {

            lblSon.Text = "";
            
        }
        //---------Silme İşlemi--------
        string satir;
        private void btnSil_Click(object sender, EventArgs e)
        {
            if (satir != null)
            {

                Yardimci.Tablo(Yardimci.Sil(satir));
                grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());

                lblSon.Text = "SİLME İŞLEMİ BAŞARILI..";
                lblSon.ForeColor = Color.DarkRed;



                /*
                OleDbDataAdapter da = new OleDbDataAdapter(sorgu, olCon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                grdData.DataSource = dt;
                lblSonuc.Text = "İşlem Başarı ile Gerçekleştirildi.";
                lblSonuc.ForeColor = Color.Green;*/
            }
            else
            {
                //lblSonuc.Text = "Lütfen Seçim Yapınız..";
                //lblSonuc.ForeColor = Color.Red;
                
            }
            
        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            btnSil.Visible = true;
            btnKaydet.Text = "Güncelle";
            btnKayıt.Visible = true;

            int rowIndex = e.RowIndex;

            if (rowIndex!=-1)
            {
                satir = grdData.Rows[rowIndex].Cells[0].Value.ToString();
                txtAd.Text = grdData.Rows[rowIndex].Cells[1].Value.ToString();
                txtSoyad.Text = grdData.Rows[rowIndex].Cells[2].Value.ToString();
                txtUnvan.Text = grdData.Rows[rowIndex].Cells[3].Value.ToString();
                txtTarih.Text = grdData.Rows[rowIndex].Cells[4].Value.ToString();
                txtTel.Text = grdData.Rows[rowIndex].Cells[5].Value.ToString();
                txtGsm.Text= grdData.Rows[rowIndex].Cells[6].Value.ToString();
                txtFax.Text = grdData.Rows[rowIndex].Cells[7].Value.ToString();
                txtMail.Text = grdData.Rows[rowIndex].Cells[8].Value.ToString();
                txtAdres.Text = grdData.Rows[rowIndex].Cells[9].Value.ToString();
                txtSAdi.Text = grdData.Rows[rowIndex].Cells[10].Value.ToString();
                txtWeb.Text = grdData.Rows[rowIndex].Cells[11].Value.ToString();
                cmbKategori.Text= grdData.Rows[rowIndex].Cells[12].Value.ToString();



                if (checkBox1.Checked==true)
                {
                    btnKaydet.Text = "Geri Al";
                    satir = grdData.Rows[rowIndex].Cells[0].Value.ToString();
                    txtAd.Text = grdData.Rows[rowIndex].Cells[1].Value.ToString();
                    txtSoyad.Text = grdData.Rows[rowIndex].Cells[2].Value.ToString();
                    txtUnvan.Text = grdData.Rows[rowIndex].Cells[3].Value.ToString();
                    txtTarih.Text = grdData.Rows[rowIndex].Cells[4].Value.ToString();
                    txtTel.Text = grdData.Rows[rowIndex].Cells[5].Value.ToString();
                    txtGsm.Text = grdData.Rows[rowIndex].Cells[6].Value.ToString();
                    txtFax.Text = grdData.Rows[rowIndex].Cells[7].Value.ToString();
                    txtMail.Text = grdData.Rows[rowIndex].Cells[8].Value.ToString();
                    txtAdres.Text = grdData.Rows[rowIndex].Cells[9].Value.ToString();
                    txtSAdi.Text = grdData.Rows[rowIndex].Cells[10].Value.ToString();
                    txtWeb.Text = grdData.Rows[rowIndex].Cells[11].Value.ToString();
                    cmbKategori.Text = grdData.Rows[rowIndex].Cells[12].Value.ToString();

                }
            }


        }

        //---------------------ARAMA--------------------
        private void txtAra_TextChanged(object sender, EventArgs e)
        {
            (grdData.DataSource as DataTable).DefaultView.RowFilter = string.Format("Ad LIKE '%" + txtAraAd.Text + "%'");
        }


        private void txtAraSoy_TextChanged(object sender, EventArgs e)
        {
            (grdData.DataSource as DataTable).DefaultView.RowFilter = string.Format("SOYAD LIKE '%" + txtAraSoy.Text + "%'");
        }

        private void txtAraSirket_TextChanged(object sender, EventArgs e)
        {
            (grdData.DataSource as DataTable).DefaultView.RowFilter = string.Format("SirketAd LIKE '%" + txtAraSirket.Text + "%'");
        }

        private void txtAraTel_TextChanged(object sender, EventArgs e)
        {
            (grdData.DataSource as DataTable).DefaultView.RowFilter = string.Format("Tel LIKE '%" + txtAraTel.Text + "%'");

        }

        private void txtKEkle_TextChanged(object sender, EventArgs e)
        {
            (grdKategori.DataSource as DataTable).DefaultView.RowFilter = string.Format("KatAdi LIKE '%" + txtKAra.Text + "%'");
        }


        private void btnKayıt_Click(object sender, EventArgs e)
        {
            txtAd.Clear();
            txtSoyad.Clear();
            txtUnvan.Clear();
            txtTarih.Clear();
            txtTel.Clear();
            txtFax.Clear();
            txtMail.Clear();
            txtAdres.Clear();
            txtSAdi.Clear();
            txtWeb.Clear();
            txtGsm.Clear();
            btnKaydet.Text = "Kaydet";
            btnKayıt.Visible = false;
            lblSonuc.Text = "";
            btnSil.Visible = false;

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)

            {
                grdData.DataSource = Yardimci.Tablo(Yardimci.SilineniGetir());

            }
            else
            {
                grdData.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());
            }
        }

        private void btnKatKayıt_Click(object sender, EventArgs e)
        {
            if (btnKatKayıt.Text == "Kaydet") 
            {
                try
                {
                    DataTable dt1 = new DataTable();
                    dt1 = Yardimci.Tablo(Yardimci.KategoriGetir());

                    Yardimci.Tablo(Yardimci.KategoriKaydet(txtKEkle.Text));
                    grdKategori.DataSource = Yardimci.Tablo(Yardimci.KategoriGetir());
                    lblKSonuc.Text = "KAYIT BAŞARILI!";

                }
                catch (Exception)
                {
                    lblKSonuc.Text = "LÜTFEN BOŞ BIRAKMAYINIZ..!";
                }
            }
        }

        private void BtnKatGün_Click(object sender, EventArgs e)
        {
            Yardimci.Tablo(Yardimci.KategoriGüncelle(txtKEkle.Text, satir));
            grdKategori.DataSource = Yardimci.Tablo(Yardimci.KategoriGetir());
            lblKSonuc.Text = "GÜNCELLEME BAŞARILI.";

        }
        private void tbcAna_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tbcAna.SelectedIndex==1)
            {
                grdKategori.DataSource = Yardimci.Tablo(Yardimci.KategoriGetir());
                btnKatSil.Visible = false;
            }
            if (tbcAna.SelectedIndex==2)
            {
                grdRapor.DataSource = Yardimci.Tablo(Yardimci.VeriGetir());
                chbAd.Checked = true;
                chbSoyad.Checked = true;
                chbUnvan.Checked = true;
                

                for (int i = 4; i < grdRapor.ColumnCount; i++)
                {
                    grdRapor.Columns[i].Visible = false;
                }
            }
        }


        //------------------RAPORLAMA---------------------

        private void chbAd_CheckedChanged(object sender, EventArgs e)
        {
            if (chbAd.Checked==true)

            {
                grdRapor.Columns[1].Visible = true;
            }
            else
            {
                grdRapor.Columns[1].Visible = false;
            }

        }

        private void chbSoyad_CheckedChanged(object sender, EventArgs e)
        {
            if (chbSoyad.Checked == true)

            {
                grdRapor.Columns[2].Visible = true;
            }
            else
            {
                grdRapor.Columns[2].Visible = false;
            }
        }

        private void chbUnvan_CheckedChanged(object sender, EventArgs e)
        {
            if (chbUnvan.Checked == true)

            {
                grdRapor.Columns[3].Visible = true;
            }
            else
            {
                grdRapor.Columns[3].Visible = false;
            }
        }

        private void chbTarih_CheckedChanged(object sender, EventArgs e)
        {
            if (chbTarih.Checked == true)

            {
                grdRapor.Columns[4].Visible = true;
            }
            else
            {
                grdRapor.Columns[4].Visible = false;
            }
        }

        private void chbTel_CheckedChanged(object sender, EventArgs e)
        {
            if (chbTel.Checked == true)

            {
                grdRapor.Columns[5].Visible = true;
            }
            else
            {
                grdRapor.Columns[5].Visible = false;
            }
        }

        private void chbGsm_CheckedChanged(object sender, EventArgs e)
        {
            if (chbGsm.Checked == true)

            {
                grdRapor.Columns[6].Visible = true;
            }
            else
            {
                grdRapor.Columns[6].Visible = false;
            }
        }

        private void chbFax_CheckedChanged(object sender, EventArgs e)
        {
            if (chbFax.Checked == true)

            {
                grdRapor.Columns[7].Visible = true;
            }
            else
            {
                grdRapor.Columns[7].Visible = false;
            }
        }

        private void chbMail_CheckedChanged(object sender, EventArgs e)
        {
            if (chbMail.Checked == true)

            {
                grdRapor.Columns[8].Visible = true;
            }
            else
            {
                grdRapor.Columns[8].Visible = false;
            }
        }

        private void chbAdres_CheckedChanged(object sender, EventArgs e)
        {
            if (chbAdres.Checked == true)

            {
                grdRapor.Columns[9].Visible = true;
            }
            else
            {
                grdRapor.Columns[9].Visible = false;
            }
        }

        private void chbSirketAdi_CheckedChanged(object sender, EventArgs e)
        {
            if (chbSirketAdi.Checked == true)

            {
                grdRapor.Columns[10].Visible = true;
            }
            else
            {
                grdRapor.Columns[10].Visible = false;
            }
        }

        private void chbWeb_CheckedChanged(object sender, EventArgs e)
        {
            if (chbWeb.Checked == true)

            {
                grdRapor.Columns[11].Visible = true;
            }
            else
            {
                grdRapor.Columns[11].Visible = false;
            }
        }

        private void chbAktif_CheckedChanged(object sender, EventArgs e)
        {
            if (chbKategori.Checked == true)

            {
                grdRapor.Columns[12].Visible = true;
            }
            else
            {
                grdRapor.Columns[12].Visible = false;
            }
        }


        private void grdKategori_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int rowIndex = e.RowIndex;
            if (rowIndex != -1)
            {
                satir = grdKategori.Rows[rowIndex].Cells[0].Value.ToString();
                txtKEkle.Text = grdKategori.Rows[rowIndex].Cells[1].Value.ToString();
                kategori = Convert.ToInt32(grdKategori.Rows[rowIndex].Cells[0].Value);
                btnKatSil.Visible = true;
            }
        }

        //---------------------YAZDIRMA------------------------------
      



        StringFormat strFormat;
        ArrayList arrColumnLefts = new ArrayList();
        ArrayList arrColumnWidhts = new ArrayList();
        ArrayList liste = new ArrayList();
        int iCellHeight = 0, iTotalWidht = 0, iRow = 0, iHeaderHeight = 0;
        bool bFirstPage = false, bNewPage = false;

        private void cmbKategori_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnKatSil_Click(object sender, EventArgs e)
        {
            DataTable dt=Yardimci.Tablo(Yardimci.SilKategoriGetir(kategori));
            if (dt.Rows.Count > 0)
            {
                MessageBox.Show("Kayıt var");
            }
            else
            {
                MessageBox.Show("Başarılı");
            }
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            try
            {

                int iLeftMargin = e.MarginBounds.Left, iTopMatgin = e.MarginBounds.Top, iTmpWidth = 0;
                bool bMorePagesToPrint = false;
                bFirstPage = true;

                if (bFirstPage)
                {
                    foreach (DataGridViewColumn GridCol in grdRapor.Columns)
                    {
                        iTmpWidth = (int)(Math.Floor((double)((double)GridCol.Width / (double)iTotalWidht * (double)iTotalWidht * ((double)e.MarginBounds.Width / (double)iTotalWidht))));

                        iHeaderHeight = (int)(e.Graphics.MeasureString(GridCol.HeaderText, GridCol.InheritedStyle.Font, iTmpWidth).Height) + 20;

                        arrColumnLefts.Add(iLeftMargin);
                        arrColumnWidhts.Add(iTmpWidth);
                        iLeftMargin += iTmpWidth;
                    }
                }

                while (iRow<=grdRapor.RowCount-1)
                {
                    DataGridViewRow gridRow = grdRapor.Rows[iRow];
                    iCellHeight = gridRow.Height + 5;
                    int iCount = 0;

                    if (iTopMatgin + iCellHeight >= e.MarginBounds.Height + e.MarginBounds.Top)
                    {
                        bNewPage = true;
                        bFirstPage = false;
                        bMorePagesToPrint = true;
                        break;
                    }
                    else
                    {
                        if (bNewPage)
                        {
                            e.Graphics.DrawString("DÖKUMANLAR", new Font(grdRapor.Font, FontStyle.Bold), Brushes.Black, e.MarginBounds.Left, e.MarginBounds.Top - e.Graphics.MeasureString("DÖKUMANLAR", new Font(grdRapor.Font, FontStyle.Bold), e.MarginBounds.Width).Height - 13);

                            String strDate = DateTime.Now.ToLongDateString() + " " + DateTime.Now.ToShortTimeString();

                            e.Graphics.DrawString(strDate, new Font(grdRapor.Font, FontStyle.Bold),
                                    Brushes.Black, e.MarginBounds.Left + (e.MarginBounds.Width -
                                    e.Graphics.MeasureString(strDate, new Font(grdRapor.Font,
                                    FontStyle.Bold), e.MarginBounds.Width).Width), e.MarginBounds.Top -
                                    e.Graphics.MeasureString("Çıktı Başlığı", new Font(new Font(grdRapor.Font,
                                    FontStyle.Bold), FontStyle.Bold), e.MarginBounds.Width).Height - 13);


                            iTopMatgin = e.MarginBounds.Top;
                            foreach (DataGridViewColumn GridCol in grdRapor.Columns)
                            {
                                e.Graphics.FillRectangle(new SolidBrush(Color.LightGray),
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMatgin,
                                    (int)arrColumnWidhts[iCount], iHeaderHeight));

                                e.Graphics.DrawRectangle(Pens.Black,
                                    new Rectangle((int)arrColumnLefts[iCount], iTopMatgin,
                                    (int)arrColumnWidhts[iCount], iHeaderHeight));

                                e.Graphics.DrawString(GridCol.HeaderText, GridCol.InheritedStyle.Font,
                                    new SolidBrush(GridCol.InheritedStyle.ForeColor),
                                    new RectangleF((int)arrColumnLefts[iCount], iTopMatgin,
                                    (int)arrColumnWidhts[iCount], iHeaderHeight), strFormat);
                                iCount++;
                            }
                            bNewPage = false;
                            iTopMatgin += iHeaderHeight;
                        }
                        iCount = 0;

                        foreach (DataGridViewCell Cel in gridRow.Cells)
                        {
                            if (Cel.Value != null)
                            {
                                e.Graphics.DrawString(Cel.Value.ToString(), Cel.InheritedStyle.Font,
                                            new SolidBrush(Cel.InheritedStyle.ForeColor),
                                            new RectangleF((int)arrColumnLefts[iCount], (float)iTopMatgin,
                                            (int)arrColumnWidhts[iCount], (float)iCellHeight), strFormat);
                            }

                            e.Graphics.DrawRectangle(Pens.Black, new Rectangle((int)arrColumnLefts[iCount],
                                    iTopMatgin, (int)arrColumnWidhts[iCount], iCellHeight));

                            iCount++;
                        }
                    }
                    iRow++;
                    iTopMatgin += iCellHeight;
                }


                if (bMorePagesToPrint)
                    e.HasMorePages = true;
                else
                    e.HasMorePages = false;
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.ToString());
                
            }
        }
        private void printDocument1_BeginPrint(object sender, System.Drawing.Printing.PrintEventArgs e)
        {
            try
            {
                strFormat = new StringFormat();
                strFormat.Alignment = StringAlignment.Near;
                strFormat.LineAlignment = StringAlignment.Center;
                strFormat.Trimming = StringTrimming.EllipsisCharacter;

                arrColumnLefts.Clear();
                arrColumnWidhts.Clear();
                iCellHeight = 0;
                iRow = 0;
                bFirstPage = true;
                bNewPage = true;

                iTotalWidht = 0;
                foreach (DataGridViewColumn dgvGridCol in grdRapor.Columns)
                {
                    iTotalWidht += dgvGridCol.Width;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }
        private void btnYazdır_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt = new DataTable();
                foreach (DataGridViewColumn col in grdRapor.Columns)
                {
                    if (grdRapor.Columns[col.Index].Visible != false)
                    {
                        dt.Columns.Add(col.Name);
                    }
                    else
                    {
                        liste.Add(col.Name);
                    }
                }

                for (int i = 0; i < liste.Count; i++)
                {
                    grdRapor.Columns.Remove(liste[i].ToString());
                }

                foreach (DataGridViewRow row in grdRapor.Rows)
                {
                    DataRow rw = dt.NewRow();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        rw[cell.ColumnIndex] = cell.Value;
                    }
                    dt.Rows.Add(rw);
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            if (MessageBox.Show("Kayıtları Görüntülemek İster misiniz?","Uyarı" ,MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                PrintPreviewDialog onizleme = new PrintPreviewDialog();
                onizleme.Document = printDocument1;
                onizleme.ShowDialog();
            }
            else
            {
                PrintDialog yazdir = new PrintDialog();
                yazdir.Document = printDocument1;
                yazdir.UseEXDialog = true;
                if (yazdir.ShowDialog()== DialogResult.OK)
                {
                    printDocument1.Print();
                }
                
            }
        }
    }
}
