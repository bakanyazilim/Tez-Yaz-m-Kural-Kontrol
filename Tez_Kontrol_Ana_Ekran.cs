using System;
using System.Windows.Forms;

using System.IO;
using System.Drawing;
using System.Text;
using NetOffice.WordApi.Enums;
using Word = Microsoft.Office.Interop.Word;

namespace Tez_Yazım_Kontrol_Giris
{
    public partial class Tez_Kontrol_Ana_Ekran : Form
    {
         
        public Tez_Kontrol_Ana_Ekran()
        {
            InitializeComponent();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void Btn_Kontrol_Et_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dosya = new OpenFileDialog();
            Dosya.Filter = "Word Dosyası |*.docx";
            Dosya.RestoreDirectory = true;
            Dosya.CheckFileExists = false;
            Dosya.Title = "Word Dosyanızı Seçiniz...";

            if (Dosya.ShowDialog() == DialogResult.OK)
            {

                OpenFileDialog Document = new OpenFileDialog();
                string dosyayolu = Dosya.FileName;
                string dosya_adi = Dosya.SafeFileName;

                {

                    label1.Text = dosya_adi + " Dosyası Kontrol Ediliyor. Lütfen Bekleyiniz...";

                    richTextBox_Dosya.Clear();


                    Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();

                    object nullobject = System.Reflection.Missing.Value;



                    Microsoft.Office.Interop.Word.Document docs = wordObject.Documents.Open(dosyayolu);

                    docs.ActiveWindow.Selection.WholeStory();
                    docs.ActiveWindow.Selection.Copy();
                    IDataObject data = Clipboard.GetDataObject();


                    string satir = "";
                    int i = 1;
                    int j = 1;

                    var docum = new Document();
                    docum = docs;


                    foreach (Microsoft.Office.Interop.Word.Paragraph objParagraph in docs.Paragraphs)
                    {




                        Microsoft.Office.Interop.Word.Font s = docs.Paragraphs[j].Range.Font;


                        
                        if (docs.Paragraphs[j].Range.Text == "ÖNSÖZ")
                        {
                            richTextBox_Dosya.Text += "\n ÖNSÖZ MEVCUTTUR " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "İÇİNDEKİLER")
                        {
                            richTextBox_Dosya.Text += "\n İÇİNDEKİLER LİSTESİ MEVCUTTUR" + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "ÖZET")
                        {
                            richTextBox_Dosya.Text += "\n ÖZET METNİ MEVCUTTUR. " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "ABSTRACT")
                        {
                            richTextBox_Dosya.Text += "\n İNGİLİZCE ÖZET METNİ MEVCUTTUR. " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "ŞEKİLLER LİSTESİ")
                        {
                            richTextBox_Dosya.Text += "\n ŞEKİLLER LİSTESİ MEVCUTTUR. " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "TABLOLAR LİSTESİ")
                        {
                            richTextBox_Dosya.Text += "\n TABLOLAR LİSTESİ MEVCUTTUR. " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "EKLER LİSTESİ")
                        {
                            richTextBox_Dosya.Text += "\n EKLER LİSTESİ MEVCUTTUR. " + i;
                        }
                        else if (docs.Paragraphs[j].Range.ToString() == "SİMGELER VE KISALTMALAR")
                        {
                            richTextBox_Dosya.Text += "\n SİMGELER VE KISALTMLAR MEVCUTTUR." + i;
                        }

                        if (s.Size == 12F)
                        {


                            if (s.Position.ToString() != "wdVerticalAlignmentLeft")
                            {

                                richTextBox_Dosya.Text += "\n ara başlık sola yaslı değil satır :" + i;
                            }
                            if (s.ColorIndex.ToString() != "wdBlack")
                            {

                                richTextBox_Dosya.Text += "\n yazı rengi yanlış satır:" + i;
                            }
                            if (s.Name.ToString() != "Times New Roman")
                            {

                                richTextBox_Dosya.Text += "\n yazı stili yanlış satır:" + i;
                            }
                        }
                        else
                        if (s.Size == 16F)
                        {


                            if (s.Position.ToString() != "WdVerticalAlignmentCenter")
                            {

                                richTextBox_Dosya.Text += "\n ana başlık iki yana yaslı değil satır:  " + i;
                            }
                            if (s.ColorIndex.ToString() != "wdBlack")
                            {
 
                                richTextBox_Dosya.Text += "\n yazı rengi yanlış satır:" + i;
                            }
                            if (s.Name.ToString() != "Times New Roman")
                            {

                                richTextBox_Dosya.Text += "\n yazı stili yanlış satır:" + i;
                            }

                        }
                        else
                        if (s.Size == 11F)
                        {


                            if (s.Position.ToString() != "WdVerticalAlignmentCenter")
                            {
 
                                richTextBox_Dosya.Text += "\n ana başlık iki yana yaslı değil satır:  " + i;
                            }
                            if (s.ColorIndex.ToString() != "wdBlack")
                            {

                                richTextBox_Dosya.Text += "\n yazı rengi yanlış satır:" + i;
                            }
                            if (s.Name.ToString() != "Times New Roman")
                            {

                                richTextBox_Dosya.Text += "\n yazı stili yanlış satır:" + i;
                            }

                        }

                        else
                        {

                            richTextBox_Dosya.Text += "\n yazı boyutu 11 punto değil satır:" + i;

                            if (s.Name.ToString() != "Times New Roman")
                            {
                                richTextBox_Dosya.Text += "\n yazı stili yanlış satır:" + i;
                            }
                            if (s.ColorIndex.ToString() != "wdBlack")
                            {
                                richTextBox_Dosya.Text += "\n yazı rengi yanlış satır:" + i;
                            }
                        }


                        i++;
                        j++;
                    }

                    i = 1;
                    j=1;

                    if (docum.PageSetup.TopMargin.ToString() != "85,05")
                    {
                        richTextBox_Dosya.Text += "\n üst boşluk yanlış:";
                    }

                    if (docum.PageSetup.LeftMargin.ToString() != "92,15")
                    {
                        richTextBox_Dosya.Text += "\n sol boşluk yanlış:";
                    }

                    if (docum.PageSetup.RightMargin.ToString() != "70,9")
                    {
                        richTextBox_Dosya.Text += "\n sağ boşluk yanlış:";
                    }

                    if (docum.PageSetup.BottomMargin.ToString() != "70,9")
                    {
                        richTextBox_Dosya.Text += "\n alt boşluk yanlış:";
                    }

                    if (docum.Paragraphs.Alignment != WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        richTextBox_Dosya.Text += "\n iki yana yaslı değil";

                        label1.Text = dosya_adi + " Dosyası Kontrol Edildi. Ayrıntılar Aşağıdaki Kısımdadır.";
                    }

                    MessageBox.Show("Tarama İşlemi Tamamlandı", "BİLGİ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);


                git:
                    return;


                }

            }
        }

        private void Btn_Cikis_Click(object sender, EventArgs e)
        {
            this.Hide();
            Frm_Login Giris = new Frm_Login();
            Giris.ShowDialog();
        }

        private void Btn_Dosya_Sec_Click(object sender, EventArgs e)
        {
            OpenFileDialog Dosya = new OpenFileDialog();
            Dosya.Filter = "Word Dosyası |*.docx";
            Dosya.RestoreDirectory = true;
            Dosya.CheckFileExists = false;
            Dosya.Title = "Word Dosyanızı Seçiniz...";
        }
    }
}