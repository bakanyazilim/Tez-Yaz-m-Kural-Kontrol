using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Tez_Yazım_Kontrol_Giris
{
    public partial class Üye_Ol : Form
    {
        public Üye_Ol()
        {
            InitializeComponent();
        }
        OleDbConnection baglantı = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=Login_Kayıtlar.accdb");
        OleDbCommand cmd;
        OleDbDataReader dr;
        private void Btn_Kayıt_Click(object sender, EventArgs e)
        {
            OleDbConnection baglantı = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=Login_Kayıtlar.accdb");
            OleDbCommand cmd;
            OleDbDataReader dr;
            string isim = textBox_Kullanıcı_İsim.Text;
            string Sifreİ = textBox_İlk_Sifre.Text;
            string SifreS = textBox_Son_Sifre.Text;
            bool Kontrol = false;
            if (textBox_Kullanıcı_İsim.Text != "" && textBox_İlk_Sifre.Text != "" && textBox_Son_Sifre.Text != "")
            {
                Kontrol = true;
            }
            else
            {

                MessageBox.Show("Alanların Tamamı Doldurulmalıdır. Aksi Takdirde Üyelik İşlemi Gerçekleştirilemez!",
                    "Hata", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }
    }
}
