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
    public partial class Frm_Login : Form
    {
        public Frm_Login()
        {
            InitializeComponent();
        }
        OleDbConnection baglantı = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=Login_Kayıtlar.accdb");
        OleDbCommand cmd;
        OleDbDataReader dr;
        private void Btn_Giris_Yap_Click(object sender, EventArgs e)
        {
            string ad = Txtbox_Kullanıcı_Ad.Text;
            string sifre = Txtbox_Sifre.Text;
            baglantı = new OleDbConnection("Provider=Microsoft.ACE.Oledb.12.0;Data Source=Login_Kayıtlar.accdb");
            cmd = new OleDbCommand();
            baglantı.Open();
            cmd.Connection = baglantı;
            cmd.CommandText = "SELECT * FROM Login_Kayıt where Kullanıcı_Adı='" + Txtbox_Kullanıcı_Ad.Text + "' AND Sifre='" + Txtbox_Sifre.Text + "'";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                Tez_Kontrol_Ana_Ekran TK = new Tez_Kontrol_Ana_Ekran();
                TK.Show();
            }
            else
            {
                MessageBox.Show("Kullanıcı Adınız Veyahut Şifreniz Yanlış");
            }
            baglantı.Close();
        }

        private void Btn_Üye_Ol_Click(object sender, EventArgs e)
        {
            Üye_Ol ÜO = new Üye_Ol();
            ÜO.Show();
            }
        }
    }


