
namespace Tez_Yazım_Kontrol_Giris
{
    partial class Frm_Login
    {
        /// <summary>
        ///Gerekli tasarımcı değişkeni.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///Kullanılan tüm kaynakları temizleyin.
        /// </summary>
        ///<param name="disposing">yönetilen kaynaklar dispose edilmeliyse doğru; aksi halde yanlış.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer üretilen kod

        /// <summary>
        /// Tasarımcı desteği için gerekli metot - bu metodun 
        ///içeriğini kod düzenleyici ile değiştirmeyin.
        /// </summary>
        private void InitializeComponent()
        {
            this.Lbl_Üst_Bilgi = new System.Windows.Forms.Label();
            this.Lbl_Kullanıcı_Ad = new System.Windows.Forms.Label();
            this.Lbl_Sifre = new System.Windows.Forms.Label();
            this.Txtbox_Kullanıcı_Ad = new System.Windows.Forms.TextBox();
            this.Txtbox_Sifre = new System.Windows.Forms.TextBox();
            this.Btn_Giris_Yap = new System.Windows.Forms.Button();
            this.Btn_Üye_Ol = new System.Windows.Forms.Button();
            this.Btn_Unuttum = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Lbl_Üst_Bilgi
            // 
            this.Lbl_Üst_Bilgi.AutoSize = true;
            this.Lbl_Üst_Bilgi.Font = new System.Drawing.Font("Lucida Handwriting", 22.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Lbl_Üst_Bilgi.ForeColor = System.Drawing.Color.Red;
            this.Lbl_Üst_Bilgi.Location = new System.Drawing.Point(127, 22);
            this.Lbl_Üst_Bilgi.Name = "Lbl_Üst_Bilgi";
            this.Lbl_Üst_Bilgi.Size = new System.Drawing.Size(445, 49);
            this.Lbl_Üst_Bilgi.TabIndex = 0;
            this.Lbl_Üst_Bilgi.Text = "TEZ KONTROL GİRİŞ";
            // 
            // Lbl_Kullanıcı_Ad
            // 
            this.Lbl_Kullanıcı_Ad.AutoSize = true;
            this.Lbl_Kullanıcı_Ad.BackColor = System.Drawing.Color.Yellow;
            this.Lbl_Kullanıcı_Ad.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Lbl_Kullanıcı_Ad.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Lbl_Kullanıcı_Ad.Location = new System.Drawing.Point(12, 115);
            this.Lbl_Kullanıcı_Ad.Name = "Lbl_Kullanıcı_Ad";
            this.Lbl_Kullanıcı_Ad.Size = new System.Drawing.Size(352, 29);
            this.Lbl_Kullanıcı_Ad.TabIndex = 1;
            this.Lbl_Kullanıcı_Ad.Text = "KULLANICI ADINIZI GİRİNİZ :";
            // 
            // Lbl_Sifre
            // 
            this.Lbl_Sifre.AutoSize = true;
            this.Lbl_Sifre.BackColor = System.Drawing.Color.Yellow;
            this.Lbl_Sifre.Font = new System.Drawing.Font("Microsoft Sans Serif", 19.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Lbl_Sifre.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.Lbl_Sifre.Location = new System.Drawing.Point(27, 178);
            this.Lbl_Sifre.Name = "Lbl_Sifre";
            this.Lbl_Sifre.Size = new System.Drawing.Size(337, 38);
            this.Lbl_Sifre.TabIndex = 2;
            this.Lbl_Sifre.Text = "ŞİFRENİZİ GİRİNİZ :";
            // 
            // Txtbox_Kullanıcı_Ad
            // 
            this.Txtbox_Kullanıcı_Ad.BackColor = System.Drawing.Color.MediumPurple;
            this.Txtbox_Kullanıcı_Ad.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Txtbox_Kullanıcı_Ad.Location = new System.Drawing.Point(383, 112);
            this.Txtbox_Kullanıcı_Ad.Name = "Txtbox_Kullanıcı_Ad";
            this.Txtbox_Kullanıcı_Ad.Size = new System.Drawing.Size(364, 34);
            this.Txtbox_Kullanıcı_Ad.TabIndex = 3;
            // 
            // Txtbox_Sifre
            // 
            this.Txtbox_Sifre.BackColor = System.Drawing.Color.MediumPurple;
            this.Txtbox_Sifre.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Txtbox_Sifre.Location = new System.Drawing.Point(383, 182);
            this.Txtbox_Sifre.Name = "Txtbox_Sifre";
            this.Txtbox_Sifre.PasswordChar = '?';
            this.Txtbox_Sifre.Size = new System.Drawing.Size(364, 34);
            this.Txtbox_Sifre.TabIndex = 4;
            // 
            // Btn_Giris_Yap
            // 
            this.Btn_Giris_Yap.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.Btn_Giris_Yap.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Btn_Giris_Yap.ForeColor = System.Drawing.Color.Yellow;
            this.Btn_Giris_Yap.Location = new System.Drawing.Point(634, 272);
            this.Btn_Giris_Yap.Name = "Btn_Giris_Yap";
            this.Btn_Giris_Yap.Size = new System.Drawing.Size(113, 71);
            this.Btn_Giris_Yap.TabIndex = 5;
            this.Btn_Giris_Yap.Text = "Giriş Yap";
            this.Btn_Giris_Yap.UseVisualStyleBackColor = false;
            this.Btn_Giris_Yap.Click += new System.EventHandler(this.Btn_Giris_Yap_Click);
            // 
            // Btn_Üye_Ol
            // 
            this.Btn_Üye_Ol.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.Btn_Üye_Ol.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Btn_Üye_Ol.ForeColor = System.Drawing.Color.Yellow;
            this.Btn_Üye_Ol.Location = new System.Drawing.Point(459, 272);
            this.Btn_Üye_Ol.Name = "Btn_Üye_Ol";
            this.Btn_Üye_Ol.Size = new System.Drawing.Size(113, 71);
            this.Btn_Üye_Ol.TabIndex = 6;
            this.Btn_Üye_Ol.Text = "Üye Ol";
            this.Btn_Üye_Ol.UseVisualStyleBackColor = false;
            this.Btn_Üye_Ol.Click += new System.EventHandler(this.Btn_Üye_Ol_Click);
            // 
            // Btn_Unuttum
            // 
            this.Btn_Unuttum.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.Btn_Unuttum.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.Btn_Unuttum.ForeColor = System.Drawing.Color.Yellow;
            this.Btn_Unuttum.Location = new System.Drawing.Point(278, 275);
            this.Btn_Unuttum.Name = "Btn_Unuttum";
            this.Btn_Unuttum.Size = new System.Drawing.Size(113, 71);
            this.Btn_Unuttum.TabIndex = 7;
            this.Btn_Unuttum.Text = "Şifremi Unuttum";
            this.Btn_Unuttum.UseVisualStyleBackColor = false;
            // 
            // Frm_Login
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.Btn_Unuttum);
            this.Controls.Add(this.Btn_Üye_Ol);
            this.Controls.Add(this.Btn_Giris_Yap);
            this.Controls.Add(this.Txtbox_Sifre);
            this.Controls.Add(this.Txtbox_Kullanıcı_Ad);
            this.Controls.Add(this.Lbl_Sifre);
            this.Controls.Add(this.Lbl_Kullanıcı_Ad);
            this.Controls.Add(this.Lbl_Üst_Bilgi);
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.Name = "Frm_Login";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Login";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label Lbl_Üst_Bilgi;
        private System.Windows.Forms.Label Lbl_Kullanıcı_Ad;
        private System.Windows.Forms.Label Lbl_Sifre;
        private System.Windows.Forms.TextBox Txtbox_Kullanıcı_Ad;
        private System.Windows.Forms.TextBox Txtbox_Sifre;
        private System.Windows.Forms.Button Btn_Giris_Yap;
        private System.Windows.Forms.Button Btn_Üye_Ol;
        private System.Windows.Forms.Button Btn_Unuttum;
    }
}

