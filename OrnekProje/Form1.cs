using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OrnekProje
{
    public partial class Form1 : Form
    {
        public event EventHandler TextChanged;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void buttonLogin_Click(object sender, EventArgs e)
        {
          /*  if(textUsername.Text == "spark" && textPassword.Text == "1234") // Sifre ayarlamalari ve kullanici girisleri form1'de yapildi 
            {
           */

                Form2 yeni = new Form2();
                yeni.Show();

                this.Owner = yeni;

                this.Hide();
            //}
            
        }

        private void materialSingleLineTextField1_Click(object sender, EventArgs e)
        {
            if (textUsername.Text == "Username")
            {
                textUsername.Clear();
            }
        }
        private void textPassword_TextChanged(object sender, EventArgs e)
        {
            textPassword.UseSystemPasswordChar = true;

        }
        private void textPassword_Click(object sender, EventArgs e)
        {
            if (textPassword.Text == "Password")
            {
                textPassword.Clear();
            }
        }

        private void SPanel_Paint(object sender, PaintEventArgs e)
        {

        }

        private void sPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
