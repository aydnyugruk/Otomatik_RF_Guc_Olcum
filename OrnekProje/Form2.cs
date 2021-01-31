using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OrnekProje
{
    public partial class Form2 : Form
    {
        Form3 form3;
        Form4 form4;
        Form5 form5;
        Form6 form6;
        Form7 form7;

        public Form2()
        {
            InitializeComponent(); 
            form3 = new Form3();
            form4 = new Form4();
            form5 = new Form5();
            form6 = new Form6();
            form7 = new Form7();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private Form activeForm = null;
        private void openChildFormInPanel(Form childForm)
        {
            activeForm = childForm;
            childForm.TopLevel = false;
            childForm.FormBorderStyle = FormBorderStyle.None;
            childForm.Dock = DockStyle.Fill;
            panelSide.Controls.Add(childForm);
            panelSide.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();
        }
        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            button1.BackColor = Color.FromArgb(19, 116, 180);
            button2.BackColor = Color.FromArgb(20, 63, 181);
            button3.BackColor = Color.FromArgb(20, 63, 181);
            button4.BackColor = Color.FromArgb(20, 63, 181);
            button5.BackColor = Color.FromArgb(20, 63, 181);

            form3.measuredPowers = form4.readings;
            openChildFormInPanel(form3);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.FromArgb(19, 116, 180);
            button1.BackColor = Color.FromArgb(20, 63, 181);
            button3.BackColor = Color.FromArgb(20, 63, 181);
            button4.BackColor = Color.FromArgb(20, 63, 181);
            button5.BackColor = Color.FromArgb(20, 63, 181);
            openChildFormInPanel(form4);
        }

        private void panelSide_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.BackColor = Color.FromArgb(19, 116, 180);
            button1.BackColor = Color.FromArgb(20, 63, 181);
            button2.BackColor = Color.FromArgb(20, 63, 181);
            button4.BackColor = Color.FromArgb(20, 63, 181);
            button5.BackColor = Color.FromArgb(20, 63, 181);
            openChildFormInPanel(form5);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.BackColor = Color.FromArgb(19, 116, 180);
            button1.BackColor = Color.FromArgb(20, 63, 181);
            button2.BackColor = Color.FromArgb(20, 63, 181);
            button3.BackColor = Color.FromArgb(20, 63, 181);
            button5.BackColor = Color.FromArgb(20, 63, 181);

            form6.PRf = form3.PRfAnalitik;
            form6.Std = form3.StdUnc;

            openChildFormInPanel(form6);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            button5.BackColor = Color.FromArgb(19, 116, 180);
            button1.BackColor = Color.FromArgb(20, 63, 181);
            button2.BackColor = Color.FromArgb(20, 63, 181);
            button3.BackColor = Color.FromArgb(20, 63, 181);
            button4.BackColor = Color.FromArgb(20, 63, 181);

            form7.PRf = form3.PRf;
            form7.Std = form3.StdPRf;
            form7.PRfs = form3.PRfs;

            openChildFormInPanel(form7);

        }
    }
}
