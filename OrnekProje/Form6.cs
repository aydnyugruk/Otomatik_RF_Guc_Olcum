using Syncfusion.XlsIO.Implementation.PivotAnalysis;
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
    public partial class Form6 : Form
    {
        public double PRf;
        public double Std;
        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Load(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            double z = 0;
            double right = -1;
            double left = -1;

            chart1.Series["Analitik Ölçümler"].Points.Clear();
            chart1.Series["Güvenilirlik Düzeyi"].Points.Clear();
            chart1.Series["Mean"].Points.Clear();
            chart1.Series["Left"].Points.Clear();
            chart1.Series["Right"].Points.Clear();

            if (comboBox1.Text != "")
            {
                if (comboBox1.Text == "%0")
                {
                    z = 0;
                }
                else if (comboBox1.Text == "%68.27")
                {
                    z = 1.0;
                }
                else if (comboBox1.Text == "%90")
                {
                    z = 1.64;
                }
                else if (comboBox1.Text == "%95")
                {
                    z = 1.96;
                }
                else if (comboBox1.Text == "%95.45")
                {
                    z = 2.0;
                }
                else if (comboBox1.Text == "%99")
                {
                    z = 2.58;
                }
                else if (comboBox1.Text == "%99.73")
                {
                    z = 3.0;
                }
                right = PRf + z * (Std / Math.Sqrt(10));
                left = PRf - z * (Std / Math.Sqrt(10));
            }
            if (PRf != 0 && Std != 0)
            {
                bool leftFlag = false;
                bool rightFlag = false;
                double genisletilmis_belirsizlik = 0;
                for (int i = 99; i >= 1; i--)
                {

                    this.chart1.Series["Analitik Ölçümler"].Points.AddXY(Math.Round(PRf - (2 * i - 2) * 0.0001, 5), F(PRf - (2 * i - 2) * 0.0001, PRf, Std));
                    if (left != -1 && right != -1 && z != 0)
                    {
                        if (PRf - (2 * i - 2) * 0.0001 >= left && PRf - (2 * i - 2) * 0.0001 <= right)
                        {
                            this.chart1.Series["Güvenilirlik Düzeyi"].Points.AddXY(Math.Round(PRf - (2 * i - 2) * 0.0001, 5), F(PRf - (2 * i - 2) * 0.0001, PRf, Std));
                            if (!leftFlag)
                            {
                                this.chart1.Series["Left"].Points.AddXY(Math.Round(left, 5), F(PRf - (2 * i - 2) * 0.0001, PRf, Std));
                                leftFlag = true;
                            }
                        }
                    }
                }
                for (int i = 1; i < 100; i++)
                {
                    this.chart1.Series["Analitik Ölçümler"].Points.AddXY(Math.Round(PRf + 2 * i * 0.0001, 5), F(PRf + 2 * i * 0.0001, PRf, Std));
                    if (left != -1 && right != -1 && z != 0)
                    {
                        if (PRf + 2 * i * 0.0001 >= left && PRf + 2 * i * 0.0001 <= right)
                        {
                            this.chart1.Series["Güvenilirlik Düzeyi"].Points.AddXY(Math.Round(PRf + 2 * i * 0.0001, 5), F(PRf + 2 * i * 0.0001, PRf, Std));
                        }
                        else
                        {
                            if(rightFlag == false)
                            {
                                genisletilmis_belirsizlik = Math.Round(PRf + 2 * (i - 1) * 0.0001, 5) - PRf;
                                this.chart1.Series["Right"].Points.AddXY(Math.Round(right, 5),  F(PRf + 2 * i * 0.0001, PRf, Std));
                                rightFlag = true;
                            }
                        }
                    }

                    chart1.Visible = true;
                    label1.Visible = true;
                    label2.Visible = true;
                    label3.Visible = true;
                    label4.Visible = true;
                    label5.Visible = true;

                    textBox1.Visible = true;
                    textBox2.Visible = true;
                    textBox3.Visible = true;
                }
                this.chart1.Series["Mean"].Points.AddXY(Math.Round(PRf, 5), F(PRf, PRf, Std));

                textBox1.Text = Math.Round(PRf, 5).ToString();
                textBox2.Text = Math.Round(genisletilmis_belirsizlik, 5).ToString();
                textBox3.Text = Math.Round(Std, 5).ToString();

            }
        }
        private double F(double x, double mean, double stddev)
        {
            return (double)((1 / (Math.Sqrt(2 * Math.PI) * stddev)) * Math.Exp(-1 * (x - mean) * (x - mean) / (2 * stddev * stddev)));
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
