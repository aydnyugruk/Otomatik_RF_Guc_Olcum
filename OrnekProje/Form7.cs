using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ivi.Visa;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Syncfusion.XlsIO.Implementation.PivotAnalysis;
using System.Windows.Forms.DataVisualization.Charting;
namespace OrnekProje
{
    public partial class Form7 : Form
    {
        public double PRf;
        public double Std;
        public List<double> PRfs;

        public Form7()
        {
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {

        }

        private void chart2_Click(object sender, EventArgs e)
        {

        }
        void get_data_from_xl()
        {
            PRfs = new List<double>();
            var mySheet = Path.Combine(Directory.GetCurrentDirectory(), "Book1.xls");
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;
            object misValue = System.Reflection.Missing.Value;

            if (xlApp == null)
            {
                return;
            }

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(mySheet);
            Excel._Worksheet workSheet = xlApp.ActiveSheet;

            int row_num = int.Parse(workSheet.Cells[3, 2].Value.ToString());

            for (int i = 5; i < row_num + 5; i++)
            {
                PRfs.Add(double.Parse(Math.Round(workSheet.Cells[i, 8].Value, 4).ToString()));
            }

            PRf = double.Parse(Math.Round(workSheet.Cells[5, 12].Value, 4).ToString());
            Std = double.Parse(Math.Round(workSheet.Cells[6, 12].Valie, 4).ToString());

            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double z = 0;
            double left = -1;
            double right = -1;
            double genisletilmis_belirsizlik = 0;


            chart2.Series["Okunan Değerler"].Points.Clear();
            chart2.Series["Güvenilirlik Düzeyi"].Points.Clear();
            chart2.Series["Left"].Points.Clear();
            chart2.Series["Right"].Points.Clear();

            if (PRfs == null)
            {
                get_data_from_xl();
            }

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
                right = PRf + z * (Std / Math.Sqrt(PRfs.Count()));
                left = PRf - z * (Std / Math.Sqrt(PRfs.Count()));
            }

            if (PRf != 0 && Std != 0 && PRfs != null)
            {
                PRfs.Sort();


                for(int i = 0; i < PRfs.Count(); i++)
                {
                    PRfs[i] = Math.Round(PRfs[i], 4);
                }
                bool leftFlag = false;
                bool rightFlag = false;
                int prev_count = 0;

                for (int i = 0; i < PRfs.Count(); i++)
                {
                    int count = PRfs.Where(temp => temp.Equals(PRfs[i]))
                             .Select(temp => temp)
                             .Count();
                    try
                    {
                        this.chart2.Series["Okunan Değerler"].Points.AddXY(PRfs[i], count);
                        if (left != -1 && right != -1 && z != 0) {
                            if (PRfs[i] >= left && PRfs[i] <= right)
                            {
                                this.chart2.Series["Güvenilirlik Düzeyi"].Points.AddXY(PRfs[i], count);
                                if (leftFlag == false)
                                {
                                    genisletilmis_belirsizlik = PRf - PRfs[i];
                                    this.chart2.Series["Left"].Points.AddXY(PRfs[i], 0/*count*/);
                                    leftFlag = true;
                                }
                            }
                            if(PRfs[i] > right)
                            {
                                if (rightFlag == false)
                                {
                                    this.chart2.Series["Right"].Points.AddXY(PRfs[i - 1], 0/*prev_count*/);
                                    rightFlag = true;
                                }
                            }
                        }
                    }
                    catch
                    {

                    }
                    prev_count = count;
                }

                chart2.Visible = true;
                label1.Visible = true;
                label2.Visible = true;

                label3.Visible = true;
                label4.Visible = true;
                label5.Visible = true;
                label6.Visible = true;
                label7.Visible = true;

                textBox1.Visible = true;
                textBox2.Visible = true;
                textBox3.Visible = true;
                textBox4.Visible = true;
                textBox5.Visible = true;
            }
            PRfs.Sort();
            textBox1.Text = Math.Round(PRf, 5).ToString();
            textBox2.Text = Math.Round(genisletilmis_belirsizlik, 5).ToString();
            textBox3.Text = Math.Round(Std, 5).ToString();
            textBox4.Text = Math.Round(PRfs[0], 5).ToString();
            textBox5.Text = Math.Round(PRfs[PRfs.Count - 1], 5).ToString();
        }
    }
}
