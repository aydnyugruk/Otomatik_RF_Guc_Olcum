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
    public partial class Form5 : Form
    {
        public List<double> measuredPowers;
        double ortalama_guc;
        double standart_sapma;
        double min, max;

        public Form5()
        {
            InitializeComponent();
            chart1.Visible = false;
        }

        private void Form5_Load(object sender, EventArgs e)
        {

        }
        // Excelden degerlerin alinmasi
        void get_data_from_xl()
        {
            measuredPowers = new List<double>();
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
                measuredPowers.Add(double.Parse(Math.Round(workSheet.Cells[i, 2].Value, 5).ToString()));
            }
            ortalama_guc = Double.Parse(workSheet.Cells[5, 6].Value.ToString());
            standart_sapma = Double.Parse(workSheet.Cells[6, 6].Value.ToString());
            min = Double.Parse(workSheet.Cells[7, 6].Value.ToString());
            max = Double.Parse(workSheet.Cells[8, 6].Value.ToString());

            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.chart1.Series["Okunan Değerler"].Points.Clear();  // onceden kalan noktalarin silinmesi 
            get_data_from_xl();

            for (int i = 0; i < measuredPowers.Count(); i++)
            {
                int count = measuredPowers.Where(temp => temp.Equals(measuredPowers[i]))
                         .Select(temp => temp)
                         .Count(); // measuredPowers'da measuredPowers[i]'den kac tane var
                try
                {
                    this.chart1.Series["Okunan Değerler"].Points.AddXY(measuredPowers[i], count);
                }
                catch
                {

                }
            }

            chart1.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            label6.Visible = true;

            textBox1.Visible = true;
            textBox2.Visible = true;
            textBox3.Visible = true;
            textBox4.Visible = true;

            textBox1.Text = Math.Round(ortalama_guc, 5).ToString();     // textboxlara sonuclarin son 5 hanesine kadar yuvarlayarak girilmesi
            textBox2.Text = Math.Round(standart_sapma, 5).ToString();
            textBox3.Text = Math.Round(min, 5).ToString();
            textBox4.Text = Math.Round(max, 5).ToString();

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
