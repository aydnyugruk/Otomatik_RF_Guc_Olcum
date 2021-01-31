using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Ivi.Visa.Interop;         // cihazla baglanti kurmayi sagliyor
using Excel = Microsoft.Office.Interop.Excel; // excelde islem gormeyi sagliyor
using System.IO;
using System.Runtime.InteropServices;           // excele yardimci kutuphaneler

namespace OrnekProje
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }
        FormattedIO488 gpibDevice;
        FormattedIO488 gpibDevice2; // kutuphaneden gelen gpib objeleri
        ResourceManager manager;
        ResourceManager manager2;
        string show20 = "";
        string sample_size;
        double deger;
        public List<double> readings;

        // ölçülen değerlerin excele aktarılması
        private void create_xl(List<double> arr)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();               //Excel Dosyasının Oluşturulması
            var mySheet = Path.Combine(Directory.GetCurrentDirectory(), "Book1.xls");
            int i;

            if (File.Exists(mySheet))
            {
                File.Delete(mySheet);
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);



           

            xlWorkSheet.Cells.ColumnWidth = 15;
           // xlWorkSheet.Cells[3, 4].Font.Size = 20;
           // xlWorkSheet.Cells[3, 4].Font.Bold = true;
            //xlWorkSheet.Cells[3,4] = "OTOMATİK RF GÜÇ ÖLÇÜMÜ";
          //  xlWorkSheet.get_Range("D3:G3", Type.Missing).Merge(Type.Missing);

            xlWorkSheet.Cells[3, 1] = "Ölçüm Sayısı";
            xlWorkSheet.Cells[3, 2] = sample_size;

            xlWorkSheet.Cells[5, 1] = "Frekanslar(Hz)";


            int örnek_sayısı = Convert.ToInt32(sample_size);
            Double ornek_miktarı = Convert.ToDouble(sample_size);

            int satir = 5;
            int sutun = 2;
            int k=1;

            for (i = 1; i <= Convert.ToInt32(sample_size); i++)         
            {
                xlWorkSheet.Cells[5, i+1] = String.Format("{0}.Ölçüm(W)", i );
               
                

            }



          //SATIRLARA GÜÇLERİ YAZDIRMA


            for (int j = 1; j <= frekanslar.Count; j++)                  //satır
            {

                for (i = 1; i <= Convert.ToInt32(sample_size); i++)
                {



                     
                    //WATT VE DBM DEĞERLERİNİ ÖLÇÜP EXCELE YAZMA



                    gpibDevice2.WriteString("UNIT:POWer?");          //GÜÇ BİRİMİNİ BELİRLEME EĞER DBM İSE DÖNÜŞÜM YAPIYOR...

                    string guc_birimi = gpibDevice2.ReadString();

                    List<string> okunan_birim = guc_birimi.Split('\n').ToList();             //WATT BİRİMİ OKUNDUĞUNDA

                    if (okunan_birim[0]=="DBM")
                    {


                        //güç birimi DBM ise


                        xlWorkSheet.Cells[satir + j, sutun + i - 1] = Math.Round(Math.Pow(10, (readings[k - 1] - 30) / 10), 7);        //okunan değerleri watt cinsinden yan yan bastırdım.

                        k++;

                        xlWorkSheet.Cells[5 + j, 1] = frekanslar[j - 1];               //frekans değerlerini alt alta basmak için




                    }





                    else if (okunan_birim[0] == "W")
                    {


                        //güç birimi DBM ise


                        xlWorkSheet.Cells[satir + j, sutun + i - 1] = Math.Round( readings[k - 1] , 7);        //okunan değerleri watt cinsinden yan yan bastırdım.

                        k++;

                        xlWorkSheet.Cells[5 + j, 1] = frekanslar[j - 1];               //frekans değerlerini alt alta basmak için

                    }


                    }

                }




                double toplam = 0.0;
                double ortalama = 0.0;
                double std_sapma = 0.0;
                double s_deger = 0.0;
                double fark = 0.0;




            //ORTALAMA GÜÇLERİ YAZDIRMA


                for (int j = 0; j < frekanslar.Count; j++)                  //satır
                {

                    for (i = 0; i < Convert.ToInt32(sample_size); i++)
                    {





                        gpibDevice2.WriteString("UNIT:POWer?");          //GÜÇ BİRİMİNİ BELİRLEME EĞER DBM İSE DÖNÜŞÜM YAPIYOR...

                        string guc_birimi = gpibDevice2.ReadString();

                        List<string> okunan_birim = guc_birimi.Split('\n').ToList();             //WATT BİRİMİ OKUNDUĞUNDA

                        if (okunan_birim[0] == "DBM")
                        {


                            //güç birimi DBM ise



                            xlWorkSheet.Cells[satir, örnek_sayısı + 3] = "Ortalama Güç(W)";

                            toplam = toplam + Math.Round(Math.Pow(10, (readings[örnek_sayısı * j + i] - 30) / 10), 7);

                        }



                        else if (okunan_birim[0] == "W")
                        {


                            //güç birimi W ise

                            xlWorkSheet.Cells[satir, örnek_sayısı + 3] = "Ortalama Güç(W)";

                            toplam = toplam + Math.Round(readings[örnek_sayısı * j + i],7);


                        }


                    }


                    ortalama = toplam / örnek_sayısı;

                    toplam = 0;


                    xlWorkSheet.Cells[satir + j + 1, örnek_sayısı + 3] = ortalama;







                    for (int m = 0; m < Convert.ToInt32(sample_size); m++)
                    {
                        s_deger = Math.Round(readings[örnek_sayısı * j + m], 7);

                        fark = s_deger - ortalama;



                        toplam = toplam + Math.Pow(fark, 2);

                        toplam = Math.Round(toplam, 10);
                    }

                    toplam = toplam / (ornek_miktarı - 1);

                    std_sapma = Math.Sqrt(toplam);
                    std_sapma = Math.Round(std_sapma, 7);



                    xlWorkSheet.Cells[satir, örnek_sayısı + 4] = "Standart Sapma";

                    xlWorkSheet.Cells[satir + j + 1, örnek_sayısı + 4] = std_sapma;



                    std_sapma = 0.0;
                    toplam = 0.0;
                    ortalama = 0.0;
                    s_deger = 0.0;
                    fark = 0.0;






                }

            





            
              
 

          

            xlWorkBook.SaveAs(mySheet, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            pictureBox9.Visible = false;
            pictureBox7.Visible = false;    // yesil tiklerin gorunmez yapilmasi
            pictureBox5.Visible = false;
            pictureBox3.Visible = false;
            pictureBox1.Visible = false;
            pictureBox4.Visible = false;
            pictureBox6.Visible = false;
            pictureBox10.Visible = false;


            label9.Visible = false;

            string waveform, show1 = "", show2 = "", show5 = "";
            string show9 = "", frequence;
            double sayi1;

            //YEŞİL TİKLER
            if (comboBox2.Text != "")
            {
                pictureBox9.Visible = true;
            }
            if (comboBox1.Text != "")
            {
                pictureBox3.Visible = true;
            }
            if (comboBox4.Text != "" && textBox4.Text != "")
            {
                pictureBox7.Visible = true;
            }
            if (comboBox3.Text != "" && textBox3.Text != "")
            {
                pictureBox5.Visible = true;
            }
            if (textBox1.Text != "")
            {
                pictureBox1.Visible = true;
            }




            if (comboBox5.Text != "" && materialSingleLineTextField1.Text != "")
            {
                pictureBox4.Visible = true;
            }
            if (comboBox6.Text != "" && materialSingleLineTextField2.Text != "")
            {
                pictureBox6.Visible = true;
            }
            if (comboBox7.Text != "" && materialSingleLineTextField3.Text != "")
            {
                pictureBox10.Visible = true;
            }











            // Secilen cihazin belirlenmesi ve uygun komutlar gonderilmesi
            if (comboBox1.Text == "Agilent 33250A (1 µHz to 80 MHz)")
            {

                string GPIBAddress;

                manager = new ResourceManager();
                gpibDevice = new FormattedIO488();

                GPIBAddress = textBox1.Text;
                try
                {
                    gpibDevice.IO = (IMessage)manager.Open(GPIBAddress);
                }
                catch
                {
                    label9.Text = "Cihaza erişim sağlanamadı";
                    label9.Visible = true;
                    return;
                }

                //IDN sorma
                gpibDevice.WriteString("*IDN?");

                show9 = gpibDevice.ReadString();

                //frekansı alma
                waveform = comboBox2.Text;
                if (waveform != "")
                {
                    gpibDevice.WriteString("FUNCtion:SHAPe " + waveform);
                    gpibDevice.WriteString("FUNC:SHAPe?");
                    show1 = gpibDevice.ReadString();
                }

                frequence = comboBox3.Text;
                if (frequence != "")
                {
                    deger = 0.0;
                    switch (frequence)
                    {
                        case "Hz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1;
                            break;
                        case "KHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000;                                                   //FREKANS BİRİMLERİ BURADA TANIMLANIYOR.
                            break;
                        case "MHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000000;
                            break;
                        case "GHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000000000;
                            break;
                    }
                    gpibDevice.WriteString("SOURce:FREQuency:CW " + deger);
                    gpibDevice.WriteString("FREQuency?");
                    show2 = gpibDevice.ReadString();
                }

                // gücün ayarlanması
                if (comboBox4.Text == "dBm")
                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " DBM");
                gpibDevice.WriteString("VOLTage?");
                show5 = gpibDevice.ReadString();

                listbox.Items.Add("Cihaz İsmi:" + show9);
                // listbox.Items.Add("Frekans:" + show2);                                    //LİSTBOXT'TA CİHAZ BİLGİLERİ ,GÜÇ VE FREKANS DEĞERLERİ BURADA GÖSTERİLİYOR.
                // listbox.Items.Add("Güç:" + show5);
                listbox.Items.Add("Dalga Şekli:" + show1);

                for (int i = 1; i < 21; i++)
                {
                    gpibDevice.WriteString("SYSTem:ERRor?");
                    show20 = gpibDevice.ReadString();
                    if (show20 == "+0,\"No error\"\n")                                           //SİSTEM HATALARINI BURDA GÖSTERİLİYOR.
                        break;
                    if (show20 != "+0,\"No error\"\n")
                        listbox.Items.Add("ERROR:" + show20);
                }
                listbox.Items.Add("");

            }
            else if (comboBox1.Text == "Keysight E8257D(100 kHz to 67 GHz)")               //E8257D scpi komutlari ve ayarlari
            {
                string GPIBAddress;

                manager = new ResourceManager();
                gpibDevice = new FormattedIO488();

                GPIBAddress = textBox1.Text;
                try
                {
                    gpibDevice.IO = (IMessage)manager.Open(GPIBAddress);
                }
                catch
                {
                    label9.Text = "Cihaza erişim sağlanamadı";
                    label9.Visible = true;
                    return;
                }

                //IDN SORMA
                gpibDevice.WriteString("*IDN?");

                show9 = gpibDevice.ReadString();

                //FREKANS ALMA
                show1 = comboBox2.Text;

                frequence = comboBox3.Text;
                if (frequence != "")
                {
                    deger = 0.0;
                    switch (frequence)
                    {
                        case "Hz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1;
                            break;
                        case "KHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000;
                            break;
                        case "MHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000000;
                            break;
                    }
                    gpibDevice.WriteString("SOURce:FREQuency:CW " + deger);
                    gpibDevice.WriteString("SOURce:FREQuency:CW?");
                    show2 = gpibDevice.ReadString();
                }


                // GÜCÜN AYARLANMASI
                if (comboBox4.Text == "dBm")
                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " DBM");
                if (comboBox4.Text == "W")
                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " W");
                if (comboBox4.Text == "mW")
                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " MW");
                if (comboBox4.Text == "UW")
                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " UW");
                gpibDevice.WriteString("VOLTage?");
                show5 = gpibDevice.ReadString();

                listbox.Items.Add("Cihaz İsmi:" + show9);
                listbox.Items.Add("Frekans:" + show2);
                listbox.Items.Add("Güç:" + show5);
                listbox.Items.Add("Dalga Şekli:" + show1);

                for (int i = 1; i < 21; i++)
                {
                    gpibDevice.WriteString("SYSTem:ERRor?");
                    show20 = gpibDevice.ReadString();
                    if (show20 == "+0,\"No error\"\n")
                        break;
                    if (show20 != "+0,\"No error\"\n")
                        listbox.Items.Add("ERROR:" + show20);
                }
                listbox.Items.Add("");
            }


            if (comboBox1.Text == "Agilent 33120A (100 μHz - 15 MHz)")                        //33120A CİHAZININ BİLGİLERİ BURADA TANIMLANDI..
            {

                string GPIBAddress;

                manager = new ResourceManager();
                gpibDevice = new FormattedIO488();

                GPIBAddress = textBox1.Text;
                try
                {
                    gpibDevice.IO = (IMessage)manager.Open(GPIBAddress);
                }
                catch
                {
                    label9.Text = "Cihaza erişim sağlanamadı";
                    label9.Visible = true;
                    return;
                }

                //IDN sorma
                gpibDevice.WriteString("*IDN?");

                show9 = gpibDevice.ReadString();

                //frekansı alma
                waveform = comboBox2.Text;
                if (waveform != "")
                {
                    gpibDevice.WriteString("FUNCtion:SHAPe " + waveform);
                    gpibDevice.WriteString("FUNC:SHAPe?");
                    show1 = gpibDevice.ReadString();
                }

                frequence = comboBox3.Text;
                if (frequence != "")
                {
                    deger = 0.0;
                    switch (frequence)
                    {
                        case "Hz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1;                                                                    //SAYİ1=OKUNAN FREKANS DEĞERİ
                            break;
                        case "KHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000;
                            break;
                        case "MHz":
                            sayi1 = Convert.ToDouble(textBox3.Text);
                            deger = sayi1 * 1000000;
                            break;
                    }
                    // gpibDevice.WriteString("SOURce:FREQuency:CW " + deger);                                   //FREKANS DEĞERİNİ CİHAZA GÖNDERME BURDA.
                    // gpibDevice.WriteString("FREQuency?");
                    // show2 = gpibDevice.ReadString();
                }

                // GÜCÜN AYARLANMASI





                /* gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude " + textBox4.Text + " DBM");             // GÜCÜN AYARLANMASI(DEĞİŞİKLİK BURDA YAPILACAK)
                 gpibDevice.WriteString("VOLTage?");
                 show5 = gpibDevice.ReadString();
                 */
                listbox.Items.Add("Cihaz İsmi:" + show9);
                //  listbox.Items.Add("Frekans:" + show2);
                // listbox.Items.Add("Güç:" + show5);
                listbox.Items.Add("Dalga Şekli:" + show1);

                /*  for (int i = 1; i < 21; i++)
                  {
                      gpibDevice.WriteString("SYSTem:ERRor?");
                      show20 = gpibDevice.ReadString();
                      if (show20 == "+0,\"No error\"\n")                               //SİSTEM HATALARI BURADA TANIMLANMIŞTIR.
                          break;
                      if (show20 != "+0,\"No error\"\n")
                          listbox.Items.Add("ERROR:" + show20);
                  }
                 */

                listbox.Items.Add("");

            }


        }





        // Sil butonu
        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (gpibDevice != null)
                {
                    gpibDevice.WriteString("*RST");
                    gpibDevice.WriteString("*CLS");



                }
            }
            catch
            {

            }


            try
            {
                if (gpibDevice2 != null)
                {
                    gpibDevice2.WriteString("*RST");
                    gpibDevice2.WriteString("*CLS");



                }
            }
            catch
            {

            }













            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";


            textBox3.Clear();
            textBox4.Clear();
            textBox1.Clear();

            label9.Visible = false;

            pictureBox9.Visible = false;
            pictureBox7.Visible = false;
            pictureBox5.Visible = false;
            pictureBox3.Visible = false;
            pictureBox1.Visible = false;

            listbox.Items.Clear();
            listBox2.Items.Clear();

            gucler.Clear();                       //güç ve frekans listesini silmek için kullandım.
            frekanslar.Clear();
            matrix.Clear();

            label9.Visible = false;
            label19.Visible = false;



        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }


        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

       // int baslangıc_deger = 0;


        // Gucmetrenin ayarlari buradan yapildi
        private void button3_Click(object sender, EventArgs e)
        {
            pictureBox2.Visible = false;
            pictureBox8.Visible = false;


            label10.Visible = false;

            if (textBox2.Text != "")
            {
                pictureBox2.Visible = true;
            }
            /* if (comboBox5.Text != "")
             {
                 pictureBox4.Visible = true;
             }
             * */
            if (textBox19.Text != "")
            {
                pictureBox8.Visible = true;
            }

            if (textBox2.Text == "" || textBox19.Text == "")
            {
                label10.Text = "Girdi değerleri eksik";
                label10.Visible = true;
                return;
            }
            // progressbar duzenlenmesi 
            sample_size = textBox19.Text;
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 100;
            progressBar1.Value = 0;
            frekans_sayisi = listBox2.Items.Count;  
            progressBar1.Step = (100 / (Convert.ToInt32(sample_size) * frekans_sayisi));         // progressbarın hızını belirledim.



            string GPIBAddress2;            //Güçmetrenin GPIB ayarlaması burada yapılmıştır
            string show3;
            readings = new List<double>();

            manager2 = new ResourceManager();
            gpibDevice2 = new FormattedIO488();

            GPIBAddress2 = textBox2.Text;
            try
            {
                gpibDevice2.IO = (IMessage)manager2.Open(GPIBAddress2);
            }
            catch
            {
                label10.Text = "Cihaza erişim sağlanamadı";
                label10.Visible = true;
                return;
            }



            double frekans_deger1 = 0.0;     //sinyal generatorunun frekansı
            double guc_deger1 = 0.0;        //sinyal generatorunun gücü için






            if (comboBox4.Text == "dBm")
            {


                DataColumn kolon1 = new DataColumn("İşaret Üreteci Güç Değeri(dBm)");

                DataColumn kolon2 = new DataColumn("İşaret Üreteci Güç Değeri(W)");

                DataColumn kolon3 = new DataColumn("İşaret Üreteci Frekans Değeri(Hz)");

                DataColumn kolon4 = new DataColumn("Güçmetreden Okunan Güç Değeri(dBm)");                          //DATAGRİD TABLODA SÜTUNLAR 

                DataColumn kolon5 = new DataColumn("Güçmetreden Okunan Güç Değeri(W)");




                datatablosu.Columns.Add(kolon1);
                datatablosu.Columns.Add(kolon2);
                datatablosu.Columns.Add(kolon3);
                datatablosu.Columns.Add(kolon4);
                datatablosu.Columns.Add(kolon5);


                string val1 = "", val2 = "", val3 = "", val4 = "", val5 = "";

                string val2_string;

                double dönüşüm,ölçülen_watt;


                for (int k = 0; k < gucler.Count; k++)
                {



                    try
                    {
                        val1 = gucler[k].ToString();
                        
                    }
                    catch (ArgumentOutOfRangeException)
                    {

                        val1 = "";
                    }



                    
                    try
                    {

                        dönüşüm = ((Convert.ToDouble(val1)-30)/10);
                        val2_string = Math.Pow(10,dönüşüm).ToString();
                        double double_val2 = Convert.ToDouble(val2_string);
                        ölçülen_watt = Math.Round(double_val2, 7);

                        val2 = ölçülen_watt.ToString();
                        
                    }
                    catch (ArgumentOutOfRangeException)
                    {

                        val2 = "";
                    }







                    guc_deger1 = gucler[k];

                    gpibDevice.WriteString("SOURce:VOLTage:LEVel:IMMediate:AMPLitude" + " " + guc_deger1.ToString() + " DBM");
                    // gpibDevice2.WriteString("VOLTage?");

                    //System.Threading.Thread.Sleep(1000);

                    for (int j = 0; j < frekanslar.Count; j++)
                    {
                        if (matrix[j][k] == 1)
                        {
                            frekans_deger1 = frekanslar[j];
                            gpibDevice.WriteString("SOURce:FREQuency:CW  " + frekans_deger1.ToString() + " HZ");          //SENSe:FREQuency:CW 

                            gpibDevice2.WriteString("SENSe:FREQuency:CW " + frekans_deger1.ToString() + " HZ");

                            System.Threading.Thread.Sleep(2000);          //2sn delay verildi.


                            try
                            {
                                val3 = frekanslar[j].ToString();
                            }
                            catch (ArgumentOutOfRangeException)
                            {

                                val3 = "";
                            }

                            for (int i = 0; i < Convert.ToInt32(sample_size); i++)
                            {
                                gpibDevice2.WriteString("MEASure:SCAlar:POWer:AC?");
                                System.Threading.Thread.Sleep(1000);
                                show3 = gpibDevice2.ReadString();
                                progressBar1.PerformStep();

                                double show4 = Convert.ToDouble(show3);
                                show4 = Math.Round(show4, 7);
                                okunan.Add(show4.ToString());
                                readings.Add(Convert.ToDouble(show4));

                                try
                                {
                                    val4 = show4.ToString();
                                }
                                catch (ArgumentOutOfRangeException)
                                {

                                    val4 = "";
                                }


                                double güçmetre_dönüşüm, double_val5, güçmetre_ölçülen_watt;

                                string val5_string;

                                try
                                {
                                    güçmetre_dönüşüm = ((show4 - 30) / 10);
                                    val5_string = Math.Pow(10,  güçmetre_dönüşüm).ToString();

                                    double_val5 = Convert.ToDouble(val5_string);
                                   
                                    güçmetre_ölçülen_watt= Math.Round(double_val5, 7);


                                    val5 = güçmetre_ölçülen_watt.ToString();
                                }
                                catch (ArgumentOutOfRangeException)
                                {

                                    val5 = "";
                                }









                                datatablosu.Rows.Add(val1, val2, val3,val4,val5);




                                //listBox1.Items.Add("Güç Değeri=" + " " + gucler[k].ToString() + "dBm" + "  " + "&&" + " " + " Frekans Değeri=" + " " + frekanslar[j].ToString() + "Hz" + "  " + "Okunan Güç Değeri=" + " " + okunan.ToString("0.00000") + " " + "dBm");
                                // progressBar1.PerformStep();





                                if (i % 100 == 0)
                                {
                                    gpibDevice2.WriteString("VOLTage?");           //cihaz belirli bir okuma sayısına ulaştığında okuma yapmıyor cihaza voltaj sormamız gerekiyor.
                                }


                            }   //sample size

                            dataGridView1.DataSource = datatablosu;
                        }

                       

                    }
                }
            



            progressBar1.Value = 100;



            try
            {
                create_xl(readings);
            }
            catch
            {
                label10.Text = "Excel dosyası oluşturulamadı";
                label10.Visible = true;
            }
          
        }
            
        }
         
          

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            progressBar1.Minimum = 0;
            progressBar1.Maximum = 1000000;
            progressBar1.Value = 0;

            // comboBox5.Text = "";

           

            textBox2.Clear();
            textBox19.Clear();
            pictureBox2.Visible = false;
            pictureBox8.Visible = false;
            //pictureBox4.Visible = false;


            label10.Visible = false;

            dataGridView1.Columns.Clear();             //datagrid temizleme
            datatablosu.Clear();
            dataGridView1.Refresh();


        }

        private void listbox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            listbox.Items.Clear();
            gucler.Clear();                       //güç ve frekans listesini silmek için kullandım.
            frekanslar.Clear();
            matrix.Clear();





            sayıcı = 0;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            
            progressBar1.Value = 0;
            gucler.Clear();                       //güç ve frekans listesini silmek için kullandım.
            frekanslar.Clear();
            matrix.Clear();

            dataGridView1.Columns.Clear();             //datagrid temizleme
            datatablosu.Clear();
            dataGridView1.Refresh();

        }

        private void sPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        //  private void pictureBox4_Click(object sender, EventArgs e)
        //{

        // }
        // Cihaz secildiginde uygun fonksyon turlerinin cikmasi
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (comboBox1.Text == "Agilent 33250A (1 µHz to 80 MHz)")
            {
                comboBox4.Items.Clear();
                comboBox4.Items.Add("dBm");

                comboBox2.Items.Clear();
                comboBox2.Items.Add("SINUSOID");
                comboBox2.Items.Add("SQUARE");
                comboBox2.Items.Add("TRIANGLE");
                comboBox2.Items.Add("RAMP");
            }
            else if (comboBox1.Text == "Keysight E8257D (100 kHz to 67 GHz)")
            {
                comboBox4.Items.Clear();
                comboBox4.Items.Add("W");
                comboBox4.Items.Add("mW");
                comboBox4.Items.Add("dBm");
                comboBox4.Items.Add("UW");

                comboBox2.Items.Clear();
                comboBox2.Items.Add("SINUSOID");
            }

            if (comboBox1.Text == "Agilent 33120A (100 μHz - 15 MHz)")
            {
                comboBox4.Items.Clear();
                comboBox4.Items.Add("dBm");

                comboBox2.Items.Clear();
                comboBox2.Items.Add("SINUSOID");
                comboBox2.Items.Add("SQUARE");
                comboBox2.Items.Add("TRIANGLE");
                comboBox2.Items.Add("RAMP");
            }





        }

        private void sPanel1_Paint(object sender, PaintEventArgs e)
        {




        }

        private void textBox3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {

        }

        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox3.Clear();
            textBox4.Clear();
        }

        private void button7_Click(object sender, EventArgs e)
        {

        }


        int sayıcı = 0;



        List<List<double>> matrix = new List<List<double>>();                  //matrix eklendi
        List<double> frekanslar = new List<double>();
        List<double> gucler = new List<double>();
                        


        List<string> okunan = new List<string>();




        DataTable datatablosu = new DataTable();

        int frekans_sayisi;



        //GİRDİ DEĞERİ EKSİK HATASI



        private void button7_Click_1(object sender, EventArgs e)
        {




            if (comboBox5.Text != "" && comboBox6.Text != "" &&  comboBox7.Text != "" && materialSingleLineTextField1.Text != "" && materialSingleLineTextField2.Text != "" && materialSingleLineTextField3.Text != "")        //çoklu frekans eksik veri kontrolü
            {

                double sonfrekans = 0.00;
                double ilkfrekans = 0.00;                                                //OTOMATİK ADIM ADIM FREKANS TARAMAK İÇİN AYARLAMALAR YAPILDI..
                double artismiktari = 0.00;



                if (comboBox6.Text == "Hz")

                    sonfrekans = Convert.ToDouble(materialSingleLineTextField2.Text);
                else if (comboBox6.Text == "KHz")
                                                                                                              //son frekans değeri
                    sonfrekans = Convert.ToDouble(materialSingleLineTextField2.Text) * 1000;

                else if (comboBox6.Text == "MHz")

                    sonfrekans = Convert.ToDouble(materialSingleLineTextField2.Text) * 1000000;
                else if (comboBox6.Text == "GHz")

                    sonfrekans = Convert.ToDouble(materialSingleLineTextField2.Text) * 1000000000;





                if (comboBox5.Text == "Hz")

                    ilkfrekans = Convert.ToDouble(materialSingleLineTextField1.Text);
                else if (comboBox5.Text == "KHz")
                                                                                                               //  ilk  frekans değeri
                    ilkfrekans = Convert.ToDouble(materialSingleLineTextField1.Text) * 1000;

                else if (comboBox5.Text == "MHz")

                    ilkfrekans = Convert.ToDouble(materialSingleLineTextField1.Text) * 1000000;
                else if (comboBox5.Text == "GHz")

                    ilkfrekans = Convert.ToDouble(materialSingleLineTextField1.Text) * 1000000000;







                if (comboBox7.Text == "Hz")

                    artismiktari = Convert.ToDouble(materialSingleLineTextField3.Text);

                else if (comboBox7.Text == "KHz")
                                                                                                             //  Artış frekans değeri
                    artismiktari = Convert.ToDouble(materialSingleLineTextField3.Text) * 1000;

                else if (comboBox7.Text == "MHz")
                    artismiktari = Convert.ToDouble(materialSingleLineTextField3.Text) * 1000000;
                else if (comboBox7.Text == "GHz")

                    artismiktari = Convert.ToDouble(materialSingleLineTextField3.Text) * 1000000000;









                int Güç_degerleri = listBox2.Items.Count;                                                    //Güç_degerlerilistBox2.Items.Count;              


                double guc = 0.00;


                if (comboBox4.Text == "dBm")

                    guc = Convert.ToDouble(textBox4.Text);

                else if (comboBox4.Text == "mW")

                    guc = Convert.ToDouble(textBox4.Text);


                else if (comboBox4.Text == "W")

                    guc = Convert.ToDouble(textBox4.Text) * 1000;








                if (materialSingleLineTextField1.Enabled == true && materialSingleLineTextField2.Enabled == true && materialSingleLineTextField3.Enabled == true && textBox3.Enabled == false && comboBox3.Enabled == false && comboBox1.Text != "" && comboBox2.Text != "" && comboBox4.Text != "" && textBox4.Text != "" && comboBox5.Text != "" && comboBox6.Text != "" && comboBox7.Text != "" && materialSingleLineTextField1.Text != "" && materialSingleLineTextField2.Text != "" && materialSingleLineTextField3.Text != "")               //EKLE BUTONU İLE FREKANS VE GÜÇ DEĞERLERİNİ LİSTBOX'A YAZDIRMA
                {


                    double dönen_frekans_miktarı = (((sonfrekans - ilkfrekans) / artismiktari) + 1);                         //toplam frekans sayısı formülü





                   if (ilkfrekans >= sonfrekans || artismiktari > sonfrekans)
                    {
                        MessageBox.Show("Frekans Aralıklarını Kontrol Ediniz !!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);              //hataları ekrana bastırma ve frensları tarama
                    }


                    for (int i = 0; i < dönen_frekans_miktarı; i++)
                    {

                        if (ilkfrekans <= sonfrekans && artismiktari < sonfrekans)                                    // OTOMATİK OLARAK FREKANSLARI TARATTIM.
                        {
                            listBox2.Items.Add("Güç Seviyesi=" + " " + textBox4.Text + " " + comboBox4.Text + " " + "&&" + " " + "Frekans=" + " " + ilkfrekans + " " + "Hz");

                         


                            int row = -1;

                            int column = -1;

                            for (int j = 0; j < frekanslar.Count; j++)
                            {

                                if (frekanslar[j] == ilkfrekans)
                                {


                                    row = j;
                                }                                                                  //BENZER GÜÇ VE FREKANSIN OLMAMASI İÇİN KÖŞEGEN MATRİSİNİ KULLANDIM.

                            }




                            for (int  j= 0; j < gucler.Count; j++)
                            {

                                if (gucler[j] == guc)
                                {


                                    column = j;
                                }

                            }


                            if (row == -1)
                            {

                                List<double> temp = new List<double>();

                                frekanslar.Add(ilkfrekans);

                                for (int j = 0; j < gucler.Count; j++)
                                {

                                    temp.Add(0.0);

                                }

                                matrix.Add(temp);

                                row = frekanslar.Count - 1;

                            }

                            if (column != -1)
                            {

                                matrix[row][column] = 1.0;

                            }
                            else
                            {


                                gucler.Add(guc);

                                for (int j = 0; j< frekanslar.Count; j++)
                                {

                                    if (j == row)
                                    {

                                        matrix[j].Add(1.0);
                                    }
                                    else
                                    {

                                        matrix[j].Add(0.0);
                                    }

                                }



                            }




















                            ilkfrekans = ilkfrekans + artismiktari;

                            

                            
                         }

                    }

                  }

           }
               





             if  (materialSingleLineTextField1.Enabled==true&&materialSingleLineTextField2.Enabled==true&&materialSingleLineTextField3.Enabled==true)
            {
                
                
                if (comboBox5.Text == "" || comboBox6.Text == "" || comboBox7.Text == "" || comboBox1.Text == "" || comboBox2.Text == "" || comboBox4.Text == "" || textBox4.Text == "" || materialSingleLineTextField1.Text == "" || materialSingleLineTextField2.Text == "" || materialSingleLineTextField3.Text == "")

                    MessageBox.Show("Frekans Aralıklarını Kontrol Ediniz !!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);          //FREKANS ARALIĞINI KONTROL ETTİM.
            }

           
                        //TEK TEK FREKANSI BASTIRMAK İÇİN
            
            else if (comboBox5.Enabled == false && comboBox6.Enabled == false && comboBox7.Enabled == false && materialSingleLineTextField1.Enabled == false && materialSingleLineTextField2.Enabled == false && materialSingleLineTextField3.Enabled == false && comboBox3.Enabled == true && textBox3.Enabled == true && comboBox3.Text != "" && textBox3.Text != "" && comboBox1.Text != "" && comboBox2.Text != "" && textBox4.Text != "" && comboBox4.Text != "")               //EKLE BUTONU İLE FREKANS VE GÜÇ DEĞERLERİNİ LİSTBOX'A YAZDIRMA
            {
                //sayıcı++;

               listBox2.Items.Add("Güç Seviyesi=" + " " + textBox4.Text + " " + comboBox4.Text + " " + "&&" + " " + "Frekans=" + " " + textBox3.Text + " " + comboBox3.Text);




                frekans_sayisi = listBox2.Items.Count;                                     //TEK TEK BASTIRILAN FREKASNS SAYSINI ÖĞRENMEK İÇİN YAZDIM.
                
                double frekans = 0.00;


                if (comboBox3.Text == "Hz")
                                                                                                   // FREKANSIN BÜYÜKLÜĞÜNÜ HESAPLAMAK İÇİN
                    frekans = Convert.ToDouble(textBox3.Text);
                else if (comboBox3.Text == "KHz")

                    frekans = Convert.ToDouble(textBox3.Text) * 1000;

                else if (comboBox3.Text == "MHz")

                    frekans = Convert.ToDouble(textBox3.Text) * 1000000;
                else if (comboBox3.Text == "GHz")

                    frekans = Convert.ToDouble(textBox3.Text) * 1000000000;








                int Güç_degerleri = listBox2.Items.Count;                 //Güç_degerlerilistBox2.Items.Count;              


                double guc = 0.00;


                if (comboBox4.Text == "dBm")

                    guc = Convert.ToDouble(textBox4.Text);

                else if (comboBox4.Text == "mW")

                    guc = Convert.ToDouble(textBox4.Text);


                else if (comboBox4.Text == "W")

                    guc = Convert.ToDouble(textBox4.Text) * 1000;



               
                                                                     

                int row = -1;

                int column = -1;

                for (int i = 0; i < frekanslar.Count; i++)
                {

                    if (frekanslar[i] == frekans)
                    {


                        row = i;
                    }                                                                  //BENZER GÜÇ VE FREKANSIN OLMAMASI İÇİN KÖŞEGEN MATRİSİNİ KULLANDIM.

                }




                for (int i = 0; i < gucler.Count; i++)
                {

                    if (gucler[i] == guc)
                    {


                        column = i;
                    }

                }


                if (row == -1)
                {

                    List<double> temp = new List<double>();

                    frekanslar.Add(frekans);

                    for (int i = 0; i < gucler.Count; i++)
                    {

                        temp.Add(0.0);

                    }

                    matrix.Add(temp);

                    row = frekanslar.Count - 1;

                }

                if (column != -1)
                {

                    matrix[row][column] = 1.0;

                }
                else
                {


                    gucler.Add(guc);

                    for (int i = 0; i < frekanslar.Count; i++)
                    {

                        if (i == row)
                        {

                            matrix[i].Add(1.0);
                        }
                        else
                        {

                            matrix[i].Add(0.0);
                        }

                    }



                }


            }
            
           
            else if(textBox3.Text== "" ||comboBox3.Text==""||textBox4.Text==""||comboBox4.Text==""||comboBox1.Text==""||comboBox2.Text=="")
                
                 if(textBox3.Enabled=true && comboBox3.Enabled==true)
                   MessageBox.Show("Eksik veri ,giriş değerlerinizi kontrol ediniz !!!", "UYARI", MessageBoxButtons.OK, MessageBoxIcon.Error);           //EKSİK VERİ VE GİRİŞ KONTROLÜ




             }


        

        private void button8_Click_1(object sender, EventArgs e)
        {
            int indeks = listBox2.SelectedIndex;
            listBox2.Items.RemoveAt(indeks);
            sayıcı--;

        }

        private void button9_Click(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {





        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void checkBox1_CheckedChanged_1(object sender, EventArgs e)
        {



        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {


        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {




        }

        private void button10_Click(object sender, EventArgs e)
        {

            comboBox5.Enabled = false;
            comboBox6.Enabled = false;
            comboBox7.Enabled = false;

            materialSingleLineTextField1.Enabled = false;
            materialSingleLineTextField2.Enabled = false;
            materialSingleLineTextField3.Enabled = false;




            materialSingleLineTextField1.BackColor = Color.Red;
            materialSingleLineTextField2.BackColor = Color.Red;
            materialSingleLineTextField3.BackColor = Color.Red;


            materialSingleLineTextField1.Text = "";
            materialSingleLineTextField2.Text = "";
            materialSingleLineTextField3.Text = "";

            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox7.Text = "";
            textBox3.Enabled = true;
            comboBox3.Enabled = true;
            textBox3.BackColor = Color.LightGray;

        }

        private void button9_Click_1(object sender, EventArgs e)
        {


            comboBox5.Enabled = true;
            comboBox6.Enabled = true;
            comboBox7.Enabled = true;
            comboBox3.Enabled = false;

            materialSingleLineTextField1.Enabled = true;
            materialSingleLineTextField2.Enabled = true;
            materialSingleLineTextField3.Enabled = true;
            textBox3.Enabled = false;




            materialSingleLineTextField1.BackColor = Color.LightGray;
            materialSingleLineTextField2.BackColor = Color.LightGray;
            materialSingleLineTextField3.BackColor = Color.LightGray;
            textBox3.BackColor = Color.Red;




            pictureBox5.Visible = false;
            textBox3.Text = "";
            comboBox7.Text = "";

           


        }

        private void button11_Click(object sender, EventArgs e)
        {

           
        

        }
    }
}