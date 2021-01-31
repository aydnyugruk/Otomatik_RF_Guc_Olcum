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
using System.Runtime.InteropServices;

namespace OrnekProje
{
    public partial class Form3 : Form
    {
        // private string show12;
        // string show20 = "";

        // PARAMETRELER
        double frekans;
        double Po; // guc olcer ile yapilan okuma degeri
        double DeltaPmacc; //Error of powermeter accuracy
        double DeltaPRfPow; // Guc olcerin referans rf guc olcumunden gelen belirsizlik yeni olusturulan
        double CFstd; // guc algilayicinin kalibrasyon faktoru
        double rhoGB; //Guc bolucu yansima katsayisi,
        double rhoStd; // Guc algilayicinin yansima katsayisi
        double thetaStd; // Guc algilayicinin yansima katsayisinin fazi
        double thetaGB; // Guc bolucusunun esdeger yansima katsayisinin fazi
        double sigma = 0, N = 0; //standart sapma, aritmetik ortalama

        // KISMI BELIRSIZLIKLER // bolenleri
        double uPo; //k=1
        double uDeltaPmacc; // k = root3
        double uDeltaPRfPow; // k = 2
        double uCFstd; //k = 2
        double urhoStd; //k = 2
        double urhoGB; // k = 2
        double uthetaStd; // k = 2
        double uthetaGB; // k = 2 


        // DUYARLILIK KATSAYISI
        double cPo; 
        double cDeltaPmacc; 
        double cDeltaPRfPow; 
        double cCFstd; 
        double crhoStd; 
        double crhoGB; 
        double cthetaStd; 
        double cthetaGB;

        // KISMI VARYANSLAR
        // KV_i = Ui^2*k^2*Ci^2
        double KV_Po;
        double KV_DeltaPmacc;
        double KV_DeltaPRfPow;
        double KV_CFstd;
        double KV_rhoStd;
        double KV_rhoGB;
        double KV_thetaStd;
        double KV_thetaGB;

        double TV; //Toplam Varyans
        public double PRfAnalitik;
        public double StdUnc; //Standart Belirsizlik
        double ExtUnc; //Genisletilmis Belirsizlik

        // MONTE CARLO SONUCLARI
        public double PRf;
        public double StdPRf;
    
        double trustabilityRatio = 0; // Guvenlilik duzeyi

        List<long> knownFrequencies = new List<long>(); // Bilinen frekanslarin listesi 
        List<double> calFactors = new List<double>(); // Guc algilayicisinin frekanslara gore kalibrasyon faktorleri
        List<double> calFacUncerts = new List<double>(); // Kalibrasyon faktorlerinin belirsizlikleri
        public List<double> measuredPowers = new List<double>(); // Gucmetreden okunan guc degerleri 
        List<double> reflectionMagnitudes = new List<double>(); // Guc algilayicisinin yansima katsayilari 
        List<double> coefficientDegrees = new List<double>(); // Guc algilayicisinin faz acilari (radyan)
        public List<double> PRfs = new List<double>(); // Monte-Carlo uygulanilan PV olcumleri

        string warning;
        public Form3()
        {
            InitializeComponent();
        }

        // Sensor secildiginde olusturulan bilinen degerler listesi
        void MakeKnownList(List<long> kf, List<double> cf, List<double> cfu, List<double> rm, List<double> cd)
        {
            
            // N8481A icin
            if (comboBox3.Text == "N8481A [10 MHz - 18 GHz]")
            {
                kf.Add(10000000); cf.Add(0.9863); cfu.Add(0.0060); rm.Add(0.115); cd.Add(4.9131018444); //10MHz
                kf.Add(30000000); cf.Add(0.9987); cfu.Add(0.0045); rm.Add(0.040); cd.Add(5.1487212934);

                kf.Add(50000000); cf.Add(1.0000); cfu.Add(0.0000); rm.Add(0.025); cd.Add(5.0474921968);
                kf.Add(100000000); cf.Add(1.0007); cfu.Add(0.0041); rm.Add(0.015); cd.Add(5.2482050607);//1MHz
                kf.Add(300000000); cf.Add(0.9997); cfu.Add(0.0052); rm.Add(0.009); cd.Add(5.5728363016);
                kf.Add(500000000); cf.Add(0.9982); cfu.Add(0.0052); rm.Add(0.009); cd.Add(5.5850536064);

                kf.Add(800000000); cf.Add(0.9966); cfu.Add(0.0052); rm.Add(0.009); cd.Add(5.4541539125);
                kf.Add(1000000000); cf.Add(0.9961); cfu.Add(0.0052); rm.Add(0.009); cd.Add(5.3249995478);//1GHz
                kf.Add(1200000000); cf.Add(0.9938); cfu.Add(0.0052); rm.Add(0.009); cd.Add(5.1731559029);

                kf.Add(1500000000); cf.Add(0.9935); cfu.Add(0.0052); rm.Add(0.010); cd.Add(4.9096111859);
                kf.Add(2000000000); cf.Add(0.9925); cfu.Add(0.0052); rm.Add(0.011); cd.Add(4.419173666);
                kf.Add(3000000000); cf.Add(0.9873); cfu.Add(0.0056); rm.Add(0.011); cd.Add(3.1834805556);

                kf.Add(4000000000); cf.Add(0.9803); cfu.Add(0.0056); rm.Add(0.011); cd.Add(2.0053833105);
                kf.Add(5000000000); cf.Add(0.9775); cfu.Add(0.0063); rm.Add(0.009); cd.Add(0.9965830029);
                kf.Add(6000000000); cf.Add(0.9698); cfu.Add(0.0065); rm.Add(0.005); cd.Add(5.6199601914);
                kf.Add(7000000000); cf.Add(0.9693); cfu.Add(0.0065); rm.Add(0.005); cd.Add(1.5917402778);

                kf.Add(8000000000); cf.Add(0.9741); cfu.Add(0.0067); rm.Add(0.005); cd.Add(1.6737707527);
                kf.Add(9000000000); cf.Add(0.9715); cfu.Add(0.0077); rm.Add(0.011); cd.Add(0.7784168464);
                kf.Add(10000000000); cf.Add(0.9742); cfu.Add(0.0078); rm.Add(0.017); cd.Add(6.0562925044);//10GHz
                kf.Add(11000000000); cf.Add(0.9755); cfu.Add(0.0073); rm.Add(0.020); cd.Add(4.9043751981);
                kf.Add(12000000000); cf.Add(0.9804); cfu.Add(0.0077); rm.Add(0.022); cd.Add(3.4417892849);
                kf.Add(12400000000); cf.Add(0.9795); cfu.Add(0.0078); rm.Add(0.024); cd.Add(2.8012534495);
                kf.Add(13000000000); cf.Add(0.9757); cfu.Add(0.0075); rm.Add(0.029); cd.Add(1.8675022996);
                kf.Add(14000000000); cf.Add(0.9771); cfu.Add(0.0073); rm.Add(0.043); cd.Add(0.5305800926);
                kf.Add(15000000000); cf.Add(0.9773); cfu.Add(0.0072); rm.Add(0.054); cd.Add(5.6339228254);
                kf.Add(16000000000); cf.Add(0.9781); cfu.Add(0.0081); rm.Add(0.056); cd.Add(4.4854961776);
                kf.Add(17000000000); cf.Add(0.9778); cfu.Add(0.0086); rm.Add(0.047); cd.Add(3.317870908);
                kf.Add(18000000000); cf.Add(0.9796); cfu.Add(0.0082); rm.Add(0.032); cd.Add(1.8901915799);//18GHz
            }
            else if(comboBox3.Text == "N8482A [100 kHz - 6 GHz]") // 100 kHz - 6 GHz
            {
                kf.Add(100000); cf.Add(0.9673); cfu.Add(0.0065); rm.Add(0.182); cd.Add(4.9253191491);// 0.1 MHz
                kf.Add(300000); cf.Add(0.9950); cfu.Add(0.0065); rm.Add(0.063); cd.Add(4.8153634063);

                kf.Add(500000); cf.Add(0.9990); cfu.Add(0.0065); rm.Add(0.038); cd.Add(4.80663676);
                kf.Add(1000000); cf.Add(1.0005); cfu.Add(0.0065); rm.Add(0.019); cd.Add(4.813618077);
                kf.Add(3000000); cf.Add(1.0005); cfu.Add(0.0064); rm.Add(0.007); cd.Add(4.9043751981);
                kf.Add(5000000); cf.Add(0.9998); cfu.Add(0.0064); rm.Add(0.004); cd.Add(4.9968776485);

                kf.Add(10000000); cf.Add(1.0009); cfu.Add(0.0060); rm.Add(0.002); cd.Add(5.1836278784);
                kf.Add(30000000); cf.Add(1.0003); cfu.Add(0.0045); rm.Add(0.001); cd.Add(5.8695422745);
                kf.Add(50000000); cf.Add(1.0000); cfu.Add(0.0000); rm.Add(0.001); cd.Add(0.1343903524);// 50 MHz

                kf.Add(100000000); cf.Add(0.9995); cfu.Add(0.0041); rm.Add(0.001); cd.Add(0.7801621756);
                kf.Add(300000000); cf.Add(0.9967); cfu.Add(0.0052); rm.Add(0.004); cd.Add(0.7627088831);
                kf.Add(500000000); cf.Add(0.9932); cfu.Add(0.0052); rm.Add(0.005); cd.Add(0.5393067389);

                kf.Add(1000000000); cf.Add(0.9875); cfu.Add(0.0052); rm.Add(0.010); cd.Add(6.1383229793);// 1 GHz
                kf.Add(1500000000); cf.Add(0.9801); cfu.Add(0.0052); rm.Add(0.014); cd.Add(5.3913220594);
                kf.Add(2000000000); cf.Add(0.9739); cfu.Add(0.0052); rm.Add(0.016); cd.Add(4.61988653);
                kf.Add(2500000000); cf.Add(0.9663); cfu.Add(0.0056); rm.Add(0.015); cd.Add(3.8083084279);

                kf.Add(3000000000); cf.Add(0.9632); cfu.Add(0.0056); rm.Add(0.015); cd.Add(3.0054569719);
                kf.Add(3500000000); cf.Add(0.9612); cfu.Add(0.0056); rm.Add(0.015); cd.Add(2.2130774915);
                kf.Add(3700000000); cf.Add(0.9602); cfu.Add(0.0056); rm.Add(0.014); cd.Add(1.9128808602);
                kf.Add(4000000000); cf.Add(0.9558); cfu.Add(0.0056); rm.Add(0.013); cd.Add(1.4416419621);
                kf.Add(4200000000); cf.Add(0.9548); cfu.Add(0.0061); rm.Add(0.012); cd.Add(1.1763519158);
                kf.Add(5000000000); cf.Add(0.9490); cfu.Add(0.0063); rm.Add(0.011); cd.Add(6.2360614174);
                kf.Add(6000000000); cf.Add(0.9368); cfu.Add(0.0065); rm.Add(0.012); cd.Add(4.1102503884);// 6 GHz
            }
            else if(comboBox3.Text == "N8485A [10 MHz - 26.5 GHz]") // 10 MHz - 26.5 GHz
            {
                kf.Add(10000000); cf.Add(0.9856); cfu.Add(0.0090); rm.Add(0.116); cd.Add(4.8153634063);// 10 MHz
                kf.Add(30000000); cf.Add(0.9981); cfu.Add(0.0074); rm.Add(0.040); cd.Add(4.6774823953);

                kf.Add(50000000); cf.Add(1.0000); cfu.Add(0.0000); rm.Add(0.024); cd.Add(4.5867252742);
                kf.Add(100000000); cf.Add(1.0009); cfu.Add(0.0066); rm.Add(0.012); cd.Add(4.3929937273);
                kf.Add(300000000); cf.Add(0.9994); cfu.Add(0.0073); rm.Add(0.004); cd.Add(3.5604716741);
                kf.Add(500000000); cf.Add(0.9979); cfu.Add(0.0075); rm.Add(0.003); cd.Add(2.6057765732);

                kf.Add(800000000); cf.Add(0.9973); cfu.Add(0.0079); rm.Add(0.004); cd.Add(1.6947147037);
                kf.Add(1000000000); cf.Add(0.9967); cfu.Add(0.0075); rm.Add(0.005); cd.Add(1.3264502315); // 1 GHz
                kf.Add(1200000000); cf.Add(0.9957); cfu.Add(0.0079); rm.Add(0.006); cd.Add(1.0314895879);

                kf.Add(1500000000); cf.Add(0.9944); cfu.Add(0.0075); rm.Add(0.008); cd.Add(0.6632251158);
                kf.Add(2000000000); cf.Add(0.9928); cfu.Add(0.0076); rm.Add(0.010); cd.Add(0.1064650844);
                kf.Add(3000000000); cf.Add(0.9881); cfu.Add(0.0078); rm.Add(0.013); cd.Add(5.2516957193);

                kf.Add(4000000000); cf.Add(0.9822); cfu.Add(0.0085); rm.Add(0.017); cd.Add(3.947934768);
                kf.Add(5000000000); cf.Add(0.9786); cfu.Add(0.0088); rm.Add(0.021); cd.Add(2.9530970944); // 5 GHz
                kf.Add(6000000000); cf.Add(0.9751); cfu.Add(0.0095); rm.Add(0.026); cd.Add(1.9408061282);
                kf.Add(7000000000); cf.Add(0.9689); cfu.Add(0.0097); rm.Add(0.026); cd.Add(1.0855947947);

                kf.Add(8000000000); cf.Add(0.9722); cfu.Add(0.0106); rm.Add(0.029); cd.Add(0.2443460953);
                kf.Add(9000000000); cf.Add(0.9747); cfu.Add(0.0117); rm.Add(0.031); cd.Add(5.5571283383);
                kf.Add(10000000000); cf.Add(0.9712); cfu.Add(0.0117); rm.Add(0.031); cd.Add(4.5099307872); // 10 GHz
                kf.Add(11000000000); cf.Add(0.9706); cfu.Add(0.0117); rm.Add(0.034); cd.Add(3.4103733584);
                kf.Add(12000000000); cf.Add(0.9719); cfu.Add(0.0117); rm.Add(0.039); cd.Add(2.3736477827);
                kf.Add(12400000000); cf.Add(0.9715); cfu.Add(0.0117); rm.Add(0.041); cd.Add(1.9757127133);
                kf.Add(13000000000); cf.Add(0.9705); cfu.Add(0.0118); rm.Add(0.045); cd.Add(1.3980087308);
                kf.Add(14000000000); cf.Add(0.9695); cfu.Add(0.0127); rm.Add(0.048); cd.Add(0.44331363);
                kf.Add(15000000000); cf.Add(0.9680); cfu.Add(0.0127); rm.Add(0.048); cd.Add(5.7648225193);
                kf.Add(16000000000); cf.Add(0.9692); cfu.Add(0.0136); rm.Add(0.047); cd.Add(4.8118727477);
                kf.Add(17000000000); cf.Add(0.9708); cfu.Add(0.0137); rm.Add(0.044); cd.Add(3.8275070496);
                kf.Add(18000000000); cf.Add(0.9680); cfu.Add(0.0137); rm.Add(0.040); cd.Add(2.839650693);

                kf.Add(18500000000); cf.Add(0.9694); cfu.Add(0.0160); rm.Add(0.037); cd.Add(2.3701571242);
                kf.Add(19000000000); cf.Add(0.9700); cfu.Add(0.0160); rm.Add(0.032); cd.Add(1.8867009214);
                kf.Add(19500000000); cf.Add(0.9708); cfu.Add(0.0160); rm.Add(0.027); cd.Add(1.3718287921);
                kf.Add(20000000000); cf.Add(0.9716); cfu.Add(0.0160); rm.Add(0.022); cd.Add(0.79587013891); // 20 GHz
                kf.Add(20500000000); cf.Add(0.9726); cfu.Add(0.0159); rm.Add(0.017); cd.Add(0.11344640138);
                kf.Add(21000000000); cf.Add(0.9733); cfu.Add(0.0159); rm.Add(0.013); cd.Add(5.5012778023);
                kf.Add(21500000000); cf.Add(0.9751); cfu.Add(0.0159); rm.Add(0.015); cd.Add(4.4575709096);
                kf.Add(22000000000); cf.Add(0.9755); cfu.Add(0.0159); rm.Add(0.020); cd.Add(2.0001473228);
                kf.Add(22500000000); cf.Add(0.9756); cfu.Add(0.0160); rm.Add(0.027); cd.Add(3.061307508);
                kf.Add(23000000000); cf.Add(0.9768); cfu.Add(0.0160); rm.Add(0.032); cd.Add(2.5673793297);
                kf.Add(23500000000); cf.Add(0.9779); cfu.Add(0.0160); rm.Add(0.036); cd.Add(2.1066124072);
                kf.Add(24000000000); cf.Add(0.9778); cfu.Add(0.0160); rm.Add(0.038); cd.Add(1.6510814724);
                kf.Add(24500000000); cf.Add(0.9788); cfu.Add(0.0160); rm.Add(0.039); cd.Add(1.1920598791);
                kf.Add(25000000000); cf.Add(0.9825); cfu.Add(0.0160); rm.Add(0.037); cd.Add(0.72082098107);
                kf.Add(25500000000); cf.Add(0.9831); cfu.Add(0.0160); rm.Add(0.034); cd.Add(0.19722220548);
                kf.Add(26000000000); cf.Add(0.9854); cfu.Add(0.0160); rm.Add(0.029); cd.Add(5.8782689207);
                kf.Add(26500000000); cf.Add(0.9859); cfu.Add(0.0160); rm.Add(0.026); cd.Add(5.1644292567); // 26.5 GHz
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textUsername_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            knownFrequencies = new List<long>();
            calFactors = new List<double>();
            calFacUncerts = new List<double>();           // hesapla tiklaninca listeler icin olusturulan objeler
            reflectionMagnitudes = new List<double>();
            coefficientDegrees = new List<double>();

            MakeKnownList(knownFrequencies, calFactors, calFacUncerts, reflectionMagnitudes, coefficientDegrees);
            // eksik ya da hatali girildiginde verilen uyari
            label20.Visible = false;

            // veriler gonderilip alindiginda beliren onay isareti (tikler)
            if(textBox1.Text != "" && comboBox1.Text != "")
            {
                pictureBox1.Visible = true;
            }
            if (textBox8.Text != "")
            {
                pictureBox3.Visible = true;
            }
            if (textBox9.Text != "")
            {
                pictureBox9.Visible = true;
            }
            if (textBox2.Text != "")
            {
                pictureBox6.Visible = true;
            }
            if (textBox6.Text != "")
            {
                pictureBox8.Visible = true;
            }
            if (textBox5.Text != "")
            {
                pictureBox4.Visible = true;
            }
            if (textBox7.Text != "")
            {
                pictureBox10.Visible = true;
            }
            if (textBox10.Text != "")
            {
                pictureBox12.Visible = true;
            }
            if (textBox11.Text != "")
            {
                pictureBox13.Visible = true;
            }
            if (textBox20.Text != "")
            {
                pictureBox7.Visible = true;
            }
            if (comboBox2.Text != "")
            {
                pictureBox11.Visible = true;
            }
            if (comboBox3.Text != "")
            {
                pictureBox5.Visible = true;
            }

            try
            {
                measuredPowers = new List<double>();
                get_data_from_xl();  // excele aktarilmis olan guc degerlerinin koda aktarilmasi 
            }
            catch
            {
                warning = "Excel dosyasına erişirken hata oluştu";  
                label20.Visible = true;
                label20.Text = warning;
                return;
            }

            if (measuredPowers.Count < 10) // powermetreden okunan olcum degeri 10dan azsa gelen uyari
            {
                warning = "10'dan fazla örnek sayısı alınız!!!";
                label20.Visible = true;
                label20.Text = warning;
                return;
            }
            else
            {
                if (textBox1.Text.Equals("") ||  textBox9.Text.Equals("") || textBox8.Text.Equals("") || comboBox1.Text == "" ||
                    comboBox2.Text == "" || textBox2.Text.Equals("") || textBox6.Text.Equals("") || textBox6.Text.Equals("") ||
                    textBox7.Text.Equals("") || textBox10.Text.Equals("") || textBox11.Text.Equals("") || textBox20.Text.Equals(""))
                {
                    warning = "Girdi değerleri eksik";
                    label20.Visible = true;    // eksik veri girilirse olusan hata
                    label20.Text = warning;
                    return;

                }
                else
                {
                    // frekans birimlerinin tanimlanmasi
                    
                    if (comboBox1.Text.Equals("Hz")) frekans = Convert.ToDouble(textBox1.Text);
                    else if (comboBox1.Text.Equals("kHz")) frekans = Convert.ToDouble(textBox1.Text) * Math.Pow(10, 3);
                    else if (comboBox1.Text.Equals("MHz")) frekans = Convert.ToDouble(textBox1.Text) * Math.Pow(10, 6);
                    else if (comboBox1.Text.Equals("GHz")) frekans = Convert.ToDouble(textBox1.Text) * Math.Pow(10, 9);


                    WriteCalibrationFactorAndUncertaintyPlus();
        
                    
                    Po = measuredPowers.Average(); // olculen degerlerin ortalamasi
                    DeltaPmacc = Convert.ToDouble(textBox5.Text);
                    DeltaPRfPow = Convert.ToDouble(textBox6.Text);
                    rhoGB = Convert.ToDouble(textBox2.Text);
                    sigma = Math.Sqrt(measuredPowers.Sum(x => Math.Pow(x - Po, 2)) / (measuredPowers.Count - 1)); // standart sapmayi hesapladik
                    thetaGB = Convert.ToDouble(textBox20.Text) * 2 * Math.PI / 360; // gucbolucusunun faz acisi(rad) degisiklik gosterebilir (kullanici tarafindan girilebilir)


                    //KISMI BELIRSIZLIKLERIN HESAPLANMASI
                    uPo = sigma / Math.Sqrt(measuredPowers.Count);
                    uDeltaPmacc = Po * Convert.ToDouble(textBox7.Text);
                    uDeltaPRfPow = Po * Convert.ToDouble(textBox8.Text) / Convert.ToDouble(textBox9.Text);
                    urhoStd = 0.0045; // Sertifikadan alindi (2'ye bolunerek alindi)
                    urhoGB = Convert.ToDouble(textBox10.Text) / 2; // Sertifikadan alindi (2'ye bolunerek alindi)
                    uthetaStd = Math.Asin(urhoStd / rhoStd) * 2 * Math.PI / 360;
                    uthetaGB = Convert.ToDouble(textBox11.Text) * (2 * Math.PI / 360) / 2; // Sertifikadan alindi (2'ye bolunerek alindi)

                    //DUYARLILIK KATSAYILARI HESAPLANMASI
                    cPo = (1 + rhoStd * rhoStd * rhoGB * rhoGB - 2 * rhoGB * rhoStd * Math.Cos(thetaStd + thetaGB)) / CFstd;
                    cDeltaPmacc = (1 + rhoStd * rhoStd * rhoGB * rhoGB - 2 * rhoGB * rhoStd * Math.Cos(thetaStd + thetaGB)) / CFstd;
                    cDeltaPRfPow = (1 + rhoStd * rhoStd * rhoGB * rhoGB - 2 * rhoGB * rhoStd * Math.Cos(thetaStd + thetaGB)) / CFstd;
                    cCFstd = -1 * (Po + DeltaPmacc + DeltaPRfPow) * ((1 + rhoStd * rhoStd * rhoGB * rhoGB - 2 * rhoGB * rhoStd * Math.Cos(thetaStd + thetaGB)) / (CFstd * CFstd));
                    crhoStd = (Po + DeltaPmacc + DeltaPRfPow) * ((2 * rhoStd * rhoGB * rhoGB - 2 * rhoGB * Math.Cos(thetaStd + thetaGB)) / CFstd);
                    crhoGB = (Po + DeltaPmacc + DeltaPRfPow) * ((2 * rhoStd * rhoStd * rhoGB - 2 * rhoStd * Math.Cos(thetaStd + thetaGB)) / CFstd);
                    cthetaStd = (Po + DeltaPmacc + DeltaPRfPow) * ((2 * rhoStd * rhoGB * Math.Sin(thetaStd + thetaGB)) / CFstd);
                    cthetaGB = (Po + DeltaPmacc + DeltaPRfPow) * ((2 * rhoStd * rhoGB * Math.Sin(thetaStd + thetaGB)) / CFstd);


                    //KISMI VARYANSLAR
                    KV_Po = uPo * uPo * cPo * cPo * 1 * 1;
                    KV_DeltaPmacc = uDeltaPmacc * uDeltaPmacc * cDeltaPmacc * cDeltaPmacc * 0.577 * 0.577;
                    KV_DeltaPRfPow = uDeltaPRfPow * uDeltaPRfPow * cDeltaPRfPow * cDeltaPRfPow * 0.5 * 0.5;
                    KV_CFstd = uCFstd * uCFstd * cCFstd * cCFstd * 0.5 * 0.5;
                    KV_rhoStd = urhoStd * urhoStd * crhoStd * crhoStd * 0.5 * 0.5;
                    KV_rhoGB = urhoGB * urhoGB * crhoGB * crhoGB * 0.5 * 0.5;
                    KV_thetaStd = uthetaStd * uthetaStd * cthetaStd * cthetaStd * 0.5 * 0.5;
                    KV_thetaGB = uthetaGB * uthetaGB * cthetaGB * cthetaGB * 0.5 * 0.5;

                    //TOPLAM VARYANS
                    TV = KV_Po + KV_DeltaPmacc + KV_DeltaPRfPow + KV_CFstd + KV_rhoStd + KV_rhoGB + KV_thetaStd + KV_thetaGB;

                    //STANDART BELIRSIZLIK
                    StdUnc = Math.Sqrt(TV);

                    // guvenilirlik duzeyinin belirlenmesi
                    if (comboBox2.Text.Equals("%68.27")) trustabilityRatio = 1;
                    else if (comboBox2.Text.Equals("%90")) trustabilityRatio = 1.645;
                    else if (comboBox2.Text.Equals("%95")) trustabilityRatio = 1.96;
                    else if (comboBox2.Text.Equals("%95.45 ( k=2 )")) trustabilityRatio = 2;
                    else if (comboBox2.Text.Equals("%99")) trustabilityRatio = 2.576;
                    else if (comboBox2.Text.Equals("%99.73")) trustabilityRatio = 3;

                    //GENISLETILMIS BELIRSIZLIK
                    ExtUnc = trustabilityRatio * StdUnc;

                    textBox15.Text = Math.Round(TV, 5).ToString(); // Total varyans
                    //System.Threading.Thread.Sleep(200);
                    textBox16.Text = Math.Round(StdUnc, 5).ToString(); // Standart belirsizlik
                    //System.Threading.Thread.Sleep(200);
                    textBox17.Text = Math.Round(ExtUnc, 5).ToString(); // Genisletilmis belirsizlik
                    //System.Threading.Thread.Sleep(200);
                    textBox12.Text = Math.Round(Po, 5).ToString(); // Ortalama guc
                    //System.Threading.Thread.Sleep(200);
                    textBox13.Text = Math.Round(sigma, 5).ToString(); // Standart sapma
                    //System.Threading.Thread.Sleep(200);
                    textBox14.Text = Math.Round(CFstd, 5).ToString(); //Kalibrasyon faktoru
                    //System.Threading.Thread.Sleep(200);
                    PRfAnalitik = ((Po + DeltaPmacc + DeltaPRfPow) * (1 + rhoStd * rhoStd * rhoGB * rhoGB -
                        2 * rhoStd * rhoGB * Math.Cos(thetaStd + thetaGB)) / CFstd); // PRfAnalitik hesaplandi
                    textBox18.Text = Math.Round(PRfAnalitik, 5).ToString();
                    //System.Threading.Thread.Sleep(200);
                    textBox19.Text = Math.Round(ExtUnc, 5).ToString(); // Genisletilmis belirsizlik
                    //System.Threading.Thread.Sleep(200);

                    // Monte-Carlo uygulanmasi
                    Random r = new Random();
                    PRfs.Clear();
                    for (int i = 0; i < measuredPowers.Count; i++)
                    {
                        double DeltaPmacctemp; // Temp Error of powermeter accuracy
                        double DeltaPRfPowtemp; // Temp Guc olcerin referans rf guc olcumunden gelen belirsizlik yeni olusturulan
                        double CFstdtemp; // Temp Guc algilayicinin kalibrasyon faktoru
                        double rhoGBtemp; // Temp Guc bolucu yansima katsayisi,
                        double rhoStdtemp; // Temp Guc algilayicinin yansima katsayisi
                        double thetaStdtemp; // Temp Guc algilayicinin yansima katsayisinin fazi
                        double thetaGBtemp; // Temp Guc bolucusunun esdeger yansima katsayisinin fazi

                        double min = DeltaPmacc - Math.Sqrt(uDeltaPmacc * uDeltaPmacc * 12) / 2;
                        double max = DeltaPmacc + Math.Sqrt(uDeltaPmacc * uDeltaPmacc * 12) / 2; // diktorgen dagilimin min ve max degerleri
                        DeltaPmacctemp = r.NextDouble() * (max - min) + min; // 0 ve 1 arasi dikdortgen dagilimli random deger olustulup min ve max arasi bir degeri cekilmesi

                        double u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        double u2 = 1.0 - r.NextDouble();
                        double randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        DeltaPRfPowtemp = DeltaPRfPow + uDeltaPRfPow * randStdNormal; //random normal(mean,stdDev^2)

                        u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        u2 = 1.0 - r.NextDouble();
                        randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        CFstdtemp = CFstd + uCFstd * randStdNormal; //random normal(mean,stdDev^2)
                    
                        u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        u2 = 1.0 - r.NextDouble();
                        randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        rhoGBtemp = rhoGB + urhoGB * randStdNormal; //random normal(mean,stdDev^2)

                        u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        u2 = 1.0 - r.NextDouble();
                        randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        rhoStdtemp = rhoStd + urhoStd * randStdNormal; //random normal(mean,stdDev^2)

                        u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        u2 = 1.0 - r.NextDouble();
                        randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        thetaStdtemp = thetaStd + uthetaStd * randStdNormal; //random normal(mean,stdDev^2)

                        u1 = 1.0 - r.NextDouble(); //uniform(0,1] random doubles
                        u2 = 1.0 - r.NextDouble();
                        randStdNormal = Math.Sqrt(-2.0 * Math.Log(u1)) * Math.Sin(2.0 * Math.PI * u2); //random normal(0,1)
                        thetaGBtemp = thetaGB + uthetaGB * randStdNormal; //random normal(mean,stdDev^2)

                        PRfs.Add((measuredPowers[i] + DeltaPmacctemp + DeltaPRfPowtemp) * (1 + rhoStdtemp * rhoStdtemp * rhoGBtemp * rhoGBtemp -
                        2 * rhoStdtemp * rhoGBtemp * Math.Cos(thetaStdtemp + thetaGBtemp)) / CFstdtemp); // denklemde yer alan parametrelerin dagilimlari goz onune 
                        // alinarak her bir veri bir-biriyle isleme sokulup bulunan degerler listeye atilmistir
                    }

                    PRf = PRfs.Average(); // listenin ortalamasi


                    
                    // Perform the Sum of (value-avg)_2_2.      
                    double sum = PRfs.Sum(d => Math.Pow(d - PRf, 2)); // Standart sapma hesaplamasi icin formul uygulanmistir

                    // Put it all together.      
                    StdPRf = Math.Sqrt((sum) / (PRfs.Count() - 1)); // devami

                    try
                    {
                        write_to_xl(); // excele yazmak icin kullanilan fonksyon
                    }
                    catch
                    {
                        label20.Visible = true;
                        label20.Text = "Excel dosyasına erişim sağlanamadı, dosyayı kapatınız!";
                        return;
                    }
                    
                }
            }
        }
        // Excelden veri cekilmek icin kullanilmistir
        void get_data_from_xl()
        {
            var mySheet = Path.Combine(Directory.GetCurrentDirectory(), "Book1.xls");
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;                      
            object misValue = System.Reflection.Missing.Value;
            Excel.Workbook xlWorkbook;

            try
            {
                xlWorkbook = xlApp.Workbooks.Open(mySheet);
            }
            catch
            {
                warning = "Excel dosyasına yazdırılamadı";
                label20.Visible = true;
                label20.Text = warning;
                return;
            }
            Excel._Worksheet workSheet = xlApp.ActiveSheet;

            int row_num = int.Parse(workSheet.Cells[3, 2].Value.ToString()); // olculen degerlerin sayisi aliniyor

            for(int i = 5; i < row_num + 5; i++) // olculen degerler measuredPowers listesine atiliyor
            {
                measuredPowers.Add(double.Parse(workSheet.Cells[i, 2].Value.ToString()));
            }

            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();
            Marshal.ReleaseComObject(workSheet);  //excel dosyasinin kapatilmasi
            Marshal.ReleaseComObject(xlWorkbook);
        }

        void write_to_xl()
        {
            var mySheet = Path.Combine(Directory.GetCurrentDirectory(), "Book1.xls");
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;


            if (xlApp == null)
            {
                return;
            }

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(mySheet);
            Excel._Worksheet workSheet = xlApp.ActiveSheet;
            object misValue = System.Reflection.Missing.Value;

            workSheet.Columns["J:U"].Delete(); // onceki hesaplamalarda kalan monte-carlo girdilerinin silinmesi

            workSheet.Cells[3, 11] = "Monte Carlo Sonuçları";
            for (int i = 1; i <= PRfs.Count; i++)
            {
                workSheet.Cells[i + 4, 11] = "Simülasyon " + i.ToString();
                workSheet.Cells[i + 4, 12] = PRfs[i - 1];                   // Monte Carlodan bulunan degerlerin excele aktarilmasi
                workSheet.Cells[i + 4, 13] = "mW";
            }

            workSheet.Cells[5, 15] = "Ortalama Güç";
            workSheet.Cells[5, 16] = PRf;
            workSheet.Cells[5, 17] = "mW";

            workSheet.Cells[6, 15] = "Standart Sapma";
            workSheet.Cells[6, 16] = StdPRf;

            workSheet.Cells[7, 15] = "Minimum";
            workSheet.Cells[7, 16] = PRfs.Min();
            workSheet.Cells[7, 17] = "mW";

            workSheet.Cells[8, 15] = "Maximum";
            workSheet.Cells[8, 16] = PRfs.Max();
            workSheet.Cells[8, 17] = "mW";

            xlApp.DisplayAlerts = false;
            workSheet.Columns.AutoFit();
            workSheet.Rows.AutoFit();
            
            xlWorkbook.SaveAs(mySheet);
            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();                                   // Excel dosyasinin kapatilmasi ve save edilmesi
            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(xlWorkbook);
        }

        bool WriteCalibrationFactorAndUncertaintyPlus()
        {
            double cf = 0;
            double cfu = 0;
            for (int i = 0; i < knownFrequencies.Count; i++)
            {
                if (frekans == knownFrequencies[i]) // girilien frekans bilinen degerler listesinde varsa
                {
                    cf = calFactors[i];
                    cfu = calFacUncerts[i];
                    
                    CFstd = cf;
                    uCFstd = cfu / 2;
                    thetaStd = coefficientDegrees[i]; // guc algilayicisinin faz acisi(rad) degisiklik gosterebilir
                    rhoStd = reflectionMagnitudes[i];
                    textBox14.Text = CFstd.ToString();
                    return true;
                }
            }

            for (int i = 0; i < knownFrequencies.Count - 1; i++) // frekans bilinen degerler listesinde yoksa interpolasyon fonksyonu kullanildi
            {
                // Frekans araligini bulma
                if (frekans > knownFrequencies[i] && frekans < knownFrequencies[i + 1])
                {
                    cf = Interpolate(knownFrequencies[i], knownFrequencies[i + 1], calFactors[i], calFactors[i + 1], frekans);
                    cfu = Interpolate(knownFrequencies[i], knownFrequencies[i + 1], calFacUncerts[i], calFacUncerts[i + 1], frekans);
                    uCFstd = cfu / 2;
                    CFstd = cf;
                    rhoStd = Interpolate(knownFrequencies[i], knownFrequencies[i + 1], reflectionMagnitudes[i], reflectionMagnitudes[i + 1], frekans);
                    thetaStd = Interpolate(knownFrequencies[i], knownFrequencies[i + 1], coefficientDegrees[i], coefficientDegrees[i + 1], frekans);
                    textBox14.Text = CFstd.ToString("0.00000E00");
                    return true;
                }
            }
            CFstd = cf;
            textBox14.Text = CFstd.ToString();
            return true;
        }

        // interpolasyon fonksyonu
        double Interpolate(long x1, long x2, double y1, double y2, double f)
        {
            double result = 0;
            double w1 = Convert.ToDouble(x1), w2 = Convert.ToDouble(x2);
            if (y2 >= y1)
            {
                result = y1 + (f - w1) * (y2 - y1) / (w2 - w1);
            }
            else                                                      // ara frekanslarin kalibrasyon degerleri bulundu
            {
                result = y1 - (f - w1) * (y1 - y2) / (w2 - w1);
            }
            return result;
        }

        private void sPanel3_Paint(object sender, PaintEventArgs e)
        {

        }
        

        private void sPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void sPanel2_Paint(object sender, PaintEventArgs e)
        {

        }


        private void label10_Click(object sender, EventArgs e)
        {

        }


        private void label6_Click(object sender, EventArgs e)
        {

        }


        private void label11_Click(object sender, EventArgs e)
        {

        }

        private void textBox15_Click(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void textBox18_Click(object sender, EventArgs e)
        {

        }

        private void textBox19_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        // Sil butonu
        private void button4_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox6.Text = "";
            textBox5.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox20.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";

            pictureBox1.Visible = false;
            pictureBox5.Visible = false;
            pictureBox6.Visible = false;
            pictureBox8.Visible = false;
            pictureBox4.Visible = false;
            pictureBox10.Visible = false;
            pictureBox3.Visible = false;
            pictureBox9.Visible = false;
            pictureBox12.Visible = false;
            pictureBox13.Visible = false;
            pictureBox7.Visible = false;
            pictureBox11.Visible = false;

            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";

        }
    }
}
