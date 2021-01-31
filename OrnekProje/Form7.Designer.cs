namespace OrnekProje
{
    partial class Form7
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series3 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            System.Windows.Forms.DataVisualization.Charting.Series series5 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.chart1 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.label14 = new System.Windows.Forms.Label();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.chart2 = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.button1 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox3 = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.textBox1 = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.textBox2 = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox4 = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.textBox5 = new MaterialSkin.Controls.MaterialSingleLineTextField();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).BeginInit();
            this.SuspendLayout();
            // 
            // chart1
            // 
            chartArea1.Name = "ChartArea1";
            this.chart1.ChartAreas.Add(chartArea1);
            this.chart1.Dock = System.Windows.Forms.DockStyle.Fill;
            legend1.Name = "Legend1";
            this.chart1.Legends.Add(legend1);
            this.chart1.Location = new System.Drawing.Point(0, 0);
            this.chart1.Name = "chart1";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.chart1.Series.Add(series1);
            this.chart1.Size = new System.Drawing.Size(1329, 784);
            this.chart1.TabIndex = 1;
            this.chart1.Text = "chart1";
            this.chart1.Visible = false;
            // 
            // label14
            // 
            this.label14.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label14.AutoSize = true;
            this.label14.BackColor = System.Drawing.Color.White;
            this.label14.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label14.Location = new System.Drawing.Point(476, 750);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(135, 18);
            this.label14.TabIndex = 147;
            this.label14.Text = "Güvenilirlik düzeyi:";
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "%0",
            "%68.27",
            "%90",
            "%95",
            "%95.45",
            "%99",
            "%99.73"});
            this.comboBox1.Location = new System.Drawing.Point(617, 747);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(137, 21);
            this.comboBox1.TabIndex = 146;
            // 
            // chart2
            // 
            chartArea2.Name = "ChartArea1";
            this.chart2.ChartAreas.Add(chartArea2);
            this.chart2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chart2.IsSoftShadows = false;
            legend2.Name = "Legend1";
            this.chart2.Legends.Add(legend2);
            this.chart2.Location = new System.Drawing.Point(0, 0);
            this.chart2.Name = "chart2";
            this.chart2.Palette = System.Windows.Forms.DataVisualization.Charting.ChartColorPalette.Excel;
            series2.ChartArea = "ChartArea1";
            series2.Color = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(119)))), ((int)(((byte)(197)))));
            series2.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            series2.LabelBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(119)))), ((int)(((byte)(197)))));
            series2.LabelForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(119)))), ((int)(((byte)(197)))));
            series2.Legend = "Legend1";
            series2.Name = "Okunan Değerler";
            series3.ChartArea = "ChartArea1";
            series3.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.SplineArea;
            series3.Color = System.Drawing.Color.Red;
            series3.Legend = "Legend1";
            series3.Name = "Güvenilirlik Düzeyi";
            series4.ChartArea = "ChartArea1";
            series4.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series4.Color = System.Drawing.Color.Black;
            series4.CustomProperties = "LabelStyle=Left";
            series4.IsValueShownAsLabel = true;
            series4.IsVisibleInLegend = false;
            series4.Label = "#VALX";
            series4.Legend = "Legend1";
            series4.Name = "Left";
            series5.ChartArea = "ChartArea1";
            series5.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            series5.Color = System.Drawing.Color.Black;
            series5.CustomProperties = "LabelStyle=Right";
            series5.IsValueShownAsLabel = true;
            series5.IsVisibleInLegend = false;
            series5.Label = "#VALX";
            series5.Legend = "Legend1";
            series5.Name = "Right";
            this.chart2.Series.Add(series2);
            this.chart2.Series.Add(series3);
            this.chart2.Series.Add(series4);
            this.chart2.Series.Add(series5);
            this.chart2.Size = new System.Drawing.Size(1329, 784);
            this.chart2.TabIndex = 145;
            this.chart2.Text = "chart2";
            this.chart2.Visible = false;
            this.chart2.Click += new System.EventHandler(this.chart2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(119)))), ((int)(((byte)(197)))));
            this.button1.FlatAppearance.BorderSize = 0;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.Control;
            this.button1.Location = new System.Drawing.Point(1187, 727);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(130, 45);
            this.button1.TabIndex = 148;
            this.button1.Text = "Başla";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(47, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 16);
            this.label2.TabIndex = 152;
            this.label2.Text = "Sıklık [Kere]";
            this.label2.Visible = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(1073, 750);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 16);
            this.label1.TabIndex = 153;
            this.label1.Text = "PRf [mW]";
            this.label1.Visible = false;
            // 
            // textBox3
            // 
            this.textBox3.Depth = 0;
            this.textBox3.Hint = "";
            this.textBox3.Location = new System.Drawing.Point(1241, 220);
            this.textBox3.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox3.Name = "textBox3";
            this.textBox3.PasswordChar = '\0';
            this.textBox3.SelectedText = "";
            this.textBox3.SelectionLength = 0;
            this.textBox3.SelectionStart = 0;
            this.textBox3.Size = new System.Drawing.Size(76, 23);
            this.textBox3.TabIndex = 163;
            this.textBox3.UseSystemPasswordChar = false;
            this.textBox3.Visible = false;
            // 
            // textBox1
            // 
            this.textBox1.Depth = 0;
            this.textBox1.Hint = "";
            this.textBox1.Location = new System.Drawing.Point(1241, 177);
            this.textBox1.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox1.Name = "textBox1";
            this.textBox1.PasswordChar = '\0';
            this.textBox1.SelectedText = "";
            this.textBox1.SelectionLength = 0;
            this.textBox1.SelectionStart = 0;
            this.textBox1.Size = new System.Drawing.Size(76, 23);
            this.textBox1.TabIndex = 162;
            this.textBox1.UseSystemPasswordChar = false;
            this.textBox1.Visible = false;
            // 
            // textBox2
            // 
            this.textBox2.Depth = 0;
            this.textBox2.Hint = "";
            this.textBox2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.textBox2.Location = new System.Drawing.Point(1241, 198);
            this.textBox2.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox2.Name = "textBox2";
            this.textBox2.PasswordChar = '\0';
            this.textBox2.SelectedText = "";
            this.textBox2.SelectionLength = 0;
            this.textBox2.SelectionStart = 0;
            this.textBox2.Size = new System.Drawing.Size(76, 23);
            this.textBox2.TabIndex = 161;
            this.textBox2.UseSystemPasswordChar = false;
            this.textBox2.Visible = false;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 8.75F);
            this.label5.Location = new System.Drawing.Point(1101, 228);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(96, 15);
            this.label5.TabIndex = 160;
            this.label5.Text = "Standart Sapma";
            this.label5.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 8.75F);
            this.label4.Location = new System.Drawing.Point(1101, 206);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(135, 15);
            this.label4.TabIndex = 159;
            this.label4.Text = "Genişletilmiş Belirsizlik";
            this.label4.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 8.75F);
            this.label3.Location = new System.Drawing.Point(1101, 185);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 15);
            this.label3.TabIndex = 158;
            this.label3.Text = "Ortalama";
            this.label3.Visible = false;
            // 
            // textBox4
            // 
            this.textBox4.Depth = 0;
            this.textBox4.Hint = "";
            this.textBox4.Location = new System.Drawing.Point(1241, 243);
            this.textBox4.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox4.Name = "textBox4";
            this.textBox4.PasswordChar = '\0';
            this.textBox4.SelectedText = "";
            this.textBox4.SelectionLength = 0;
            this.textBox4.SelectionStart = 0;
            this.textBox4.Size = new System.Drawing.Size(76, 23);
            this.textBox4.TabIndex = 173;
            this.textBox4.UseSystemPasswordChar = false;
            this.textBox4.Visible = false;
            // 
            // textBox5
            // 
            this.textBox5.Depth = 0;
            this.textBox5.Hint = "";
            this.textBox5.Location = new System.Drawing.Point(1241, 266);
            this.textBox5.MouseState = MaterialSkin.MouseState.HOVER;
            this.textBox5.Name = "textBox5";
            this.textBox5.PasswordChar = '\0';
            this.textBox5.SelectedText = "";
            this.textBox5.SelectionLength = 0;
            this.textBox5.SelectionStart = 0;
            this.textBox5.Size = new System.Drawing.Size(76, 23);
            this.textBox5.TabIndex = 172;
            this.textBox5.UseSystemPasswordChar = false;
            this.textBox5.Visible = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Arial", 8.75F);
            this.label7.Location = new System.Drawing.Point(1101, 274);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(68, 15);
            this.label7.TabIndex = 171;
            this.label7.Text = "Maksimum";
            this.label7.Visible = false;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 8.75F);
            this.label6.Location = new System.Drawing.Point(1101, 251);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 15);
            this.label6.TabIndex = 170;
            this.label6.Text = "Minimim";
            this.label6.Visible = false;
            // 
            // Form7
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(1329, 784);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.chart2);
            this.Controls.Add(this.chart1);
            this.Name = "Form7";
            this.Text = "Otomatik RF Güç Ölçümü";
            this.Load += new System.EventHandler(this.Form7_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chart1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chart2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataVisualization.Charting.Chart chart1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.DataVisualization.Charting.Chart chart2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox3;
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox1;
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox4;
        private MaterialSkin.Controls.MaterialSingleLineTextField textBox5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
    }
}