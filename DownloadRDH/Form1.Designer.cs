namespace DownloadCompass
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button3 = new System.Windows.Forms.Button();
            this.DownRDH = new System.Windows.Forms.Button();
            this.mes = new System.Windows.Forms.ComboBox();
            this.ano = new System.Windows.Forms.NumericUpDown();
            this.button5 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.DownAcomph = new System.Windows.Forms.Button();
            this.dia = new System.Windows.Forms.ComboBox();
            this.DownChuvaVazao = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.weol = new System.Windows.Forms.Button();
            this.Carga_Hr = new System.Windows.Forms.Button();
            this.bt_cargaDiaria = new System.Windows.Forms.Button();
            this.DownDessem = new System.Windows.Forms.Button();
            this.bt_IPDO = new System.Windows.Forms.Button();
            this.vazObsDown = new System.Windows.Forms.Button();
            this.DownGifEuro = new System.Windows.Forms.Button();
            this.DownECMWF = new System.Windows.Forms.Button();
            this.bt_NOA = new System.Windows.Forms.Button();
            this.bt_previs = new System.Windows.Forms.Button();
            this.bt_Sat = new System.Windows.Forms.Button();
            this.DownGevazp = new System.Windows.Forms.Button();
            this.EntradaSaidaPrevivaz = new System.Windows.Forms.Button();
            this.DownTemp = new System.Windows.Forms.Button();
            this.DownVE = new System.Windows.Forms.Button();
            this.DownCFS = new System.Windows.Forms.Button();
            this.DownPmoNewave = new System.Windows.Forms.Button();
            this.DownPmoDecomp = new System.Windows.Forms.Button();
            this.DownNoticias = new System.Windows.Forms.Button();
            this.DownSemanal = new System.Windows.Forms.Button();
            this.DownMensal = new System.Windows.Forms.Button();
            this.DownGifObs = new System.Windows.Forms.Button();
            this.DownGifsGefs = new System.Windows.Forms.Button();
            this.DownGifsEta = new System.Windows.Forms.Button();
            this.DownGefs = new System.Windows.Forms.Button();
            this.DownEta = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ano)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button3
            // 
            this.button3.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.button3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(6, 12);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(150, 23);
            this.button3.TabIndex = 4;
            this.button3.Text = "Autentificação";
            this.button3.UseVisualStyleBackColor = false;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // DownRDH
            // 
            this.DownRDH.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownRDH.Enabled = false;
            this.DownRDH.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownRDH.Location = new System.Drawing.Point(2, 20);
            this.DownRDH.Name = "DownRDH";
            this.DownRDH.Size = new System.Drawing.Size(71, 24);
            this.DownRDH.TabIndex = 5;
            this.DownRDH.Text = "RDH";
            this.DownRDH.UseVisualStyleBackColor = false;
            this.DownRDH.Click += new System.EventHandler(this.DownRDH_Click);
            // 
            // mes
            // 
            this.mes.FormattingEnabled = true;
            this.mes.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12"});
            this.mes.Location = new System.Drawing.Point(82, 41);
            this.mes.Name = "mes";
            this.mes.Size = new System.Drawing.Size(74, 21);
            this.mes.TabIndex = 8;
            // 
            // ano
            // 
            this.ano.Location = new System.Drawing.Point(6, 68);
            this.ano.Maximum = new decimal(new int[] {
            2023,
            0,
            0,
            0});
            this.ano.Minimum = new decimal(new int[] {
            2015,
            0,
            0,
            0});
            this.ano.Name = "ano";
            this.ano.Size = new System.Drawing.Size(150, 20);
            this.ano.TabIndex = 9;
            this.ano.Value = new decimal(new int[] {
            2015,
            0,
            0,
            0});
            // 
            // button5
            // 
            this.button5.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button5.Location = new System.Drawing.Point(3, 611);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(150, 23);
            this.button5.TabIndex = 11;
            this.button5.Text = "Complete Run";
            this.button5.UseVisualStyleBackColor = false;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 428);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 12;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.webBrowser1.Location = new System.Drawing.Point(162, 12);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(492, 626);
            this.webBrowser1.TabIndex = 15;
            // 
            // DownAcomph
            // 
            this.DownAcomph.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownAcomph.Enabled = false;
            this.DownAcomph.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownAcomph.Location = new System.Drawing.Point(75, 20);
            this.DownAcomph.Name = "DownAcomph";
            this.DownAcomph.Size = new System.Drawing.Size(71, 24);
            this.DownAcomph.TabIndex = 16;
            this.DownAcomph.Text = "Acomph";
            this.DownAcomph.UseVisualStyleBackColor = false;
            this.DownAcomph.Click += new System.EventHandler(this.DownAcomph_Click);
            // 
            // dia
            // 
            this.dia.FormattingEnabled = true;
            this.dia.ItemHeight = 13;
            this.dia.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6",
            "7",
            "8",
            "9",
            "10",
            "11",
            "12",
            "13",
            "14",
            "15",
            "16",
            "17",
            "18",
            "19",
            "20",
            "21",
            "22",
            "23",
            "24",
            "25",
            "26",
            "27",
            "28",
            "29",
            "30",
            "31"});
            this.dia.Location = new System.Drawing.Point(6, 41);
            this.dia.Name = "dia";
            this.dia.Size = new System.Drawing.Size(75, 21);
            this.dia.TabIndex = 19;
            // 
            // DownChuvaVazao
            // 
            this.DownChuvaVazao.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownChuvaVazao.Enabled = false;
            this.DownChuvaVazao.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownChuvaVazao.Location = new System.Drawing.Point(2, 47);
            this.DownChuvaVazao.Name = "DownChuvaVazao";
            this.DownChuvaVazao.Size = new System.Drawing.Size(144, 24);
            this.DownChuvaVazao.TabIndex = 20;
            this.DownChuvaVazao.Text = "Modelos Chuva Vazão";
            this.DownChuvaVazao.UseVisualStyleBackColor = false;
            this.DownChuvaVazao.Click += new System.EventHandler(this.DownChuvaVazao_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.weol);
            this.panel1.Controls.Add(this.Carga_Hr);
            this.panel1.Controls.Add(this.bt_cargaDiaria);
            this.panel1.Controls.Add(this.DownDessem);
            this.panel1.Controls.Add(this.bt_IPDO);
            this.panel1.Controls.Add(this.vazObsDown);
            this.panel1.Controls.Add(this.DownGifEuro);
            this.panel1.Controls.Add(this.DownECMWF);
            this.panel1.Controls.Add(this.bt_NOA);
            this.panel1.Controls.Add(this.bt_previs);
            this.panel1.Controls.Add(this.bt_Sat);
            this.panel1.Controls.Add(this.DownGevazp);
            this.panel1.Controls.Add(this.EntradaSaidaPrevivaz);
            this.panel1.Controls.Add(this.DownTemp);
            this.panel1.Controls.Add(this.DownVE);
            this.panel1.Controls.Add(this.DownCFS);
            this.panel1.Controls.Add(this.DownPmoNewave);
            this.panel1.Controls.Add(this.DownPmoDecomp);
            this.panel1.Controls.Add(this.DownNoticias);
            this.panel1.Controls.Add(this.DownSemanal);
            this.panel1.Controls.Add(this.DownMensal);
            this.panel1.Controls.Add(this.DownGifObs);
            this.panel1.Controls.Add(this.DownGifsGefs);
            this.panel1.Controls.Add(this.DownGifsEta);
            this.panel1.Controls.Add(this.DownGefs);
            this.panel1.Controls.Add(this.DownEta);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.DownRDH);
            this.panel1.Controls.Add(this.DownChuvaVazao);
            this.panel1.Controls.Add(this.DownAcomph);
            this.panel1.Location = new System.Drawing.Point(6, 92);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(150, 513);
            this.panel1.TabIndex = 21;
            this.panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.Panel1_Paint);
            // 
            // weol
            // 
            this.weol.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.weol.Enabled = false;
            this.weol.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.weol.Location = new System.Drawing.Point(75, 479);
            this.weol.Name = "weol";
            this.weol.Size = new System.Drawing.Size(71, 24);
            this.weol.TabIndex = 47;
            this.weol.Text = "WEOL";
            this.weol.UseVisualStyleBackColor = false;
            this.weol.Click += new System.EventHandler(this.weol_Click);
            // 
            // Carga_Hr
            // 
            this.Carga_Hr.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.Carga_Hr.Enabled = false;
            this.Carga_Hr.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Carga_Hr.Location = new System.Drawing.Point(3, 479);
            this.Carga_Hr.Name = "Carga_Hr";
            this.Carga_Hr.Size = new System.Drawing.Size(71, 24);
            this.Carga_Hr.TabIndex = 46;
            this.Carga_Hr.Text = "Carga Hr";
            this.Carga_Hr.UseVisualStyleBackColor = false;
            this.Carga_Hr.Click += new System.EventHandler(this.Carga_Hr_Click);
            // 
            // bt_cargaDiaria
            // 
            this.bt_cargaDiaria.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_cargaDiaria.Enabled = false;
            this.bt_cargaDiaria.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_cargaDiaria.Location = new System.Drawing.Point(75, 449);
            this.bt_cargaDiaria.Name = "bt_cargaDiaria";
            this.bt_cargaDiaria.Size = new System.Drawing.Size(71, 24);
            this.bt_cargaDiaria.TabIndex = 45;
            this.bt_cargaDiaria.Text = "Diario";
            this.bt_cargaDiaria.UseVisualStyleBackColor = false;
            this.bt_cargaDiaria.Click += new System.EventHandler(this.bt_cargaDiaria_Click);
            // 
            // DownDessem
            // 
            this.DownDessem.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownDessem.Enabled = false;
            this.DownDessem.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownDessem.Location = new System.Drawing.Point(2, 449);
            this.DownDessem.Name = "DownDessem";
            this.DownDessem.Size = new System.Drawing.Size(71, 24);
            this.DownDessem.TabIndex = 44;
            this.DownDessem.Text = "Dessem";
            this.DownDessem.UseVisualStyleBackColor = false;
            this.DownDessem.Click += new System.EventHandler(this.DownDessem_Click);
            // 
            // bt_IPDO
            // 
            this.bt_IPDO.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_IPDO.Enabled = false;
            this.bt_IPDO.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_IPDO.Location = new System.Drawing.Point(75, 419);
            this.bt_IPDO.Name = "bt_IPDO";
            this.bt_IPDO.Size = new System.Drawing.Size(71, 24);
            this.bt_IPDO.TabIndex = 43;
            this.bt_IPDO.Text = "IPDO";
            this.bt_IPDO.UseVisualStyleBackColor = false;
            this.bt_IPDO.Click += new System.EventHandler(this.bt_IPDO_Click);
            // 
            // vazObsDown
            // 
            this.vazObsDown.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.vazObsDown.Enabled = false;
            this.vazObsDown.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.vazObsDown.Location = new System.Drawing.Point(2, 419);
            this.vazObsDown.Name = "vazObsDown";
            this.vazObsDown.Size = new System.Drawing.Size(71, 24);
            this.vazObsDown.TabIndex = 42;
            this.vazObsDown.Text = "Vaz-Obs";
            this.vazObsDown.UseVisualStyleBackColor = false;
            this.vazObsDown.Click += new System.EventHandler(this.vazObsDown_Click);
            // 
            // DownGifEuro
            // 
            this.DownGifEuro.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGifEuro.Enabled = false;
            this.DownGifEuro.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGifEuro.Location = new System.Drawing.Point(75, 364);
            this.DownGifEuro.Name = "DownGifEuro";
            this.DownGifEuro.Size = new System.Drawing.Size(71, 24);
            this.DownGifEuro.TabIndex = 41;
            this.DownGifEuro.Text = "Gifs EURO";
            this.DownGifEuro.UseVisualStyleBackColor = false;
            this.DownGifEuro.Click += new System.EventHandler(this.DownGifEuro_Click);
            // 
            // DownECMWF
            // 
            this.DownECMWF.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownECMWF.Enabled = false;
            this.DownECMWF.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownECMWF.Location = new System.Drawing.Point(2, 364);
            this.DownECMWF.Name = "DownECMWF";
            this.DownECMWF.Size = new System.Drawing.Size(71, 24);
            this.DownECMWF.TabIndex = 40;
            this.DownECMWF.Text = "ECMWF";
            this.DownECMWF.UseVisualStyleBackColor = false;
            this.DownECMWF.Click += new System.EventHandler(this.DownECMWF_Click);
            // 
            // bt_NOA
            // 
            this.bt_NOA.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_NOA.Enabled = false;
            this.bt_NOA.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_NOA.Location = new System.Drawing.Point(75, 391);
            this.bt_NOA.Name = "bt_NOA";
            this.bt_NOA.Size = new System.Drawing.Size(71, 24);
            this.bt_NOA.TabIndex = 39;
            this.bt_NOA.Text = "NOA";
            this.bt_NOA.UseVisualStyleBackColor = false;
            this.bt_NOA.Click += new System.EventHandler(this.bt_NOA_Click);
            // 
            // bt_previs
            // 
            this.bt_previs.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_previs.Enabled = false;
            this.bt_previs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_previs.Location = new System.Drawing.Point(2, 338);
            this.bt_previs.Name = "bt_previs";
            this.bt_previs.Size = new System.Drawing.Size(144, 23);
            this.bt_previs.TabIndex = 38;
            this.bt_previs.Text = "Previs Precip";
            this.bt_previs.UseVisualStyleBackColor = false;
            this.bt_previs.Click += new System.EventHandler(this.bt_previs_Click);
            // 
            // bt_Sat
            // 
            this.bt_Sat.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_Sat.Enabled = false;
            this.bt_Sat.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Sat.Location = new System.Drawing.Point(2, 392);
            this.bt_Sat.Name = "bt_Sat";
            this.bt_Sat.Size = new System.Drawing.Size(71, 23);
            this.bt_Sat.TabIndex = 37;
            this.bt_Sat.Text = "Satelite";
            this.bt_Sat.UseVisualStyleBackColor = false;
            this.bt_Sat.Click += new System.EventHandler(this.bt_Sat_Click);
            // 
            // DownGevazp
            // 
            this.DownGevazp.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGevazp.Enabled = false;
            this.DownGevazp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGevazp.Location = new System.Drawing.Point(2, 312);
            this.DownGevazp.Name = "DownGevazp";
            this.DownGevazp.Size = new System.Drawing.Size(144, 24);
            this.DownGevazp.TabIndex = 36;
            this.DownGevazp.Text = "Gevazp";
            this.DownGevazp.UseVisualStyleBackColor = false;
            this.DownGevazp.Click += new System.EventHandler(this.DownGevazp_Click);
            // 
            // EntradaSaidaPrevivaz
            // 
            this.EntradaSaidaPrevivaz.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.EntradaSaidaPrevivaz.Enabled = false;
            this.EntradaSaidaPrevivaz.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.EntradaSaidaPrevivaz.Location = new System.Drawing.Point(2, 286);
            this.EntradaSaidaPrevivaz.Name = "EntradaSaidaPrevivaz";
            this.EntradaSaidaPrevivaz.Size = new System.Drawing.Size(144, 24);
            this.EntradaSaidaPrevivaz.TabIndex = 35;
            this.EntradaSaidaPrevivaz.Text = "EntradaSaidaPrevivaz";
            this.EntradaSaidaPrevivaz.UseVisualStyleBackColor = false;
            this.EntradaSaidaPrevivaz.Click += new System.EventHandler(this.EntradaSaidaPrevivaz_Click);
            // 
            // DownTemp
            // 
            this.DownTemp.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownTemp.Enabled = false;
            this.DownTemp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownTemp.Location = new System.Drawing.Point(2, 180);
            this.DownTemp.Name = "DownTemp";
            this.DownTemp.Size = new System.Drawing.Size(144, 24);
            this.DownTemp.TabIndex = 34;
            this.DownTemp.Text = "Temperatura";
            this.DownTemp.UseVisualStyleBackColor = false;
            this.DownTemp.Click += new System.EventHandler(this.DownTemp_Click);
            // 
            // DownVE
            // 
            this.DownVE.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownVE.Enabled = false;
            this.DownVE.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownVE.Location = new System.Drawing.Point(75, 233);
            this.DownVE.Name = "DownVE";
            this.DownVE.Size = new System.Drawing.Size(71, 24);
            this.DownVE.TabIndex = 33;
            this.DownVE.Text = "VE";
            this.DownVE.UseVisualStyleBackColor = false;
            this.DownVE.Click += new System.EventHandler(this.DownVE_Click);
            // 
            // DownCFS
            // 
            this.DownCFS.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownCFS.Enabled = false;
            this.DownCFS.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownCFS.Location = new System.Drawing.Point(2, 233);
            this.DownCFS.Name = "DownCFS";
            this.DownCFS.Size = new System.Drawing.Size(71, 24);
            this.DownCFS.TabIndex = 32;
            this.DownCFS.Text = "CFS";
            this.DownCFS.UseVisualStyleBackColor = false;
            this.DownCFS.Click += new System.EventHandler(this.DownCFS_Click);
            // 
            // DownPmoNewave
            // 
            this.DownPmoNewave.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownPmoNewave.Enabled = false;
            this.DownPmoNewave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownPmoNewave.Location = new System.Drawing.Point(75, 207);
            this.DownPmoNewave.Name = "DownPmoNewave";
            this.DownPmoNewave.Size = new System.Drawing.Size(71, 24);
            this.DownPmoNewave.TabIndex = 31;
            this.DownPmoNewave.Text = "Newave";
            this.DownPmoNewave.UseVisualStyleBackColor = false;
            this.DownPmoNewave.Click += new System.EventHandler(this.DownPmoNewave_Click);
            // 
            // DownPmoDecomp
            // 
            this.DownPmoDecomp.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownPmoDecomp.Enabled = false;
            this.DownPmoDecomp.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownPmoDecomp.Location = new System.Drawing.Point(2, 207);
            this.DownPmoDecomp.Name = "DownPmoDecomp";
            this.DownPmoDecomp.Size = new System.Drawing.Size(71, 24);
            this.DownPmoDecomp.TabIndex = 30;
            this.DownPmoDecomp.Text = "Decomp";
            this.DownPmoDecomp.UseVisualStyleBackColor = false;
            this.DownPmoDecomp.Click += new System.EventHandler(this.DownPmoDecomp_Click);
            // 
            // DownNoticias
            // 
            this.DownNoticias.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownNoticias.Enabled = false;
            this.DownNoticias.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownNoticias.Location = new System.Drawing.Point(2, 153);
            this.DownNoticias.Name = "DownNoticias";
            this.DownNoticias.Size = new System.Drawing.Size(144, 24);
            this.DownNoticias.TabIndex = 29;
            this.DownNoticias.Text = "Notícias";
            this.DownNoticias.UseVisualStyleBackColor = false;
            this.DownNoticias.Click += new System.EventHandler(this.DownNoticias_Click);
            // 
            // DownSemanal
            // 
            this.DownSemanal.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownSemanal.Enabled = false;
            this.DownSemanal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownSemanal.Location = new System.Drawing.Point(75, 260);
            this.DownSemanal.Name = "DownSemanal";
            this.DownSemanal.Size = new System.Drawing.Size(71, 24);
            this.DownSemanal.TabIndex = 28;
            this.DownSemanal.Text = "Semanal";
            this.DownSemanal.UseVisualStyleBackColor = false;
            this.DownSemanal.Click += new System.EventHandler(this.DownSemanal_Click);
            // 
            // DownMensal
            // 
            this.DownMensal.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownMensal.Enabled = false;
            this.DownMensal.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownMensal.Location = new System.Drawing.Point(2, 260);
            this.DownMensal.Name = "DownMensal";
            this.DownMensal.Size = new System.Drawing.Size(71, 24);
            this.DownMensal.TabIndex = 27;
            this.DownMensal.Text = "Mensal";
            this.DownMensal.UseVisualStyleBackColor = false;
            this.DownMensal.Click += new System.EventHandler(this.DownMensal_Click);
            // 
            // DownGifObs
            // 
            this.DownGifObs.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGifObs.Enabled = false;
            this.DownGifObs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGifObs.Location = new System.Drawing.Point(2, 127);
            this.DownGifObs.Name = "DownGifObs";
            this.DownGifObs.Size = new System.Drawing.Size(144, 24);
            this.DownGifObs.TabIndex = 26;
            this.DownGifObs.Text = "Gif Observado";
            this.DownGifObs.UseVisualStyleBackColor = false;
            this.DownGifObs.Click += new System.EventHandler(this.DownGifObs_Click);
            // 
            // DownGifsGefs
            // 
            this.DownGifsGefs.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGifsGefs.Enabled = false;
            this.DownGifsGefs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGifsGefs.Location = new System.Drawing.Point(75, 100);
            this.DownGifsGefs.Name = "DownGifsGefs";
            this.DownGifsGefs.Size = new System.Drawing.Size(71, 24);
            this.DownGifsGefs.TabIndex = 25;
            this.DownGifsGefs.Text = "Gifs GEFS";
            this.DownGifsGefs.UseVisualStyleBackColor = false;
            this.DownGifsGefs.Click += new System.EventHandler(this.DownGifsGefs_Click);
            // 
            // DownGifsEta
            // 
            this.DownGifsEta.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGifsEta.Enabled = false;
            this.DownGifsEta.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGifsEta.Location = new System.Drawing.Point(2, 100);
            this.DownGifsEta.Name = "DownGifsEta";
            this.DownGifsEta.Size = new System.Drawing.Size(71, 24);
            this.DownGifsEta.TabIndex = 24;
            this.DownGifsEta.Text = "Gifs ETA";
            this.DownGifsEta.UseVisualStyleBackColor = false;
            this.DownGifsEta.Click += new System.EventHandler(this.DownGifsEta_Click);
            // 
            // DownGefs
            // 
            this.DownGefs.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownGefs.Enabled = false;
            this.DownGefs.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownGefs.Location = new System.Drawing.Point(75, 74);
            this.DownGefs.Name = "DownGefs";
            this.DownGefs.Size = new System.Drawing.Size(71, 24);
            this.DownGefs.TabIndex = 23;
            this.DownGefs.Text = "Mapa GEFS";
            this.DownGefs.UseVisualStyleBackColor = false;
            this.DownGefs.Click += new System.EventHandler(this.DownGefs_Click);
            // 
            // DownEta
            // 
            this.DownEta.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.DownEta.Enabled = false;
            this.DownEta.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DownEta.Location = new System.Drawing.Point(2, 74);
            this.DownEta.Name = "DownEta";
            this.DownEta.Size = new System.Drawing.Size(71, 24);
            this.DownEta.TabIndex = 22;
            this.DownEta.Text = "Mapa ETA";
            this.DownEta.UseVisualStyleBackColor = false;
            this.DownEta.Click += new System.EventHandler(this.DownEta_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(139, 16);
            this.label2.TabIndex = 21;
            this.label2.Text = "Donwload Sintegre";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(666, 646);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dia);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.ano);
            this.Controls.Add(this.mes);
            this.Controls.Add(this.button3);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Download Enercore";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ano)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button DownRDH;
        private System.Windows.Forms.ComboBox mes;
        private System.Windows.Forms.NumericUpDown ano;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.Button DownAcomph;
        private System.Windows.Forms.ComboBox dia;
        private System.Windows.Forms.Button DownChuvaVazao;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button DownGefs;
        private System.Windows.Forms.Button DownEta;
        private System.Windows.Forms.Button DownGifsGefs;
        private System.Windows.Forms.Button DownGifsEta;
        private System.Windows.Forms.Button DownGifObs;
        private System.Windows.Forms.Button DownMensal;
        private System.Windows.Forms.Button DownSemanal;
        private System.Windows.Forms.Button DownNoticias;
        private System.Windows.Forms.Button DownPmoDecomp;
        private System.Windows.Forms.Button DownPmoNewave;
        private System.Windows.Forms.Button DownCFS;
        private System.Windows.Forms.Button DownVE;
        private System.Windows.Forms.Button DownTemp;
        private System.Windows.Forms.Button EntradaSaidaPrevivaz;
        private System.Windows.Forms.Button DownGevazp;
        private System.Windows.Forms.Button bt_Sat;
        private System.Windows.Forms.Button bt_previs;
        private System.Windows.Forms.Button bt_NOA;
        private System.Windows.Forms.Button DownECMWF;
        private System.Windows.Forms.Button DownGifEuro;
        private System.Windows.Forms.Button vazObsDown;
        private System.Windows.Forms.Button bt_IPDO;
        private System.Windows.Forms.Button DownDessem;
        private System.Windows.Forms.Button bt_cargaDiaria;
        private System.Windows.Forms.Button Carga_Hr;
        private System.Windows.Forms.Button weol;
    }
}

