namespace GCScript_for_Excel
{
    partial class frm_SetTitle
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
            this.btn_Executar = new System.Windows.Forms.Button();
            this.cmb_Mes = new System.Windows.Forms.ComboBox();
            this.cmb_Titulo = new System.Windows.Forms.ComboBox();
            this.cmb_Ano = new System.Windows.Forms.ComboBox();
            this.cmb_Compra = new System.Windows.Forms.ComboBox();
            this.chk_Titulo = new System.Windows.Forms.CheckBox();
            this.chk_Compra = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btn_Executar
            // 
            this.btn_Executar.Location = new System.Drawing.Point(9, 113);
            this.btn_Executar.Name = "btn_Executar";
            this.btn_Executar.Size = new System.Drawing.Size(280, 40);
            this.btn_Executar.TabIndex = 4;
            this.btn_Executar.Text = "APLICAR";
            this.btn_Executar.UseVisualStyleBackColor = true;
            this.btn_Executar.Click += new System.EventHandler(this.btn_Executar_Click);
            // 
            // cmb_Mes
            // 
            this.cmb_Mes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_Mes.FormattingEnabled = true;
            this.cmb_Mes.Items.AddRange(new object[] {
            "Janeiro",
            "Fevereiro",
            "Março",
            "Abril",
            "Maio",
            "Junho",
            "Julho",
            "Agosto",
            "Setembro",
            "Outubro",
            "Novembro",
            "Dezembro"});
            this.cmb_Mes.Location = new System.Drawing.Point(9, 45);
            this.cmb_Mes.Name = "cmb_Mes";
            this.cmb_Mes.Size = new System.Drawing.Size(174, 28);
            this.cmb_Mes.TabIndex = 1;
            // 
            // cmb_Titulo
            // 
            this.cmb_Titulo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_Titulo.FormattingEnabled = true;
            this.cmb_Titulo.Items.AddRange(new object[] {
            "AMOEDO",
            "BAJA",
            "BIOB",
            "BRASIL BROKERS",
            "BRT",
            "BTCC",
            "BTG PACTUAL",
            "CEI",
            "CONFEITARIA",
            "DAVITA",
            "ESCOLA ELEVA",
            "ESTEVAO",
            "FLEX LTDA",
            "FLEX LTDA - FLEX RIO - HERCULES",
            "FLEX RIO",
            "FORCA AMBIENTAL",
            "HERCULES",
            "HOSPITAL",
            "KATTAK",
            "KIK",
            "KLES",
            "L2R",
            "LIBANO",
            "LIGHT",
            "MERCADOS BRAGA",
            "MOITINHO",
            "MR PHARMA",
            "OI NORTE LESTE",
            "PAGGO",
            "PAVIBRAS",
            "PRIME",
            "PRIME - ASG",
            "PRIME - PORTEIROS",
            "PROTEC",
            "REVIVER",
            "RIOSHOP",
            "SEREDE",
            "SMART CAFE",
            "SUPER PRENDAS",
            "TUISE",
            "UNIAO NORTE",
            "VIDA",
            "VIGBAN",
            "Z2010"});
            this.cmb_Titulo.Location = new System.Drawing.Point(9, 11);
            this.cmb_Titulo.Name = "cmb_Titulo";
            this.cmb_Titulo.Size = new System.Drawing.Size(252, 28);
            this.cmb_Titulo.TabIndex = 0;
            // 
            // cmb_Ano
            // 
            this.cmb_Ano.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_Ano.FormattingEnabled = true;
            this.cmb_Ano.Location = new System.Drawing.Point(189, 45);
            this.cmb_Ano.Name = "cmb_Ano";
            this.cmb_Ano.Size = new System.Drawing.Size(100, 28);
            this.cmb_Ano.TabIndex = 2;
            // 
            // cmb_Compra
            // 
            this.cmb_Compra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_Compra.FormattingEnabled = true;
            this.cmb_Compra.Items.AddRange(new object[] {
            "Mensal",
            "1ª Compra",
            "2ª Compra",
            "3ª Compra",
            "4ª Compra",
            "5ª Compra"});
            this.cmb_Compra.Location = new System.Drawing.Point(9, 79);
            this.cmb_Compra.Name = "cmb_Compra";
            this.cmb_Compra.Size = new System.Drawing.Size(252, 28);
            this.cmb_Compra.TabIndex = 3;
            // 
            // chk_Titulo
            // 
            this.chk_Titulo.AutoSize = true;
            this.chk_Titulo.Location = new System.Drawing.Point(267, 15);
            this.chk_Titulo.Name = "chk_Titulo";
            this.chk_Titulo.Size = new System.Drawing.Size(22, 21);
            this.chk_Titulo.TabIndex = 5;
            this.chk_Titulo.UseVisualStyleBackColor = true;
            this.chk_Titulo.CheckedChanged += new System.EventHandler(this.chk_Titulo_CheckedChanged);
            // 
            // chk_Compra
            // 
            this.chk_Compra.AutoSize = true;
            this.chk_Compra.Location = new System.Drawing.Point(267, 83);
            this.chk_Compra.Name = "chk_Compra";
            this.chk_Compra.Size = new System.Drawing.Size(22, 21);
            this.chk_Compra.TabIndex = 6;
            this.chk_Compra.UseVisualStyleBackColor = true;
            this.chk_Compra.CheckedChanged += new System.EventHandler(this.chk_Compra_CheckedChanged);
            // 
            // frm_SetTitle
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(298, 164);
            this.Controls.Add(this.chk_Compra);
            this.Controls.Add(this.chk_Titulo);
            this.Controls.Add(this.cmb_Compra);
            this.Controls.Add(this.cmb_Ano);
            this.Controls.Add(this.cmb_Titulo);
            this.Controls.Add(this.cmb_Mes);
            this.Controls.Add(this.btn_Executar);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_SetTitle";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Definir Título";
            this.Load += new System.EventHandler(this.frm_DefinirTitulo_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_Executar;
        private System.Windows.Forms.ComboBox cmb_Mes;
        private System.Windows.Forms.ComboBox cmb_Titulo;
        private System.Windows.Forms.ComboBox cmb_Ano;
        private System.Windows.Forms.ComboBox cmb_Compra;
        private System.Windows.Forms.CheckBox chk_Titulo;
        private System.Windows.Forms.CheckBox chk_Compra;
    }
}