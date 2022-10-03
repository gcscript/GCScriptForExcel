namespace GCScript_for_Excel.Views
{
    partial class frm_PurchaseCreator
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
            this.tlp_Main = new System.Windows.Forms.TableLayoutPanel();
            this.btn_Start = new System.Windows.Forms.Button();
            this.flp_Tab = new System.Windows.Forms.FlowLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.rbtn_Tab_CustomName = new System.Windows.Forms.RadioButton();
            this.txt_Tab_CustomName = new System.Windows.Forms.TextBox();
            this.rbtn_Tab_Empresa = new System.Windows.Forms.RadioButton();
            this.rbtn_Tab_Uf = new System.Windows.Forms.RadioButton();
            this.rbtn_Tab_Operadora = new System.Windows.Forms.RadioButton();
            this.rbtn_Tab_CUnid = new System.Windows.Forms.RadioButton();
            this.flp_Subtotal = new System.Windows.Forms.FlowLayoutPanel();
            this.label1 = new System.Windows.Forms.Label();
            this.rbtn_Subtotal_Empresa = new System.Windows.Forms.RadioButton();
            this.rbtn_Subtotal_Uf = new System.Windows.Forms.RadioButton();
            this.rbtn_Subtotal_Operadora = new System.Windows.Forms.RadioButton();
            this.rbtn_Subtotal_CUnid = new System.Windows.Forms.RadioButton();
            this.rbtn_Subtotal_CDepto = new System.Windows.Forms.RadioButton();
            this.rbtn_Subtotal_Depto = new System.Windows.Forms.RadioButton();
            this.flp_SplitPurchase = new System.Windows.Forms.FlowLayoutPanel();
            this.label3 = new System.Windows.Forms.Label();
            this.rbtn_SplitPurchase_1x = new System.Windows.Forms.RadioButton();
            this.rbtn_SplitPurchase_2x = new System.Windows.Forms.RadioButton();
            this.rbtn_SplitPurchase_3x = new System.Windows.Forms.RadioButton();
            this.tlp_Main.SuspendLayout();
            this.flp_Tab.SuspendLayout();
            this.flp_Subtotal.SuspendLayout();
            this.flp_SplitPurchase.SuspendLayout();
            this.SuspendLayout();
            // 
            // tlp_Main
            // 
            this.tlp_Main.ColumnCount = 1;
            this.tlp_Main.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tlp_Main.Controls.Add(this.flp_Tab, 0, 0);
            this.tlp_Main.Controls.Add(this.flp_Subtotal, 0, 1);
            this.tlp_Main.Controls.Add(this.flp_SplitPurchase, 0, 2);
            this.tlp_Main.Controls.Add(this.btn_Start, 0, 3);
            this.tlp_Main.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tlp_Main.Location = new System.Drawing.Point(0, 0);
            this.tlp_Main.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.tlp_Main.Name = "tlp_Main";
            this.tlp_Main.RowCount = 4;
            this.tlp_Main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 38F));
            this.tlp_Main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 38F));
            this.tlp_Main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 24F));
            this.tlp_Main.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50F));
            this.tlp_Main.Size = new System.Drawing.Size(278, 744);
            this.tlp_Main.TabIndex = 0;
            // 
            // btn_Start
            // 
            this.btn_Start.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btn_Start.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Start.Location = new System.Drawing.Point(3, 695);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(272, 46);
            this.btn_Start.TabIndex = 1;
            this.btn_Start.Text = "EXECUTAR";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // flp_Tab
            // 
            this.flp_Tab.Controls.Add(this.label2);
            this.flp_Tab.Controls.Add(this.rbtn_Tab_CustomName);
            this.flp_Tab.Controls.Add(this.txt_Tab_CustomName);
            this.flp_Tab.Controls.Add(this.rbtn_Tab_Empresa);
            this.flp_Tab.Controls.Add(this.rbtn_Tab_Uf);
            this.flp_Tab.Controls.Add(this.rbtn_Tab_Operadora);
            this.flp_Tab.Controls.Add(this.rbtn_Tab_CUnid);
            this.flp_Tab.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flp_Tab.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flp_Tab.Location = new System.Drawing.Point(3, 3);
            this.flp_Tab.Name = "flp_Tab";
            this.flp_Tab.Size = new System.Drawing.Size(272, 257);
            this.flp_Tab.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Bold);
            this.label2.Location = new System.Drawing.Point(5, 10);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 10, 3, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 23);
            this.label2.TabIndex = 8;
            this.label2.Text = "Tabs based on:";
            // 
            // rbtn_Tab_CustomName
            // 
            this.rbtn_Tab_CustomName.AutoSize = true;
            this.rbtn_Tab_CustomName.Checked = true;
            this.rbtn_Tab_CustomName.Location = new System.Drawing.Point(10, 46);
            this.rbtn_Tab_CustomName.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Tab_CustomName.Name = "rbtn_Tab_CustomName";
            this.rbtn_Tab_CustomName.Size = new System.Drawing.Size(79, 23);
            this.rbtn_Tab_CustomName.TabIndex = 12;
            this.rbtn_Tab_CustomName.TabStop = true;
            this.rbtn_Tab_CustomName.Text = "Name:";
            this.rbtn_Tab_CustomName.UseVisualStyleBackColor = true;
            this.rbtn_Tab_CustomName.CheckedChanged += new System.EventHandler(this.rbtn_Tab_CustomName_CheckedChanged);
            // 
            // txt_Tab_CustomName
            // 
            this.txt_Tab_CustomName.Location = new System.Drawing.Point(10, 72);
            this.txt_Tab_CustomName.Margin = new System.Windows.Forms.Padding(10, 0, 3, 3);
            this.txt_Tab_CustomName.MaxLength = 30;
            this.txt_Tab_CustomName.Name = "txt_Tab_CustomName";
            this.txt_Tab_CustomName.Size = new System.Drawing.Size(253, 26);
            this.txt_Tab_CustomName.TabIndex = 13;
            this.txt_Tab_CustomName.Text = "Compra";
            this.txt_Tab_CustomName.Leave += new System.EventHandler(this.txt_Tab_CustomName_Leave);
            // 
            // rbtn_Tab_Empresa
            // 
            this.rbtn_Tab_Empresa.AutoSize = true;
            this.rbtn_Tab_Empresa.Enabled = false;
            this.rbtn_Tab_Empresa.Location = new System.Drawing.Point(10, 104);
            this.rbtn_Tab_Empresa.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Tab_Empresa.Name = "rbtn_Tab_Empresa";
            this.rbtn_Tab_Empresa.Size = new System.Drawing.Size(97, 23);
            this.rbtn_Tab_Empresa.TabIndex = 9;
            this.rbtn_Tab_Empresa.Text = "Empresa";
            this.rbtn_Tab_Empresa.UseVisualStyleBackColor = true;
            this.rbtn_Tab_Empresa.CheckedChanged += new System.EventHandler(this.rbtn_Tab_Empresa_CheckedChanged);
            // 
            // rbtn_Tab_Uf
            // 
            this.rbtn_Tab_Uf.AutoSize = true;
            this.rbtn_Tab_Uf.Enabled = false;
            this.rbtn_Tab_Uf.Location = new System.Drawing.Point(10, 133);
            this.rbtn_Tab_Uf.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Tab_Uf.Name = "rbtn_Tab_Uf";
            this.rbtn_Tab_Uf.Size = new System.Drawing.Size(52, 23);
            this.rbtn_Tab_Uf.TabIndex = 10;
            this.rbtn_Tab_Uf.Text = "UF";
            this.rbtn_Tab_Uf.UseVisualStyleBackColor = true;
            this.rbtn_Tab_Uf.CheckedChanged += new System.EventHandler(this.rbtn_Tab_Uf_CheckedChanged);
            // 
            // rbtn_Tab_Operadora
            // 
            this.rbtn_Tab_Operadora.AutoSize = true;
            this.rbtn_Tab_Operadora.Enabled = false;
            this.rbtn_Tab_Operadora.Location = new System.Drawing.Point(10, 162);
            this.rbtn_Tab_Operadora.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Tab_Operadora.Name = "rbtn_Tab_Operadora";
            this.rbtn_Tab_Operadora.Size = new System.Drawing.Size(115, 23);
            this.rbtn_Tab_Operadora.TabIndex = 11;
            this.rbtn_Tab_Operadora.Text = "Operadora";
            this.rbtn_Tab_Operadora.UseVisualStyleBackColor = true;
            this.rbtn_Tab_Operadora.CheckedChanged += new System.EventHandler(this.rbtn_Tab_Operadora_CheckedChanged);
            // 
            // rbtn_Tab_CUnid
            // 
            this.rbtn_Tab_CUnid.AutoSize = true;
            this.rbtn_Tab_CUnid.Enabled = false;
            this.rbtn_Tab_CUnid.Location = new System.Drawing.Point(10, 191);
            this.rbtn_Tab_CUnid.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Tab_CUnid.Name = "rbtn_Tab_CUnid";
            this.rbtn_Tab_CUnid.Size = new System.Drawing.Size(88, 23);
            this.rbtn_Tab_CUnid.TabIndex = 14;
            this.rbtn_Tab_CUnid.Text = "C.Unid";
            this.rbtn_Tab_CUnid.UseVisualStyleBackColor = true;
            this.rbtn_Tab_CUnid.CheckedChanged += new System.EventHandler(this.rbtn_Tab_CUnid_CheckedChanged);
            // 
            // flp_Subtotal
            // 
            this.flp_Subtotal.Controls.Add(this.label1);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_Empresa);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_Uf);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_Operadora);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_CUnid);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_CDepto);
            this.flp_Subtotal.Controls.Add(this.rbtn_Subtotal_Depto);
            this.flp_Subtotal.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flp_Subtotal.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flp_Subtotal.Location = new System.Drawing.Point(3, 266);
            this.flp_Subtotal.Name = "flp_Subtotal";
            this.flp_Subtotal.Size = new System.Drawing.Size(272, 257);
            this.flp_Subtotal.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(5, 10);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 10, 3, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(208, 23);
            this.label1.TabIndex = 9;
            this.label1.Text = "Subtotal based on:";
            // 
            // rbtn_Subtotal_Empresa
            // 
            this.rbtn_Subtotal_Empresa.AutoSize = true;
            this.rbtn_Subtotal_Empresa.Enabled = false;
            this.rbtn_Subtotal_Empresa.Location = new System.Drawing.Point(10, 46);
            this.rbtn_Subtotal_Empresa.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_Empresa.Name = "rbtn_Subtotal_Empresa";
            this.rbtn_Subtotal_Empresa.Size = new System.Drawing.Size(97, 23);
            this.rbtn_Subtotal_Empresa.TabIndex = 1;
            this.rbtn_Subtotal_Empresa.Text = "Empresa";
            this.rbtn_Subtotal_Empresa.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_Empresa.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_Empresa_CheckedChanged);
            // 
            // rbtn_Subtotal_Uf
            // 
            this.rbtn_Subtotal_Uf.AutoSize = true;
            this.rbtn_Subtotal_Uf.Enabled = false;
            this.rbtn_Subtotal_Uf.Location = new System.Drawing.Point(10, 75);
            this.rbtn_Subtotal_Uf.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_Uf.Name = "rbtn_Subtotal_Uf";
            this.rbtn_Subtotal_Uf.Size = new System.Drawing.Size(52, 23);
            this.rbtn_Subtotal_Uf.TabIndex = 2;
            this.rbtn_Subtotal_Uf.Text = "UF";
            this.rbtn_Subtotal_Uf.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_Uf.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_Uf_CheckedChanged);
            // 
            // rbtn_Subtotal_Operadora
            // 
            this.rbtn_Subtotal_Operadora.AutoSize = true;
            this.rbtn_Subtotal_Operadora.Enabled = false;
            this.rbtn_Subtotal_Operadora.Location = new System.Drawing.Point(10, 104);
            this.rbtn_Subtotal_Operadora.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_Operadora.Name = "rbtn_Subtotal_Operadora";
            this.rbtn_Subtotal_Operadora.Size = new System.Drawing.Size(115, 23);
            this.rbtn_Subtotal_Operadora.TabIndex = 3;
            this.rbtn_Subtotal_Operadora.Text = "Operadora";
            this.rbtn_Subtotal_Operadora.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_Operadora.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_Operadora_CheckedChanged);
            // 
            // rbtn_Subtotal_CUnid
            // 
            this.rbtn_Subtotal_CUnid.AutoSize = true;
            this.rbtn_Subtotal_CUnid.Checked = true;
            this.rbtn_Subtotal_CUnid.Enabled = false;
            this.rbtn_Subtotal_CUnid.Location = new System.Drawing.Point(10, 133);
            this.rbtn_Subtotal_CUnid.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_CUnid.Name = "rbtn_Subtotal_CUnid";
            this.rbtn_Subtotal_CUnid.Size = new System.Drawing.Size(88, 23);
            this.rbtn_Subtotal_CUnid.TabIndex = 3;
            this.rbtn_Subtotal_CUnid.TabStop = true;
            this.rbtn_Subtotal_CUnid.Text = "C.Unid";
            this.rbtn_Subtotal_CUnid.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_CUnid.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_CUnid_CheckedChanged);
            // 
            // rbtn_Subtotal_CDepto
            // 
            this.rbtn_Subtotal_CDepto.AutoSize = true;
            this.rbtn_Subtotal_CDepto.Enabled = false;
            this.rbtn_Subtotal_CDepto.Location = new System.Drawing.Point(10, 162);
            this.rbtn_Subtotal_CDepto.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_CDepto.Name = "rbtn_Subtotal_CDepto";
            this.rbtn_Subtotal_CDepto.Size = new System.Drawing.Size(97, 23);
            this.rbtn_Subtotal_CDepto.TabIndex = 4;
            this.rbtn_Subtotal_CDepto.Text = "C.Depto";
            this.rbtn_Subtotal_CDepto.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_CDepto.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_CDepto_CheckedChanged);
            // 
            // rbtn_Subtotal_Depto
            // 
            this.rbtn_Subtotal_Depto.AutoSize = true;
            this.rbtn_Subtotal_Depto.Enabled = false;
            this.rbtn_Subtotal_Depto.Location = new System.Drawing.Point(10, 191);
            this.rbtn_Subtotal_Depto.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_Subtotal_Depto.Name = "rbtn_Subtotal_Depto";
            this.rbtn_Subtotal_Depto.Size = new System.Drawing.Size(79, 23);
            this.rbtn_Subtotal_Depto.TabIndex = 5;
            this.rbtn_Subtotal_Depto.Text = "Depto";
            this.rbtn_Subtotal_Depto.UseVisualStyleBackColor = true;
            this.rbtn_Subtotal_Depto.CheckedChanged += new System.EventHandler(this.rbtn_Subtotal_Depto_CheckedChanged);
            // 
            // flp_SplitPurchase
            // 
            this.flp_SplitPurchase.Controls.Add(this.label3);
            this.flp_SplitPurchase.Controls.Add(this.rbtn_SplitPurchase_1x);
            this.flp_SplitPurchase.Controls.Add(this.rbtn_SplitPurchase_2x);
            this.flp_SplitPurchase.Controls.Add(this.rbtn_SplitPurchase_3x);
            this.flp_SplitPurchase.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flp_SplitPurchase.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flp_SplitPurchase.Location = new System.Drawing.Point(3, 529);
            this.flp_SplitPurchase.Name = "flp_SplitPurchase";
            this.flp_SplitPurchase.Size = new System.Drawing.Size(272, 160);
            this.flp_SplitPurchase.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 10F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(5, 10);
            this.label3.Margin = new System.Windows.Forms.Padding(5, 10, 3, 10);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(208, 23);
            this.label3.TabIndex = 10;
            this.label3.Text = "Split Purchase in:";
            // 
            // rbtn_SplitPurchase_1x
            // 
            this.rbtn_SplitPurchase_1x.AutoSize = true;
            this.rbtn_SplitPurchase_1x.Location = new System.Drawing.Point(10, 46);
            this.rbtn_SplitPurchase_1x.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_SplitPurchase_1x.Name = "rbtn_SplitPurchase_1x";
            this.rbtn_SplitPurchase_1x.Size = new System.Drawing.Size(52, 23);
            this.rbtn_SplitPurchase_1x.TabIndex = 11;
            this.rbtn_SplitPurchase_1x.Text = "1x";
            this.rbtn_SplitPurchase_1x.UseVisualStyleBackColor = true;
            this.rbtn_SplitPurchase_1x.CheckedChanged += new System.EventHandler(this.rbtn_SplitPurchase_1x_CheckedChanged);
            // 
            // rbtn_SplitPurchase_2x
            // 
            this.rbtn_SplitPurchase_2x.AutoSize = true;
            this.rbtn_SplitPurchase_2x.Location = new System.Drawing.Point(10, 75);
            this.rbtn_SplitPurchase_2x.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_SplitPurchase_2x.Name = "rbtn_SplitPurchase_2x";
            this.rbtn_SplitPurchase_2x.Size = new System.Drawing.Size(52, 23);
            this.rbtn_SplitPurchase_2x.TabIndex = 12;
            this.rbtn_SplitPurchase_2x.Text = "2x";
            this.rbtn_SplitPurchase_2x.UseVisualStyleBackColor = true;
            this.rbtn_SplitPurchase_2x.CheckedChanged += new System.EventHandler(this.rbtn_SplitPurchase_2x_CheckedChanged);
            // 
            // rbtn_SplitPurchase_3x
            // 
            this.rbtn_SplitPurchase_3x.AutoSize = true;
            this.rbtn_SplitPurchase_3x.Location = new System.Drawing.Point(10, 104);
            this.rbtn_SplitPurchase_3x.Margin = new System.Windows.Forms.Padding(10, 3, 3, 3);
            this.rbtn_SplitPurchase_3x.Name = "rbtn_SplitPurchase_3x";
            this.rbtn_SplitPurchase_3x.Size = new System.Drawing.Size(52, 23);
            this.rbtn_SplitPurchase_3x.TabIndex = 13;
            this.rbtn_SplitPurchase_3x.Text = "3x";
            this.rbtn_SplitPurchase_3x.UseVisualStyleBackColor = true;
            this.rbtn_SplitPurchase_3x.CheckedChanged += new System.EventHandler(this.rbtn_SplitPurchase_3x_CheckedChanged);
            // 
            // frm_PurchaseCreator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(278, 744);
            this.Controls.Add(this.tlp_Main);
            this.Font = new System.Drawing.Font("Consolas", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frm_PurchaseCreator";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Purchase Creator";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frm_PurchaseCreator_Load);
            this.tlp_Main.ResumeLayout(false);
            this.flp_Tab.ResumeLayout(false);
            this.flp_Tab.PerformLayout();
            this.flp_Subtotal.ResumeLayout(false);
            this.flp_Subtotal.PerformLayout();
            this.flp_SplitPurchase.ResumeLayout(false);
            this.flp_SplitPurchase.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tlp_Main;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_Depto;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_CDepto;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_CUnid;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_Operadora;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_Uf;
        private System.Windows.Forms.RadioButton rbtn_Subtotal_Empresa;
        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.FlowLayoutPanel flp_Tab;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RadioButton rbtn_Tab_Empresa;
        private System.Windows.Forms.RadioButton rbtn_Tab_Uf;
        private System.Windows.Forms.RadioButton rbtn_Tab_Operadora;
        private System.Windows.Forms.RadioButton rbtn_Tab_CustomName;
        private System.Windows.Forms.FlowLayoutPanel flp_Subtotal;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txt_Tab_CustomName;
        private System.Windows.Forms.RadioButton rbtn_Tab_CUnid;
        private System.Windows.Forms.FlowLayoutPanel flp_SplitPurchase;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RadioButton rbtn_SplitPurchase_1x;
        private System.Windows.Forms.RadioButton rbtn_SplitPurchase_2x;
        private System.Windows.Forms.RadioButton rbtn_SplitPurchase_3x;
    }
}