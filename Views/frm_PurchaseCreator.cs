using GCScript_for_Excel.Classes;
using System;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Views
{
    public partial class frm_PurchaseCreator : Form
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public frm_PurchaseCreator()
        {
            InitializeComponent();
        }

        private void rbtn_Subtotal_Empresa_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.Empresa;
        }

        private void rbtn_Subtotal_Uf_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.Uf;
        }

        private void rbtn_Subtotal_Operadora_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.Operadora;
        }

        private void rbtn_Subtotal_CUnid_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.CUnid;
        }

        private void rbtn_Subtotal_CDepto_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.CDepto;
        }

        private void rbtn_Subtotal_Depto_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorSubtotalOption = Enums.EPurchaseCreatorSubtotalOption.Depto;
        }

        private void btn_Start_Click(object sender, EventArgs e)
        {
            var purchaseCreator = new PurchaseCreator();
            purchaseCreator.Start();
            this.Close();
        }

        private void frm_PurchaseCreator_Load(object sender, EventArgs e)
        {
            if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.Empresa) == -1)
            {
                MessageBox.Show($"Nenhuma coluna encontrada!", "X765937", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                this.Close();
            }
            else if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.Uf) == -1)
            {
                rbtn_Tab_Empresa.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Empresa.Checked = true;
            }
            else if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.Operadora) == -1)
            {
                rbtn_Tab_Empresa.Enabled = true;
                rbtn_Tab_Uf.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Uf.Enabled = true;
                rbtn_Subtotal_Uf.Checked = true;
            }
            else if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.CUnid) == -1)
            {
                rbtn_Tab_Empresa.Enabled = true;
                rbtn_Tab_Uf.Enabled = true;
                rbtn_Tab_Operadora.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Uf.Enabled = true;
                rbtn_Subtotal_Operadora.Enabled = true;
                rbtn_Subtotal_Operadora.Checked = true;
            }
            else if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.CDepto) == -1)
            {
                rbtn_Tab_Empresa.Enabled = true;
                rbtn_Tab_Uf.Enabled = true;
                rbtn_Tab_Operadora.Enabled = true;
                rbtn_Tab_CUnid.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Uf.Enabled = true;
                rbtn_Subtotal_Operadora.Enabled = true;
                rbtn_Subtotal_CUnid.Enabled = true;
                rbtn_Subtotal_CUnid.Checked = true;
            }
            else if (ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.Depto) == -1)
            {
                rbtn_Tab_Empresa.Enabled = true;
                rbtn_Tab_Uf.Enabled = true;
                rbtn_Tab_Operadora.Enabled = true;
                rbtn_Tab_CUnid.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Uf.Enabled = true;
                rbtn_Subtotal_Operadora.Enabled = true;
                rbtn_Subtotal_CUnid.Enabled = true;
                rbtn_Subtotal_CDepto.Enabled = true;
            }
            else
            {
                rbtn_Tab_Empresa.Enabled = true;
                rbtn_Tab_Uf.Enabled = true;
                rbtn_Tab_Operadora.Enabled = true;
                rbtn_Tab_CUnid.Enabled = true;

                rbtn_Subtotal_Empresa.Enabled = true;
                rbtn_Subtotal_Uf.Enabled = true;
                rbtn_Subtotal_Operadora.Enabled = true;
                rbtn_Subtotal_CUnid.Enabled = true;
                rbtn_Subtotal_CDepto.Enabled = true;
                rbtn_Subtotal_Depto.Enabled = true;
            }
            //var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(gcsApp.ActiveSheet, ColumnsName.Empresa);
            txt_Tab_CustomName.Text = cl_Settings.PurchaseCreatorTabName;
        }

        private void rbtn_Tab_CustomName_CheckedChanged(object sender, EventArgs e)
        {
            if (rbtn_Tab_CustomName.Checked)
            {
                txt_Tab_CustomName.Enabled = true;
                cl_Settings.PurchaseCreatorTabOption = Enums.EPurchaseCreatorTabOption.CustomName;
            }
            else
            {
                txt_Tab_CustomName.Enabled = false;
            }
        }

        private void rbtn_Tab_Empresa_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorTabOption = Enums.EPurchaseCreatorTabOption.Empresa;
        }

        private void rbtn_Tab_Uf_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorTabOption = Enums.EPurchaseCreatorTabOption.Uf;
        }

        private void rbtn_Tab_Operadora_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorTabOption = Enums.EPurchaseCreatorTabOption.Operadora;
        }

        private void rbtn_Tab_CUnid_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.PurchaseCreatorTabOption = Enums.EPurchaseCreatorTabOption.CUnid;
        }

        private void txt_Tab_CustomName_Leave(object sender, EventArgs e)
        {
            if (txt_Tab_CustomName.Text.Length < 1) { txt_Tab_CustomName.Text = "Compra"; }
            cl_Settings.PurchaseCreatorTabName = txt_Tab_CustomName.Text;
        }
    }
}
