using GCScript_for_Excel.Classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GCScript_for_Excel
{
    public partial class frm_SetTitle : Form
    {
        public frm_SetTitle()
        {
            InitializeComponent();
        }

        private void frm_DefinirTitulo_Load(object sender, EventArgs e)
        {
            cmb_Ano.Items.Add(DateTime.Now.AddYears(-1).Year.ToString());
            cmb_Ano.Items.Add(DateTime.Now.Year.ToString());
            cmb_Ano.Items.Add(DateTime.Now.AddYears(1).Year.ToString());


            cmb_Titulo.SelectedIndex = 0;
            cmb_Mes.SelectedIndex = DateTime.Now.AddMonths(-1).Month;
            cmb_Ano.SelectedIndex = 1;
            cmb_Compra.SelectedIndex = 0;
        }

        private void btn_Executar_Click(object sender, EventArgs e)
        {
            try
            {
                cl_Tools.DefinirDados(cmb_Titulo.Text, cmb_Mes.Text, cmb_Ano.Text, cmb_Compra.Text);
                this.Close();
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 821749", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void chk_Titulo_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_Titulo.Checked == true)
            {
                cmb_Titulo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            }
            else
            {
                cmb_Titulo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            }
        }

        private void chk_Compra_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_Compra.Checked == true)
            {
                cmb_Compra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            }
            else
            {
                cmb_Compra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            }
        }
    }
}
