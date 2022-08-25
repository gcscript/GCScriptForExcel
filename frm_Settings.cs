using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel.Classes;

namespace GCScript_for_Excel
{
    public partial class frm_Settings : Form
    {
        public int tabPage { private get; set; }
        // 0 = Geral
        // 1 = Converter
        // 2 = Apply & Remove

        public frm_Settings()
        {
            InitializeComponent();
        }

        //----------------[OPCOES DE TEXTO]-----------------
        private void frm_Settings_Load(object sender, EventArgs e)
        {
            switch (tabPage)
            {
                case 0:
                    tbc_Main.SelectedTab = tbp_Geral;
                    break;
                case 1:
                    tbc_Main.SelectedTab = tbp_Converter;
                    break;
                case 2:
                    tbc_Main.SelectedTab = tbp_ApplyRemove;
                    break;
            }

            #region Converter
            Converter_chk_Texto_Trim.Checked = cl_Settings.Text_Trim;
            Converter_chk_Texto_RemoverEspacoDuplicado.Checked = cl_Settings.Text_RemoverEspacoDuplicado;
            Converter_chk_Texto_Acentos.Checked = cl_Settings.Text_RemoverAcentos;

            switch (cl_Settings.Text_Option)
            {
                case 0:
                    Converter_rdo_Texto_Maiusculo.Checked = true;
                    break;
                case 1:
                    Converter_rdo_Texto_Minusculo.Checked = true;
                    break;
                case 2:
                    Converter_rdo_Texto_Titulo.Checked = true;
                    break;
                case 3:
                    Converter_rdo_Text_Default.Checked = true;
                    break;
                case 4:
                    Converter_rdo_Texto_Substituir.Checked = true;
                    break;
                case 5:
                    Converter_rdo_Texto_Alinhamento.Checked = true;
                    break;
            }

            switch (cl_Settings.converter_Text_Opcao_Alinhamento)
            {
                case 0:
                    Converter_rdo_Texto_Alinhamento_Esquerda.Checked = true;
                    break;
                case 1:
                    Converter_rdo_Texto_Alinhamento_Centralizado.Checked = true;
                    break;
                case 2:
                    Converter_rdo_Texto_Alinhamento_Direita.Checked = true;
                    break;
            }

            Converter_txt_Texto_Substituir_Origem.Text = cl_Settings.converter_Text_Substituir_Origem;
            Converter_txt_Texto_Substituir_Destino.Text = cl_Settings.converter_Text_Substituir_Destino;

            Converter_nud_Texto_Alinhamento_Comprimento.Value = cl_Settings.converter_Text_Alinhamento_Comprimento;
            Converter_txt_Texto_Alinhamento_Preenchimento.Text = cl_Settings.converter_Text_Alinhamento_Preenchimento;

            Converter_chk_CPF_ZeroAEsquerda.Checked = cl_Settings.CPF_ZeroAEsquerda;
            switch (cl_Settings.CPF_Opcao)
            {
                case 0:
                    Converter_rdo_CPF_Formato01.Checked = true;
                    break;
                case 1:
                    Converter_rdo_CPF_Formato02.Checked = true;
                    break;
            }

            Converter_chk_CNPJ_ZeroAEsquerda.Checked = cl_Settings.CNPJ_ZeroAEsquerda;
            switch (cl_Settings.CNPJ_Opcao)
            {
                case 0:
                    Converter_rdo_CNPJ_Formato01.Checked = true;
                    break;
                case 1:
                    Converter_rdo_CNPJ_Formato02.Checked = true;
                    break;
            }
            #endregion

            #region APPLY & REMOVE
            ApplyRemove_chk_Apply_AllSheets.Checked = cl_Settings.ApplyRemove_Apply_AllSheets;
            ApplyRemove_chk_Apply_FontName.Checked = cl_Settings.ApplyRemove_Apply_FontName;
            ApplyRemove_cmb_Apply_FontName.SelectedItem = cl_Settings.ApplyRemove_Apply_FontNameText;
            ApplyRemove_chk_Apply_FontSize.Checked = cl_Settings.ApplyRemove_Apply_FontSize;
            ApplyRemove_cmb_Apply_FontSize.SelectedItem = cl_Settings.ApplyRemove_Apply_FontSizeText;

            ApplyRemove_chk_Apply_Align_Vertical.Checked = cl_Settings.ApplyRemove_Apply_Align_Vertical;
            switch (cl_Settings.ApplyRemove_Apply_Align_VerticalValue)
            {
                case 0:
                    ApplyRemove_cmb_Apply_Align_Vertical.SelectedItem = "Top";
                    break;
                case 1:
                    ApplyRemove_cmb_Apply_Align_Vertical.SelectedItem = "Middle";
                    break;
                case 2:
                    ApplyRemove_cmb_Apply_Align_Vertical.SelectedItem = "Bottom";
                    break;
            }

            ApplyRemove_chk_Apply_Align_Horizontal.Checked = cl_Settings.ApplyRemove_Apply_Align_Horizontal;
            switch (cl_Settings.ApplyRemove_Apply_Align_HorizontalValue)
            {
                case 0:
                    ApplyRemove_cmb_Apply_Align_Horizontal.SelectedItem = "Left";
                    break;
                case 1:
                    ApplyRemove_cmb_Apply_Align_Horizontal.SelectedItem = "Center";
                    break;
                case 2:
                    ApplyRemove_cmb_Apply_Align_Horizontal.SelectedItem = "Right";
                    break;
            }

            ApplyRemove_chk_Apply_RowHeight.Checked = cl_Settings.ApplyRemove_Apply_RowHeight;
            ApplyRemove_nud_Apply_RowHeight.Value = cl_Settings.ApplyRemove_Apply_RowHeightValue;
            ApplyRemove_chk_Apply_ColumnWidth.Checked = cl_Settings.ApplyRemove_Apply_ColumnWidth;
            ApplyRemove_nud_Apply_ColumnWidth.Value = cl_Settings.ApplyRemove_Apply_ColumnWidthValue;
            ApplyRemove_chk_Apply_Zoom.Checked = cl_Settings.ApplyRemove_Apply_Zoom;
            ApplyRemove_nud_Apply_Zoom.Value = cl_Settings.ApplyRemove_Apply_ZoomValue;

            ApplyRemove_chk_Remove_FontBold.Checked = cl_Settings.ApplyRemove_Remove_FontBold;
            ApplyRemove_chk_Remove_FontItalic.Checked = cl_Settings.ApplyRemove_Remove_FontItalic;
            ApplyRemove_chk_Remove_FontUnderline.Checked = cl_Settings.ApplyRemove_Remove_FontUnderline;
            ApplyRemove_chk_Remove_Borders.Checked = cl_Settings.ApplyRemove_Remove_Borders;
            ApplyRemove_chk_Remove_Fill.Checked = cl_Settings.ApplyRemove_Remove_Fill;
            ApplyRemove_chk_Remove_FontColor.Checked = cl_Settings.ApplyRemove_Remove_FontColor;
            ApplyRemove_chk_Remove_WrapText.Checked = cl_Settings.ApplyRemove_Remove_WrapText;
            ApplyRemove_chk_Remove_MergeCells.Checked = cl_Settings.ApplyRemove_Remove_MergeCells;
            ApplyRemove_chk_Remove_Formula.Checked = cl_Settings.ApplyRemove_Remove_Formula;
            ApplyRemove_chk_Remove_ConditionalFormatting.Checked = cl_Settings.ApplyRemove_Remove_ConditionalFormatting;
            #endregion
        }

        private void Converter_chk_Texto_Trim_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Trim = Converter_chk_Texto_Trim.Checked;
        }

        private void Converter_chk_Texto_RemoverEspacoDuplicado_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_RemoverEspacoDuplicado = Converter_chk_Texto_RemoverEspacoDuplicado.Checked;
        }

        private void Converter_chk_Texto_Acentos_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_RemoverAcentos = Converter_chk_Texto_Acentos.Checked;
        }

        //----------------[FUNCOES DE TEXTO]----------------
        private void Converter_rdo_Texto_Maiusculo_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 0;
        }

        private void Converter_rdo_Texto_Minusculo_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 1;
        }

        private void Converter_rdo_Texto_Titulo_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 2;
        }

        private void Converter_rdo_Texto_Original_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 3;
        }

        private void Converter_rdo_Texto_Substituir_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 4;

            if (Converter_rdo_Texto_Substituir.Checked)
            {
                Converter_pnl_Texto_Substituir.Enabled = true;
            }
            else
            {
                Converter_pnl_Texto_Substituir.Enabled = false;
            }
        }

        private void Converter_rdo_Texto_Alinhamento_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.Text_Option = 5;

            if (Converter_rdo_Texto_Alinhamento.Checked)
            {
                Converter_pnl_Texto_Alinhamento.Enabled = true;
            }
            else
            {
                Converter_pnl_Texto_Alinhamento.Enabled = false;
            }
        }

        //---------------[SUBSTITUIR: OPCOES]---------------
        private void Converter_txt_Texto_Substituir_Origem_TextChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Substituir_Origem = Converter_txt_Texto_Substituir_Origem.Text;
        }

        private void Converter_txt_Texto_Substituir_Destino_TextChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Substituir_Destino = Converter_txt_Texto_Substituir_Destino.Text;
        }

        //--------------[ALINHAMENTO: OPCOES]---------------
        private void Converter_rdo_Texto_Alinhamento_Esquerda_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Opcao_Alinhamento = 0;
        }

        private void Converter_rdo_Texto_Alinhamento_Centralizado_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Opcao_Alinhamento = 1;
        }

        private void Converter_rdo_Texto_Alinhamento_Direita_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Opcao_Alinhamento = 2;
        }

        private void Converter_nud_Texto_Alinhamento_Comprimento_ValueChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Alinhamento_Comprimento = (int)Converter_nud_Texto_Alinhamento_Comprimento.Value;
        }

        private void Converter_txt_Texto_Alinhamento_Preenchimento_TextChanged(object sender, EventArgs e)
        {
            cl_Settings.converter_Text_Alinhamento_Preenchimento = Converter_txt_Texto_Alinhamento_Preenchimento.Text;
        }

        //-----------------[OPCOES DE CPF]------------------
        private void Converter_chk_CPF_ZeroAEsquerda_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CPF_ZeroAEsquerda = Converter_chk_CPF_ZeroAEsquerda.Checked;
        }

        //-----------------[FORMATO DE CPF]-----------------

        private void Converter_rdo_CPF_Formato01_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CPF_Opcao = 0;
        }

        private void Converter_rdo_CPF_Formato02_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CPF_Opcao = 1;
        }

        //-----------------[OPCOES DE CNPJ]-----------------
        private void Converter_chk_CNPJ_ZeroAEsquerda_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CNPJ_ZeroAEsquerda = Converter_chk_CNPJ_ZeroAEsquerda.Checked;
        }

        //----------------[FORMATO DE CNPJ]-----------------
        private void Converter_rdo_CNPJ_Formato01_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CNPJ_Opcao = 0;
        }

        private void Converter_rdo_CNPJ_Formato02_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.CNPJ_Opcao = 1;
        }

        //-----------------[APPLY & REMOVE]-----------------
        private void ApplyRemove_chk_Apply_AllSheets_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_AllSheets = ApplyRemove_chk_Apply_AllSheets.Checked;
        }

        private void ApplyRemove_chk_Apply_FontName_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_FontName = ApplyRemove_chk_Apply_FontName.Checked;

            if (ApplyRemove_chk_Apply_FontName.Checked == false)
            {
                ApplyRemove_cmb_Apply_FontName.Enabled = false;
            }
            else
            {
                ApplyRemove_cmb_Apply_FontName.Enabled = true;
            }

        }

        private void ApplyRemove_cmb_Apply_FontName_SelectedIndexChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_FontNameText = ApplyRemove_cmb_Apply_FontName.Text;
        }

        private void ApplyRemove_chk_Apply_FontSize_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_FontSize = ApplyRemove_chk_Apply_FontSize.Checked;

            if (ApplyRemove_chk_Apply_FontSize.Checked == false)
            {
                ApplyRemove_cmb_Apply_FontSize.Enabled = false;
            }
            else
            {
                ApplyRemove_cmb_Apply_FontSize.Enabled = true;
            }
        }

        private void ApplyRemove_cmb_Apply_FontSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_FontSizeText = ApplyRemove_cmb_Apply_FontSize.Text;
        }

        private void ApplyRemove_chk_Remove_FontBold_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_FontBold = ApplyRemove_chk_Remove_FontBold.Checked;
        }

        private void ApplyRemove_chk_Remove_FontItalic_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_FontItalic = ApplyRemove_chk_Remove_FontItalic.Checked;
        }

        private void ApplyRemove_chk_Remove_FontUnderline_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_FontUnderline = ApplyRemove_chk_Remove_FontUnderline.Checked;
        }

        private void ApplyRemove_chk_Remove_Borders_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_Borders = ApplyRemove_chk_Remove_Borders.Checked;
        }

        private void ApplyRemove_chk_Remove_Fill_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_Fill = ApplyRemove_chk_Remove_Fill.Checked;
        }

        private void ApplyRemove_chk_Remove_FontColor_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_FontColor = ApplyRemove_chk_Remove_FontColor.Checked;
        }

        private void ApplyRemove_chk_Remove_WrapText_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_WrapText = ApplyRemove_chk_Remove_WrapText.Checked;
        }

        private void ApplyRemove_chk_Remove_MergeCells_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_MergeCells = ApplyRemove_chk_Remove_MergeCells.Checked;
        }

        private void ApplyRemove_chk_Remove_Formula_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_Formula = ApplyRemove_chk_Remove_Formula.Checked;
        }

        private void ApplyRemove_chk_Remove_ConditionalFormatting_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Remove_ConditionalFormatting = ApplyRemove_chk_Remove_ConditionalFormatting.Checked;
        }

        private void ApplyRemove_chk_Apply_RowHeight_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_RowHeight = ApplyRemove_chk_Apply_RowHeight.Checked;

            if (ApplyRemove_chk_Apply_RowHeight.Checked == false)
            {
                ApplyRemove_nud_Apply_RowHeight.Enabled = false;
            }
            else
            {
                ApplyRemove_nud_Apply_RowHeight.Enabled = true;
            }
        }

        private void ApplyRemove_nud_Apply_RowHeight_ValueChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_RowHeightValue = ApplyRemove_nud_Apply_RowHeight.Value;
        }

        private void ApplyRemove_chk_Apply_ColumnWidth_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_ColumnWidth = ApplyRemove_chk_Apply_ColumnWidth.Checked;

            if (ApplyRemove_chk_Apply_ColumnWidth.Checked == false)
            {
                ApplyRemove_nud_Apply_ColumnWidth.Enabled = false;
            }
            else
            {
                ApplyRemove_nud_Apply_ColumnWidth.Enabled = true;
            }
        }

        private void ApplyRemove_nud_Apply_ColumnWidth_ValueChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_ColumnWidthValue = ApplyRemove_nud_Apply_ColumnWidth.Value;
        }

        private void ApplyRemove_chk_Apply_Align_Vertical_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_Align_Vertical = ApplyRemove_chk_Apply_Align_Vertical.Checked;

            if (ApplyRemove_chk_Apply_Align_Vertical.Checked == false)
            {
                ApplyRemove_cmb_Apply_Align_Vertical.Enabled = false;
            }
            else
            {
                ApplyRemove_cmb_Apply_Align_Vertical.Enabled = true;
            }
        }

        private void ApplyRemove_cmb_Apply_Align_Vertical_SelectedIndexChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_Align_VerticalValue = ApplyRemove_cmb_Apply_Align_Vertical.SelectedIndex;
        }

        private void ApplyRemove_chk_Apply_Align_Horizontal_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_Align_Horizontal = ApplyRemove_chk_Apply_Align_Horizontal.Checked;

            if (ApplyRemove_chk_Apply_Align_Horizontal.Checked == false)
            {
                ApplyRemove_cmb_Apply_Align_Horizontal.Enabled = false;
            }
            else
            {
                ApplyRemove_cmb_Apply_Align_Horizontal.Enabled = true;
            }
        }

        private void ApplyRemove_cmb_Apply_Align_Horizontal_SelectedIndexChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_Align_HorizontalValue = ApplyRemove_cmb_Apply_Align_Horizontal.SelectedIndex;
        }

        private void ApplyRemove_chk_Apply_Zoom_CheckedChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_Zoom = ApplyRemove_chk_Apply_Zoom.Checked;

            if (ApplyRemove_chk_Apply_Zoom.Checked == false)
            {
                ApplyRemove_nud_Apply_Zoom.Enabled = false;
            }
            else
            {
                ApplyRemove_nud_Apply_Zoom.Enabled = true;
            }
        }

        private void ApplyRemove_nud_Apply_Zoom_ValueChanged(object sender, EventArgs e)
        {
            cl_Settings.ApplyRemove_Apply_ZoomValue = ApplyRemove_nud_Apply_Zoom.Value;
        }
    }
}
