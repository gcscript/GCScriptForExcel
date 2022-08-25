using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GCScript_for_Excel.Classes;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Appl = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel
{
    public partial class rbb_Main
    {

        private void btn_Tools_DefinirAreaMex_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.SetBZPA(ws, app.Selection);
        }

        private void Info_btn_ColunasUsadas_Click(object sender, RibbonControlEventArgs e)
        {
            cl_Tools.ColunasUsadas();

        }

        private void Info_btn_LinhasUsadas_Click(object sender, RibbonControlEventArgs e)
        {
            cl_Tools.LinhasUsadas();

        }

        private void Info_btn_SelecionarRange_Click(object sender, RibbonControlEventArgs e)
        {
            cl_Tools.SelecionarRange();
        }

        private void Info_btn_SelecionarTudo_Click(object sender, RibbonControlEventArgs e)
        {
            cl_Tools.SelecionarTudo();
        }

        private void btn_T1_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            MessageBox.Show(cl_ExcelFunctions.GetCellInfo(app.ActiveCell));
        }

        private void btn_T2_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;

        }

        private void btn_T3_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void btn_T4_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                cl_ExcelFunctions.SetColumnWidthByName(ws, "Nome", 50);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 117089", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }

        private void btn_T5_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                cl_ExcelFunctions.RenameSheet("Compra", "Shopping");
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 626819", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }

        private void btn_T6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void rbb_Main_Load(object sender, RibbonUIEventArgs e)
        {
            //List<string> anos = new List<string>();
            //anos.Add(DateTime.Now.AddYears(-1).Year.ToString());
            //anos.Add(DateTime.Now.Year.ToString());
            //anos.Add(DateTime.Now.AddYears(1).Year.ToString());

            //foreach (string ano in anos)
            //{
            //    RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
            //    item.Label = ano;
            //    data_drd_Ano.Items.Add(item);
            //}

            //data_drd_Ano.SelectedItemIndex = 1;
            //More_btn_CheckActiveSheet.Image = GCScript_for_Excel.Properties.Resources.check;
        }

        private void Tools_btn_SetTitle_Click(object sender, RibbonControlEventArgs e)
        {
            frm_SetTitle frm = new frm_SetTitle();
            frm.ShowDialog();
        }

        private void Info_btn_ObterTipoSelecao_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_Tools.ObterTipoSelecao(app.Selection);
        }

        private void Generate_btn_Apportionment_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_GenerateApportionment.Start(app.ActiveSheet);
        }

        private void Generate_btn_Purchase_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_GeneratePurchase.Start(app.ActiveSheet);
        }

        private void Generate_btn_FileToSend_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_GenerateFileToSend.Start();
        }

        private void Converter_btn_Text_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 428083", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 978135", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterText(ws, selecao);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 325412", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_CPF_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 306904", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 452149", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Tools.ConverterCPF(ws, selecao);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 872708", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_CNPJ_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 549661", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 105067", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterCNPJ(ws, selecao);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 737030", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_WorkSchedule_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 491774", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 138689", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterWorkSchedule(ws, selecao);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 325412", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_Settings_Click(object sender, RibbonControlEventArgs e)
        {
            frm_Settings frm = new frm_Settings();
            frm.tabPage = 1;
            frm.ShowDialog();
        }

        private void btn_Settings_Click(object sender, RibbonControlEventArgs e)
        {
            frm_Settings frm = new frm_Settings();
            frm.tabPage = 0;
            frm.ShowDialog();
        }

        private void Styles_btn_Primary_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 1);
        }

        private void Styles_btn_Secondary_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 2);
        }

        private void Styles_btn_Success_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 3);
        }

        private void Styles_btn_Danger_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 4);
        }

        private void Styles_btn_Warning_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 5);
        }

        private void Styles_btn_Info_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 6);
        }

        private void Styles_glr_Bootstrap_Primary_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 1);
        }

        private void Styles_glr_Bootstrap_Secondary_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 2);
        }

        private void Styles_glr_Bootstrap_Success_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 3);
        }

        private void Styles_glr_Bootstrap_Danger_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 4);
        }

        private void Styles_glr_Bootstrap_Warning_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 5);
        }

        private void Styles_glr_Bootstrap_Info_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 6);
        }

        private void Styles_glr_Bootstrap_Light_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 7);
        }

        private void Styles_glr_Bootstrap_Dark_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 8);
        }

        private void Styles_glr_Bootstrap_White_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Bootstrap(app.Selection, 9);
        }

        private void Styles_glr_Emphasis1_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Emphasis(app.Selection, 1);
        }

        private void Styles_glr_Emphasis2_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Emphasis(app.Selection, 2);
        }

        private void Styles_glr_Emphasis3_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Emphasis(app.Selection, 3);
        }

        private void Styles_glr_Emphasis4_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Emphasis(app.Selection, 4);
        }

        private void Styles_glr_Emphasis5_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Emphasis(app.Selection, 5);
        }

        private void Styles_btn_Default_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            cl_ExcelFunctions.CreateBackup();
            cl_ExcelFunctions.Styles_Colors(app.Selection, 0);
        }

        private void glr_More_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            string selectiontype = "";

            More_btn_CheckSelection.Image = GCScript_for_Excel.Properties.Resources.uncheck;
            More_btn_CheckActiveSheet.Image = GCScript_for_Excel.Properties.Resources.uncheck;
            More_btn_CheckAllSheets.Image = GCScript_for_Excel.Properties.Resources.uncheck;

            if (cl_Settings.More_SelectionType == 0)
            {
                selectiontype = "Selection";
                More_btn_CheckSelection.Image = GCScript_for_Excel.Properties.Resources.check;
            }
            else if (cl_Settings.More_SelectionType == 1)
            {
                selectiontype = "Active Sheet";
                More_btn_CheckActiveSheet.Image = GCScript_for_Excel.Properties.Resources.check;
            }
            else if (cl_Settings.More_SelectionType == 2)
            {
                selectiontype = "All Sheets";
                More_btn_CheckAllSheets.Image = GCScript_for_Excel.Properties.Resources.check;
            }
            else
            {
                selectiontype = "ERRO: 307714";
            }

            More_btn_OnlyValues.Label = "Only Values (" + selectiontype + ")";
            More_btn_RemoveConditionalFormatting.Label = "Remove Conditional Formatting (" + selectiontype + ")";
        }

        private void More_btn_OnlyValues_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            try
            {
                app.ScreenUpdating = false;
                Range rng = app.Selection;

                if (cl_Settings.More_SelectionType == 0) // Selection
                {
                    cl_ExcelFunctions.RemoveFormula(rng);
                }
                else if (cl_Settings.More_SelectionType == 1) // Active Sheet
                {
                    rng = ws.Cells;
                    cl_ExcelFunctions.RemoveFormula(rng);
                    app.Goto(ws.Range["A1"], true);
                }
                else if (cl_Settings.More_SelectionType == 2) // All Sheet
                {
                    if (MessageBox.Show("Essa função pode demorar um pouco!\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                    {
                        return;
                    }

                    foreach (Worksheet sheet in app.ActiveWorkbook.Worksheets)
                    {
                        rng = sheet.Cells;
                        cl_ExcelFunctions.RemoveFormula(rng);
                        app.Goto(sheet.Range["A1"], true);
                    }
                }
                else
                {
                    MessageBox.Show("Aconteceu um erro!", "ERRO: 615369", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ws.Select();

                MessageBox.Show("Valor(es) convertido(s) com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 869460", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.CutCopyMode = (XlCutCopyMode)0;
                app.ScreenUpdating = true;
            }
        }

        private void More_btn_RemoveConditionalFormatting_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            try
            {
                app.ScreenUpdating = false;
                Range rng = app.Selection;

                if (cl_Settings.More_SelectionType == 0) // Selection
                {
                    cl_ExcelFunctions.RemoveConditionalFormatting(rng);
                }
                else if (cl_Settings.More_SelectionType == 1) // Active Sheet
                {
                    rng = ws.Cells;
                    cl_ExcelFunctions.RemoveConditionalFormatting(rng);
                    app.Goto(ws.Range["A1"], true);
                }
                else if (cl_Settings.More_SelectionType == 2) // All Sheet
                {
                    if (MessageBox.Show("Essa função pode demorar um pouco!\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                    {
                        return;
                    }

                    foreach (Worksheet sheet in app.ActiveWorkbook.Worksheets)
                    {
                        rng = sheet.Cells;
                        cl_ExcelFunctions.RemoveConditionalFormatting(rng);
                        app.Goto(sheet.Range["A1"], true);
                    }
                }
                else
                {
                    MessageBox.Show("Aconteceu um erro!", "ERRO: 569013", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                ws.Select();

                MessageBox.Show("Formatação Condicional removida(s) com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 492360", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void More_btn_CheckSelection_Click(object sender, RibbonControlEventArgs e)
        {
            More_SelectionType(0);
        }

        private void More_btn_CheckActiveSheet_Click(object sender, RibbonControlEventArgs e)
        {
            More_SelectionType(1);
        }

        private void More_btn_CheckAllSheets_Click(object sender, RibbonControlEventArgs e)
        {
            More_SelectionType(2);
        }

        private void More_SelectionType(int selectiontype)
        {
            cl_Settings.More_SelectionType = selectiontype;

            More_btn_CheckSelection.Image = GCScript_for_Excel.Properties.Resources.uncheck;
            More_btn_CheckActiveSheet.Image = GCScript_for_Excel.Properties.Resources.uncheck;
            More_btn_CheckAllSheets.Image = GCScript_for_Excel.Properties.Resources.uncheck;

            if (cl_Settings.More_SelectionType == 0)
            {
                More_btn_OnlyValues.Label = "Only Values (Selection)";
                More_btn_RemoveConditionalFormatting.Label = "Remove Conditional Formatting (Selection)";
                More_btn_CheckSelection.Image = GCScript_for_Excel.Properties.Resources.check;
            }
            else if (cl_Settings.More_SelectionType == 1)
            {
                More_btn_OnlyValues.Label = "Only Values (Active Sheet)";
                More_btn_RemoveConditionalFormatting.Label = "Remove Conditional Formatting (Active Sheet)";
                More_btn_CheckActiveSheet.Image = GCScript_for_Excel.Properties.Resources.check;
            }
            else if (cl_Settings.More_SelectionType == 2)
            {
                More_btn_OnlyValues.Label = "Only Values (All Sheets)";
                More_btn_RemoveConditionalFormatting.Label = "Remove Conditional Formatting (All Sheets)";
                More_btn_CheckAllSheets.Image = GCScript_for_Excel.Properties.Resources.check;
            }

        }

        private void ApplyRemove_btn_Start_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            try
            {
                app.ScreenUpdating = false;
                if (cl_Settings.ApplyRemove_Apply_AllSheets == true)
                {
                    if (MessageBox.Show("Essa função irá Aplicar/Remover configurações definidas em todas as planilhas!\nIsso pode demorar dependendo da quantidade de dados.\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                    {
                        return;
                    }
                }
                cl_ExcelFunctions.ApplyRemove(ws);
                app.ScreenUpdating = true;
                MessageBox.Show("Operação efetuada com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                app.ScreenUpdating = true;
                MessageBox.Show(erro.Message, "ERRO: 255869", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyRemove_btn_Settings_Click(object sender, RibbonControlEventArgs e)
        {
            frm_Settings frm = new frm_Settings();
            frm.tabPage = 2;
            frm.ShowDialog();
        }

        private void Number_btn_General_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = "General";
        }

        private void Number_btn_Text_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = "@";
        }

        private void Number_btn_Accounting_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
        }

        private void Sort_btn_Preset01_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            List<string> lst_SortDataColumns = new List<string>() { "UF", "Operadora", "Nome" };
            cl_ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset02_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            List<string> lst_SortDataColumns = new List<string>() { "UF", "Operadora", "Empresa", "C.Unid", "Nome" };
            cl_ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset03_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            List<string> lst_SortDataColumns = new List<string>() { "UF", "Operadora", "Empresa", "C.Unid", "C.Depto", "Depto", "Nome" };
            cl_ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset04_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            List<string> lst_SortDataColumns = new List<string>() { "UF", "Operadora", "Empresa", "C.Unid", "C.Depto", "Depto", "Nome" };
            cl_ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sheets_btn_SortSheetsASC_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                cl_ExcelFunctions.CreateBackup();
                cl_ExcelFunctions.SheetsOrderBy();
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 709980", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Sheets_btn_SortSheetsDESC_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                cl_ExcelFunctions.CreateBackup();
                cl_ExcelFunctions.SheetsOrderBy(true);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 692050", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_Round_0DecimalPlaces_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 904976", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 656732", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Tools.ConverterDecimalPlaces(ws, selecao, 0);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 268639", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_Round_1DecimalPlaces_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 128527", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 958401", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Tools.ConverterDecimalPlaces(ws, selecao, 1);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 946717", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Converter_btn_Round_2DecimalPlaces_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 138941", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 735773", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Tools.ConverterDecimalPlaces(ws, selecao, 2);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 434892", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        private void Sheets_btn_RemoveHiddenSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                cl_ExcelFunctions.RemoveHiddenSheets();
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 411220", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }

        private void Sheets_btn_ShowHiddenSheets_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                cl_ExcelFunctions.ShowHiddenSheets();
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 166642", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }
        }

        private void Generate_btn_RandomCPF_Click(object sender, RibbonControlEventArgs e)
        {
            Appl app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            cl_ExcelFunctions.CreateBackup();
            app.ScreenUpdating = false;
            try
            {
                Range selecao = app.Selection;

                if (selecao.Cells.Count > 1000)
                {
                    MessageBox.Show("O Intervalo contém mais de 1000 células.", "ERRO: 653774", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selecao.Cells.Count > 500)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 500 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Tools.Generator_CPF(ws, selecao);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 268639", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }


    }
}
