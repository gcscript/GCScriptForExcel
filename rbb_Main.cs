using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using GCScript_for_Excel.Classes;

using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel
{
    public partial class rbb_Main
    {

        private void btn_Tools_DefinirAreaMex_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("BZPA");
            ExcelFunctions.SetBZPA(ws, app.Selection);
        }

        private void Info_btn_ColunasUsadas_Click(object sender, RibbonControlEventArgs e)
        {
            Tools.ColunasUsadas();

        }

        private void Info_btn_LinhasUsadas_Click(object sender, RibbonControlEventArgs e)
        {
            Tools.LinhasUsadas();

        }

        private void Info_btn_SelecionarRange_Click(object sender, RibbonControlEventArgs e)
        {
            Tools.SelecionarRange();
        }

        private void Info_btn_SelecionarTudo_Click(object sender, RibbonControlEventArgs e)
        {
            Tools.SelecionarTudo();
        }

        private void btn_T1_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;

            ExcelFunctions.RemoveRows(ws);
        }

        private void btn_T2_Click(object sender, RibbonControlEventArgs e)
        {
            var purchaseCreator = new PurchaseCreator();
            purchaseCreator.Start();
        }

        private void btn_T3_Click(object sender, RibbonControlEventArgs e)
        {
        }

        private void btn_T4_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                ExcelFunctions.SetColumnWidthByName(ws, ColumnsName.Nome, 50);
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
            gcsApplication app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                ExcelFunctions.RenameSheet("Compra", "Shopping");
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;

            MessageBox.Show(ws.Application.Worksheets.Count.ToString());
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Tools.ObterTipoSelecao(app.Selection);
        }

        private void Generate_btn_AdjustDescontoAndCompraFinal_Click(object sender, RibbonControlEventArgs e)
        {
            //Appl app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("AdjustDescontoAndCompraFinal");
            var adjustBalanceDaysValueColumns = new AdjustDescontoAndCompraFinal();
            adjustBalanceDaysValueColumns.Start();
        }

        private void Generate_btn_Apportionment_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("Apportionment");
            cl_GenerateApportionment.Start(app.ActiveSheet);
        }

        private void Generate_btn_Purchase_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("Purchase");
            cl_GeneratePurchase.Start(app.ActiveSheet);
        }

        private void Generate_btn_FileToSend_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("FileToSend");
            GenerateFileToSend.Start();
        }

        private void Converter_btn_Text_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("ConverterText");
            app.ScreenUpdating = false;
            try
            {
                Range selection = app.Selection;
                //Range constantsSelection = selection.SpecialCells(XlCellType.xlCellTypeConstants);
                //Range formulasSelection = selection.SpecialCells(XlCellType.xlCellTypeFormulas);
                //Range finalSelection = app.Union(constantsSelection, formulasSelection);

                if (selection.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 428083", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selection.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 978135", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterText(ws, selection);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("ConverterCPF");
            app.ScreenUpdating = false;
            try
            {
                Range selection = app.Selection;
                //Range constantsSelection = selection.SpecialCells(XlCellType.xlCellTypeConstants);
                //Range formulasSelection = selection.SpecialCells(XlCellType.xlCellTypeFormulas);
                //Range finalSelection = app.Union(constantsSelection, formulasSelection);

                if (selection.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 638734", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selection.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 978135", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                int count = Tools.ConverterCPF(ws, selection);

                MessageBox.Show($"CPF(s) alterado(s): {count}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("ConverterCNPJ");
            app.ScreenUpdating = false;
            try
            {
                Range selection = app.Selection;
                //Range constantsSelection = selection.SpecialCells(XlCellType.xlCellTypeConstants);
                //Range formulasSelection = selection.SpecialCells(XlCellType.xlCellTypeFormulas);
                //Range finalSelection = app.Union(constantsSelection, formulasSelection);

                if (selection.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 638734", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selection.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 978135", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterCNPJ(ws, selection);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("ConverterWorkSchedule");
            app.ScreenUpdating = false;
            try
            {
                Range selection = app.Selection;
                //Range constantsSelection = selection.SpecialCells(XlCellType.xlCellTypeConstants);
                //Range formulasSelection = selection.SpecialCells(XlCellType.xlCellTypeFormulas);
                //Range finalSelection = app.Union(constantsSelection, formulasSelection);

                if (selection.Cells.Count > 100000)
                {
                    MessageBox.Show("O Intervalo contém mais de 100.000 células.", "ERRO: 765843", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (selection.Cells.Count > 50000)
                {
                    if (MessageBox.Show("O Intervalo contém mais de 50.000 células.\nIsso pode travar sua aplicação!\nDeseja continuar?", "ERRO: 978135", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                    {
                        return;
                    }
                }

                cl_Settings.ConverterWorkSchedule(ws, selection);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesPrimary");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 1);
        }

        private void Styles_btn_Secondary_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesSecondary");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 2);
        }

        private void Styles_btn_Success_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesSuccess");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 3);
        }

        private void Styles_btn_Danger_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesDanger");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 4);
        }

        private void Styles_btn_Warning_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesWarning");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 5);
        }

        private void Styles_btn_Info_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesInfo");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 6);
        }

        private void Styles_glr_Bootstrap_Primary_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapPrimary");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 1);
        }

        private void Styles_glr_Bootstrap_Secondary_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapSecondary");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 2);
        }

        private void Styles_glr_Bootstrap_Success_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapSuccess");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 3);
        }

        private void Styles_glr_Bootstrap_Danger_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapDanger");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 4);
        }

        private void Styles_glr_Bootstrap_Warning_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapWarning");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 5);
        }

        private void Styles_glr_Bootstrap_Info_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapInfo");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 6);
        }

        private void Styles_glr_Bootstrap_Light_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapLight");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 7);
        }

        private void Styles_glr_Bootstrap_Dark_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapDark");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 8);
        }

        private void Styles_glr_Bootstrap_White_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesBootstrapWhite");
            ExcelFunctions.Styles_Bootstrap(app.Selection, 9);
        }

        private void Styles_glr_Emphasis1_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesEmphasis1");
            ExcelFunctions.Styles_Emphasis_OLD(app.Selection, 1);
        }

        private void Styles_glr_Emphasis2_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesEmphasis2");
            ExcelFunctions.Styles_Emphasis_OLD(app.Selection, 2);
        }

        private void Styles_glr_Emphasis3_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesEmphasis3");
            ExcelFunctions.Styles_Emphasis_OLD(app.Selection, 3);
        }

        private void Styles_glr_Emphasis4_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesEmphasis4");
            ExcelFunctions.Styles_Emphasis_OLD(app.Selection, 4);
        }

        private void Styles_glr_Emphasis5_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesEmphasis5");
            ExcelFunctions.Styles_Emphasis_OLD(app.Selection, 5);
        }

        private void Styles_btn_Default_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            ExcelFunctions.CreateBackup("StylesDefault");
            ExcelFunctions.Styles_Colors_OLD(app.Selection, 0);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("OnlyValues");
            try
            {
                app.ScreenUpdating = false;
                Range rng = app.Selection;

                if (cl_Settings.More_SelectionType == 0) // Selection
                {
                    ExcelFunctions.RemoveFormula(rng);
                }
                else if (cl_Settings.More_SelectionType == 1) // Active Sheet
                {
                    rng = ws.Cells;
                    ExcelFunctions.RemoveFormula(rng);
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
                        if (sheet.Visible == XlSheetVisibility.xlSheetHidden)
                            continue;

                        rng = sheet.Cells;
                        ExcelFunctions.RemoveFormula(rng);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("RemoveConditionalFormatting");
            try
            {
                app.ScreenUpdating = false;
                Range rng = app.Selection;

                if (cl_Settings.More_SelectionType == 0) // Selection
                {
                    ExcelFunctions.RemoveConditionalFormatting(rng);
                }
                else if (cl_Settings.More_SelectionType == 1) // Active Sheet
                {
                    rng = ws.Cells;
                    ExcelFunctions.RemoveConditionalFormatting(rng);
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
                        ExcelFunctions.RemoveConditionalFormatting(rng);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("ApplyRemove");
            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                if (cl_Settings.ApplyRemove_Apply_AllSheets == true)
                {
                    if (MessageBox.Show("Essa função irá Aplicar/Remover configurações definidas em todas as planilhas!\nIsso pode demorar dependendo da quantidade de dados.\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                    {
                        return;
                    }
                }
                ExcelFunctions.ApplyRemove(ws);
                MessageBox.Show("Operação efetuada com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                app.ScreenUpdating = true;
                MessageBox.Show(erro.Message, "ERRO: 255869", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = "General";
        }

        private void Number_btn_Text_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = "@";
        }

        private void Number_btn_Accounting_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Range rng = app.Selection;
            rng.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
        }

        private void Sort_btn_Preset01_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            ExcelFunctions.CreateBackup("SortPreset1");
            List<string> lst_SortDataColumns = new List<string>() { ColumnsName.ArquivoDeCompra, ColumnsName.Uf, ColumnsName.Operadora, ColumnsName.Nome };
            ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset02_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            ExcelFunctions.CreateBackup("SortPreset2");
            List<string> lst_SortDataColumns = new List<string>() { ColumnsName.ArquivoDeCompra, ColumnsName.Uf, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.Nome };
            ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset03_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            ExcelFunctions.CreateBackup("SortPreset3");
            List<string> lst_SortDataColumns = new List<string>() { ColumnsName.ArquivoDeCompra, ColumnsName.Uf, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.CDepto, "Depto", ColumnsName.Nome };
            ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sort_btn_Preset04_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet workSheet = app.ActiveSheet;
            ExcelFunctions.CreateBackup("SortPreset4");
            List<string> lst_SortDataColumns = new List<string>() { ColumnsName.ArquivoDeCompra, ColumnsName.Uf, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.CDepto, "Depto", ColumnsName.Nome };
            ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);
        }

        private void Sheets_btn_SortSheetsASC_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                ExcelFunctions.CreateBackup("SortSheetsASC");
                ExcelFunctions.SheetsOrderBy();
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
            gcsApplication app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                ExcelFunctions.CreateBackup("SortSheetsAESC");
                ExcelFunctions.SheetsOrderBy(true);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("Round0DecimalPlaces");
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

                Tools.ConverterDecimalPlaces(ws, selecao, 0);
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
        private void Converter_btn_Round_2DecimalPlaces_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("Round2DecimalPlaces");
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

                Tools.ConverterDecimalPlaces(ws, selecao, 2);
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
            gcsApplication app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                int count = ExcelFunctions.RemoveHiddenSheets();
                MessageBox.Show(string.Format("Planilhas ocultas removidas: {0}", count), "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            gcsApplication app = Globals.ThisAddIn.Application;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                int count = ExcelFunctions.ShowHiddenSheets();
                MessageBox.Show(string.Format("Planilhas desocultas: {0}", count), "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            gcsApplication app = Globals.ThisAddIn.Application;
            Worksheet ws = app.ActiveSheet;
            ExcelFunctions.CreateBackup("RandomCPF");
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

                Tools.Generator_CPF(ws, selecao);
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
        private void glr_Get_Content_Click(object sender, RibbonControlEventArgs e)
        {

            gcsApplication app = Globals.ThisAddIn.Application;
            Range activeCell = app.ActiveCell;
            decimal Value2 = 0;
            _ = decimal.TryParse(activeCell.Value2.ToString(), out Value2);

            MessageBox.Show($"Value: [{activeCell.Value.ToString()}]\n" +
                $"Value2: [{activeCell.Value2.ToString()}]\n" +
                $"Value2 Round: [{Math.Round(Value2, 2)}]\n" +
                $"Text: [{activeCell.Text.ToString()}]\n" +
                $"Formula: [{activeCell.Formula.ToString()}]\n" +
                //$"FormulaArray: [{activeCell.FormulaArray.ToString()}]\n" +
                $"FormulaLocal: [{activeCell.FormulaLocal.ToString()}]\n" +
                $"FormulaR1C1: [{activeCell.FormulaR1C1.ToString()}]\n" +
                $"FormulaR1C1Local: [{activeCell.FormulaR1C1Local.ToString()}]\n");
        }

        private void glr_Get_Value_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            var activeCell = app.ActiveCell.Value;
            MessageBox.Show(activeCell.ToString());
        }

        private void glr_Get_Value2_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            var activeCell = app.ActiveCell.Value2;
            MessageBox.Show(activeCell.ToString());
        }

        private void glr_Get_Text_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Range activeCell = app.ActiveCell;
            MessageBox.Show(activeCell.Text.ToString());
        }

        private void btn_IsNumeric_Click(object sender, RibbonControlEventArgs e)
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            (bool isNumeric, bool isNull, decimal value) teste = ExcelFunctions.IsNumeric(app.ActiveCell);
            MessageBox.Show($"Is Null? {teste.isNull}\nIs Numeric? {teste.isNumeric}\nValue: {teste.value}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
        }

        private void TransferData_btn_Import_Click(object sender, RibbonControlEventArgs e)
        {
            TransferData transferData = new TransferData();
            transferData.Import();
        }

        private void TransferData_btn_Export_Click(object sender, RibbonControlEventArgs e)
        {
            TransferData transferData = new TransferData();
            transferData.Export();
        }
    }
}
