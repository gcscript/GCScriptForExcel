﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Appl = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class cl_GenerateApportionment
    {
        static Appl app = Globals.ThisAddIn.Application;
        static Worksheet ws;

        public static void Start(Worksheet worksheet)
        {
            ws = worksheet;

            if (ws.Name.ToLower().Trim() == "dados")
            {
                MessageBox.Show("Esse script não pode ser executado em uma aba [DADOS]!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (ws.Name.ToLower().Trim() != "rateio")
            {
                if (MessageBox.Show("Esse script deve ser executado na aba [RATEIO]\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                {
                    return;
                }
            }

            try
            {
                app.ScreenUpdating = false;

                if (!InitialChecks()) { ResetFocus(); return; }

                MoveColumns();

                //RemoveDuplicateRows();

                SortData();

                RemoveColumns();

                //RemoveFillColumns();

                if (!GenerateSubtotal()) { ResetFocus(); return; }

                OrganizeSubtotal();

                RemoveTotalInCUnid();

                AdjustColumnsWidth();

                ResetFocus();

                void ResetFocus()
                {
                    app.ScreenUpdating = true;
                    ws.Cells[1, 1].Select();
                    cl_ExcelFunctions.AdjustScroll();
                }

                MessageBox.Show("Rateio criado com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                app.ScreenUpdating = true;
                MessageBox.Show(erro.ToString(), "ERRO: 981933", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        static bool InitialChecks()
        {
            int usedColumns = ws.UsedRange.Columns.Count;

            string[] columnsName = { "UF", "Empresa", "C.Unid", "Operadora", "Total", "Desconto", "CompraFinal" };

            foreach (string columnName in columnsName)
            {
                if (CheckColumnExistence(columnName) == false)
                {
                    MessageBox.Show("A coluna [" + columnName.Trim().ToUpper() + "] não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            bool CheckColumnExistence(string columnName)
            {
                columnName = columnName.Trim().ToLower();
                Range rng = ws.Range[app.Cells[1, 1], app.Cells[1, usedColumns]].Find(What: columnName, LookAt: XlLookAt.xlWhole, MatchCase: false);
                if (rng == null) { return false; }
                return true;
            }

            return true;
        }

        static void MoveColumns()
        {
            int ColumnUF_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "UF");
            cl_ExcelFunctions.GetRangeColumn(ws, ColumnUF_Number).Cut();
            cl_ExcelFunctions.GetRangeColumn(ws, 1).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnOperadora_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Operadora");
            cl_ExcelFunctions.GetRangeColumn(ws, ColumnOperadora_Number).Cut();
            cl_ExcelFunctions.GetRangeColumn(ws, 2).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnEmpresa_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Empresa");
            cl_ExcelFunctions.GetRangeColumn(ws, ColumnEmpresa_Number).Cut();
            cl_ExcelFunctions.GetRangeColumn(ws, 3).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnCUnid_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "C.UNID");
            cl_ExcelFunctions.GetRangeColumn(ws, ColumnCUnid_Number).Cut();
            cl_ExcelFunctions.GetRangeColumn(ws, 4).Insert(XlInsertShiftDirection.xlShiftToRight);

            Range ColumnCDepto_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "C.DEPTO");

            if (ColumnCDepto_Range != null)
            {
                ws.Columns[ColumnCDepto_Range.Cells.Column].Cut();
                ws.Columns[5].Insert(XlInsertShiftDirection.xlShiftToRight); // Shift:=xlToRight
                //colunaCDeptoExiste = true;
            }
        }

        static void RemoveDuplicateRows()
        {
            int ColumnCnpjCpfOperadora_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CNPJ + CPF + Operadora");

            Range rngInicial = ws.Cells[1048576, ColumnCnpjCpfOperadora_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
                Range rngAtual = ws.Cells[linha, ColumnCnpjCpfOperadora_Number].Offset[offSetRow, 0];

                if (rngAtual.Row < 2)
                {
                    break;
                }
                else
                {
                    if (rngAtual.Value == rngAtual.Offset[-1, 0].Value)
                    {
                        linha = rngAtual.Row;
                        rngAtual.EntireRow.Delete();
                        offSetRow = 0;
                        continue;
                    }
                }

                offSetRow--;
            }
        }

        static void SortData()
        {
            Range ColumnUF_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "UF");
            Range ColumnOperadora_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "Operadora");
            Range ColumnEmpresa_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "Empresa");
            Range ColumnCUnid_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "C.Unid");

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(Key: ColumnUF_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SortFields.Add(Key: ColumnOperadora_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SortFields.Add(Key: ColumnEmpresa_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SortFields.Add(Key: ColumnCUnid_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SetRange(ws.Cells);
            ws.Sort.Header = XlYesNoGuess.xlYes;
            ws.Sort.MatchCase = false;
            ws.Sort.Orientation = (XlSortOrientation)Constants.xlTopToBottom;
            ws.Sort.SortMethod = XlSortMethod.xlPinYin;
            ws.Sort.Apply();
        }

        static void RemoveColumns()
        {
            string[] nameColumns = { "ORG1", "CNPJ", "Depto", "Escala", "ID", "Mat", "Mat Site", "Nome", "CPF", "RG", "Data Nasc.", "Desc", 
                                     "Qvt1", "Vvt1", "Tvt1", "VvtNovo", "TvtNovo", "RecPend", "Saldo1", "Saldo", "ValorDias", "1ª Compra", "2ª Compra", 
                                     "Tipo1", "CNPJ + CPF + Operadora", "Buscador", "ORDEM", "CF -R$10", "Nr. do Cartao", "OBS"};

            foreach (string nameColumn in nameColumns)
            {
                while (true)
                {
                    Range rng = cl_ExcelFunctions.GetRangeColumnByName(ws, nameColumn);

                    if (rng != null)
                    {
                        rng.EntireColumn.Delete();
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        static void RemoveFillColumns()
        {
            Range ColumnTotal_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "Total");
            Range ColumnDesconto_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "Desconto");
            Range ColumnCompraFinal_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "CompraFinal");
            Range Column1Compra_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "1ª Compra");
            Range Column2Compra_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "2ª Compra");

            RemoveFill(ColumnTotal_Range);
            RemoveFill(ColumnDesconto_Range);
            RemoveFill(ColumnCompraFinal_Range);
            if (Column1Compra_Range != null) { RemoveFill(Column1Compra_Range); }
            if (Column2Compra_Range != null) { RemoveFill(Column2Compra_Range); }

            void RemoveFill(Range rng)
            {
                Range range = ws.Columns[rng.Column];
                range.Interior.Pattern = Constants.xlNone;
                range.Interior.TintAndShade = 0;
                range.Interior.PatternTintAndShade = 0;
            }
        }

        static bool GenerateSubtotal()
        {
            int ColumnTotal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Total");
            int ColumnDesconto_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Desconto");
            int ColumnCompraFinal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CompraFinal");

            List<int> array_ColumnsSubtotal = new List<int>();

            array_ColumnsSubtotal.Add(ColumnTotal_Number);
            array_ColumnsSubtotal.Add(ColumnDesconto_Number);
            array_ColumnsSubtotal.Add(ColumnCompraFinal_Number);

            Range rangeUF = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0]];
            rangeUF.Subtotal(GroupBy: 1, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeOperadora = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-1, 0]];
            rangeOperadora.Subtotal(GroupBy: 2, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeEmpresa = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-3, 0]];
            rangeEmpresa.Subtotal(GroupBy: 3, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeCUNID = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-5, 0]];
            rangeCUNID.Subtotal(GroupBy: 4, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);


            { // COPIAR E COLAR COMO VALOR | REMOVER SUBTOTAL
                Range selecao = ws.Cells;
                cl_Tools.CopiarColarValor(selecao);
                app.Application.CutCopyMode = 0;
                selecao.RemoveSubtotal();
            }

            { // DEFINIR ÁREA DE IMPRESSÃO | BORDAS | ZOOM
                Range areaDeImpressao = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0]];
                cl_ExcelFunctions.SetBZPA(ws, areaDeImpressao);
            }

            return true;
        }

        static void OrganizeSubtotal()
        {
            int ColumnUF_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "UF");
            int ColumnOperadora_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Operadora");
            int ColumnEmpresa_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Empresa");
            int ColumnCUnid_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "C.Unid");
            int ColumnCompraFinal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CompraFinal");

            cl_ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(ws, "Operadora", "Total Geral", "<>");
            cl_ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(ws, "Empresa", "Total Geral", "<>");
            //cl_ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(ws, "C.Unid", "Total", "<>");
            cl_ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(ws, "C.Unid", "<>*total", "<>");

            Range rngInicial = ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
                if (ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                string valorColunaUF = ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaOperadora = ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaEmpresa = ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaCUNID = ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Text.Trim().ToLower();

                if (valorColunaUF == "total geral")
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 5);
                }
                else if (valorColunaUF.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 4);

                }
                else if (valorColunaOperadora.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 3);
                }
                else if (valorColunaEmpresa.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 2);
                }
                else if (valorColunaCUNID.EndsWith(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.FontBold(rng_linha, false);
                }

                if (ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                offSetRow--;
            }
        }

        static void OrganizeSubtotal_BK()
        {
            int ColumnUF_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "UF");
            int ColumnOperadora_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Operadora");
            int ColumnEmpresa_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Empresa");
            int ColumnCUnid_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "C.Unid");
            int ColumnCompraFinal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CompraFinal");

            Range rngInicial = ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
                if (ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                string valorColunaUF = ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaOperadora = ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaEmpresa = ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaCUNID = ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Text.Trim().ToLower();

                if (valorColunaUF == "total geral")
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 5);
                }
                else if (valorColunaUF.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 4);

                }
                else if (valorColunaOperadora == "total geral")
                {
                    ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].EntireRow.Delete();
                    linha = (ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Row) - 1;
                    offSetRow = 0;
                    continue;
                }
                else if (valorColunaOperadora.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 3);
                }
                else if (valorColunaEmpresa == "total geral")
                {
                    ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].EntireRow.Delete();
                    linha = (ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Row) - 1;
                    offSetRow = 0;
                    continue;
                }
                else if (valorColunaEmpresa.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 2);
                }
                else if (valorColunaCUNID == "total geral")
                {
                    ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].EntireRow.Delete();
                    linha = (ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Row) - 1;
                    offSetRow = 0;
                    continue;
                }
                else if (valorColunaCUNID.EndsWith(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.FontBold(rng_linha, false);
                }
                else if (valorColunaCUNID != "" && valorColunaCUNID != "total geral" && !valorColunaCUNID.EndsWith(" total"))
                {
                    ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].EntireRow.Delete();
                    linha = (ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Row) - 1;
                    offSetRow = 0;
                    continue;
                }

                if (ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                offSetRow--;
            }
        }

        static void AdjustColumnsWidth()
        {
            string[] nameAdjustColumns = { "UF", "Operadora", "Empresa", "C.Unid", "Total", "Desconto", "CompraFinal" };

            foreach (string nameAdjustColumn in nameAdjustColumns)
            {
                Range rng = cl_ExcelFunctions.GetRangeColumnByName(ws, nameAdjustColumn);

                if (rng != null)
                {

                    if (nameAdjustColumn == "C.Unid")
                    {
                        rng.EntireColumn.AutoFit();
                        if (rng.ColumnWidth < 30) { rng.ColumnWidth = 30; }
                    }
                    else if (nameAdjustColumn == "Total" || nameAdjustColumn == "Desconto" || nameAdjustColumn == "CompraFinal")
                    {
                        rng.EntireColumn.AutoFit();
                        if (rng.ColumnWidth < 12) { rng.ColumnWidth = 12; }
                    }
                    else
                    {
                        rng.ColumnWidth = 0.08;
                    }
                    continue;
                }
            }
        }

        static void RemoveTotalInCUnid()
        {
            int ColumnCUnid_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "C.Unid");

            Range ColumnCUnid_Range = ws.Range[ws.Cells[2, ColumnCUnid_Number], ws.Cells[1048576, ColumnCUnid_Number].End(XlDirection.xlUp)];

            foreach (Range row in ColumnCUnid_Range.Cells)
            {
                string text = row.Text;
                if (text.ToLower().EndsWith("total"))
                {
                    row.Value = text.Substring(0, text.Length - 6);
                }
            }
        }
    }
}
