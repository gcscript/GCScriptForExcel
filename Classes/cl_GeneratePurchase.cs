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
    public static class cl_GeneratePurchase
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
            else if (ws.Name.ToLower().Trim() != "compra")
            {
                if (MessageBox.Show("Esse script deve ser executado em uma aba de [COMPRAS]\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                {
                    return;
                }
            }

            try
            {
                app.ScreenUpdating = false;

                if (!InitialChecks()) { ResetFocus(); return; }

                MoveColumns();

                RemoveDuplicateRows();
                
                SortData();

                RemoveColumns();

                RemoveFillColumns();

                if (!GenerateSubtotal()) { ResetFocus(); return; }

                OrganizeSubtotal();

                if (!SeparatePurchases()) { ResetFocus(); return; }

                AdjustHideColumns();

                ResetFocus();

                void ResetFocus()
                {
                    app.ScreenUpdating = true;
                    ws.Cells[1, 1].Select();
                    cl_ExcelFunctions.AdjustScroll();
                }

                MessageBox.Show("Compra criada com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                app.ScreenUpdating = true;
                MessageBox.Show(erro.ToString(), "ERRO: 861680", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        static bool InitialChecks()
        {
            int usedColumns = ws.UsedRange.Columns.Count;

            string[] columnsName = { "Org1", "UF", "Empresa", "C.Unid", "Nome", "Operadora", "Total", "Desconto", "CompraFinal", "CNPJ + CPF + Operadora", "ORDEM", "OBS" };

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
            Range ColumnOrg1_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "ORG1");
            Range ColumnOrdem_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "ORDEM");
            Range ColumnNome_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "Nome");

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(Key: ColumnOrg1_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SortFields.Add(Key: ColumnOrdem_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SortFields.Add(Key: ColumnNome_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            ws.Sort.SetRange(ws.Cells);
            ws.Sort.Header = XlYesNoGuess.xlYes;
            ws.Sort.MatchCase = false;
            ws.Sort.Orientation = (XlSortOrientation)Constants.xlTopToBottom;
            ws.Sort.SortMethod = XlSortMethod.xlPinYin;
            ws.Sort.Apply();
        }

        static void RemoveColumns()
        {
            string[] nameColumns = { "ORG1", "Depto", "VvtNovo", "TvtNovo", "RecPend", "Saldo1", "Saldo", "ValorDias", "CNPJ + CPF + Operadora", "Buscador", "ORDEM", "CF -R$10", "Nr. do Cartao" };

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
            Range ColumnCompraFinal_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "CompraFinal");
            Range Column1Compra_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "1ª Compra");
            Range Column2Compra_Range = cl_ExcelFunctions.GetRangeColumnByName(ws, "2ª Compra");

            int ColumnCompraFinal_Number = ColumnCompraFinal_Range.Cells.Column;
            List<int> array_ColumnsSubtotal = new List<int>();

            if (Column1Compra_Range != null && Column2Compra_Range != null)
            {
                int Column1Compra_Number = Column1Compra_Range.Cells.Column;
                int Column2Compra_Number = Column2Compra_Range.Cells.Column;

                array_ColumnsSubtotal.Add(Column1Compra_Number);
                array_ColumnsSubtotal.Add(Column2Compra_Number);
                array_ColumnsSubtotal.Add(ColumnCompraFinal_Number);
            }
            else
            {
                array_ColumnsSubtotal.Add(ColumnCompraFinal_Number);
            }

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
                Range printArea = ws.Range[ws.Cells[1, 1], ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0]];
                cl_ExcelFunctions.SetBZPA(ws, printArea);
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
            cl_ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(ws, "C.Unid", "Total Geral", "<>");

            Range rngInicial = ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
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
                //else if (valorColunaOperadora == "total geral")
                //{
                //    ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (ws.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaOperadora.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 3);
                }
                //else if (valorColunaEmpresa == "total geral")
                //{
                //    ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (ws.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaEmpresa.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 2);
                }
                //else if (valorColunaCUNID == "total geral")
                //{
                //    ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (ws.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaCUNID.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 1);
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
                else if (valorColunaCUNID.Contains(" total"))
                {
                    Range rng_linha = ws.Range[ws.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    cl_ExcelFunctions.Styles_Emphasis(rng_linha, 1);
                }

                if (ws.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                offSetRow--;
            }
        }

        static bool SeparatePurchases()
        {
            bool warning = false;
            int ColumnNome_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Nome");
            int ColumnCompraFinal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CompraFinal");
            int ColumnOBS_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "OBS");

            int lastUsedRow = ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Row;

            int offSetRow = 0;
            int offSetColumn = 0;

            while (true)
            {
                Range actvCell = ws.Cells[lastUsedRow, ColumnCompraFinal_Number].Offset[offSetRow, offSetColumn];

                if (actvCell.Row < 2)
                {
                    break;
                }

                //-------------------[COLUMN OBS]-------------------
                string ColumnOBS_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueCurrentRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 0, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                //------------------[COLUMN NOME]-------------------
                string ColumnNome_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, -1, 0).ToLower();
                string ColumnNome_CellValueCurrentRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, 0, 0).ToLower();
                string ColumnNome_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, 1, 0).ToLower();

                if (ColumnOBS_CellValueCurrentRow == "inativo" || ColumnOBS_CellValueCurrentRow == "sem cadastro" || ColumnOBS_CellValueCurrentRow == "cpf ativo em outro comprador")
                {
                    if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        CheckInRows(actvCell, false, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        CheckInRows(actvCell, false, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else
                    {
                        MessageBox.Show("Existe uma probabilidade não calculada!", "ERRO: 859143", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                }
                else if (ColumnOBS_CellValueCurrentRow == "novo/sem cartao")
                {
                    if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                    }
                    else if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        CheckInRows(actvCell, false, false);
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, false, false);
                    }
                    else
                    {
                        MessageBox.Show("Existe uma probabilidade não calculada!", "ERRO: 550166", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                else
                {
                    if (ColumnNome_CellValueCurrentRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA]

                        string activeCell = actvCell.Text.Replace("-", "0").Trim(); // ATUAL
                        string activeCellOffset = actvCell.Offset[-1, 0].Text.Replace("-", "0").Trim(); // EM CIMA

                        if (double.TryParse(activeCell, out double aC))
                        {
                            if (double.TryParse(activeCellOffset, out double aCOffset))
                            {
                                if (aC > 0 && aCOffset == 0)
                                {
                                    // COLUNA COMPRA FINAL NA LINHA ATUAL É MAIOR QUE ZERO
                                    // COLUNA COMPRA FINAL NA LINHA [CIMA] DA ATUAL É ZERO
                                    CheckInRows(actvCell, false, false);
                                }
                            }
                        }
                    }
                }

                offSetRow--;
            }

            void CheckInRows(Range actvCell, bool abaixo, bool problema)
            {
                //-------------------[COLUNA OBS]-------------------
                string ColumnOBS_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                if (abaixo == true && problema == true)
                {
                    if (ColumnOBS_CellValueBottomRow != "inativo" && ColumnOBS_CellValueBottomRow != "sem cadastro" && ColumnOBS_CellValueBottomRow != "cpf ativo em outro comprador")
                    {
                        ws.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == true)
                {
                    if (ColumnOBS_CellValueTopRow != "inativo" && ColumnOBS_CellValueTopRow != "sem cadastro" && ColumnOBS_CellValueTopRow != "cpf ativo em outro comprador")
                    {
                        ws.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
                else if (abaixo == true && problema == false)
                {
                    if (ColumnOBS_CellValueBottomRow != "novo/sem cartao")
                    {
                        ws.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == false)
                {
                    if (ColumnOBS_CellValueTopRow != "novo/sem cartao")
                    {
                        ws.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
            }

            if (warning == false) { cl_ExcelFunctions.TabColor(ws, 3); }

            return true;
        }

        static bool SeparatePurchases_BK()
        {
            bool warning = false;
            int ColumnNome_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "Nome");
            int ColumnCompraFinal_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "CompraFinal");
            int ColumnOBS_Number = cl_ExcelFunctions.GetNumberColumnByName(ws, "OBS");

            int lastUsedRow = ws.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Row;

            int offSetRow = 0;
            int offSetColumn = 0;

            while (true)
            {
                Range actvCell = ws.Cells[lastUsedRow, ColumnCompraFinal_Number].Offset[offSetRow, offSetColumn];

                if (actvCell.Row < 2)
                {
                    break;
                }

                //-------------------[COLUMN OBS]-------------------
                string ColumnOBS_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueCurrentRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 0, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                //------------------[COLUMN NOME]-------------------
                string ColumnNome_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, -1, 0).ToLower();
                string ColumnNome_CellValueCurrentRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, 0, 0).ToLower();
                string ColumnNome_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnNome_Number, 1, 0).ToLower();

                if (ColumnOBS_CellValueCurrentRow == "inativo" || ColumnOBS_CellValueCurrentRow == "sem cadastro" || ColumnOBS_CellValueCurrentRow == "cpf ativo em outro comprador")
                {
                    if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        CheckInRows(actvCell, false, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        CheckInRows(actvCell, false, true);
                        cl_ExcelFunctions.Styles_Colors(cl_ExcelFunctions.GetRangeCell(ws, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { cl_ExcelFunctions.TabColor(ws, 5); warning = true; }
                    }
                    else
                    {
                        MessageBox.Show("Existe uma probabilidade não calculada!", "ERRO: 859143", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                }
                else if (ColumnOBS_CellValueCurrentRow == "novo/sem cartao")
                {
                    if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                    }
                    else if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        CheckInRows(actvCell, false, false);
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, false, false);
                    }
                    else
                    {
                        MessageBox.Show("Existe uma probabilidade não calculada!", "ERRO: 550166", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                }
                else
                {
                    if (ColumnNome_CellValueCurrentRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA]

                        string activeCell = actvCell.Text.Replace("-", "0").Trim(); // ATUAL
                        string activeCellOffset = actvCell.Offset[-1, 0].Text.Replace("-", "0").Trim(); // EM CIMA

                        if (double.TryParse(activeCell, out double aC))
                        {
                            if (double.TryParse(activeCellOffset, out double aCOffset))
                            {
                                if (aC > 0 && aCOffset == 0)
                                {
                                    // COLUNA COMPRA FINAL NA LINHA ATUAL É MAIOR QUE ZERO
                                    // COLUNA COMPRA FINAL NA LINHA [CIMA] DA ATUAL É ZERO
                                    CheckInRows(actvCell, false, false);
                                }
                            }
                        }
                    }
                }

                offSetRow--;
            }

            void CheckInRows(Range actvCell, bool abaixo, bool problema)
            {
                //-------------------[COLUNA OBS]-------------------
                string ColumnOBS_CellValueTopRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = cl_ExcelFunctions.GetCellText(ws, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                if (abaixo == true && problema == true)
                {
                    if (ColumnOBS_CellValueBottomRow != "inativo" && ColumnOBS_CellValueBottomRow != "sem cadastro" && ColumnOBS_CellValueBottomRow != "cpf ativo em outro comprador")
                    {
                        ws.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == true)
                {
                    if (ColumnOBS_CellValueTopRow != "inativo" && ColumnOBS_CellValueTopRow != "sem cadastro" && ColumnOBS_CellValueTopRow != "cpf ativo em outro comprador")
                    {
                        ws.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
                else if (abaixo == true && problema == false)
                {
                    if (ColumnOBS_CellValueBottomRow != "novo/sem cartao")
                    {
                        ws.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == false)
                {
                    if (ColumnOBS_CellValueTopRow != "novo/sem cartao")
                    {
                        ws.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
            }

            if (warning == false) { cl_ExcelFunctions.TabColor(ws, 3); }

            return true;
        }

        static void AdjustHideColumns()
        {
            string[] nameAdjustColumns = { "UF", "Operadora", "Empresa", "C.Unid" };
            string[] nameHideColumns = { "C.Depto", "CNPJ", "Escala", "RG", "Data Nasc.", "Desc", "Qvt1", "Vvt1", "Tvt1", "Total", "Desconto" };

            foreach (string nameAdjustColumn in nameAdjustColumns)
            {
                Range rng = cl_ExcelFunctions.GetRangeColumnByName(ws, nameAdjustColumn);

                if (rng != null)
                {
                    rng.ColumnWidth = 0.08;
                    continue;
                }
            }

            foreach (string nameHideColumn in nameHideColumns)
            {
                Range rng = cl_ExcelFunctions.GetRangeColumnByName(ws, nameHideColumn);

                if (rng != null)
                {
                    rng.EntireColumn.Hidden = true;
                    continue;
                }
            }
        }
    }
}
