using System;
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
    public class cl_GeneratePurchase_new
    {
        public Appl gApp { private get; set; }
        public Worksheet gWs { private get; set; }

        public void Start()
        {
            if (gWs.Name.ToLower().Trim() == "dados")
            {
                MessageBox.Show("Esse script não pode ser executado em uma aba [DADOS]!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else if (gWs.Name.ToLower().Trim() != "compra")
            {
                if (MessageBox.Show("Esse script deve ser executado em uma aba de [COMPRAS]\nDeseja continuar?", "ATENÇÃO!", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                {
                    return;
                }
            }

            try
            {
                gApp.ScreenUpdating = false;

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
                    gApp.ScreenUpdating = true;
                    gWs.Cells[1, 1].Select();
                    ExcelFunctions.AdjustScroll();
                }

                MessageBox.Show("Compra criada com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                gApp.ScreenUpdating = true;
                MessageBox.Show(erro.ToString(), "ERRO: 861680", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        bool InitialChecks()
        {
            int usedColumns = gWs.UsedRange.Columns.Count;

            string[] columnsName = { ColumnsName.Org, ColumnsName.Uf, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.Nome, 
                                     ColumnsName.Operadora, ColumnsName.Total, ColumnsName.Desconto, ColumnsName.CompraFinal, ColumnsName.CnpjCpfOperadora, 
                                     ColumnsName.Ordem, ColumnsName.Obs };

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
                Range rng = gWs.Range[gApp.Cells[1, 1], gApp.Cells[1, usedColumns]].Find(What: columnName, LookAt: XlLookAt.xlWhole, MatchCase: false);
                if (rng == null) { return false; }
                return true;
            }

            return true;
        }

        void MoveColumns()
        {
            int ColumnUF_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Uf);
            ExcelFunctions.GetRangeColumn(gWs, ColumnUF_Number).Cut();
            ExcelFunctions.GetRangeColumn(gWs, 1).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnOperadora_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Operadora);
            ExcelFunctions.GetRangeColumn(gWs, ColumnOperadora_Number).Cut();
            ExcelFunctions.GetRangeColumn(gWs, 2).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnEmpresa_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Empresa);
            ExcelFunctions.GetRangeColumn(gWs, ColumnEmpresa_Number).Cut();
            ExcelFunctions.GetRangeColumn(gWs, 3).Insert(XlInsertShiftDirection.xlShiftToRight);

            int ColumnCUnid_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.CUnid);
            ExcelFunctions.GetRangeColumn(gWs, ColumnCUnid_Number).Cut();
            ExcelFunctions.GetRangeColumn(gWs, 4).Insert(XlInsertShiftDirection.xlShiftToRight);

            Range ColumnCDepto_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.CDepto);

            if (ColumnCDepto_Range != null)
            {
                gWs.Columns[ColumnCDepto_Range.Cells.Column].Cut();
                gWs.Columns[5].Insert(XlInsertShiftDirection.xlShiftToRight); // Shift:=xlToRight
                //colunaCDeptoExiste = true;
            }
        }

        void RemoveDuplicateRows()
        {
            int ColumnCnpjCpfOperadora_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.CnpjCpfOperadora);

            Range rngInicial = gWs.Cells[1048576, ColumnCnpjCpfOperadora_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
                Range rngAtual = gWs.Cells[linha, ColumnCnpjCpfOperadora_Number].Offset[offSetRow, 0];

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

        void SortData()
        {
            Range ColumnOrg1_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Org);
            Range ColumnOrdem_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Ordem);
            Range ColumnNome_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Nome);

            gWs.Sort.SortFields.Clear();
            gWs.Sort.SortFields.Add(Key: ColumnOrg1_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            gWs.Sort.SortFields.Add(Key: ColumnOrdem_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            gWs.Sort.SortFields.Add(Key: ColumnNome_Range, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
            gWs.Sort.SetRange(gWs.Cells);
            gWs.Sort.Header = XlYesNoGuess.xlYes;
            gWs.Sort.MatchCase = false;
            gWs.Sort.Orientation = (XlSortOrientation)Constants.xlTopToBottom;
            gWs.Sort.SortMethod = XlSortMethod.xlPinYin;
            gWs.Sort.Apply();
        }

        void RemoveColumns()
        {
            string[] nameColumns = { ColumnsName.Org, ColumnsName.Depto, ColumnsName.VvtNovo, ColumnsName.TvtNovo, ColumnsName.RecPendSet, 
                                     ColumnsName.SaldoSet, ColumnsName.Saldo, ColumnsName.ValorDias, ColumnsName.CnpjCpfOperadora, ColumnsName.Buscador, 
                                     ColumnsName.Ordem, ColumnsName.Cf10, ColumnsName.NrDoCartao };

            foreach (string nameColumn in nameColumns)
            {
                while (true)
                {
                    Range rng = ExcelFunctions.GetRangeColumnByName(gWs, nameColumn);

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

        void RemoveFillColumns()
        {
            Range ColumnTotal_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Total);
            Range ColumnDesconto_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Desconto);
            Range ColumnCompraFinal_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.CompraFinal);
            Range Column1Compra_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Compra1);
            Range Column2Compra_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Compra2);

            RemoveFill(ColumnTotal_Range);
            RemoveFill(ColumnDesconto_Range);
            RemoveFill(ColumnCompraFinal_Range);
            if (Column1Compra_Range != null) { RemoveFill(Column1Compra_Range); }
            if (Column2Compra_Range != null) { RemoveFill(Column2Compra_Range); }

            void RemoveFill(Range rng)
            {
                Range range = gWs.Columns[rng.Column];
                range.Interior.Pattern = Constants.xlNone;
                range.Interior.TintAndShade = 0;
                range.Interior.PatternTintAndShade = 0;
            }
        }

        bool GenerateSubtotal()
        {
            Range ColumnCompraFinal_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.CompraFinal);
            Range Column1Compra_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Compra1);
            Range Column2Compra_Range = ExcelFunctions.GetRangeColumnByName(gWs, ColumnsName.Compra2);

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

            Range rangeUF = gWs.Range[gWs.Cells[1, 1], gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0]];
            rangeUF.Subtotal(GroupBy: 1, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeOperadora = gWs.Range[gWs.Cells[1, 1], gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-1, 0]];
            rangeOperadora.Subtotal(GroupBy: 2, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeEmpresa = gWs.Range[gWs.Cells[1, 1], gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-3, 0]];
            rangeEmpresa.Subtotal(GroupBy: 3, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);

            Range rangeCUNID = gWs.Range[gWs.Cells[1, 1], gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[-5, 0]];
            rangeCUNID.Subtotal(GroupBy: 4, Function: XlConsolidationFunction.xlSum, TotalList: array_ColumnsSubtotal.ToArray(), Replace: false, PageBreaks: false, XlSummaryRow.xlSummaryBelow);


            { // COPIAR E COLAR COMO VALOR | REMOVER SUBTOTAL
                Range selecao = gWs.Cells;
                Tools.CopiarColarValor(selecao);
                gApp.Application.CutCopyMode = 0;
                selecao.RemoveSubtotal();
            }

            { // DEFINIR ÁREA DE IMPRESSÃO | BORDAS | ZOOM
                Range printArea = gWs.Range[gWs.Cells[1, 1], gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0]];
                ExcelFunctions.SetBZPA(gWs, printArea);
            }

            return true;
        }

        void OrganizeSubtotal()
        {
            int ColumnUF_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Uf);
            int ColumnOperadora_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Operadora);
            int ColumnEmpresa_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Empresa);
            int ColumnCUnid_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.CUnid);
            int ColumnCompraFinal_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.CompraFinal);

            ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(gWs, ColumnsName.Operadora, "Total Geral", "<>");
            ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(gWs, ColumnsName.Empresa, "Total Geral", "<>");
            ExcelFunctions.DeleteRowsThatContainSpecificTextInColumn(gWs, ColumnsName.CUnid, "Total Geral", "<>");

            Range rngInicial = gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Offset[0, 0];

            int offSetRow = 0;
            int linha = rngInicial.Row;

            while (true)
            {
                string valorColunaUF = gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaOperadora = gWs.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaEmpresa = gWs.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Text.Trim().ToLower();
                string valorColunaCUNID = gWs.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Text.Trim().ToLower();

                if (valorColunaUF == "total geral")
                {
                    Range rng_linha = gWs.Range[gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    ExcelFunctions.Styles_Emphasis(rng_linha, 5);
                }
                else if (valorColunaUF.Contains(" total"))
                {
                    Range rng_linha = gWs.Range[gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    ExcelFunctions.Styles_Emphasis(rng_linha, 4);

                }
                //else if (valorColunaOperadora == "total geral")
                //{
                //    gWs.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (gWs.Cells[linha, ColumnOperadora_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaOperadora.Contains(" total"))
                {
                    Range rng_linha = gWs.Range[gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    ExcelFunctions.Styles_Emphasis(rng_linha, 3);
                }
                //else if (valorColunaEmpresa == "total geral")
                //{
                //    gWs.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (gWs.Cells[linha, ColumnEmpresa_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaEmpresa.Contains(" total"))
                {
                    Range rng_linha = gWs.Range[gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    ExcelFunctions.Styles_Emphasis(rng_linha, 2);
                }
                //else if (valorColunaCUNID == "total geral")
                //{
                //    gWs.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].EntireRow.Delete();
                //    linha = (gWs.Cells[linha, ColumnCUnid_Number].Offset[offSetRow, 0].Row) - 1;
                //    offSetRow = 0;
                //    continue;
                //}
                else if (valorColunaCUNID.Contains(" total"))
                {
                    Range rng_linha = gWs.Range[gWs.Cells[linha, ColumnUF_Number].Offset[offSetRow, 0], gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0]];
                    ExcelFunctions.Styles_Emphasis(rng_linha, 1);
                }

                if (gWs.Cells[linha, ColumnCompraFinal_Number].Offset[offSetRow, 0].Row < 2)
                {
                    break;
                }

                offSetRow--;
            }
        }

        bool SeparatePurchases()
        {
            bool warning = false;
            int ColumnNome_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Nome);
            int ColumnCompraFinal_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.CompraFinal);
            int ColumnOBS_Number = ExcelFunctions.GetNumberColumnByName(gWs, ColumnsName.Obs);

            int lastUsedRow = gWs.Cells[1048576, ColumnCompraFinal_Number].End(XlDirection.xlUp).Row;

            int offSetRow = 0;
            int offSetColumn = 0;

            while (true)
            {
                Range actvCell = gWs.Cells[lastUsedRow, ColumnCompraFinal_Number].Offset[offSetRow, offSetColumn];

                if (actvCell.Row < 2)
                {
                    break;
                }

                //-------------------[COLUMN OBS]-------------------
                string ColumnOBS_CellValueTopRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueCurrentRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnOBS_Number, 0, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                //------------------[COLUMN NOME]-------------------
                string ColumnNome_CellValueTopRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnNome_Number, -1, 0).ToLower();
                string ColumnNome_CellValueCurrentRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnNome_Number, 0, 0).ToLower();
                string ColumnNome_CellValueBottomRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnNome_Number, 1, 0).ToLower();

                if (ColumnOBS_CellValueCurrentRow == "inativo" || ColumnOBS_CellValueCurrentRow == "sem cadastro" || ColumnOBS_CellValueCurrentRow == "cpf ativo em outro comprador")
                {
                    if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        ExcelFunctions.Styles_Colors(ExcelFunctions.GetRangeCell(gWs, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { ExcelFunctions.TabColor(gWs, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow == "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ VAZIA
                        CheckInRows(actvCell, false, true);
                        ExcelFunctions.Styles_Colors(ExcelFunctions.GetRangeCell(gWs, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { ExcelFunctions.TabColor(gWs, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow == "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ VAZIA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        ExcelFunctions.Styles_Colors(ExcelFunctions.GetRangeCell(gWs, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { ExcelFunctions.TabColor(gWs, 5); warning = true; }
                    }
                    else if (ColumnNome_CellValueBottomRow != "" && ColumnNome_CellValueTopRow != "")
                    {
                        // COLUNA NOME NA LINHA [CIMA] DA ATUAL ESTÁ OCUPADA
                        // COLUNA NOME NA LINHA [BAIXO] DA ATUAL ESTÁ OCUPADA
                        CheckInRows(actvCell, true, true);
                        CheckInRows(actvCell, false, true);
                        ExcelFunctions.Styles_Colors(ExcelFunctions.GetRangeCell(gWs, actvCell.Row, ColumnCompraFinal_Number), 5);
                        if (warning == false) { ExcelFunctions.TabColor(gWs, 5); warning = true; }
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
                string ColumnOBS_CellValueTopRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnOBS_Number, -1, 0).ToLower();
                string ColumnOBS_CellValueBottomRow = ExcelFunctions.GetCellText(gWs, actvCell.Row, ColumnOBS_Number, 1, 0).ToLower();

                if (abaixo == true && problema == true)
                {
                    if (ColumnOBS_CellValueBottomRow != "inativo" && ColumnOBS_CellValueBottomRow != "sem cadastro" && ColumnOBS_CellValueBottomRow != "cpf ativo em outro comprador")
                    {
                        gWs.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == true)
                {
                    if (ColumnOBS_CellValueTopRow != "inativo" && ColumnOBS_CellValueTopRow != "sem cadastro" && ColumnOBS_CellValueTopRow != "cpf ativo em outro comprador")
                    {
                        gWs.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
                else if (abaixo == true && problema == false)
                {
                    if (ColumnOBS_CellValueBottomRow != "novo/sem cartao")
                    {
                        gWs.Rows[actvCell.Offset[1, 0].Row].Insert();
                    }
                }
                else if (abaixo == false && problema == false)
                {
                    if (ColumnOBS_CellValueTopRow != "novo/sem cartao")
                    {
                        gWs.Rows[actvCell.Row].Insert();
                        offSetRow++;
                    }
                }
            }

            if (warning == false) { ExcelFunctions.TabColor(gWs, 3); }

            return true;
        }

        void AdjustHideColumns()
        {
            string[] nameAdjustColumns = { ColumnsName.Uf, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid };
            string[] nameHideColumns = { ColumnsName.CDepto, ColumnsName.Cnpj, ColumnsName.Escala, ColumnsName.Rg, ColumnsName.DataNasc, ColumnsName.Desc, ColumnsName.Qvt, ColumnsName.Vvt, ColumnsName.Tvt, ColumnsName.Total, ColumnsName.Desconto };

            foreach (string nameAdjustColumn in nameAdjustColumns)
            {
                Range rng = ExcelFunctions.GetRangeColumnByName(gWs, nameAdjustColumn);

                if (rng != null)
                {
                    rng.ColumnWidth = 0.08;
                    continue;
                }
            }

            foreach (string nameHideColumn in nameHideColumns)
            {
                Range rng = ExcelFunctions.GetRangeColumnByName(gWs, nameHideColumn);

                if (rng != null)
                {
                    rng.EntireColumn.Hidden = true;
                    continue;
                }
            }
        }
    }
}
