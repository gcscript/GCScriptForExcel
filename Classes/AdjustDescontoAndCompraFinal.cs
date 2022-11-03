using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public class AdjustDescontoAndCompraFinal
    {
        gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                var ws = ExcelFunctions.SearchWorksheet(gcsApp, "Dados");

                if (ws == null)
                {
                    MessageBox.Show($"A aba Dados não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                ws.Select();

                if (!ExcelFunctions.CheckIfColumnsExist(ws, new List<string> { ColumnsName.ValorDias, ColumnsName.Saldo }))
                {
                    ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                var totalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Total);
                if (totalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Total} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var saldoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Saldo);
                if (saldoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Saldo} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var valorDiasColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ValorDias);
                if (valorDiasColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ValorDias} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var descontoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desconto);
                if (descontoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desconto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var compraFinalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CompraFinal);
                if (compraFinalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CompraFinal} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }


                int lastUsedRowBySaldo = ws.Cells[1048576, saldoColumnNumber].End(XlDirection.xlUp).Row;

                var offSetRow = 0;
                var countSaldo = 0;
                var countValorDias = 0;
                var countDiscountGreaterThan0AndLessThan10 = 0;
                var countPurchaseGreaterThan0AndLessThan10 = 0;

                while (true)
                {
                    Range activeCellByTotal = ws.Cells[lastUsedRowBySaldo, totalColumnNumber].Offset[offSetRow, 0];

                    Range activeCellBySaldo = ws.Cells[lastUsedRowBySaldo, saldoColumnNumber].Offset[offSetRow, 0];
                    Range activeCellByValorDias = ws.Cells[lastUsedRowBySaldo, valorDiasColumnNumber].Offset[offSetRow, 0];

                    Range activeCellByDesconto = ws.Cells[lastUsedRowBySaldo, descontoColumnNumber].Offset[offSetRow, 0];
                    Range activeCellByCompraFinal = ws.Cells[lastUsedRowBySaldo, compraFinalColumnNumber].Offset[offSetRow, 0];

                    if (activeCellBySaldo.Row < 2) { break; }

                    string activeCellBySaldoText = Regex.Replace(activeCellBySaldo.Text.ToString(), @"\s", "");
                    string activeCellByValorDiasText = Regex.Replace(activeCellByValorDias.Text.ToString(), @"\s", "");

                    if (activeCellBySaldo.Value2 == 0 || activeCellBySaldo.Value2 == -2146826246
                        || activeCellBySaldoText == "#N/D" || activeCellBySaldoText == "0,00" || activeCellBySaldoText == "-0,00"
                        || activeCellByValorDias.Value2 == 0 || activeCellByValorDias.Value2 == -2146826246
                        || activeCellByValorDiasText == "#N/D" || activeCellByValorDiasText == "0,00" || activeCellByValorDiasText == "-0,00")
                    {
                        activeCellBySaldo.Value2 = 0;
                        activeCellByValorDias.Value2 = 0;
                        offSetRow--;
                        countSaldo++;
                        countValorDias++;
                        continue;
                    }

                    var saldoIsNumeric = ExcelFunctions.IsNumeric(activeCellBySaldo);

                    if (saldoIsNumeric.isNull)
                    {
                        activeCellBySaldo.Value2 = 0;
                        activeCellByValorDias.Value2 = 0;
                        offSetRow--;
                        countSaldo++;
                        countValorDias++;
                        continue;
                    }
                    else
                    {
                        if (!saldoIsNumeric.isNumeric)
                        {
                            MessageBox.Show($"Saldo com erro na linha {activeCellBySaldo.Row}",
                                           "ERROR: 104927",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Error,
                                           MessageBoxDefaultButton.Button1);
                            return;
                        }
                    }

                    var valorDiasIsNumeric = ExcelFunctions.IsNumeric(activeCellByValorDias);

                    if (valorDiasIsNumeric.isNull)
                    {
                        activeCellBySaldo.Value2 = 0;
                        activeCellByValorDias.Value2 = 0;
                        offSetRow--;
                        countSaldo++;
                        countValorDias++;
                        continue;
                    }
                    else
                    {
                        if (!valorDiasIsNumeric.isNumeric)
                        {
                            MessageBox.Show($"ValorDias com erro na linha {activeCellByValorDias.Row}",
                                           "ERROR: 974244",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Error,
                                           MessageBoxDefaultButton.Button1);
                            return;
                        }
                    }

                    string activeCellByTotalText = Regex.Replace(activeCellByTotal.Text.ToString(), @"\s", "");
                    string activeCellByDescontoText = Regex.Replace(activeCellByDesconto.Text.ToString(), @"\s", "");
                    string activeCellByCompraFinalText = Regex.Replace(activeCellByCompraFinal.Text.ToString(), @"\s", "");


                    if (activeCellByDescontoText == "0,00" || activeCellByDescontoText == "-0,00")
                    {
                        MessageBox.Show($"Desconto: {activeCellByDescontoText}", $"ROW: {activeCellByDesconto.Row}", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                        activeCellByDesconto.Value2 = 0;
                    }

                    if (activeCellByDesconto.Value2 > 0 && activeCellByDesconto.Value2 < 10) // SE [DESCONTO] FOR MAIOR QUE [0] E MENOR QUE [10]
                    {
                        if (activeCellByDesconto.Value2 != activeCellByTotal.Value2) // SE [DESCONTO] FOR DIFERENTE DE [TOTAL]
                        {
                            activeCellBySaldo.Value2 = activeCellByValorDias.Value2; // [SALDO] VAI SER IGUAL A [VALOR DIAS]
                            countSaldo++;
                            countDiscountGreaterThan0AndLessThan10++;
                        }
                    }

                    if (activeCellByCompraFinalText == "0,00" || activeCellByCompraFinalText == "-0,00") { activeCellByCompraFinal.Value2 = 0; }
                    if (activeCellByCompraFinalText == "0,00" || activeCellByCompraFinalText == "-0,00")
                    {
                        MessageBox.Show($"CompraFinal: {activeCellByCompraFinalText}", $"ROW: {activeCellByCompraFinal.Row}", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                        activeCellByCompraFinal.Value2 = 0;
                    }

                    if (activeCellByCompraFinal.Value2 > 0 && activeCellByCompraFinal.Value2 < 10) // SE [COMPRA FINAL] FOR MAIOR QUE [0] E MENOR QUE [10]
                    {
                        if (activeCellByTotal.Value2 > 10) // SE [TOTAL] FOR MAIOR QUE [10]
                        {
                            activeCellBySaldo.Value2 -= (10 - activeCellByCompraFinal.Value2); // [SALDO] VAI SER IGUAL A [SALDO] MENOS O RESULTADO DE [10] MENOS [COMPRA FINAL]
                            countSaldo++;
                            countPurchaseGreaterThan0AndLessThan10++;
                        }

                    }

                    offSetRow--;
                }

                MessageBox.Show($"Saldo(s) ajustado(s): {countSaldo}\n" +
                                     $"ValorDias ajustado(s): {countValorDias}\n" +
                                     $"Desconto(s) [>0] & [<10] ajustado(s): {countDiscountGreaterThan0AndLessThan10}\n" +
                                     $"Compra Final [>0] & [<10] ajustada(s): {countPurchaseGreaterThan0AndLessThan10}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "ERROR: 967040", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }
        }
    }
}
