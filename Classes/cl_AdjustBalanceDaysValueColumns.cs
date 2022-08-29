using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public class cl_AdjustBalanceDaysValueColumns
    {
        gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;

                var ws = cl_ExcelFunctions.SearchWorksheet(gcsApp, "Dados");

                if (ws == null)
                {
                    MessageBox.Show($"A aba Dados não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cl_ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                ws.Select();

                if (!cl_ExcelFunctions.CheckIfColumnsExist(ws, new List<string> { ColumnsName.ValorDias, ColumnsName.Saldo }))
                {
                    cl_ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                var totalColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Total);
                if (totalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Total} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var saldoColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Saldo);
                if (saldoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Saldo} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var valorDiasColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ValorDias);
                if (valorDiasColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ValorDias} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var descontoColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desconto);
                if (descontoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desconto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var compraFinalColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CompraFinal);
                if (compraFinalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CompraFinal} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }


                int lastUsedRowBySaldo = ws.Cells[1048576, saldoColumnNumber].End(XlDirection.xlUp).Row;

                var offSetRow = 0;
                //lastUsedRowBySaldo.Select();

                while (true)
                {
                    Range activeCellByTotal = ws.Cells[lastUsedRowBySaldo, totalColumnNumber].Offset[offSetRow, 0];

                    Range activeCellBySaldo = ws.Cells[lastUsedRowBySaldo, saldoColumnNumber].Offset[offSetRow, 0];
                    Range activeCellByValorDias = ws.Cells[lastUsedRowBySaldo, valorDiasColumnNumber].Offset[offSetRow, 0];

                    Range activeCellByDesconto = ws.Cells[lastUsedRowBySaldo, descontoColumnNumber].Offset[offSetRow, 0];
                    Range activeCellByCompraFinal = ws.Cells[lastUsedRowBySaldo, compraFinalColumnNumber].Offset[offSetRow, 0];

                    //MessageBox.Show($"lastUsedRowBySaldo: {lastUsedRowBySaldo.ToString()}\n" +
                    //                     $"activeCellBySaldo: {activeCellBySaldo.Row.ToString()}\n" +
                    //                     $"offSetRow: {offSetRow.ToString()}");

                    //return;

                    if (activeCellBySaldo.Row < 2) { break; }


                    string activeCellByTotalText = Regex.Replace(activeCellByTotal.Text.ToString(), @"\s", "");
                    string activeCellBySaldoText = Regex.Replace(activeCellBySaldo.Text.ToString(), @"\s", "");
                    string activeCellByValorDiasText = Regex.Replace(activeCellByValorDias.Text.ToString(), @"\s", "");
                    string activeCellByDescontoText = Regex.Replace(activeCellByDesconto.Text.ToString(), @"\s", "");
                    string activeCellByCompraFinalText = Regex.Replace(activeCellByCompraFinal.Text.ToString(), @"\s", "");



                    if (activeCellBySaldo.Value2 == 0 || activeCellBySaldo.Value2 == -2146826246
                        || activeCellBySaldoText == "#N/D" || activeCellBySaldoText == "0,00" || activeCellBySaldoText == "-0,00"
                        || activeCellByValorDias.Value2 == 0 || activeCellByValorDias.Value2 == -2146826246
                        || activeCellByValorDiasText == "#N/D" || activeCellByValorDiasText == "0,00" || activeCellByValorDiasText == "-0,00")
                    {
                        activeCellBySaldo.Value2 = 0;
                        activeCellByValorDias.Value2 = 0;
                        offSetRow--;
                        continue;
                    }
                    //string valoranterior = activeCellByCompraFinal.Value2.ToString();



                    if (activeCellByDescontoText == "0,00" || activeCellByDescontoText == "-0,00") { activeCellByDesconto.Value2 = 0; }

                    if (activeCellByDesconto.Value2 > 0 && activeCellByDesconto.Value2 < 10) // SE [DESCONTO] FOR MAIOR QUE [0] E MENOR QUE [10]
                    {
                        if (activeCellByDesconto.Value2 != activeCellByTotal.Value2) // SE [DESCONTO] FOR DIFERENTE DE [TOTAL]
                        {
                            activeCellBySaldo.Value2 = activeCellByValorDias.Value2; // [SALDO] VAI SER IGUAL A [VALOR DIAS]
                            //if (activeCellByCompraFinal.Value2 < 10)
                            //{
                            //    MessageBox.Show($"Linha: {activeCellByCompraFinal.Row.ToString()} Valor Anterior: {valoranterior} Valor: {activeCellByCompraFinal.Value2.ToString()}");

                            //}
                        }
                    }

                    if (activeCellByCompraFinalText == "0,00" || activeCellByCompraFinalText == "-0,00") { activeCellByCompraFinal.Value2 = 0; }

                    if (activeCellByCompraFinal.Value2 > 0 && activeCellByCompraFinal.Value2 < 10) // SE [COMPRA FINAL] FOR MAIOR QUE [0] E MENOR QUE [10]
                    {
                        if (activeCellByTotal.Value2 > 10) // SE [TOTAL] FOR MAIOR QUE [10]
                        {
                            activeCellBySaldo.Value2 = activeCellBySaldo.Value2 - (10 - activeCellByCompraFinal.Value2); // [SALDO] VAI SER IGUAL A [SALDO] MENOS O RESULTADO DE [10] MENOS [COMPRA FINAL]
                        }

                    }





                    offSetRow--;


                }

                MessageBox.Show("Terminou");

            }
            catch (Exception erro)
            {
                gcsApp.ScreenUpdating = true;
                MessageBox.Show(erro.ToString(), "ERROR: 812058", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }
        }




    }
}
