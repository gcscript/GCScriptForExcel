using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System;
using System.Data.Common;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public class GetMatbyCpf
    {
        private readonly string _saldoTabName;

        public GetMatbyCpf(string saldoTabName)
        {
            _saldoTabName = saldoTabName;
        }

        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;
        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;
                gcsApp.Calculation = XlCalculation.xlCalculationManual;

                Worksheet ws = gcsApp.ActiveSheet;

                var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
                if (cnpjColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cnpj} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var matSiteColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.MatSite);
                if (matSiteColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.MatSite} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                
                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cpf);
                if (cpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var nrDoCartaoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.NrDoCartao);
                if (nrDoCartaoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.NrDoCartao} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var contSeCpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ContSeCpf);
                if (contSeCpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ContSeCpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                Stopwatch stopwatch = Stopwatch.StartNew();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;
                if (lastUsedRowByNome < 2) lastUsedRowByNome = 2;

                int count = 0;

                for (int row = 1; row < lastUsedRowByNome; row++)
                {
                    Range rngCnpj = ws.Cells[row+1, cnpjColumnNumber];
                    Range rngNome = ws.Cells[row+1, nomeColumnNumber];
                    Range rngNrDoCartao = ws.Cells[row+1, nrDoCartaoColumnNumber];
                    Range rngContSeCpf = ws.Cells[row+1, contSeCpfColumnNumber];

                    if (rngNome.Text.Length > 2)
                    {
                        if (rngCnpj.Text.Length == 18)
                        {
                            if (rngNrDoCartao.Text == "_CARTAO NAO ENCONTRADO")
                            {
                                if (rngContSeCpf.Value2 == 1)
                                {
                                    string columnLetterCpf = Regex.Replace(ws.Cells[1, cpfColumnNumber].Address, @"[^a-zA-Z]", "");
                                    Range rngMatSite = ws.Cells[row+1, matSiteColumnNumber];
                                    rngMatSite.NumberFormat = "General";
                                    rngMatSite.FormulaLocal = $"=PROCX({columnLetterCpf}{row+1};{_saldoTabName}!G:G;{_saldoTabName}!E:E)";
                                    count++;
                                }
                            }
                        }
                    }
                    continue;
                }

                stopwatch.Stop();
                MessageBox.Show($"Matrículas Corrigidas pelo CPF: {count}\nTempo: {stopwatch.Elapsed:hh\\:mm\\:ss\\.ff}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }

            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "ERROR: 443905", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gcsApp.Calculation = XlCalculation.xlCalculationAutomatic;
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }
        }
    }
}
