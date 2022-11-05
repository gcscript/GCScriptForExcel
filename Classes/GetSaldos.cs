using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{

    public class GetSaldos
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;
        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;
                gcsApp.Calculation = XlCalculation.xlCalculationManual;

                Worksheet ws = gcsApp.ActiveSheet;

                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cpf);
                if (cpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var recPendSetColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.RecPendSet);
                if (recPendSetColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.RecPendSet} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var saldoSetColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.SaldoSet);
                if (saldoSetColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.SaldoSet} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var buscaValorDiasColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.BuscaValorDias);
                if (buscaValorDiasColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.BuscaValorDias} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var buscaCartaoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.BuscaCartao);
                if (buscaCartaoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.BuscaCartao} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var nrDoCartaoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.NrDoCartao);
                if (nrDoCartaoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.NrDoCartao} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var obsColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Obs);
                if (obsColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Obs} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                string initialDirectory = Path.GetDirectoryName(gcsApp.ActiveWorkbook.FullName);
                if (Directory.Exists(Path.Combine(initialDirectory, "SALDOS")))
                    initialDirectory = Path.Combine(initialDirectory, "SALDOS");
                else if (Directory.Exists(Path.Combine(initialDirectory, "SALDO")))
                    initialDirectory = Path.Combine(initialDirectory, "SALDO");

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = initialDirectory;
                ofd.Title = "Selecionar Arquivo de Saldo";
                ofd.Filter = "Excel File (*.xlsx)|*.xlsx";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.Multiselect = false;
                ofd.ShowDialog();

                if (ofd.FileName == "") { return; }
                var fileName = Path.GetFileName(ofd.FileName);

                if (!CheckWorkbookIsOpen(fileName))
                {
                    MessageBox.Show($"A Planilha de Saldos precisa estar aberta!", "Result", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                    return;
                }

                Stopwatch stopwatch = Stopwatch.StartNew();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;
                var fileNameToPROCX = Tools.ReplaceInvalidCharactersExcel(fileName);

                string columnLetterCpf = Regex.Replace(ws.Cells[2, cpfColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterBuscaValorDias = Regex.Replace(ws.Cells[2, buscaValorDiasColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterBuscaCartao = Regex.Replace(ws.Cells[2, buscaCartaoColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterNrDoCartao = Regex.Replace(ws.Cells[2, nrDoCartaoColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterCpfDoCartao = Regex.Replace(ws.Cells[2, obsColumnNumber + 1].Address, @"[^a-zA-Z]", "");

                // NR DO CARTAO
                Range startNrDoCartao = ws.Cells[2, nrDoCartaoColumnNumber];
                startNrDoCartao.FormulaLocal = $"=PROCX({columnLetterBuscaCartao}{startNrDoCartao.Row};'[{fileNameToPROCX}]Saldos'!$C:$C;'{fileNameToPROCX}'!$D:$D;\"_CARTAO NAO ENCONTRADO\";0;1)";
                startNrDoCartao.AutoFill(Destination: ws.Range[startNrDoCartao, ws.Cells[lastUsedRowByNome, startNrDoCartao.Column]]);

                // CPF DO CARTAO
                var cpfDoCartaoColumnNumber = obsColumnNumber + 1;
                ws.Cells[1, cpfDoCartaoColumnNumber].Value = $"CPF do Cartao";
                Range startcpfDoCartao = ws.Cells[2, cpfDoCartaoColumnNumber];
                startcpfDoCartao.FormulaLocal = $"=SE(PROCX({columnLetterNrDoCartao}{startcpfDoCartao.Row};'{fileNameToPROCX}'!$D:$D;'{fileNameToPROCX}'!$G:$G;\"_ARRUME O CARTAO\";0;1)=0;\"_ARRUME O CARTAO\";PROCX({columnLetterNrDoCartao}{startcpfDoCartao.Row};'{fileNameToPROCX}'!$D:$D;'{fileNameToPROCX}'!$G:$G;\"_ARRUME O CARTAO\";0;1))";
                startcpfDoCartao.AutoFill(Destination: ws.Range[startcpfDoCartao, ws.Cells[lastUsedRowByNome, startcpfDoCartao.Column]]);

                // DIF CPF
                var difCpfColumnNumber = obsColumnNumber + 2;
                ws.Cells[1, difCpfColumnNumber].Value = $"Dif CPF";
                Range startdifCpf = ws.Cells[2, difCpfColumnNumber];
                startdifCpf.FormulaLocal = $"=SE({columnLetterCpfDoCartao}{startdifCpf.Row}={columnLetterCpf}{startdifCpf.Row};\"IGUAL\";\"DIFERENTE\")";
                startdifCpf.AutoFill(Destination: ws.Range[startdifCpf, ws.Cells[lastUsedRowByNome, startdifCpf.Column]]);

                // REC PEND SET
                Range startRecPendSet = ws.Cells[2, recPendSetColumnNumber];
                startRecPendSet.FormulaLocal = $"=SE({columnLetterBuscaValorDias}{startcpfDoCartao.Row}={columnLetterBuscaValorDias}{startcpfDoCartao.Row - 1};0;PROCX({columnLetterNrDoCartao}{startcpfDoCartao.Row};'{fileNameToPROCX}'!NR_DO_CARTAO;'{fileNameToPROCX}'!SALDO;-999999;0;1))";
                startRecPendSet.AutoFill(Destination: ws.Range[startRecPendSet, ws.Cells[lastUsedRowByNome, startRecPendSet.Column]]);

                // SALDO SET
                Range startSaldoSet = ws.Cells[2, saldoSetColumnNumber]; //=PROCX(AG2;'_RIOCARD (2022-10-27).xlsx'!NR_DO_CARTAO;'_RIOCARD (2022-10-27).xlsx'!REC_PEND;-999999;0;1)
                startSaldoSet.FormulaLocal = $"=SE({columnLetterBuscaValorDias}{startcpfDoCartao.Row}={columnLetterBuscaValorDias}{startcpfDoCartao.Row - 1};0;PROCX({columnLetterNrDoCartao}{startcpfDoCartao.Row};'{fileNameToPROCX}'!NR_DO_CARTAO;'{fileNameToPROCX}'!REC_PEND;-999999;0;1))";
                startSaldoSet.AutoFill(Destination: ws.Range[startSaldoSet, ws.Cells[lastUsedRowByNome, startSaldoSet.Column]]);

                Range rngAllCells = ws.Cells;
                gcsApp.Calculation = XlCalculation.xlCalculationAutomatic;
                ExcelFunctions.RowHeight(rngAllCells);
                ExcelFunctions.ColumnWidth(rngAllCells);

                stopwatch.Stop();
                MessageBox.Show($"Tempo: {stopwatch.Elapsed:hh\\:mm\\:ss\\.ff}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "ERROR: 360425", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gcsApp.Calculation = XlCalculation.xlCalculationAutomatic;
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }
        }

        private static bool CheckWorkbookIsOpen(string fileName)
        {
            foreach (var window in OpenWindowGetter.GetOpenWindows())
            {
                //IntPtr handle = window.Key;
                string title = window.Value;

                if (title.Contains(fileName))
                {
                    return true;
                }
            }
            return false;
        }

    }
}
