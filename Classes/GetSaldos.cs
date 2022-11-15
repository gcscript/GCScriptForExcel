using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
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

                var cpfDoCartaoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CpfDoCartao);
                if (cpfDoCartaoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CpfDoCartao} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var difCpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.DifCpf);
                if (difCpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.DifCpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var contSeCpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ContSeCpf);
                if (contSeCpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ContSeCpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var contSeNomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ContSeNome);
                if (contSeNomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ContSeNome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                string initialDirectory = Path.GetDirectoryName(gcsApp.ActiveWorkbook.FullName);
                if (Directory.Exists(Path.Combine(initialDirectory, "SALDOS")))
                    initialDirectory = Path.Combine(initialDirectory, "SALDOS");
                else if (Directory.Exists(Path.Combine(initialDirectory, "SALDO")))
                    initialDirectory = Path.Combine(initialDirectory, "SALDO");

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = initialDirectory;
                ofd.Title = "Selecionar Arquivo de Saldo";
                ofd.Filter = "Text File (*.txt)|*.txt";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.Multiselect = false;
                ofd.ShowDialog();

                if (ofd.FileName == "") { return; }

                Stopwatch stopwatch = Stopwatch.StartNew();

                var fileName = Path.GetFileName(ofd.FileName);

                string activeTabName = ws.Name;
                string saldoTabName = $"Saldos_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}";
                gcsApp.Worksheets.Add(After: gcsApp.Worksheets[gcsApp.Worksheets.Count]);
                gcsApp.ActiveSheet.Name = saldoTabName;

                Worksheet wsSaldo = gcsApp.Worksheets[saldoTabName];
                wsSaldo.QueryTables.Add(Connection: $"TEXT;{ofd.FileName}", Destination: wsSaldo.Range["A1"]);

                wsSaldo.QueryTables[1].Name = saldoTabName; // Nome da tabela de consulta.
                wsSaldo.QueryTables[1].TextFileCommaDelimiter = false; // true se a vírgula for o delimitador ao importar um arquivo de texto para uma tabela de consulta. false se você quiser usar algum outro caractere como delimitador. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].FieldNames = true; // true se a primeira linha do arquivo de texto contiver nomes de campo. false se a primeira linha do arquivo de texto não contiver nomes de campo. O valor padrão é true. Ler/gravar booleano.
                wsSaldo.QueryTables[1].RowNumbers = false; // true se a primeira coluna do arquivo de texto contiver números de linha. false se a primeira coluna do arquivo de texto não contiver números de linha. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].FillAdjacentFormulas = false; // true se você quiser preencher as fórmulas adjacentes. false se você não quiser preencher as fórmulas adjacentes. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].PreserveFormatting = true; // true se você quiser preservar o formato do arquivo de texto. false se você não quiser preservar o formato do arquivo de texto. O valor padrão é true. Ler/gravar booleano.
                wsSaldo.QueryTables[1].RefreshOnFileOpen = false; // true se você quiser atualizar a tabela de consulta quando o arquivo de texto for aberto. false se você não quiser atualizar a tabela de consulta quando o arquivo de texto for aberto. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].RefreshStyle = XlCellInsertionMode.xlInsertDeleteCells; // Especifica como atualizar os dados na tabela de consulta. O valor padrão é xlInsertDeleteCells. Ler/gravar XlCellInsertionMode.
                wsSaldo.QueryTables[1].SavePassword = false; // true se você quiser salvar a senha do arquivo de texto. false se você não quiser salvar a senha do arquivo de texto. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].SaveData = true; // true se você quiser salvar os dados da tabela de consulta. false se você não quiser salvar os dados da tabela de consulta. O valor padrão é true. Ler/gravar booleano.
                wsSaldo.QueryTables[1].AdjustColumnWidth = true; // true se você quiser ajustar a largura da coluna para caber o conteúdo. false se você não quiser ajustar a largura da coluna para caber o conteúdo. O valor padrão é true. Ler/gravar booleano.
                wsSaldo.QueryTables[1].RefreshPeriod = 0; // Especifica o intervalo de tempo em minutos entre atualizações da tabela de consulta. O valor padrão é 0. Ler/gravar inteiro.
                wsSaldo.QueryTables[1].TextFilePromptOnRefresh = false; // true se você quiser exibir uma caixa de diálogo para atualizar a tabela de consulta. false se você não quiser exibir uma caixa de diálogo para atualizar a tabela de consulta. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].TextFilePlatform = 437; // Especifica o código de página do sistema operacional para o arquivo de texto. O valor padrão é 437. Ler/gravar inteiro.
                wsSaldo.QueryTables[1].TextFileStartRow = 1; // Especifica a linha inicial do arquivo de texto. O valor padrão é 1. Ler/gravar inteiro.
                wsSaldo.QueryTables[1].TextFileParseType = XlTextParsingType.xlDelimited; // Especifica o tipo de arquivo de texto. O valor padrão é xlDelimited. Ler/gravar XlTextParsingType.
                wsSaldo.QueryTables[1].TextFileTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote; // Especifica o qualificador de texto do arquivo de texto. O valor padrão é xlTextQualifierDoubleQuote. Ler/gravar XlTextQualifier.
                wsSaldo.QueryTables[1].TextFileConsecutiveDelimiter = false; // true se você quiser considerar delimitadores consecutivos como um delimitador único. false se você não quiser considerar delimitadores consecutivos como um delimitador único. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].TextFileTabDelimiter = false; // true se você quiser usar o caractere de tabulação como delimitador. false se você não quiser usar o caractere de tabulação como delimitador. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].TextFileSemicolonDelimiter = false; // true se você quiser usar o ponto e vírgula como delimitador. false se você não quiser usar o ponto e vírgula como delimitador. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].TextFileSpaceDelimiter = false; // true se você quiser usar o espaço como delimitador. false se você não quiser usar o espaço como delimitador. O valor padrão é false. Ler/gravar booleano.
                wsSaldo.QueryTables[1].TextFileOtherDelimiter = "\t"; // Especifica o delimitador de texto do arquivo de texto. O valor padrão é "". Ler/gravar string.
                wsSaldo.QueryTables[1].TextFileColumnDataTypes = new XlColumnDataType[]
                {
                    XlColumnDataType.xlTextFormat, // CNPJ
                    XlColumnDataType.xlTextFormat, // EMPRESA
                    XlColumnDataType.xlTextFormat, // BUSCADOR
                    XlColumnDataType.xlTextFormat, // NR. DO CARTAO
                    XlColumnDataType.xlTextFormat, // MAT
                    XlColumnDataType.xlTextFormat, // NOME
                    XlColumnDataType.xlTextFormat, // CPF
                    XlColumnDataType.xlTextFormat, // TP CARTAO
                    XlColumnDataType.xlGeneralFormat, // SALDO
                    XlColumnDataType.xlDMYFormat, // ATT SALDO
                    XlColumnDataType.xlGeneralFormat, // REC PEND
                    XlColumnDataType.xlDMYFormat, // DATA PGMT REC PEND
                }; // Especifica o tipo de dados de cada coluna. O valor padrão é xlGeneralFormat. Ler/gravar XlColumnDataType.
                wsSaldo.QueryTables[1].Refresh(false); // true se você quiser atualizar a tabela de consulta. false se você não quiser atualizar a tabela de consulta. O valor padrão é false. Ler/gravar booleano.
                
                gcsApp.Worksheets[activeTabName].Activate();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;
                if (lastUsedRowByNome < 2) lastUsedRowByNome = 2;

                var fileNameToPROCX = Tools.ReplaceInvalidCharactersExcel(fileName);

                string columnLetterCpf = Regex.Replace(ws.Cells[2, cpfColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterNome = Regex.Replace(ws.Cells[2, nomeColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterBuscaValorDias = Regex.Replace(ws.Cells[2, buscaValorDiasColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterBuscaCartao = Regex.Replace(ws.Cells[2, buscaCartaoColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterNrDoCartao = Regex.Replace(ws.Cells[2, nrDoCartaoColumnNumber].Address, @"[^a-zA-Z]", "");
                string columnLetterCpfDoCartao = Regex.Replace(ws.Cells[2, obsColumnNumber + 1].Address, @"[^a-zA-Z]", "");

                // NR DO CARTAO
                Range startNrDoCartao = ws.Cells[2, nrDoCartaoColumnNumber];
                startNrDoCartao.FormulaLocal = $"=PROCX({columnLetterBuscaCartao}{startNrDoCartao.Row};{saldoTabName}!$C:$C;{saldoTabName}!$D:$D;\"_CARTAO NAO ENCONTRADO\";0;1)";
                startNrDoCartao.AutoFill(Destination: ws.Range[startNrDoCartao, ws.Cells[lastUsedRowByNome, startNrDoCartao.Column]]);

                // CPF DO CARTAO
                Range startCpfDoCartao = ws.Cells[2, cpfDoCartaoColumnNumber];
                startCpfDoCartao.FormulaLocal = $"=SE(PROCX({columnLetterNrDoCartao}{startCpfDoCartao.Row};{saldoTabName}!$D:$D;{saldoTabName}!$G:$G;\"_ARRUME O CARTAO\";0;1)=0;\"_ARRUME O CARTAO\";PROCX({columnLetterNrDoCartao}{startCpfDoCartao.Row};{saldoTabName}!$D:$D;{saldoTabName}!$G:$G;\"_ARRUME O CARTAO\";0;1))";
                startCpfDoCartao.AutoFill(Destination: ws.Range[startCpfDoCartao, ws.Cells[lastUsedRowByNome, startCpfDoCartao.Column]]);

                // DIF CPF
                Range startDifCpf = ws.Cells[2, difCpfColumnNumber];
                startDifCpf.FormulaLocal = $"=SE({columnLetterCpfDoCartao}{startDifCpf.Row}={columnLetterCpf}{startDifCpf.Row};\"IGUAL\";\"DIFERENTE\")";
                startDifCpf.AutoFill(Destination: ws.Range[startDifCpf, ws.Cells[lastUsedRowByNome, startDifCpf.Column]]);

                // REC PEND SET
                Range startRecPendSet = ws.Cells[2, recPendSetColumnNumber];
                startRecPendSet.FormulaLocal = $"=SE({columnLetterBuscaValorDias}{startCpfDoCartao.Row}={columnLetterBuscaValorDias}{startCpfDoCartao.Row - 1};0;PROCX({columnLetterNrDoCartao}{startCpfDoCartao.Row};{saldoTabName}!$D:$D;{saldoTabName}!$I:$I;-999999;0;1))";
                startRecPendSet.AutoFill(Destination: ws.Range[startRecPendSet, ws.Cells[lastUsedRowByNome, startRecPendSet.Column]]);

                // SALDO SET
                Range startSaldoSet = ws.Cells[2, saldoSetColumnNumber];
                startSaldoSet.FormulaLocal = $"=SE({columnLetterBuscaValorDias}{startCpfDoCartao.Row}={columnLetterBuscaValorDias}{startCpfDoCartao.Row - 1};0;PROCX({columnLetterNrDoCartao}{startCpfDoCartao.Row};{saldoTabName}!$D:$D;{saldoTabName}!$K:$K;-999999;0;1))";
                startSaldoSet.AutoFill(Destination: ws.Range[startSaldoSet, ws.Cells[lastUsedRowByNome, startSaldoSet.Column]]);

                // CONT.SE CPF
                Range startContSeCpf = ws.Cells[2, contSeCpfColumnNumber];
                startContSeCpf.FormulaLocal = $"=CONT.SE({saldoTabName}!G:G;{columnLetterCpf}{startContSeCpf.Row})";
                startContSeCpf.AutoFill(Destination: ws.Range[startContSeCpf, ws.Cells[lastUsedRowByNome, startContSeCpf.Column]]);

                // CONT.SE NOME
                Range startContSeNome = ws.Cells[2, contSeNomeColumnNumber];
                startContSeNome.FormulaLocal = $"=CONT.SE({saldoTabName}!F:F;{columnLetterNome}{startContSeNome.Row})";
                startContSeNome.AutoFill(Destination: ws.Range[startContSeNome, ws.Cells[lastUsedRowByNome, startContSeNome.Column]]);

                Range rngAllCells = ws.Cells;
                gcsApp.Calculation = XlCalculation.xlCalculationAutomatic;
                ExcelFunctions.RowHeight(rngAllCells);
                ExcelFunctions.ColumnWidth(rngAllCells);

                stopwatch.Stop();

                if (MessageBox.Show($"Saldos Carregados com Sucesso!\nTempo: {stopwatch.Elapsed:hh\\:mm\\:ss\\.ff}\nDeseja Corrigir Matrículas?", "ATENCAO!", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                {
                    var beta = new GetMat(saldoTabName);
                    beta.Start();
                }
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
    }
}
