using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.XPath;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    internal class ModelTransferData
    {
        public string Cnpj { get; set; }
        public string ArquivoDeCompra { get; set; }
        public string Uf { get; set; }
        public string Empresa { get; set; }
        public string CUnid { get; set; }
        public string CDepto { get; set; }
        public string Depto { get; set; }
        public string Escala { get; set; }
        public string Id { get; set; }
        public string Mat { get; set; }
        public string MatSite { get; set; }
        public string Nome { get; set; }
        public string Cpf { get; set; }
        public string Rg { get; set; }
        public string DataNasc { get; set; }
        public string Operadora { get; set; }
        public decimal Desc { get; set; }
        public int Qvt { get; set; }
        public decimal Vvt { get; set; }
        public decimal Total { get; set; }
        public decimal CompraFinal { get; set; }
        public string Obs { get; set; }
    }

    public class TransferData
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public void Import()
        {
            try
            {
                string initialDirectory = Path.GetDirectoryName(gcsApp.ActiveWorkbook.FullName);
                if (Directory.Exists(Path.Combine(initialDirectory, "ARQUIVOS DE COMPRA")))
                    initialDirectory = Path.Combine(initialDirectory, "ARQUIVOS DE COMPRA");
                else if (Directory.Exists(Path.Combine(initialDirectory, "ARQUIVO DE COMPRA")))
                    initialDirectory = Path.Combine(initialDirectory, "ARQUIVO DE COMPRA");

                OpenFileDialog ofd = new OpenFileDialog();
                ofd.InitialDirectory = initialDirectory;
                ofd.Title = "Selecionar Purchase File";
                ofd.Filter = "Json File (*.json)|*.json";
                ofd.CheckFileExists = true;
                ofd.CheckPathExists = true;
                ofd.Multiselect = false;
                ofd.ShowDialog();

                if (ofd.FileName == "") { return; }

                Stopwatch stopwatch = Stopwatch.StartNew();

                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;
                gcsApp.Calculation = XlCalculation.xlCalculationManual;

                Worksheet ws = gcsApp.ActiveSheet;

                var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
                if (cnpjColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cnpj} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var aCColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ArquivoDeCompra);
                if (aCColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.ArquivoDeCompra} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Uf);
                if (ufColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Uf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Empresa);
                if (empresaColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Empresa} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cUnidColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CUnid);
                if (cUnidColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CUnid} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cDeptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CDepto);
                if (cDeptoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CDepto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var deptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Depto);
                if (deptoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Depto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var escalaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Escala);
                if (escalaColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Escala} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var idColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Id);
                if (idColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Id} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var matColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Mat);
                if (matColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Mat} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var matSiteColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.MatSite);
                if (matSiteColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.MatSite} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CpfDel);
                if (cpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CpfDel} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var rgColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Rg);
                if (rgColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Rg} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var dataNascimentoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.DataNasc);
                if (dataNascimentoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.DataNasc} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var operadoraColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.OperadoraDel);
                if (operadoraColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.OperadoraDel} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var descColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.DescDel);
                if (descColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.DescDel} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var qvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.QvtDel);
                if (qvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.QvtDel} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var vvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.VvtDel);
                if (vvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.VvtDel} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var obsColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Obs);
                if (obsColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Obs} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var row = 2;
                var count = 0;
                Range rngAllCells = ws.Cells;

                using (var sr = new StreamReader(ofd.FileName))
                {
                    string json = sr.ReadToEnd();
                    List<ModelTransferData> jsonItems = JsonConvert.DeserializeObject<List<ModelTransferData>>(json);

                    #region DELETE REMAINING COLUMNS
                    Range entireRows = ws.Range[ws.Cells[jsonItems.Count + 2, 1], ws.Cells[999999, 1]];
                    entireRows.EntireRow.Delete();
                    #endregion

                    foreach (var item in jsonItems)
                    {
                        SetText(ws, row, cnpjColumnNumber, item.Cnpj);
                        SetText(ws, row, aCColumnNumber, item.ArquivoDeCompra);
                        SetText(ws, row, ufColumnNumber, item.Uf);
                        SetText(ws, row, empresaColumnNumber, item.Empresa);
                        SetText(ws, row, cUnidColumnNumber, item.CUnid);
                        SetText(ws, row, cDeptoColumnNumber, item.CDepto);
                        SetText(ws, row, deptoColumnNumber, item.Depto);
                        SetText(ws, row, escalaColumnNumber, item.Escala);
                        SetText(ws, row, idColumnNumber, item.Id);

                        if (item.Mat == null)
                            SetText(ws, row, matColumnNumber, "0");
                        else
                            SetText(ws, row, matColumnNumber, item.Mat);

                        if (item.MatSite == null)
                            SetText(ws, row, matSiteColumnNumber, "0");
                        else
                            SetText(ws, row, matSiteColumnNumber, item.MatSite);

                        SetText(ws, row, nomeColumnNumber, item.Nome);
                        SetText(ws, row, cpfColumnNumber, item.Cpf);
                        SetText(ws, row, rgColumnNumber, item.Rg);
                        SetText(ws, row, dataNascimentoColumnNumber, item.DataNasc);
                        SetText(ws, row, operadoraColumnNumber, item.Operadora);
                        SetDecimal(ws, row, descColumnNumber, item.Desc);
                        SetInt(ws, row, qvtColumnNumber, item.Qvt);
                        SetDecimal(ws, row, vvtColumnNumber, item.Vvt);
                        SetText(ws, row, obsColumnNumber, item.Obs);
                        row++;
                        count++;


                    }
                }
                gcsApp.Calculation = XlCalculation.xlCalculationAutomatic;
                ExcelFunctions.RowHeight(rngAllCells);
                ExcelFunctions.ColumnWidth(rngAllCells);

                //gcsApp.ScreenUpdating = true;
                //gcsApp.DisplayAlerts = true;
                stopwatch.Stop();
                //MessageBox.Show($"Dados Importados: {count}\nTempo: {stopwatch.Elapsed.Duration()}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                MessageBox.Show($"Dados Importados: {count}\nTempo: {stopwatch.Elapsed:hh\\:mm\\:ss\\.ff}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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

        public void Export()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                Worksheet ws = gcsApp.ActiveSheet;

                if (Path.GetExtension(gcsApp.ActiveWorkbook.FullName) == ".xls")
                {
                    MessageBox.Show($"Função não compatível com versões antigas do Excel.", "X425719", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1); return;
                }

                // REQUIRED FIELDS
                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                var qvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Qvt);
                if (qvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Qvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                var vvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Vvt);
                if (vvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Vvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
                var aCColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ArquivoDeCompra);
                var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Uf);
                var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Empresa);
                var cUnidColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CUnid);
                var cDeptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CDepto);
                var deptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Depto);
                var escalaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Escala);
                var idColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Id);
                var matColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Mat);
                var matSiteColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.MatSite);
                var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cpf);
                var rgColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Rg);
                var dataNascColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.DataNasc);
                var operadoraColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Operadora);
                var descColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desc);
                var obsColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Obs);

                var lstModelTransferData = new List<ModelTransferData>();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;

                var offSetRow = 0;
                var count = 0;

                while (true)
                {
                    var modelTransferData = new ModelTransferData();

                    if (ws.Cells[lastUsedRowByNome, nomeColumnNumber].Offset[offSetRow, 0].Row < 2) { break; }

                    // REQUIRED FIELDS
                    modelTransferData.Nome = GetTextAndTreat(ws, lastUsedRowByNome, nomeColumnNumber, offSetRow, 0);
                    if (modelTransferData.Nome is null) { offSetRow--; continue; }

                    Range activeCellByQvt = ws.Cells[lastUsedRowByNome, qvtColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByQvt.Value2 is null || activeCellByQvt.Value2 == 0) { offSetRow--; continue; }
                    modelTransferData.Qvt = (int)activeCellByQvt.Value2;

                    Range activeCellByVvt = ws.Cells[lastUsedRowByNome, vvtColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByVvt.Value2 is null || activeCellByVvt.Value2 == 0) { offSetRow--; continue; }
                    modelTransferData.Vvt = Math.Round((decimal)activeCellByVvt.Value2, 2);

                    // OPTIONAL FIELDS
                    modelTransferData.Cnpj = GetTextAndTreat(ws, lastUsedRowByNome, cnpjColumnNumber, offSetRow, 0);
                    modelTransferData.ArquivoDeCompra = GetTextAndTreat(ws, lastUsedRowByNome, aCColumnNumber, offSetRow, 0);
                    modelTransferData.Uf = GetTextAndTreat(ws, lastUsedRowByNome, ufColumnNumber, offSetRow, 0);
                    modelTransferData.Empresa = GetTextAndTreat(ws, lastUsedRowByNome, empresaColumnNumber, offSetRow, 0);
                    modelTransferData.CUnid = GetTextAndTreat(ws, lastUsedRowByNome, cUnidColumnNumber, offSetRow, 0);
                    modelTransferData.CDepto = GetTextAndTreat(ws, lastUsedRowByNome, cDeptoColumnNumber, offSetRow, 0);
                    modelTransferData.Depto = GetTextAndTreat(ws, lastUsedRowByNome, deptoColumnNumber, offSetRow, 0);
                    modelTransferData.Escala = GetWorkScheduleAndTreat(ws, lastUsedRowByNome, escalaColumnNumber, offSetRow, 0);
                    modelTransferData.Id = GetTextAndTreat(ws, lastUsedRowByNome, idColumnNumber, offSetRow, 0);
                    modelTransferData.Mat = GetTextAndTreat(ws, lastUsedRowByNome, matColumnNumber, offSetRow, 0);
                    modelTransferData.MatSite = GetTextAndTreat(ws, lastUsedRowByNome, matSiteColumnNumber, offSetRow, 0);
                    modelTransferData.Cpf = GetCpfAndTreat(ws, lastUsedRowByNome, cpfColumnNumber, offSetRow, 0);
                    modelTransferData.Rg = GetTextAndTreat(ws, lastUsedRowByNome, rgColumnNumber, offSetRow, 0);
                    modelTransferData.DataNasc = GetTextAndTreat(ws, lastUsedRowByNome, dataNascColumnNumber, offSetRow, 0);
                    modelTransferData.Operadora = GetTextAndTreat(ws, lastUsedRowByNome, operadoraColumnNumber, offSetRow, 0);
                    modelTransferData.Obs = GetTextAndTreat(ws, lastUsedRowByNome, obsColumnNumber, offSetRow, 0);

                    if (descColumnNumber != -1)
                    {
                        Range activeCellByDesc = ws.Cells[lastUsedRowByNome, descColumnNumber].Offset[offSetRow, 0];
                        if (activeCellByDesc.Value2 != null) { modelTransferData.Desc = Math.Round((decimal)activeCellByDesc.Value2, 2); }
                    }

                    lstModelTransferData.Add(modelTransferData);
                    count++;
                    offSetRow--;
                }

                var orderedCustomers = lstModelTransferData.OrderBy(c => c.ArquivoDeCompra)
                                                                                        .ThenBy(c => c.Uf)
                                                                                        .ThenBy(c => c.Operadora)
                                                                                        .ThenBy(c => c.Empresa)
                                                                                        .ThenBy(c => c.CUnid)
                                                                                        .ThenBy(c => c.CDepto)
                                                                                        .ThenBy(c => c.Depto)
                                                                                        .ThenBy(c => c.Nome);
                string json = JsonConvert.SerializeObject(orderedCustomers.ToArray());
                string fullPath = Path.Combine(Path.GetDirectoryName(gcsApp.ActiveWorkbook.FullName), $"_PurchaseFile_{Tools.GetDateTime()}.json");
                System.IO.File.WriteAllText(fullPath, json);

                MessageBox.Show($"Dados Exportados: {count}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "ERROR: 360425", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }

        }

        private string GetTextAndTreat(Worksheet ws, int row, int column, int offSR, int offSC = 0)
        {
            if (column != -1)
            {
                Range rng = ws.Cells[row, column].Offset[offSR, offSC];

                if (rng != null)
                {
                    string text = Tools.TreatText(rng.Text);
                    if (text != "")
                    {
                        return text;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        private string GetCpfAndTreat(Worksheet ws, int row, int column, int offSR, int offSC = 0)
        {
            if (column != -1)
            {
                Range rng = ws.Cells[row, column].Offset[offSR, offSC];

                if (rng != null)
                {
                    string text = Tools.TreatCpf(rng.Text);
                    return text;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        private string GetWorkScheduleAndTreat(Worksheet ws, int row, int column, int offSR, int offSC = 0)
        {
            if (column != -1)
            {
                Range rng = ws.Cells[row, column].Offset[offSR, offSC];

                if (rng != null)
                {
                    string text = Tools.TreatWorkSchedule(rng.Text);
                    if (text != "")
                    {
                        return text;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        private void SetText(Worksheet ws, int row, int column, string value)
        {
            if (value == null)
                return;
            Range rng = ws.Cells[row, column];
            rng.NumberFormat = "@";
            rng.Value2 = value;
        }

        private void SetInt(Worksheet ws, int row, int column, int value)
        {
            Range rng = ws.Cells[row, column];
            //rng.NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)";
            rng.Value2 = value;
        }

        private void SetDecimal(Worksheet ws, int row, int column, decimal value)
        {
            Range rng = ws.Cells[row, column];
            //rng.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
            rng.Value2 = value;
        }
    }
}
