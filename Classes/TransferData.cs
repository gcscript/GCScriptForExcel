using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    internal class ModelTransferData
    {
        public string Cnpj { get; set; }
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
        public string DataNascimento { get; set; }
        public string Operadora { get; set; }
        public decimal Desc { get; set; }
        public int Qvt { get; set; }
        public decimal Vvt { get; set; }
        public string Obs { get; set; }
    }

    public class TransferData
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public void Save()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                Worksheet ws = gcsApp.ActiveSheet;

                // REQUIRED FIELDS
                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                var qvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Qvt);
                if (qvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Qvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }
                var vvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Vvt);
                if (vvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Vvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
                var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.UF);
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
                var dataNascimentoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.DataNascimento);
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
                    modelTransferData.Vvt = Math.Round((decimal)activeCellByVvt.Value2,2);

                    // OPTIONAL FIELDS
                    modelTransferData.Cnpj = GetTextAndTreat(ws, lastUsedRowByNome, cnpjColumnNumber, offSetRow, 0);
                    modelTransferData.Uf = GetTextAndTreat(ws, lastUsedRowByNome, ufColumnNumber, offSetRow, 0);
                    modelTransferData.Empresa = GetTextAndTreat(ws, lastUsedRowByNome, empresaColumnNumber, offSetRow, 0);
                    modelTransferData.CUnid = GetTextAndTreat(ws, lastUsedRowByNome, cUnidColumnNumber, offSetRow, 0);
                    modelTransferData.CDepto = GetTextAndTreat(ws, lastUsedRowByNome, cDeptoColumnNumber, offSetRow, 0);
                    modelTransferData.Depto = GetTextAndTreat(ws, lastUsedRowByNome, deptoColumnNumber, offSetRow, 0);
                    modelTransferData.Escala = GetTextAndTreat(ws, lastUsedRowByNome, escalaColumnNumber, offSetRow, 0);
                    modelTransferData.Id = GetTextAndTreat(ws, lastUsedRowByNome, idColumnNumber, offSetRow, 0);
                    modelTransferData.Mat = GetTextAndTreat(ws, lastUsedRowByNome, matColumnNumber, offSetRow, 0);
                    modelTransferData.MatSite = GetTextAndTreat(ws, lastUsedRowByNome, matSiteColumnNumber, offSetRow, 0);
                    modelTransferData.Cpf = GetTextAndTreat(ws, lastUsedRowByNome, cpfColumnNumber, offSetRow, 0);
                    modelTransferData.Rg = GetTextAndTreat(ws, lastUsedRowByNome, rgColumnNumber, offSetRow, 0);
                    modelTransferData.DataNascimento = GetTextAndTreat(ws, lastUsedRowByNome, dataNascimentoColumnNumber, offSetRow, 0);
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

                var orderedCustomers = lstModelTransferData.OrderBy(c => c.Uf)
                                                                                        .ThenBy(c => c.Operadora)
                                                                                        .ThenBy(c => c.Empresa)
                                                                                        .ThenBy(c => c.CUnid)
                                                                                        .ThenBy(c => c.CDepto)
                                                                                        .ThenBy(c => c.Depto)
                                                                                        .ThenBy(c => c.Nome);
                string json = JsonConvert.SerializeObject(orderedCustomers.ToArray());
                string fullPath = Path.Combine(Path.GetDirectoryName(gcsApp.ActiveWorkbook.FullName), $"_PurchaseFile_{Tools.GetDateTime()}.json");
                System.IO.File.WriteAllText(fullPath, json);

                MessageBox.Show($"Linhas Salvas: {count}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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
                    //string text = rng.Text;

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

    }
}
