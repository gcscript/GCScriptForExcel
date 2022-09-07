using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public class TransferColumnData
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
        gcsApplication gcsApp = Globals.ThisAddIn.Application;

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

                var lstTransferColumnData = new List<TransferColumnData>();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;

                var offSetRow = 0;

                while (true)
                {
                    var transferColumnData = new TransferColumnData();

                    // REQUIRED FIELDS
                    Range activeCellByNome = ws.Cells[lastUsedRowByNome, nomeColumnNumber].Offset[offSetRow, 0]; if (activeCellByNome.Row < 2) { break; }
                    transferColumnData.Nome = activeCellByNome.Text;

                    Range activeCellByQvt = ws.Cells[lastUsedRowByNome, qvtColumnNumber].Offset[offSetRow, 0];
                    transferColumnData.Qvt = (int)activeCellByQvt.Value2;

                    Range activeCellByVvt = ws.Cells[lastUsedRowByNome, vvtColumnNumber].Offset[offSetRow, 0];
                    transferColumnData.Vvt = (decimal)activeCellByVvt.Value2;

                    // OPTIONAL FIELDS

                    if (cnpjColumnNumber != -1) {
                        Range activeCellByCnpj = ws.Cells[lastUsedRowByNome, cnpjColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Cnpj = activeCellByCnpj.Text;
                    }

                    if (ufColumnNumber != -1)
                    {
                        Range activeCellByUf = ws.Cells[lastUsedRowByNome, ufColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Uf = activeCellByUf.Text;
                    }

                    if (empresaColumnNumber != -1)
                    {
                        Range activeCellByEmpresa = ws.Cells[lastUsedRowByNome, empresaColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Empresa = activeCellByEmpresa.Text;
                    }

                    if (cUnidColumnNumber != -1)
                    {
                        Range activeCellByCUnid = ws.Cells[lastUsedRowByNome, cUnidColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.CUnid = activeCellByCUnid.Text;
                    }

                    if (cDeptoColumnNumber != -1)
                    {
                        Range activeCellByCDepto = ws.Cells[lastUsedRowByNome, cDeptoColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.CDepto = activeCellByCDepto.Text;
                    }

                    if (deptoColumnNumber != -1)
                    {
                        Range activeCellByDepto = ws.Cells[lastUsedRowByNome, deptoColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Depto = activeCellByDepto.Text;
                    }

                    if (escalaColumnNumber != -1)
                    {
                        Range activeCellByEscala = ws.Cells[lastUsedRowByNome, escalaColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Escala = activeCellByEscala.Text;
                    }

                    if (idColumnNumber != -1)
                    {
                        Range activeCellById = ws.Cells[lastUsedRowByNome, idColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Id = activeCellById.Text;
                    }

                    if (matColumnNumber != -1)
                    {
                        Range activeCellByMat = ws.Cells[lastUsedRowByNome, matColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Mat = activeCellByMat.Text;
                    }

                    if (matSiteColumnNumber != -1)
                    {
                        Range activeCellByMatSite = ws.Cells[lastUsedRowByNome, matSiteColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.MatSite = activeCellByMatSite.Text;
                    }

                    if (cpfColumnNumber != -1)
                    {
                        Range activeCellByCpf = ws.Cells[lastUsedRowByNome, cpfColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Cpf = activeCellByCpf.Text;
                    }

                    if (rgColumnNumber != -1)
                    {
                        Range activeCellByRg = ws.Cells[lastUsedRowByNome, rgColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Rg = activeCellByRg.Text;
                    }

                    if (dataNascimentoColumnNumber != -1)
                    {
                        Range activeCellByDataNascimento = ws.Cells[lastUsedRowByNome, dataNascimentoColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.DataNascimento = activeCellByDataNascimento.Text;
                    }

                    if (operadoraColumnNumber != -1)
                    {
                        Range activeCellByOperadora = ws.Cells[lastUsedRowByNome, operadoraColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Operadora = activeCellByOperadora.Text;
                    }

                    if (descColumnNumber != -1)
                    {
                        Range activeCellByDesc = ws.Cells[lastUsedRowByNome, descColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Desc = (decimal)activeCellByDesc.Value2;
                    }

                    if (obsColumnNumber != -1)
                    {
                        Range activeCellByObs = ws.Cells[lastUsedRowByNome, obsColumnNumber].Offset[offSetRow, 0];
                        transferColumnData.Obs = activeCellByObs.Text;
                    }

                    lstTransferColumnData.Add(transferColumnData);
                    offSetRow--;
                }
                string json = JsonConvert.SerializeObject(lstTransferColumnData.ToArray());
                System.IO.File.WriteAllText(@"D:\teste.json", json);

                MessageBox.Show($"Terminou!", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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

    }
}
