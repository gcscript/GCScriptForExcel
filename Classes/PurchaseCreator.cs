using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    internal class ModelPurchase
    {
        public string Empresa { get; set; }
        public string Uf { get; set; }
        public string Operadora { get; set; }
        public string CUnid { get; set; }
        public string CDepto { get; set; }
        public string Depto { get; set; }
        public string Cnpj { get; set; }
        public string Id { get; set; }
        public string Mat { get; set; }
        public string MatSite { get; set; }
        public string Nome { get; set; }
        public string Cpf { get; set; }
        public decimal Desc { get; set; }
        public int Qvt { get; set; }
        public decimal Vvt { get; set; }
        public decimal Tvt { get; set; }
        //public decimal Total { get; set; }
        public decimal Desconto { get; set; }
        public decimal CompraFinal { get; set; }
        public string Obs { get; set; }
    }

    public class PurchaseCreator
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;

        private enum ETypePurchase
        {
            Empresa = 0,
            Uf = 1,
            Operadora = 2,
            CUnid = 3,
            CDepto = 4,
            Depto = 5
        }

        private enum EColumnOrder
        {
            Empresa = 1,
            Uf = 2,
            Operadora = 3,
            CUnid = 4,
            CDepto = 5,
            Depto = 6,
            Cnpj = 7,
            Id = 8,
            Mat = 9,
            MatSite = 10,
            Nome = 11,
            Cpf = 12,
            Desc = 13,
            Qvt = 14,
            Vvt = 15,
            Tvt = 16,
            Desconto = 17,
            CompraFinal = 18,
            Obs = 19,
            Count = 19
        }

        public void CreateBackup()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                Worksheet ws = gcsApp.ActiveSheet;

                var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Empresa);
                if (empresaColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Empresa} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Uf);
                if (ufColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Uf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var operadoraColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Operadora);
                if (operadoraColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Operadora} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cUnidColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CUnid);
                if (cUnidColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CUnid} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cDeptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CDepto);

                var deptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Depto);

                var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
                if (cnpjColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cnpj} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var idColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Id);

                var matColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Mat);
                if (matColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Mat} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var matSiteColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.MatSite);
                if (matSiteColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.MatSite} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
                if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cpf);
                if (cpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var descColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desc);
                if (descColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desc} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var qvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Qvt);
                if (qvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Qvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var vvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Vvt);
                if (vvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Vvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var tvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Tvt);
                if (tvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Tvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var totalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Total);
                if (totalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Total} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var descontoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desconto);
                if (descontoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desconto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var compraFinalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CompraFinal);
                if (compraFinalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CompraFinal} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var obsColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Obs);
                if (obsColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Obs} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var lstGeneralData = new List<ModelPurchase>();

                int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;

                var offSetRow = 0;
                var count = 0;

                while (true)
                {
                    var modelPurchase = new ModelPurchase();

                    if (ws.Cells[lastUsedRowByNome, nomeColumnNumber].Offset[offSetRow, 0].Row < 2) { break; }

                    modelPurchase.Nome = GetTextAndTreat(ws, lastUsedRowByNome, nomeColumnNumber, offSetRow, 0);
                    if (modelPurchase.Nome is null) { offSetRow--; continue; }

                    Range activeCellByTotal = ws.Cells[lastUsedRowByNome, totalColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByTotal.Value2 is null || activeCellByTotal.Value2 == 0) { offSetRow--; continue; }

                    modelPurchase.Empresa = GetTextAndTreat(ws, lastUsedRowByNome, empresaColumnNumber, offSetRow, 0);
                    modelPurchase.Uf = GetTextAndTreat(ws, lastUsedRowByNome, ufColumnNumber, offSetRow, 0);
                    modelPurchase.Operadora = GetTextAndTreat(ws, lastUsedRowByNome, operadoraColumnNumber, offSetRow, 0);
                    modelPurchase.CUnid = GetTextAndTreat(ws, lastUsedRowByNome, cUnidColumnNumber, offSetRow, 0);
                    modelPurchase.CDepto = GetTextAndTreat(ws, lastUsedRowByNome, cDeptoColumnNumber, offSetRow, 0);
                    modelPurchase.Depto = GetTextAndTreat(ws, lastUsedRowByNome, deptoColumnNumber, offSetRow, 0);
                    modelPurchase.Cnpj = GetTextAndTreat(ws, lastUsedRowByNome, cnpjColumnNumber, offSetRow, 0);
                    modelPurchase.Id = GetTextAndTreat(ws, lastUsedRowByNome, idColumnNumber, offSetRow, 0);
                    modelPurchase.Mat = GetTextAndTreat(ws, lastUsedRowByNome, matColumnNumber, offSetRow, 0);
                    modelPurchase.MatSite = GetTextAndTreat(ws, lastUsedRowByNome, matSiteColumnNumber, offSetRow, 0);
                    modelPurchase.Cpf = GetCpfAndTreat(ws, lastUsedRowByNome, cpfColumnNumber, offSetRow, 0);
                    modelPurchase.Obs = GetTextAndTreat(ws, lastUsedRowByNome, obsColumnNumber, offSetRow, 0);

                    Range activeCellByDesc = ws.Cells[lastUsedRowByNome, descColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByDesc.Value2 != null) { modelPurchase.Desc = Math.Round((decimal)activeCellByDesc.Value2, 2); }

                    Range activeCellByQvt = ws.Cells[lastUsedRowByNome, qvtColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByQvt.Value2 != null) { modelPurchase.Qvt = (int)activeCellByQvt.Value2; }

                    Range activeCellByVvt = ws.Cells[lastUsedRowByNome, vvtColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByVvt.Value2 != null) { modelPurchase.Vvt = Math.Round((decimal)activeCellByVvt.Value2, 2); }

                    Range activeCellByTvt = ws.Cells[lastUsedRowByNome, tvtColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByTvt.Value2 != null) { modelPurchase.Tvt = Math.Round((decimal)activeCellByTvt.Value2, 2); }

                    Range activeCellByDesconto = ws.Cells[lastUsedRowByNome, descontoColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByDesconto.Value2 != null) { modelPurchase.Desconto = Math.Round((decimal)activeCellByDesconto.Value2, 2); }

                    Range activeCellByCompraFinal = ws.Cells[lastUsedRowByNome, compraFinalColumnNumber].Offset[offSetRow, 0];
                    if (activeCellByCompraFinal.Value2 != null) { modelPurchase.CompraFinal = Math.Round((decimal)activeCellByCompraFinal.Value2, 2); }


                    lstGeneralData.Add(modelPurchase);
                    count++;
                    offSetRow--;
                }

                var orderedCustomers = new List<ModelPurchase>();

                var createByEmpresa = true;
                var createByUf = true;
                var createByOperadora = true;
                var createByCUnid = true;
                var createByCDepto = false;
                var createByDepto = false;

                if (createByEmpresa && createByUf && createByOperadora && createByCUnid && createByCDepto && createByDepto)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.CDepto)
                                                     .ThenBy(c => c.Depto)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByEmpresa && createByUf && createByOperadora && createByCUnid && createByCDepto)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.CDepto)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByEmpresa && createByUf && createByOperadora && createByCUnid)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByEmpresa && createByUf && createByOperadora)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByEmpresa && createByUf)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Uf)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByEmpresa)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Empresa)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else
                {
                    MessageBox.Show($"Aconteceu um erro!", "ERROR: 838574", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }

                var lstFinal = new List<ModelPurchase>();

                List<ModelPurchase> distinctEmpresas = orderedCustomers.GroupBy(p => p.Empresa)
                                                                   .Select(g => g.First())
                                                                   .ToList();

                foreach (var distinctEmpresa in distinctEmpresas)
                {
                    if (createByUf)
                    {
                        List<ModelPurchase> distinctUfs = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa))
                                                                   .GroupBy(p => new { p.Empresa, p.Uf })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                        foreach (var distinctUf in distinctUfs)
                        {
                            if (createByOperadora)
                            {
                                List<ModelPurchase> distinctOperadoras = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf))
                                                                   .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                foreach (var distinctOperadora in distinctOperadoras)
                                {
                                    if (createByCUnid)
                                    {
                                        List<ModelPurchase> distinctCUnids = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf) && (w.Operadora == distinctOperadora.Operadora))
                                                                   .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora, p.CUnid })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                        foreach (var distinctCUnid in distinctCUnids)
                                        {
                                            if (createByCDepto)
                                            {
                                                List<ModelPurchase> distinctCDeptos = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf) && (w.Operadora == distinctOperadora.Operadora) && (w.CUnid == distinctCUnid.CUnid))
                                                                   .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora, p.CUnid, p.CDepto })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                                foreach (var distinctCDepto in distinctCDeptos)
                                                {
                                                    List<ModelPurchase> lstCDepto = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid) && (x.CDepto == distinctCDepto.CDepto))
                                                                                                    .ToList();
                                                    var subTotalCDepto = SubTotalGCS(lstCDepto, ETypeSubTotal.CDepto, distinctCDepto.CDepto, false);
                                                    lstFinal.AddRange(subTotalCDepto.filteredModel);
                                                }

                                                List<ModelPurchase> lstCUnid = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid))
                                                                                               .ToList();
                                                var subTotalCUnid = SubTotalGCS(lstCUnid, ETypeSubTotal.CDepto, distinctCUnid.CUnid, true);
                                                lstFinal.Add(new ModelPurchase { CUnid = $"{distinctCUnid.CUnid.ToUpper()} Total", CompraFinal = subTotalCUnid.modelSum });
                                            }
                                            else
                                            {
                                                List<ModelPurchase> lstCUnid = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid))
                                                                                               .ToList();
                                                var subTotalCUnid = SubTotalGCS(lstCUnid, ETypeSubTotal.CUnid, distinctCUnid.CUnid, false);
                                                lstFinal.AddRange(subTotalCUnid.filteredModel);
                                            }
                                        }

                                        List<ModelPurchase> lstOperadora = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora))
                                                                                           .ToList();
                                        var subTotalOperadora = SubTotalGCS(lstOperadora, ETypeSubTotal.CDepto, distinctOperadora.Operadora, true);
                                        lstFinal.Add(new ModelPurchase { Operadora = $"{distinctOperadora.Operadora.ToUpper()} Total", CompraFinal = subTotalOperadora.modelSum });
                                    }
                                    else
                                    {
                                        List<ModelPurchase> lstOperadora = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora))
                                                                                           .ToList();
                                        var subTotalOperadora = SubTotalGCS(lstOperadora, ETypeSubTotal.Operadora, distinctOperadora.Operadora, false);
                                        lstFinal.AddRange(subTotalOperadora.filteredModel);
                                    }
                                }
                                List<ModelPurchase> lstUf = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf))
                                                                            .ToList();
                                var subTotalUf = SubTotalGCS(lstUf, ETypeSubTotal.CDepto, distinctUf.Uf, true);
                                lstFinal.Add(new ModelPurchase { Uf = $"{distinctUf.Uf.ToUpper()} Total", CompraFinal = subTotalUf.modelSum });
                            }
                            else
                            {
                                List<ModelPurchase> lstUf = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf))
                                                                            .ToList();
                                var subTotalUf = SubTotalGCS(lstUf, ETypeSubTotal.Uf, distinctUf.Uf, false);
                                lstFinal.AddRange(subTotalUf.filteredModel);
                            }
                        }

                        List<ModelPurchase> lstEmpresa = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa)).ToList();
                        var subTotalEmpresa = SubTotalGCS(lstEmpresa, ETypeSubTotal.CDepto, distinctEmpresa.Empresa, true);
                        lstFinal.Add(new ModelPurchase { Empresa = $"{distinctEmpresa.Empresa.ToUpper()} Total", CompraFinal = subTotalEmpresa.modelSum });
                    }
                    else
                    {
                        List<ModelPurchase> lstEmpresa = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa))
                                                                            .ToList();
                        var subTotalEmpresa = SubTotalGCS(lstEmpresa, ETypeSubTotal.Empresa, distinctEmpresa.Empresa, false);
                        lstFinal.AddRange(subTotalEmpresa.filteredModel);
                    }
                }
                lstFinal.Add(new ModelPurchase { Empresa = $"Total Geral", CompraFinal = SubTotal(orderedCustomers) });

                MessageBox.Show($"OK", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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

        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                var getData = GetData(); if (!getData.success) { return; }

                var separatePurchase = SeparatePurchase(getData.data, ETypePurchase.CUnid);

                string sheetName = "Compra";

                if (ExcelFunctions.ChecksIfSheetExist(sheetName))
                {
                    MessageBox.Show($"A aba {sheetName} já existe!", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }
                gcsApp.Worksheets.Add(After: gcsApp.Worksheets[gcsApp.Worksheets.Count]);
                Worksheet sheet = gcsApp.ActiveSheet;

                sheet.Name = sheetName;

                if (separatePurchase.Count < 1) { return; }

                Range allCells = sheet.Cells;

                ExcelFunctions.FontName(allCells, "Consolas");
                ExcelFunctions.FontSize(allCells, 10);
                ExcelFunctions.VerticalAlignment(allCells, 1);

                Range rngEmpresa = sheet.Range[sheet.Cells[1, EColumnOrder.Empresa], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Empresa]];
                Range rngUf = sheet.Range[sheet.Cells[1, EColumnOrder.Uf], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Uf]];
                Range rngOperadora = sheet.Range[sheet.Cells[1, EColumnOrder.Operadora], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Operadora]];
                Range rngCUnid = sheet.Range[sheet.Cells[1, EColumnOrder.CUnid], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.CUnid]];
                Range rngCDepto = sheet.Range[sheet.Cells[1, EColumnOrder.CDepto], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.CDepto]];
                Range rngDepto = sheet.Range[sheet.Cells[1, EColumnOrder.Depto], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Depto]];
                Range rngCnpj = sheet.Range[sheet.Cells[1, EColumnOrder.Cnpj], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Cnpj]];
                Range rngId = sheet.Range[sheet.Cells[1, EColumnOrder.Id], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Id]];
                Range rngMat = sheet.Range[sheet.Cells[1, EColumnOrder.Mat], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Mat]];
                Range rngMatSite = sheet.Range[sheet.Cells[1, EColumnOrder.MatSite], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.MatSite]];
                Range rngNome = sheet.Range[sheet.Cells[1, EColumnOrder.Nome], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Nome]];
                Range rngCpf = sheet.Range[sheet.Cells[1, EColumnOrder.Cpf], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Cpf]];

                Range rngDesc = sheet.Range[sheet.Cells[1, EColumnOrder.Desc], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Desc]];
                Range rngQvt = sheet.Range[sheet.Cells[1, EColumnOrder.Qvt], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Qvt]];
                Range rngVvt = sheet.Range[sheet.Cells[1, EColumnOrder.Vvt], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Vvt]];
                Range rngTvt = sheet.Range[sheet.Cells[1, EColumnOrder.Tvt], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Tvt]];
                Range rngDesconto = sheet.Range[sheet.Cells[1, EColumnOrder.Desconto], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Desconto]];
                Range rngCompraFinal = sheet.Range[sheet.Cells[1, EColumnOrder.CompraFinal], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.CompraFinal]];
                Range rngObs = sheet.Range[sheet.Cells[1, EColumnOrder.Obs], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.Obs]];

                Range textColumns = gcsApp.Union(rngEmpresa,
                                                 rngUf,
                                                 rngOperadora,
                                                 rngCUnid,
                                                 rngDepto,
                                                 rngDepto,
                                                 rngCnpj,
                                                 rngId,
                                                 rngMat,
                                                 rngMatSite,
                                                 rngNome,
                                                 rngCpf,
                                                 rngObs);

                textColumns.NumberFormat = "@";

                rngQvt.NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)";

                Range decimalColumns = gcsApp.Union(rngDesc,
                                                 rngVvt,
                                                 rngTvt,
                                                 rngDesconto,
                                                 rngCompraFinal);

                decimalColumns.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";

                rngObs.Font.Color = ColorTranslator.FromHtml("#FF0000");
                rngObs.Font.Bold = true;

                gcsApp.ActiveWindow.SplitRow = 1;
                gcsApp.ActiveWindow.FreezePanes = true;

                ColumnsHeader(sheet);
                FillDataColumns(separatePurchase, sheet);

                allCells.EntireColumn.AutoFit();

                Range columnsTitle = gcsApp.Union(rngEmpresa, rngUf, rngOperadora, rngCUnid, rngCDepto, rngDepto);
                columnsTitle.ColumnWidth = 0.08;

                Range columnsHide = gcsApp.Union(rngCnpj, rngId, rngDesc, rngQvt, rngVvt, rngTvt, rngDesconto);
                columnsHide.EntireColumn.Hidden = true;

                Range rngBZPA = sheet.Range[sheet.Cells[1, 1], sheet.Cells[separatePurchase.Count + 1, EColumnOrder.CompraFinal]];

                ExcelFunctions.SetBZPA(sheet, rngBZPA);


                MessageBox.Show($"Terminou", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "x118400", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }

        }

        private static void FillDataColumns(List<ModelPurchase> data, Worksheet sheet)
        {
            int row = 2;
            bool containsProblem = false;

            foreach (var item in data)
            {
                if (item.Nome != null && item.Nome != "")
                {
                    if (item.Nome == "[[[]]]") { row++; continue; }

                    if (item.Obs != null)
                    {
                        if (!item.Obs.Contains("NOVO/SEM CARTAO") && !item.Obs.Contains("2ª VIA"))
                        {
                            Range rng = sheet.Cells[row, EColumnOrder.CompraFinal];
                            ExcelFunctions.Styles_Colors(rng, ExcelFunctions.EStylesColors.Warning);
                            containsProblem = true;
                        }
                    }

                    sheet.Cells[row, EColumnOrder.Empresa].Value2 = item.Empresa;
                    sheet.Cells[row, EColumnOrder.Uf].Value2 = item.Uf;
                    sheet.Cells[row, EColumnOrder.Operadora].Value2 = item.Operadora;
                    sheet.Cells[row, EColumnOrder.CUnid].Value2 = item.CUnid;
                    sheet.Cells[row, EColumnOrder.CDepto].Value2 = item.CDepto;
                    sheet.Cells[row, EColumnOrder.Depto].Value2 = item.Depto;
                    sheet.Cells[row, EColumnOrder.Cnpj].Value2 = item.Cnpj;
                    sheet.Cells[row, EColumnOrder.Id].Value2 = item.Id;
                    sheet.Cells[row, EColumnOrder.Mat].Value2 = item.Mat;
                    sheet.Cells[row, EColumnOrder.MatSite].Value2 = item.MatSite;
                    sheet.Cells[row, EColumnOrder.Nome].Value2 = item.Nome;
                    sheet.Cells[row, EColumnOrder.Cpf].Value2 = item.Cpf;
                    sheet.Cells[row, EColumnOrder.Desc].Value2 = item.Desc;
                    sheet.Cells[row, EColumnOrder.Qvt].Value2 = item.Qvt;
                    sheet.Cells[row, EColumnOrder.Vvt].Value2 = item.Vvt;
                    sheet.Cells[row, EColumnOrder.Tvt].Value2 = item.Tvt;
                    sheet.Cells[row, EColumnOrder.Desconto].Value2 = item.Desconto;
                    sheet.Cells[row, EColumnOrder.CompraFinal].Value2 = item.CompraFinal;
                    sheet.Cells[row, EColumnOrder.Obs].Value2 = item.Obs;
                    row++;
                }
                else
                {
                    if (item.Empresa != null || item.Uf != null || item.Operadora != null || item.CUnid != null || item.CDepto != null || item.Depto != null)
                    {
                        Range rngRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, EColumnOrder.CompraFinal]];

                        if (item.Empresa != null && item.Empresa == "Total Geral")
                        {
                            sheet.Cells[row, EColumnOrder.Empresa].Value2 = item.Empresa;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.TotalGeral);
                        }
                        else if (item.Empresa != null && item.Empresa.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.Empresa].Value2 = item.Empresa;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Empresa);
                        }
                        else if (item.Uf != null && item.Uf.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.Uf].Value2 = item.Uf;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Uf);
                        }
                        else if (item.Operadora != null && item.Operadora.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.Operadora].Value2 = item.Operadora;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Operadora);
                        }
                        else if (item.CUnid != null && item.CUnid.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.CUnid].Value2 = item.CUnid;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.CUnid);
                        }
                        else if (item.CDepto != null && item.CDepto.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.CDepto].Value2 = item.CDepto;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.CDepto);
                        }
                        else if (item.Depto != null && item.Depto.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnOrder.Depto].Value2 = item.Depto;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Depto);
                        }

                        sheet.Cells[row, EColumnOrder.CompraFinal].Value2 = item.CompraFinal;
                        row++; continue;
                    }
                }
            }

            if (containsProblem) { ExcelFunctions.TabColor(sheet, 5); }
        }

        private static void ColumnsHeader(Worksheet sheet)
        {
            sheet.Cells[1, EColumnOrder.Empresa].Value2 = ColumnsName.Empresa;
            sheet.Cells[1, EColumnOrder.Uf].Value2 = ColumnsName.Uf;
            sheet.Cells[1, EColumnOrder.Operadora].Value2 = ColumnsName.Operadora;
            sheet.Cells[1, EColumnOrder.CUnid].Value2 = ColumnsName.CUnid;
            sheet.Cells[1, EColumnOrder.CDepto].Value2 = ColumnsName.CDepto;
            sheet.Cells[1, EColumnOrder.Depto].Value2 = ColumnsName.Depto;
            sheet.Cells[1, EColumnOrder.Cnpj].Value2 = ColumnsName.Cnpj;
            sheet.Cells[1, EColumnOrder.Id].Value2 = ColumnsName.Id;
            sheet.Cells[1, EColumnOrder.Mat].Value2 = ColumnsName.Mat;
            sheet.Cells[1, EColumnOrder.MatSite].Value2 = ColumnsName.MatSite;
            sheet.Cells[1, EColumnOrder.Nome].Value2 = ColumnsName.Nome;
            sheet.Cells[1, EColumnOrder.Cpf].Value2 = ColumnsName.Cpf;
            sheet.Cells[1, EColumnOrder.Desc].Value2 = ColumnsName.Desc;
            sheet.Cells[1, EColumnOrder.Qvt].Value2 = ColumnsName.Qvt;
            sheet.Cells[1, EColumnOrder.Vvt].Value2 = ColumnsName.Vvt;
            sheet.Cells[1, EColumnOrder.Tvt].Value2 = ColumnsName.Tvt;
            sheet.Cells[1, EColumnOrder.Desconto].Value2 = ColumnsName.Desconto;
            sheet.Cells[1, EColumnOrder.CompraFinal].Value2 = ColumnsName.CompraFinal;
            sheet.Cells[1, EColumnOrder.Obs].Value2 = ColumnsName.Obs;
            Range header = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, EColumnOrder.Obs]];
            ExcelFunctions.FontBold(header, true);
        }

        private (List<ModelPurchase> data, bool success) GetData()
        {
            Worksheet ws = gcsApp.ActiveSheet;
            var data = new List<ModelPurchase>();

            var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Empresa);
            if (empresaColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Empresa} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Uf);
            if (ufColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Uf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var operadoraColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Operadora);
            if (operadoraColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Operadora} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var cUnidColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CUnid);
            if (cUnidColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CUnid} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var cDeptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CDepto);

            var deptoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Depto);

            var cnpjColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cnpj);
            if (cnpjColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cnpj} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var idColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Id);

            var matColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Mat);
            if (matColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Mat} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var matSiteColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.MatSite);
            if (matSiteColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.MatSite} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var nomeColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Nome);
            if (nomeColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Nome} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var cpfColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Cpf);
            if (cpfColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Cpf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var descColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desc);
            if (descColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desc} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var qvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Qvt);
            if (qvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Qvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var vvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Vvt);
            if (vvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Vvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var tvtColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Tvt);
            if (tvtColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Tvt} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var totalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Total);
            if (totalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Total} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var descontoColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Desconto);
            if (descontoColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Desconto} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var compraFinalColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.CompraFinal);
            if (compraFinalColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.CompraFinal} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            var obsColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Obs);
            if (obsColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Obs} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return (data, false); }

            int lastUsedRowByNome = ws.Cells[1048576, nomeColumnNumber].End(XlDirection.xlUp).Row;

            var offSetRow = 0;
            var count = 0;

            while (true)
            {
                var currentData = new ModelPurchase();

                if (ws.Cells[lastUsedRowByNome, nomeColumnNumber].Offset[offSetRow, 0].Row < 2) { break; }

                currentData.Nome = GetTextAndTreat(ws, lastUsedRowByNome, nomeColumnNumber, offSetRow, 0);
                if (currentData.Nome is null) { offSetRow--; continue; }

                Range activeCellByTotal = ws.Cells[lastUsedRowByNome, totalColumnNumber].Offset[offSetRow, 0];
                if (activeCellByTotal.Value2 is null || activeCellByTotal.Value2 == 0) { offSetRow--; continue; }

                currentData.Empresa = GetTextAndTreat(ws, lastUsedRowByNome, empresaColumnNumber, offSetRow, 0);
                currentData.Uf = GetTextAndTreat(ws, lastUsedRowByNome, ufColumnNumber, offSetRow, 0);
                currentData.Operadora = GetTextAndTreat(ws, lastUsedRowByNome, operadoraColumnNumber, offSetRow, 0);
                currentData.CUnid = GetTextAndTreat(ws, lastUsedRowByNome, cUnidColumnNumber, offSetRow, 0);
                currentData.CDepto = GetTextAndTreat(ws, lastUsedRowByNome, cDeptoColumnNumber, offSetRow, 0);
                currentData.Depto = GetTextAndTreat(ws, lastUsedRowByNome, deptoColumnNumber, offSetRow, 0);
                currentData.Cnpj = GetTextAndTreat(ws, lastUsedRowByNome, cnpjColumnNumber, offSetRow, 0);
                currentData.Id = GetTextAndTreat(ws, lastUsedRowByNome, idColumnNumber, offSetRow, 0);
                currentData.Mat = GetTextAndTreat(ws, lastUsedRowByNome, matColumnNumber, offSetRow, 0);
                currentData.MatSite = GetTextAndTreat(ws, lastUsedRowByNome, matSiteColumnNumber, offSetRow, 0);
                currentData.Cpf = GetCpfAndTreat(ws, lastUsedRowByNome, cpfColumnNumber, offSetRow, 0);
                currentData.Obs = GetTextAndTreat(ws, lastUsedRowByNome, obsColumnNumber, offSetRow, 0);

                Range activeCellByDesc = ws.Cells[lastUsedRowByNome, descColumnNumber].Offset[offSetRow, 0];
                if (activeCellByDesc.Value2 != null) { currentData.Desc = Math.Round((decimal)activeCellByDesc.Value2, 2); }

                Range activeCellByQvt = ws.Cells[lastUsedRowByNome, qvtColumnNumber].Offset[offSetRow, 0];
                if (activeCellByQvt.Value2 != null) { currentData.Qvt = (int)activeCellByQvt.Value2; }

                Range activeCellByVvt = ws.Cells[lastUsedRowByNome, vvtColumnNumber].Offset[offSetRow, 0];
                if (activeCellByVvt.Value2 != null) { currentData.Vvt = Math.Round((decimal)activeCellByVvt.Value2, 2); }

                Range activeCellByTvt = ws.Cells[lastUsedRowByNome, tvtColumnNumber].Offset[offSetRow, 0];
                if (activeCellByTvt.Value2 != null) { currentData.Tvt = Math.Round((decimal)activeCellByTvt.Value2, 2); }

                Range activeCellByDesconto = ws.Cells[lastUsedRowByNome, descontoColumnNumber].Offset[offSetRow, 0];
                if (activeCellByDesconto.Value2 != null) { currentData.Desconto = Math.Round((decimal)activeCellByDesconto.Value2, 2); }

                Range activeCellByCompraFinal = ws.Cells[lastUsedRowByNome, compraFinalColumnNumber].Offset[offSetRow, 0];
                if (activeCellByCompraFinal.Value2 != null) { currentData.CompraFinal = Math.Round((decimal)activeCellByCompraFinal.Value2, 2); }

                data.Add(currentData);
                count++;
                offSetRow--;
            }

            return (data, true);
        }

        private List<ModelPurchase> SeparatePurchase(List<ModelPurchase> model, ETypePurchase typeSeparation)
        {
            var orderedCustomers = new List<ModelPurchase>();

            switch (typeSeparation)
            {
                case ETypePurchase.Empresa:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case ETypePurchase.Uf:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case ETypePurchase.Operadora:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case ETypePurchase.CUnid:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.CUnid)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case ETypePurchase.CDepto:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.CUnid)
                                            .ThenBy(c => c.CDepto)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case ETypePurchase.Depto:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.CUnid)
                                            .ThenBy(c => c.CDepto)
                                            .ThenBy(c => c.Depto)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
            }

            var lstFinal = new List<ModelPurchase>();

            List<ModelPurchase> distinctEmpresas = orderedCustomers.GroupBy(p => p.Empresa)
                                                               .Select(g => g.First())
                                                               .ToList();

            foreach (var distinctEmpresa in distinctEmpresas)
            {
                if (typeSeparation == ETypePurchase.Empresa)
                {
                    List<ModelPurchase> lstEmpresa = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa))
                                                                        .ToList();
                    var subTotalEmpresa = SubTotalGCS(lstEmpresa, ETypeSubTotal.Empresa, distinctEmpresa.Empresa, false);
                    lstFinal.AddRange(subTotalEmpresa.filteredModel);
                }
                else
                {
                    List<ModelPurchase> distinctUfs = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa))
                                                               .GroupBy(p => new { p.Empresa, p.Uf })
                                                               .Select(g => g.First())
                                                               .ToList();

                    foreach (var distinctUf in distinctUfs)
                    {
                        if (typeSeparation == ETypePurchase.Uf)
                        {
                            List<ModelPurchase> lstUf = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf))
                                                                        .ToList();
                            var subTotalUf = SubTotalGCS(lstUf, ETypeSubTotal.Uf, distinctUf.Uf, false);
                            lstFinal.AddRange(subTotalUf.filteredModel);
                        }
                        else
                        {
                            List<ModelPurchase> distinctOperadoras = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf))
                                                               .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora })
                                                               .Select(g => g.First())
                                                               .ToList();

                            foreach (var distinctOperadora in distinctOperadoras)
                            {
                                if (typeSeparation == ETypePurchase.Operadora)
                                {
                                    List<ModelPurchase> lstOperadora = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora))
                                                                                       .ToList();
                                    var subTotalOperadora = SubTotalGCS(lstOperadora, ETypeSubTotal.Operadora, distinctOperadora.Operadora, false);
                                    lstFinal.AddRange(subTotalOperadora.filteredModel);
                                }
                                else
                                {
                                    List<ModelPurchase> distinctCUnids = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf) && (w.Operadora == distinctOperadora.Operadora))
                                                               .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora, p.CUnid })
                                                               .Select(g => g.First())
                                                               .ToList();

                                    foreach (var distinctCUnid in distinctCUnids)
                                    {
                                        if (typeSeparation == ETypePurchase.CUnid)
                                        {
                                            List<ModelPurchase> lstCUnid = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid))
                                                                                           .ToList();
                                            var subTotalCUnid = SubTotalGCS(lstCUnid, ETypeSubTotal.CUnid, distinctCUnid.CUnid, false);
                                            lstFinal.AddRange(subTotalCUnid.filteredModel);
                                        }
                                        else
                                        {
                                            List<ModelPurchase> distinctCDeptos = orderedCustomers.Where(w => (w.Empresa == distinctEmpresa.Empresa) && (w.Uf == distinctUf.Uf) && (w.Operadora == distinctOperadora.Operadora) && (w.CUnid == distinctCUnid.CUnid))
                                                               .GroupBy(p => new { p.Empresa, p.Uf, p.Operadora, p.CUnid, p.CDepto })
                                                               .Select(g => g.First())
                                                               .ToList();

                                            foreach (var distinctCDepto in distinctCDeptos)
                                            {
                                                List<ModelPurchase> lstCDepto = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid) && (x.CDepto == distinctCDepto.CDepto))
                                                                                                .ToList();
                                                var subTotalCDepto = SubTotalGCS(lstCDepto, ETypeSubTotal.CDepto, distinctCDepto.CDepto, false);
                                                lstFinal.AddRange(subTotalCDepto.filteredModel);
                                            }

                                            List<ModelPurchase> lstCUnid = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora) && (x.CUnid == distinctCUnid.CUnid))
                                                                                           .ToList();
                                            var subTotalCUnid = SubTotalGCS(lstCUnid, ETypeSubTotal.CDepto, distinctCUnid.CUnid, true);
                                            lstFinal.Add(new ModelPurchase { CUnid = $"{distinctCUnid.CUnid.ToUpper()} Total", CompraFinal = subTotalCUnid.modelSum });
                                        }
                                    }

                                    List<ModelPurchase> lstOperadora = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora))
                                                                                       .ToList();
                                    var subTotalOperadora = SubTotalGCS(lstOperadora, ETypeSubTotal.CDepto, distinctOperadora.Operadora, true);
                                    lstFinal.Add(new ModelPurchase { Operadora = $"{distinctOperadora.Operadora.ToUpper()} Total", CompraFinal = subTotalOperadora.modelSum });
                                }
                            }
                            List<ModelPurchase> lstUf = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf))
                                                                        .ToList();
                            var subTotalUf = SubTotalGCS(lstUf, ETypeSubTotal.CDepto, distinctUf.Uf, true);
                            lstFinal.Add(new ModelPurchase { Uf = $"{distinctUf.Uf.ToUpper()} Total", CompraFinal = subTotalUf.modelSum });
                        }
                    }

                    List<ModelPurchase> lstEmpresa = orderedCustomers.Where(x => x.Empresa == distinctEmpresa.Empresa)
                                                                     .ToList();
                    var subTotalEmpresa = SubTotalGCS(lstEmpresa, ETypeSubTotal.CDepto, distinctEmpresa.Empresa, true);
                    lstFinal.Add(new ModelPurchase { Empresa = $"{distinctEmpresa.Empresa.ToUpper()} Total", CompraFinal = subTotalEmpresa.modelSum });
                }
            }
            lstFinal.Add(new ModelPurchase { Empresa = $"Total Geral", CompraFinal = SubTotal(orderedCustomers) });
            return lstFinal;
        }

        private decimal SubTotal(List<ModelPurchase> origin)
        {
            var lstSum = origin.Sum(x => x.CompraFinal);
            return lstSum;
        }

        private enum ETypeSubTotal
        {
            Empresa = 0,
            Uf = 1,
            Operadora = 2,
            CUnid = 3,
            CDepto = 4,
            Depto = 5
        }

        private (List<ModelPurchase> filteredModel, decimal modelSum) SubTotalGCS(List<ModelPurchase> modelPurchase, ETypeSubTotal type, string name, bool onlySum = false)
        {
            List<ModelPurchase> lst = new List<ModelPurchase>();
            decimal modelSum = 0;

            if (modelPurchase.Count < 1) { return (lst, modelSum); }

            modelSum = modelPurchase.Sum(x => x.CompraFinal);

            if (onlySum)
                return (lst, modelSum);

            #region ZEROED
            List<ModelPurchase> lstZeroed = modelPurchase.Where(x => x.CompraFinal == 0)
                                                  .ToList();

            if (lstZeroed.Count > 0)
            {
                lst.AddRange(lstZeroed);
                lst.Add(new ModelPurchase { Nome = "[[[]]]" });
            }
            #endregion

            #region PROBLEMS
            List<ModelPurchase> lstProblems = modelPurchase.Where(x => (x.Obs != null) && (!x.Obs.Contains("NOVO/SEM CARTAO") || x.Obs.Contains("2ª VIA")))
                                                           .OrderBy(c => c.Obs)
                                                           .ThenBy(c => c.Nome)
                                                           .ToList();

            if (lstProblems.Count > 0)
            {
                lst.AddRange(lstProblems);
                lst.Add(new ModelPurchase { Nome = "[[[]]]" });
            }
            #endregion

            #region NEWS & 2ª VIA
            List<ModelPurchase> lstNews = modelPurchase.Where(x => (x.Obs != null) && (x.Obs.Contains("NOVO/SEM CARTAO") || x.Obs.Contains("2ª VIA")))
                                                       .OrderBy(c => c.Nome)
                                                       .ToList();

            if (lstNews.Count > 0)
                lst.AddRange(lstNews);
            #endregion

            #region PURCHASE
            List<ModelPurchase> lstPurchase = modelPurchase.Where(x => (x.CompraFinal != 0) && (x.Obs == null))
                                                           .OrderBy(c => c.Nome)
                                                           .ToList();

            lst.AddRange(lstPurchase);
            #endregion

            switch (type)
            {
                case ETypeSubTotal.Empresa:
                    lst.Add(new ModelPurchase { Empresa = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                case ETypeSubTotal.Uf:
                    lst.Add(new ModelPurchase { Uf = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                case ETypeSubTotal.Operadora:
                    lst.Add(new ModelPurchase { Operadora = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                case ETypeSubTotal.CUnid:
                    lst.Add(new ModelPurchase { CUnid = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                case ETypeSubTotal.CDepto:
                    lst.Add(new ModelPurchase { CDepto = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                case ETypeSubTotal.Depto:
                    lst.Add(new ModelPurchase { Depto = $"{name.ToUpper()} Total", CompraFinal = modelSum });
                    break;
                default:
                    break;
            }

            return (lst, modelSum);
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
