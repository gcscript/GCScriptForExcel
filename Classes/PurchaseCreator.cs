using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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

        public void Create()
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
                lst.Add(new ModelPurchase { Nome = @"==//==\\==" });
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
                lst.Add(new ModelPurchase { Nome = @"==//==\\==" });
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
