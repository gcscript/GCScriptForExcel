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
        public string Uf { get; set; }
        public string Operadora { get; set; }
        public string Empresa { get; set; }
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
        public decimal Total { get; set; }
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

                var ufColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Uf);
                if (ufColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Uf} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var operadoraColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Operadora);
                if (operadoraColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Operadora} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

                var empresaColumnNumber = ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Empresa);
                if (empresaColumnNumber == -1) { MessageBox.Show($"A coluna {ColumnsName.Empresa} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

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

                    modelPurchase.Uf = GetTextAndTreat(ws, lastUsedRowByNome, ufColumnNumber, offSetRow, 0);
                    modelPurchase.Operadora = GetTextAndTreat(ws, lastUsedRowByNome, operadoraColumnNumber, offSetRow, 0);
                    modelPurchase.Empresa = GetTextAndTreat(ws, lastUsedRowByNome, empresaColumnNumber, offSetRow, 0);
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

                var createByUf = true;
                var createByOperadora = true;
                var createByEmpresa = true;
                var createByCUnid = true;
                var createByCDepto = true;
                var createByDepto = false;

                if (createByUf && createByOperadora && createByEmpresa && createByCUnid && createByCDepto && createByDepto)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Empresa)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.CDepto)
                                                     .ThenBy(c => c.Depto)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByUf && createByOperadora && createByEmpresa && createByCUnid && createByCDepto)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Empresa)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.CDepto)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByUf && createByOperadora && createByEmpresa && createByCUnid)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Empresa)
                                                     .ThenBy(c => c.CUnid)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByUf && createByOperadora && createByEmpresa)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Empresa)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByUf && createByOperadora)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Operadora)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else if (createByUf)
                {
                    orderedCustomers = lstGeneralData.OrderBy(c => c.Uf)
                                                     .ThenBy(c => c.Nome)
                                                     .ToList();
                }
                else
                {
                    MessageBox.Show($"Aconteceu um erro!", "ERROR: 838574", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
                    return;
                }

                //var distinct = new List<ModelPurchase>();
                var lstFinal = new List<ModelPurchase>();

                List<ModelPurchase> distinctUfs = orderedCustomers.GroupBy(p => p.Uf)
                                                                   .Select(g => g.First())
                                                                   .ToList();

                //List<ModelPurchase> distinctOperadoras = orderedCustomers.GroupBy(p => new { p.Uf, p.Operadora })
                //                                                   .Select(g => g.First())
                //                                                   .ToList();

                //List<ModelPurchase> distinctEmpresas = orderedCustomers.GroupBy(p => new { p.Uf, p.Operadora, p.Empresa })
                //                                                   .Select(g => g.First())
                //                                                   .ToList();

                //List<ModelPurchase> distinctCUnids = orderedCustomers.GroupBy(p => new { p.Uf, p.Operadora, p.Empresa, p.CUnid })
                //                                                   .Select(g => g.First())
                //                                                   .ToList();

                //List<ModelPurchase> distinctCDeptos = orderedCustomers.GroupBy(p => new { p.Uf, p.Operadora, p.Empresa, p.CUnid, p.CDepto })
                //                                                   .Select(g => g.First())
                //                                                   .ToList();

                //List<ModelPurchase> distinctDeptos = orderedCustomers.GroupBy(p => new { p.Uf, p.Operadora, p.Empresa, p.CUnid, p.CDepto, p.Depto })
                //                                                   .Select(g => g.First())
                //                                                   .ToList();

                foreach (var uf in distinctUfs)
                {
                    if (createByOperadora)
                    {
                        List<ModelPurchase> distinctOperadoras = orderedCustomers.Where(w => (w.Uf == uf.Uf))
                                                                   .GroupBy(p => new { p.Uf, p.Operadora })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                        foreach (var operadora in distinctOperadoras)
                        {
                            if (createByEmpresa)
                            {
                                List<ModelPurchase> distinctEmpresas = orderedCustomers.Where(w => (w.Uf == uf.Uf) && (w.Operadora == operadora.Operadora))
                                                                   .GroupBy(p => new { p.Uf, p.Operadora, p.Empresa})
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                foreach (var empresa in distinctEmpresas)
                                {
                                    if (createByCUnid)
                                    {
                                        List<ModelPurchase> distinctCUnids = orderedCustomers.Where(w => (w.Uf == uf.Uf) && (w.Operadora == operadora.Operadora) && (w.Empresa == empresa.Empresa))
                                                                   .GroupBy(p => new { p.Uf, p.Operadora, p.Empresa, p.CUnid })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                        foreach (var cunid in distinctCUnids)
                                        {
                                            if (createByCDepto)
                                            {
                                                List<ModelPurchase> distinctCDeptos = orderedCustomers.Where(w => (w.Uf == uf.Uf) && (w.Operadora == operadora.Operadora) && (w.Empresa == empresa.Empresa) && (w.CUnid == cunid.CUnid))
                                                                   .GroupBy(p => new { p.Uf, p.Operadora, p.Empresa, p.CUnid, p.CDepto })
                                                                   .Select(g => g.First())
                                                                   .ToList();

                                                foreach (var cdepto in distinctCDeptos)
                                                {
                                                    var (subTotalCDepto, _) = SubTotalCDepto(orderedCustomers, uf.Uf, operadora.Operadora, empresa.Empresa, cunid.CUnid, cdepto.CDepto);
                                                    lstFinal.AddRange(subTotalCDepto);
                                                }

                                                var (_, subTotalCUnidSum) = SubTotalCUnid(orderedCustomers, uf.Uf, operadora.Operadora, empresa.Empresa, cunid.CUnid);
                                                lstFinal.Add(new ModelPurchase { CUnid = $"{cunid.CUnid.ToUpper()} Total", CompraFinal = subTotalCUnidSum });
                                            }
                                            else
                                            {
                                                var (subTotalCUnid, _) = SubTotalCUnid(orderedCustomers, uf.Uf, operadora.Operadora, empresa.Empresa, cunid.CUnid);
                                                lstFinal.AddRange(subTotalCUnid);
                                            }
                                        }

                                        var (_, subTotalEmpresaSum) = SubTotalEmpresa(orderedCustomers, uf.Uf, operadora.Operadora, empresa.Empresa);
                                        lstFinal.Add(new ModelPurchase { Empresa = $"{empresa.Empresa.ToUpper()} Total", CompraFinal = subTotalEmpresaSum });
                                    }
                                    else
                                    {
                                        var (subTotalEmpresa, _) = SubTotalEmpresa(orderedCustomers, uf.Uf, operadora.Operadora, empresa.Empresa);
                                        lstFinal.AddRange(subTotalEmpresa);
                                    }
                                }
                                var (_, subTotalOperadoraSum) = SubTotalOperadora(orderedCustomers, uf.Uf, operadora.Operadora);
                                lstFinal.Add(new ModelPurchase { Operadora = $"{operadora.Operadora.ToUpper()} Total", CompraFinal = subTotalOperadoraSum });
                            }
                            else
                            {
                                var (subTotalOperadora, _) = SubTotalOperadora(orderedCustomers, uf.Uf, operadora.Operadora);
                                lstFinal.AddRange(subTotalOperadora);
                            }
                        }
                        var (_, subTotalUfSum) = SubTotalUf(orderedCustomers, uf.Uf);
                        lstFinal.Add(new ModelPurchase { Uf = $"{uf.Uf.ToUpper()} Total", CompraFinal = subTotalUfSum });
                    }
                    else
                    {
                        var (subTotalUf, _) = SubTotalUf(orderedCustomers, uf.Uf);
                        lstFinal.AddRange(subTotalUf);
                    }
                }


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

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalUf(List<ModelPurchase> origin, string uf)
        {
            List<ModelPurchase> lst = origin.Where(x => x.Uf == uf).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { Uf = $"{uf.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
        }

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalOperadora(List<ModelPurchase> origin, string uf, string operadora)
        {
            List<ModelPurchase> lst = origin.Where(x => (x.Uf == uf) && (x.Operadora == operadora)).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { Operadora = $"{operadora.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
        }

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalEmpresa(List<ModelPurchase> origin, string uf, string operadora, string empresa)
        {
            List<ModelPurchase> lst = origin.Where(x => (x.Uf == uf) && (x.Operadora == operadora) && (x.Empresa == empresa)).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { Empresa = $"{empresa.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
        }

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalCUnid(List<ModelPurchase> origin, string uf, string operadora, string empresa, string cunid)
        {
            List<ModelPurchase> lst = origin.Where(x => (x.Uf == uf) && (x.Operadora == operadora) && (x.Empresa == empresa) && (x.CUnid == cunid)).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { CUnid = $"{cunid.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
        }

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalCDepto(List<ModelPurchase> origin, string uf, string operadora, string empresa, string cunid, string cdepto)
        {
            List<ModelPurchase> lst = origin.Where(x => (x.Uf == uf) && (x.Operadora == operadora) && (x.Empresa == empresa) && (x.CUnid == cunid) && (x.CDepto == cdepto)).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { CDepto = $"{cdepto.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
        }

        private (List<ModelPurchase> subTotal, decimal compraFinalSum) SubTotalDepto(List<ModelPurchase> origin, string uf, string operadora, string empresa, string cunid, string cdepto, string depto)
        {
            List<ModelPurchase> lst = origin.Where(x => (x.Uf == uf) && (x.Operadora == operadora) && (x.Empresa == empresa) && (x.CUnid == cunid) && (x.CDepto == cdepto) && (x.Depto == depto)).ToList();
            var lstSum = lst.Sum(x => x.CompraFinal);
            lst.Add(new ModelPurchase { Depto = $"{depto.ToUpper()} Total", CompraFinal = lstSum });
            return (lst, lstSum);
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
