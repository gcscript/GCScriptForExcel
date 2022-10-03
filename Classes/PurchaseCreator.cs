using GCScript_for_Excel.Models;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static GCScript_for_Excel.Models.Enums;
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
        public decimal Desconto { get; set; }
        public decimal Parcela1 { get; set; }
        public decimal Parcela2 { get; set; }
        public decimal Parcela3 { get; set; }
        public decimal CompraFinal { get; set; }
        public string Obs { get; set; }
    }

    public class PurchaseCreator
    {
        readonly gcsApplication gcsApp = Globals.ThisAddIn.Application;

        private enum EColumnIndex
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
            Parcela1 = 18,
            Parcela2 = 19,
            Parcela3 = 20,
            CompraFinal = 21,
            Obs = 22,
        }

        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;
                gcsApp.DisplayAlerts = false;

                Stopwatch stopwatch = Stopwatch.StartNew();
                var getData = GetData(); if (!getData.success) { return; }
                var tabNameList = new List<string>();

                switch (cl_Settings.PurchaseCreatorTabOption)
                {
                    case EPurchaseCreatorTabOption.Empresa:
                        tabNameList = getData.data.Select(x => x.Empresa).Distinct().OrderBy(o => o).ToList();
                        break;
                    case EPurchaseCreatorTabOption.Uf:
                        tabNameList = getData.data.Select(x => x.Uf).Distinct().OrderBy(o => o).ToList();
                        break;
                    case EPurchaseCreatorTabOption.Operadora:
                        tabNameList = getData.data.Select(x => x.Operadora).Distinct().OrderBy(o => o).ToList();
                        break;
                    case EPurchaseCreatorTabOption.CUnid:
                        tabNameList = getData.data.Select(x => x.CUnid).Distinct().OrderBy(o => o).ToList();
                        break;
                    case EPurchaseCreatorTabOption.CustomName:
                        tabNameList.Add(cl_Settings.PurchaseCreatorTabName);
                        break;
                }

                var separatePurchase = new List<ModelPurchase>();

                foreach (var item in tabNameList)
                {
                    switch (cl_Settings.PurchaseCreatorTabOption)
                    {
                        case EPurchaseCreatorTabOption.Empresa:
                            separatePurchase = SeparatePurchase(getData.data.Where(x => x.Empresa == item).ToList(), cl_Settings.PurchaseCreatorSubtotalOption);
                            break;
                        case EPurchaseCreatorTabOption.Uf:
                            separatePurchase = SeparatePurchase(getData.data.Where(x => x.Uf == item).ToList(), cl_Settings.PurchaseCreatorSubtotalOption);
                            break;
                        case EPurchaseCreatorTabOption.Operadora:
                            separatePurchase = SeparatePurchase(getData.data.Where(x => x.Operadora == item).ToList(), cl_Settings.PurchaseCreatorSubtotalOption);
                            break;
                        case EPurchaseCreatorTabOption.CUnid:
                            separatePurchase = SeparatePurchase(getData.data.Where(x => x.CUnid == item).ToList(), cl_Settings.PurchaseCreatorSubtotalOption);
                            break;
                        default:
                            separatePurchase = SeparatePurchase(getData.data, cl_Settings.PurchaseCreatorSubtotalOption);
                            break;
                    }

                    gcsApp.Worksheets.Add(After: gcsApp.Worksheets[gcsApp.Worksheets.Count]);
                    Worksheet sheet = gcsApp.ActiveSheet;

                    sheet.Name = TreatTabName(item);

                    if (separatePurchase.Count < 1) { return; }

                    Range allCells = sheet.Cells;

                    ExcelFunctions.FontName(allCells, "Consolas");
                    ExcelFunctions.FontSize(allCells, 10);
                    ExcelFunctions.VerticalAlignment(allCells, 1);

                    var columnsRange = GetColumnsRange(sheet, separatePurchase);

                    SetFormatting(columnsRange);
                    SetHeader(sheet);
                    SetBody(sheet, separatePurchase);
                    SetFormattingInView(columnsRange);

                    Range rngBZPA = sheet.Range[sheet.Cells[1, 1], sheet.Cells[separatePurchase.Count + 1, EColumnIndex.CompraFinal]];

                    ExcelFunctions.SetBZPA(sheet, rngBZPA);
                    RemoveEmptyColumns(separatePurchase, columnsRange);
                }

                stopwatch.Stop();
                MessageBox.Show($"Compra Criada com Sucesso!\nTempo: {stopwatch.Elapsed:hh\\:mm\\:ss\\.ff}", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1);
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

        private static void RemoveEmptyColumns(List<ModelPurchase> data, ColumnsRangeOnPurchase rangeAddress)
        {
            List<Range> emptyColumnsList = new List<Range>();
            if (data.All(x => x.CUnid == null))
                emptyColumnsList.Add(rangeAddress.CUnid);
            if (data.All(x => x.CDepto == null))
                emptyColumnsList.Add(rangeAddress.CDepto);
            if (data.All(x => x.Depto == null))
                emptyColumnsList.Add(rangeAddress.Depto);
            if (data.All(x => x.Id == null))
                emptyColumnsList.Add(rangeAddress.Id);

            if (cl_Settings.PurchaseCreatorSplitPurchaseOption == EPurchaseCreatorSplitPurchaseOption.One)
            {
                emptyColumnsList.Add(rangeAddress.Parcela1);
                emptyColumnsList.Add(rangeAddress.Parcela2);
                emptyColumnsList.Add(rangeAddress.Parcela3);
            }
            else if (cl_Settings.PurchaseCreatorSplitPurchaseOption == EPurchaseCreatorSplitPurchaseOption.Two)
            {
                emptyColumnsList.Add(rangeAddress.Parcela3);
            }

            for (int i = emptyColumnsList.Count - 1; i >= 0; i--)
                emptyColumnsList[i].EntireColumn.Delete();
        }

        private string TreatTabName(string text)
        {
            string finalText = Regex.Replace(text, @"[^0-9a-zA-Z\s-]+", "");
            finalText = Tools.TreatText(finalText);

            if (finalText.Length > 30)
            {
                finalText = finalText.Substring(0, 30).Trim();
            }

            if (ExcelFunctions.ChecksIfSheetExist(finalText))
            {
                Random rnd = new Random();
                finalText = $"{finalText.Substring(0, 20).Trim()}... {rnd.Next(100000, 999999)}";
            }

            return finalText;
        }

        private ColumnsRangeOnPurchase GetColumnsRange(Worksheet sheet, List<ModelPurchase> modelPurchase)
        {
            Range GetRange(EColumnIndex column)
            {
                return sheet.Range[sheet.Cells[1, column],
                                   sheet.Cells[modelPurchase.Count + 1, column]];
            }

            var columnsRange = new ColumnsRangeOnPurchase
            {
                Empresa = GetRange(EColumnIndex.Empresa),
                Uf = GetRange(EColumnIndex.Uf),
                Operadora = GetRange(EColumnIndex.Operadora),
                CUnid = GetRange(EColumnIndex.CUnid),
                CDepto = GetRange(EColumnIndex.CDepto),
                Depto = GetRange(EColumnIndex.Depto),
                Cnpj = GetRange(EColumnIndex.Cnpj),
                Id = GetRange(EColumnIndex.Id),
                Mat = GetRange(EColumnIndex.Mat),
                MatSite = GetRange(EColumnIndex.MatSite),
                Nome = GetRange(EColumnIndex.Nome),
                Cpf = GetRange(EColumnIndex.Cpf),
                Desc = GetRange(EColumnIndex.Desc),
                Qvt = GetRange(EColumnIndex.Qvt),
                Vvt = GetRange(EColumnIndex.Vvt),
                Tvt = GetRange(EColumnIndex.Tvt),
                Desconto = GetRange(EColumnIndex.Desconto),
                Parcela1 = GetRange(EColumnIndex.Parcela1),
                Parcela2 = GetRange(EColumnIndex.Parcela2),
                Parcela3 = GetRange(EColumnIndex.Parcela3),
                CompraFinal = GetRange(EColumnIndex.CompraFinal),
                Obs = GetRange(EColumnIndex.Obs),
            };
            return columnsRange;
        }

        private void SetFormatting(ColumnsRangeOnPurchase columnsRange)
        {
            Range textColumns = gcsApp.Union(columnsRange.Empresa,
                                             columnsRange.Uf,
                                             columnsRange.Operadora,
                                             columnsRange.CUnid,
                                             columnsRange.CDepto,
                                             columnsRange.Depto,
                                             columnsRange.Cnpj,
                                             columnsRange.Id,
                                             columnsRange.Mat,
                                             columnsRange.MatSite,
                                             columnsRange.Nome,
                                             columnsRange.Cpf,
                                             columnsRange.Obs);

            textColumns.NumberFormat = "@";

            columnsRange.Qvt.NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)";

            Range decimalColumns = gcsApp.Union(columnsRange.Desc,
                                                columnsRange.Vvt,
                                                columnsRange.Tvt,
                                                columnsRange.Desconto,
                                                columnsRange.Parcela1,
                                                columnsRange.Parcela2,
                                                columnsRange.Parcela3,
                                                columnsRange.CompraFinal);

            decimalColumns.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";

            Range rngObs = columnsRange.Obs;

            rngObs.Font.Color = ColorTranslator.FromHtml("#FF0000");
            rngObs.Font.Bold = true;

            gcsApp.ActiveWindow.SplitRow = 1;
            gcsApp.ActiveWindow.FreezePanes = true;

            Range columnsTitle = gcsApp.Union(columnsRange.Empresa,
                                              columnsRange.Uf,
                                              columnsRange.Operadora,
                                              columnsRange.CUnid,
                                              columnsRange.CDepto,
                                              columnsRange.Depto);
            columnsTitle.ColumnWidth = 0.08;

            Range columnsHide = gcsApp.Union(columnsRange.Cnpj,
                                             columnsRange.Id,
                                             columnsRange.Desc,
                                             columnsRange.Qvt,
                                             columnsRange.Vvt,
                                             columnsRange.Tvt,
                                             columnsRange.Desconto);

            columnsHide.EntireColumn.Hidden = true;


            ExcelFunctions.HorizontalAlignment(columnsRange.Mat, 2);
            ExcelFunctions.HorizontalAlignment(columnsRange.MatSite, 2);
            ExcelFunctions.HorizontalAlignment(columnsRange.Cpf, 1);
        }
        private void SetFormattingInView(ColumnsRangeOnPurchase columnsRange)
        {
            Range columnsView = gcsApp.Union(columnsRange.Mat,
                                             columnsRange.MatSite,
                                             columnsRange.Nome,
                                             columnsRange.Cpf,
                                             columnsRange.Parcela1,
                                             columnsRange.Parcela2,
                                             columnsRange.Parcela3,
                                             columnsRange.CompraFinal);
            columnsView.EntireColumn.AutoFit();
        }

        private void SetBody(Worksheet sheet, List<ModelPurchase> data)
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
                            Range rng = sheet.Cells[row, EColumnIndex.CompraFinal];
                            ExcelFunctions.Styles_Colors(rng, ExcelFunctions.EStylesColors.Warning);
                            containsProblem = true;
                        }
                    }

                    sheet.Cells[row, EColumnIndex.Empresa].Value2 = item.Empresa;
                    sheet.Cells[row, EColumnIndex.Uf].Value2 = item.Uf;
                    sheet.Cells[row, EColumnIndex.Operadora].Value2 = item.Operadora;
                    sheet.Cells[row, EColumnIndex.CUnid].Value2 = item.CUnid;
                    sheet.Cells[row, EColumnIndex.CDepto].Value2 = item.CDepto;
                    sheet.Cells[row, EColumnIndex.Depto].Value2 = item.Depto;
                    sheet.Cells[row, EColumnIndex.Cnpj].Value2 = item.Cnpj;
                    sheet.Cells[row, EColumnIndex.Id].Value2 = item.Id;
                    sheet.Cells[row, EColumnIndex.Mat].Value2 = item.Mat;
                    sheet.Cells[row, EColumnIndex.MatSite].Value2 = item.MatSite;
                    sheet.Cells[row, EColumnIndex.Nome].Value2 = item.Nome;
                    sheet.Cells[row, EColumnIndex.Cpf].Value2 = item.Cpf;
                    sheet.Cells[row, EColumnIndex.Desc].Value2 = item.Desc;
                    sheet.Cells[row, EColumnIndex.Qvt].Value2 = item.Qvt;
                    sheet.Cells[row, EColumnIndex.Vvt].Value2 = item.Vvt;
                    sheet.Cells[row, EColumnIndex.Tvt].Value2 = item.Tvt;
                    sheet.Cells[row, EColumnIndex.Desconto].Value2 = item.Desconto;
                    sheet.Cells[row, EColumnIndex.Parcela1].Value2 = item.Parcela1;
                    sheet.Cells[row, EColumnIndex.Parcela2].Value2 = item.Parcela2;
                    sheet.Cells[row, EColumnIndex.Parcela3].Value2 = item.Parcela3;
                    sheet.Cells[row, EColumnIndex.CompraFinal].Value2 = item.CompraFinal;
                    sheet.Cells[row, EColumnIndex.Obs].Value2 = item.Obs;
                    row++;
                }
                else
                {
                    if (item.Empresa != null || item.Uf != null || item.Operadora != null || item.CUnid != null || item.CDepto != null || item.Depto != null)
                    {
                        Range rngRow = sheet.Range[sheet.Cells[row, 1], sheet.Cells[row, EColumnIndex.CompraFinal]];

                        if (item.Empresa != null && item.Empresa == "Total Geral")
                        {
                            sheet.Cells[row, EColumnIndex.Empresa].Value2 = item.Empresa;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.TotalGeral);
                        }
                        else if (item.Empresa != null && item.Empresa.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.Empresa].Value2 = item.Empresa;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Empresa);
                        }
                        else if (item.Uf != null && item.Uf.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.Uf].Value2 = item.Uf;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Uf);
                        }
                        else if (item.Operadora != null && item.Operadora.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.Operadora].Value2 = item.Operadora;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Operadora);
                        }
                        else if (item.CUnid != null && item.CUnid.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.CUnid].Value2 = item.CUnid;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.CUnid);
                        }
                        else if (item.CDepto != null && item.CDepto.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.CDepto].Value2 = item.CDepto;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.CDepto);
                        }
                        else if (item.Depto != null && item.Depto.EndsWith(" Total"))
                        {
                            sheet.Cells[row, EColumnIndex.Depto].Value2 = item.Depto;
                            ExcelFunctions.Styles_Emphasis(rngRow, ExcelFunctions.EStylesEmphasis.Depto);
                        }

                        sheet.Cells[row, EColumnIndex.Parcela1].Value2 = item.Parcela1;
                        sheet.Cells[row, EColumnIndex.Parcela2].Value2 = item.Parcela2;
                        sheet.Cells[row, EColumnIndex.Parcela3].Value2 = item.Parcela3;

                        sheet.Cells[row, EColumnIndex.CompraFinal].Value2 = item.CompraFinal;
                        row++; continue;
                    }
                }
            }

            if (containsProblem)
            {
                ExcelFunctions.TabColor(sheet, 5);
            }
            else
            {
                ExcelFunctions.TabColor(sheet, 3);
            }
        }

        private void SetHeader(Worksheet sheet)
        {
            sheet.Cells[1, EColumnIndex.Empresa].Value2 = ColumnsName.Empresa;
            sheet.Cells[1, EColumnIndex.Uf].Value2 = ColumnsName.Uf;
            sheet.Cells[1, EColumnIndex.Operadora].Value2 = ColumnsName.Operadora;
            sheet.Cells[1, EColumnIndex.CUnid].Value2 = ColumnsName.CUnid;
            sheet.Cells[1, EColumnIndex.CDepto].Value2 = ColumnsName.CDepto;
            sheet.Cells[1, EColumnIndex.Depto].Value2 = ColumnsName.Depto;
            sheet.Cells[1, EColumnIndex.Cnpj].Value2 = ColumnsName.Cnpj;
            sheet.Cells[1, EColumnIndex.Id].Value2 = ColumnsName.Id;
            sheet.Cells[1, EColumnIndex.Mat].Value2 = ColumnsName.Mat;
            sheet.Cells[1, EColumnIndex.MatSite].Value2 = ColumnsName.MatSite;
            sheet.Cells[1, EColumnIndex.Nome].Value2 = ColumnsName.Nome;
            sheet.Cells[1, EColumnIndex.Cpf].Value2 = ColumnsName.Cpf;
            sheet.Cells[1, EColumnIndex.Desc].Value2 = ColumnsName.Desc;
            sheet.Cells[1, EColumnIndex.Qvt].Value2 = ColumnsName.Qvt;
            sheet.Cells[1, EColumnIndex.Vvt].Value2 = ColumnsName.Vvt;
            sheet.Cells[1, EColumnIndex.Tvt].Value2 = ColumnsName.Tvt;
            sheet.Cells[1, EColumnIndex.Desconto].Value2 = ColumnsName.Desconto;
            sheet.Cells[1, EColumnIndex.Parcela1].Value2 = ColumnsName.Parcela1;
            sheet.Cells[1, EColumnIndex.Parcela2].Value2 = ColumnsName.Parcela2;
            sheet.Cells[1, EColumnIndex.Parcela3].Value2 = ColumnsName.Parcela3;
            sheet.Cells[1, EColumnIndex.CompraFinal].Value2 = ColumnsName.CompraFinal;
            sheet.Cells[1, EColumnIndex.Obs].Value2 = ColumnsName.Obs;
            Range header = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, EColumnIndex.Obs]];
            ExcelFunctions.FontBold(header, true);
            header.NumberFormat = "@";
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

                decimal CompraFinal = currentData.CompraFinal;
                decimal Parcela1 = 0;
                decimal Parcela2 = 0;
                decimal Parcela3 = 0;

                if (cl_Settings.PurchaseCreatorSplitPurchaseOption == EPurchaseCreatorSplitPurchaseOption.Two)
                {
                    if (CompraFinal / 2 < 10)
                    {
                        Parcela1 = CompraFinal;
                        Parcela2 = 0;
                        Parcela3 = 0;
                    }
                    else
                    {
                        Parcela1 = Math.Round(CompraFinal / 2, 2);
                        Parcela2 = CompraFinal - Parcela1;
                    }

                    if (Parcela1 + Parcela2 != CompraFinal)
                    {
                        MessageBox.Show($"Aconteceu um erro!", "E653982", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return (data, false);
                    }
                }
                else if (cl_Settings.PurchaseCreatorSplitPurchaseOption == EPurchaseCreatorSplitPurchaseOption.Three)
                {
                    if (CompraFinal / 3 < 10)
                    {
                        Parcela1 = CompraFinal;
                        Parcela2 = 0;
                        Parcela3 = 0;
                    }
                    else
                    {
                        Parcela1 = Math.Round(CompraFinal / 3, 2);
                        Parcela2 = Math.Round(CompraFinal / 3, 2);
                        Parcela3 = CompraFinal - (Parcela1 + Parcela2);
                    }

                    if (Parcela1 + Parcela2 + Parcela3 != CompraFinal)
                    {
                        MessageBox.Show($"Aconteceu um erro!", "E653982", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1);
                        return (data, false);
                    }
                }


                currentData.Parcela1 = Parcela1;
                currentData.Parcela2 = Parcela2;
                currentData.Parcela3 = Parcela3;

                data.Add(currentData);
                count++;
                offSetRow--;
            }

            return (data, true);
        }

        private List<ModelPurchase> SeparatePurchase(List<ModelPurchase> model, EPurchaseCreatorSubtotalOption typeSeparation)
        {
            var orderedCustomers = new List<ModelPurchase>();

            switch (typeSeparation)
            {
                case EPurchaseCreatorSubtotalOption.Empresa:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case EPurchaseCreatorSubtotalOption.Uf:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case EPurchaseCreatorSubtotalOption.Operadora:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case EPurchaseCreatorSubtotalOption.CUnid:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.CUnid)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case EPurchaseCreatorSubtotalOption.CDepto:
                    orderedCustomers = model.OrderBy(c => c.Empresa)
                                            .ThenBy(c => c.Uf)
                                            .ThenBy(c => c.Operadora)
                                            .ThenBy(c => c.CUnid)
                                            .ThenBy(c => c.CDepto)
                                            .ThenBy(c => c.Nome)
                                            .ToList();
                    break;
                case EPurchaseCreatorSubtotalOption.Depto:
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
                if (typeSeparation == EPurchaseCreatorSubtotalOption.Empresa)
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
                        if (typeSeparation == EPurchaseCreatorSubtotalOption.Uf)
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
                                if (typeSeparation == EPurchaseCreatorSubtotalOption.Operadora)
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
                                        if (typeSeparation == EPurchaseCreatorSubtotalOption.CUnid)
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
                                            lstFinal.Add(new ModelPurchase { CUnid = $"{distinctCUnid.CUnid.ToUpper()} Total", Parcela1 = subTotalCUnid.parcela1Sum, Parcela2 = subTotalCUnid.parcela2Sum, Parcela3 = subTotalCUnid.parcela3Sum, CompraFinal = subTotalCUnid.compraFinalSum });
                                        }
                                    }

                                    List<ModelPurchase> lstOperadora = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf) && (x.Operadora == distinctOperadora.Operadora))
                                                                                       .ToList();
                                    var subTotalOperadora = SubTotalGCS(lstOperadora, ETypeSubTotal.CDepto, distinctOperadora.Operadora, true);
                                    lstFinal.Add(new ModelPurchase { Operadora = $"{distinctOperadora.Operadora.ToUpper()} Total", Parcela1 = subTotalOperadora.parcela1Sum, Parcela2 = subTotalOperadora.parcela2Sum, Parcela3 = subTotalOperadora.parcela3Sum, CompraFinal = subTotalOperadora.compraFinalSum });
                                }
                            }
                            List<ModelPurchase> lstUf = orderedCustomers.Where(x => (x.Empresa == distinctEmpresa.Empresa) && (x.Uf == distinctUf.Uf))
                                                                        .ToList();
                            var subTotalUf = SubTotalGCS(lstUf, ETypeSubTotal.CDepto, distinctUf.Uf, true);
                            lstFinal.Add(new ModelPurchase { Uf = $"{distinctUf.Uf.ToUpper()} Total", Parcela1 = subTotalUf.parcela1Sum, Parcela2 = subTotalUf.parcela2Sum, Parcela3 = subTotalUf.parcela3Sum, CompraFinal = subTotalUf.compraFinalSum });
                        }
                    }

                    List<ModelPurchase> lstEmpresa = orderedCustomers.Where(x => x.Empresa == distinctEmpresa.Empresa)
                                                                     .ToList();
                    var subTotalEmpresa = SubTotalGCS(lstEmpresa, ETypeSubTotal.CDepto, distinctEmpresa.Empresa, true);
                    lstFinal.Add(new ModelPurchase { Empresa = $"{distinctEmpresa.Empresa.ToUpper()} Total", Parcela1 = subTotalEmpresa.parcela1Sum, Parcela2 = subTotalEmpresa.parcela2Sum, Parcela3 = subTotalEmpresa.parcela3Sum, CompraFinal = subTotalEmpresa.compraFinalSum });
                }
            }

            var subTotalGeral = SubTotalGeral(orderedCustomers);
            lstFinal.Add(new ModelPurchase { Empresa = $"Total Geral", Parcela1 = subTotalGeral.parcela1Sum, Parcela2 = subTotalGeral.parcela2Sum, Parcela3 = subTotalGeral.parcela3Sum, CompraFinal = subTotalGeral.compraFinalSum });
            return lstFinal;
        }

        private (decimal parcela1Sum, decimal parcela2Sum, decimal parcela3Sum, decimal compraFinalSum) SubTotalGeral(List<ModelPurchase> origin)
        {
            decimal compraFinal = origin.Sum(x => x.CompraFinal);
            decimal parcela1 = origin.Sum(x => x.Parcela1);
            decimal parcela2 = origin.Sum(x => x.Parcela2);
            decimal parcela3 = origin.Sum(x => x.Parcela3);
            return (parcela1, parcela2, parcela3, compraFinal);
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

        private (List<ModelPurchase> filteredModel, decimal parcela1Sum, decimal parcela2Sum, decimal parcela3Sum, decimal compraFinalSum) SubTotalGCS(List<ModelPurchase> data,
                                                                                                                                                       ETypeSubTotal type,
                                                                                                                                                       string name,
                                                                                                                                                       bool onlySum = false)
        {
            List<ModelPurchase> lst = new List<ModelPurchase>();
            decimal compraFinalSum = 0;
            decimal parcela1Sum = 0;
            decimal parcela2Sum = 0;
            decimal parcela3Sum = 0;

            if (data.Count < 1) { return (lst, parcela1Sum, parcela2Sum, parcela3Sum, compraFinalSum); }

            compraFinalSum = data.Sum(x => x.CompraFinal);
            parcela1Sum = data.Sum(x => x.Parcela1);
            parcela2Sum = data.Sum(x => x.Parcela2);
            parcela3Sum = data.Sum(x => x.Parcela3);

            if (onlySum)
                return (lst, parcela1Sum, parcela2Sum, parcela3Sum, compraFinalSum);

            #region ZEROED
            List<ModelPurchase> lstZeroed = data.Where(x => x.CompraFinal == 0)
                                                  .ToList();

            if (lstZeroed.Count > 0)
            {
                lst.AddRange(lstZeroed);
                lst.Add(new ModelPurchase { Nome = "[[[]]]" });
            }
            #endregion

            #region PROBLEMS
            List<ModelPurchase> lstProblems = data.Where(x => (x.Obs != null) && (!x.Obs.Contains("NOVO/SEM CARTAO") || x.Obs.Contains("2ª VIA")))
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
            List<ModelPurchase> lstNews = data.Where(x => (x.Obs != null) && (x.Obs.Contains("NOVO/SEM CARTAO") || x.Obs.Contains("2ª VIA")))
                                                       .OrderBy(c => c.Nome)
                                                       .ToList();

            if (lstNews.Count > 0)
                lst.AddRange(lstNews);
            #endregion

            #region PURCHASE
            List<ModelPurchase> lstPurchase = data.Where(x => (x.CompraFinal != 0) && (x.Obs == null))
                                                           .OrderBy(c => c.Nome)
                                                           .ToList();

            lst.AddRange(lstPurchase);
            #endregion

            switch (type)
            {
                case ETypeSubTotal.Empresa:
                    lst.Add(new ModelPurchase { Empresa = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                case ETypeSubTotal.Uf:
                    lst.Add(new ModelPurchase { Uf = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                case ETypeSubTotal.Operadora:
                    lst.Add(new ModelPurchase { Operadora = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                case ETypeSubTotal.CUnid:
                    lst.Add(new ModelPurchase { CUnid = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                case ETypeSubTotal.CDepto:
                    lst.Add(new ModelPurchase { CDepto = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                case ETypeSubTotal.Depto:
                    lst.Add(new ModelPurchase { Depto = $"{name.ToUpper()} Total", Parcela1 = parcela1Sum, Parcela2 = parcela2Sum, Parcela3 = parcela3Sum, CompraFinal = compraFinalSum });
                    break;
                default:
                    break;
            }

            return (lst, parcela1Sum, parcela2Sum, parcela3Sum, compraFinalSum);
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
