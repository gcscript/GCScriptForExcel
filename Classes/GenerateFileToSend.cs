using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class GenerateFileToSend

    {
        static xlApp app = Globals.ThisAddIn.Application;
        static Worksheet ws = app.ActiveSheet;

        public static void Start()
        {
            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                //app.Worksheets["Dados"].Select();

                if (!ChecksIfDadosSheetExist()) { return; }

                foreach (Worksheet workSheet in app.Worksheets)
                {
                    Range rng = workSheet.Cells;
                    ExcelFunctions.RemoveFormula(rng);

                    if (workSheet.Name.ToLower().Trim() == "dados")
                    {
                        workSheet.Select();
                        if (!ChecksIfColumnsExist(workSheet)) { return; }

                        List<string> lst_MoveColumnsName = new List<string>() { ColumnsName.Cnpj, ColumnsName.UF , ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.CDepto, ColumnsName.Depto };
                        ExcelFunctions.MoveColumns(workSheet, lst_MoveColumnsName);

                        List<string> lst_SortDataColumns = new List<string>() { ColumnsName.UF, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, ColumnsName.CDepto, ColumnsName.Depto, ColumnsName.Nome };
                        ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);

                        List<string> lst_RemoveColumns = new List<string>() { ColumnsName.Org, ColumnsName.VvtNovo, ColumnsName.TvtNovo, ColumnsName.RecPendSet, ColumnsName.SaldoSet, ColumnsName.ValorDiasSet, ColumnsName.CnpjCpfOperadora, ColumnsName.Buscador, ColumnsName.Ordem, ColumnsName.Cf10, ColumnsName.Tipo };
                        ExcelFunctions.RemoveColumns(workSheet, lst_RemoveColumns);

                        Range PrintArea = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1048576, ExcelFunctions.GetNumberColumnByName(workSheet, ColumnsName.CompraFinal)].End(XlDirection.xlUp).Offset[0, 0]];
                        ExcelFunctions.SetBZPA(workSheet, PrintArea);
                    }
                    workSheet.Application.CutCopyMode = (XlCutCopyMode)0;

                    workSheet.Application.Goto(workSheet.Range["A1"], true);
                }

                ExcelFunctions.MoveSheetOrder("Compra", 1);
                ExcelFunctions.MoveSheetOrder("Rateio", 1);
                ExcelFunctions.MoveSheetOrder("Dados", 1);

                ExcelFunctions.DeleteSheetContainsName(ColumnsName.Escala);

                app.Worksheets["Dados"].Select();

                ExcelFunctions.FileToSend();

                MessageBox.Show("Arquivo para Envio criado com sucesso!", "SUCESSO!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.ToString(), "ERRO: 843328", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
            }

        }

        static bool ChecksIfDadosSheetExist()
        {
            bool dadosExist = false;
            foreach (Worksheet sheet in app.Worksheets)
            {
                if (sheet.Name.ToLower().Trim() == "dados")
                {
                    dadosExist = true;
                }
            }

            if (!dadosExist)
            {
                MessageBox.Show("A aba [Dados] não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        static bool ChecksIfColumnsExist(Worksheet workSheet)
        {
            int usedColumns = workSheet.UsedRange.Columns.Count;

            string[] columnsName = { ColumnsName.UF, ColumnsName.Operadora, ColumnsName.Empresa, ColumnsName.CUnid, "Mat", "Mat Site", ColumnsName.Nome, ColumnsName.Desc, ColumnsName.Qvt, ColumnsName.Vvt, ColumnsName.Tvt, ColumnsName.Total, "Saldo", "ValorDias", ColumnsName.Desconto, ColumnsName.CompraFinal };

            foreach (string columnName in columnsName)
            {
                if (CheckColumnExistence(columnName) == false)
                {
                    MessageBox.Show("A coluna [" + columnName.Trim().ToUpper() + "] não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }

            bool CheckColumnExistence(string columnName)
            {
                columnName = columnName.Trim().ToLower();
                Range rng = workSheet.Range[app.Cells[1, 1], app.Cells[1, usedColumns]].Find(What: columnName, LookAt: XlLookAt.xlWhole, MatchCase: false);
                if (rng == null) { return false; }
                return true;
            }

            return true;
        }
    }
}
