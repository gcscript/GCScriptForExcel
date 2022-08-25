using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using xlApp = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class cl_GenerateFileToSend

    {
        static xlApp app = Globals.ThisAddIn.Application;
        static Worksheet ws = app.ActiveSheet;

        public static void Start()
        {
            cl_ExcelFunctions.CreateBackup();

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                //app.Worksheets["Dados"].Select();

                if (!ChecksIfDadosSheetExist()) { return; }

                foreach (Worksheet workSheet in app.Worksheets)
                {
                    Range rng = workSheet.Cells;
                    cl_ExcelFunctions.RemoveFormula(rng);

                    if (workSheet.Name.ToLower().Trim() == "dados")
                    {
                        workSheet.Select();
                        if (!ChecksIfColumnsExist(workSheet)) { return; }

                        List<string> lst_MoveColumnsName = new List<string>() { "CNPJ", "UF" , "Operadora", "Empresa", "C.Unid", "C.Depto", "Depto" };
                        cl_ExcelFunctions.MoveColumns(workSheet, lst_MoveColumnsName);

                        List<string> lst_SortDataColumns = new List<string>() { "UF", "Operadora", "Empresa", "C.Unid", "C.Depto", "Depto", "Nome" };
                        cl_ExcelFunctions.SortDataByColumn(workSheet, lst_SortDataColumns);

                        List<string> lst_RemoveColumns = new List<string>() { "ORG1", "VvtNovo", "TvtNovo", "RecPend", "Saldo1", "CNPJ + CPF + Operadora", "Buscador", "ORDEM", "CF -R$10", "Tipo1" };
                        cl_ExcelFunctions.RemoveColumns(workSheet, lst_RemoveColumns);

                        Range PrintArea = workSheet.Range[workSheet.Cells[1, 1], workSheet.Cells[1048576, cl_ExcelFunctions.GetNumberColumnByName(workSheet, "CompraFinal")].End(XlDirection.xlUp).Offset[0, 0]];
                        cl_ExcelFunctions.SetBZPA(workSheet, PrintArea);
                    }
                    workSheet.Application.CutCopyMode = (XlCutCopyMode)0;

                    workSheet.Application.Goto(workSheet.Range["A1"], true);
                }

                cl_ExcelFunctions.MoveSheetOrder("Compra", 1);
                cl_ExcelFunctions.MoveSheetOrder("Rateio", 1);
                cl_ExcelFunctions.MoveSheetOrder("Dados", 1);

                cl_ExcelFunctions.DeleteSheetContainsName("Escala");

                app.Worksheets["Dados"].Select();

                cl_ExcelFunctions.FileToSend();

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

            string[] columnsName = { "UF", "Operadora", "Empresa", "C.Unid", "Mat", "Mat Site", "Nome", "Desc", "Qvt1", "Vvt1", "Tvt1", "Total", "Saldo", "ValorDias", "Desconto", "CompraFinal" };

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
