using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public class cl_AdjustBalanceDaysValueColumns
    {
        private gcsApplication gcsApp { get; set; } = Globals.ThisAddIn.Application;

        public cl_AdjustBalanceDaysValueColumns()
        {

        }

        public void Start()
        {
            try
            {
                gcsApp.ScreenUpdating = false;

                var ws = cl_ExcelFunctions.SearchWorksheet(gcsApp, "Dados");

                if (ws == null)
                {
                    MessageBox.Show($"A aba Dados não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cl_ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                ws.Select();

                if (!cl_ExcelFunctions.CheckIfColumnsExist(ws, new List<string> { ColumnsName.ValorDias, ColumnsName.Saldo }))
                {
                    cl_ExcelFunctions.ResetApp(gcsApp);
                    return;
                }

                var saldoColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Saldo);
                var valorDiasColumnNumber = cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.ValorDias);

                var teste = ws.UsedRange.Rows.Count;
                Range aaa = ws.Cells[teste, cl_ExcelFunctions.GetNumberColumnByName(ws, ColumnsName.Saldo)].Offset[0, 0];

                if (aaa.Value2 == 0)
                    aaa.Value2 = "Gustavo";



















                
            }
            catch (Exception erro)
            {
                gcsApp.ScreenUpdating = true;
                MessageBox.Show(erro.ToString(), "ERROR: 688425", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            finally
            {
                gcsApp.ScreenUpdating = true;
                gcsApp.DisplayAlerts = true;
            }
        }




    }
}
