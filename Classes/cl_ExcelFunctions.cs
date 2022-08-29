using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Ribbon;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class cl_ExcelFunctions
    {
        static gcsApplication app = Globals.ThisAddIn.Application;

        public static void AdjustScroll(int linha = 1, int coluna = 1)
        {
            app.ActiveWindow.ScrollRow = linha;
            app.ActiveWindow.ScrollColumn = coluna;
        }

        public static Range GetRangeColumnByName(Worksheet ws, string nameColumn)
        {
            int usedColumns = ws.UsedRange.Columns.Count;
            return ws.Range[app.Cells[1, 1], app.Cells[1, usedColumns]].Find(What: nameColumn.Trim(), LookAt: XlLookAt.xlWhole, MatchCase: false);
        }

        public static int GetNumberColumnByName(Worksheet ws, string nameColumn)
        {
            int usedColumns = ws.UsedRange.Columns.Count;
            Range rng = ws.Range[app.Cells[1, 1], app.Cells[1, usedColumns]].Find(What: nameColumn.Trim(), LookAt: XlLookAt.xlWhole, MatchCase: false);
            if (rng == null)
            {
                return -1;
            }
            return rng.Cells.Column;
        }

        public static string GetCellText(Worksheet ws, int row, int column, int offSetRow, int offSetColumn)
        {
            return ws.Cells[row, column].Offset[offSetRow, offSetColumn].Text.Trim();
        }

        public static string GetCellInfo(Range rng)
        {
            return rng.Text;
        }

        public static Range GetRangeCell(Worksheet ws, int row, int column, int offSetRow = 0, int offSetColumn = 0)
        {
            return ws.Cells[row, column].Offset[offSetRow, offSetColumn];
        }

        public static Range GetRangeColumn(Worksheet ws, int column)
        {
            return ws.Columns[column];
        }

        public static void TabColor(Worksheet ws, int color)
        {
            switch (color)
            {
                case 1:
                    ws.Tab.Color = Color.FromArgb(153, 188, 255); // Primary
                    break;
                case 2:
                    ws.Tab.Color = Color.FromArgb(200, 204, 207); // Secondary
                    break;
                case 3:
                    ws.Tab.Color = Color.FromArgb(180, 227, 184); // Success
                    break;
                case 4:
                    ws.Tab.Color = Color.FromArgb(237, 170, 177); // Danger
                    break;
                case 5:
                    ws.Tab.Color = Color.FromArgb(252, 228, 155); // Warning
                    break;
                case 6:
                    ws.Tab.Color = Color.FromArgb(176, 222, 231); // Info
                    break;
                default:
                    break;
            }
        }

        public static void Styles_Colors(Range rng, int color)
        {
            // 0 = Default
            // 1 = Primary
            // 2 = Secondary
            // 3 = Success
            // 4 = Danger
            // 5 = Warning
            // 6 = Info

            switch (color)
            {
                case 0:
                    rng.Interior.Pattern = Constants.xlNone;
                    rng.Interior.TintAndShade = 0;
                    rng.Interior.PatternTintAndShade = 0;

                    rng.Font.Bold = false;
                    rng.Font.Italic = false;
                    rng.Font.Underline = false;
                    rng.Font.ColorIndex = Constants.xlAutomatic;
                    rng.Font.TintAndShade = 0;
                    break;
                case 1:
                    BackgroundColor("#99BCFF");
                    FontColor("#002365");
                    break;
                case 2:
                    BackgroundColor("#C8CCCF");
                    FontColor("#2F3336");
                    break;
                case 3:
                    BackgroundColor("#B4E3B8");
                    FontColor("#1B4A1F");
                    break;
                case 4:
                    BackgroundColor("#EDAAB1");
                    FontColor("#541118");
                    break;
                case 5:
                    BackgroundColor("#FCE49B");
                    FontColor("#634B02");
                    break;
                case 6:
                    BackgroundColor("#B0DEE7");
                    FontColor("#17454E");
                    break;
                default:
                    break;
            }

            void BackgroundColor(string htmlColor)
            {
                rng.Interior.PatternColorIndex = Constants.xlAutomatic;
                rng.Interior.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Interior.TintAndShade = 0;
                rng.Interior.PatternTintAndShade = 0;
            }

            void FontColor(string htmlColor)
            {
                rng.Font.Bold = false;
                rng.Font.Italic = false;
                rng.Font.Underline = false;
                rng.Font.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Font.TintAndShade = 0;
            }
        }

        public static void Styles_Bootstrap(Range rng, int color)
        {
            // 1 = Primary
            // 2 = Secondary
            // 3 = Success
            // 4 = Danger
            // 5 = Warning
            // 6 = Info
            // 7 = Light
            // 8 = Dark
            // 9 = White

            switch (color)
            {
                case 1:
                    BackgroundColor("#007BFF"); // bg-primary
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 2:
                    BackgroundColor("#6C757D"); // bg-secondary
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 3:
                    BackgroundColor("#28A745"); // bg-success
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 4:
                    BackgroundColor("#DC3545"); // bg-danger
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 5:
                    BackgroundColor("#FFC107"); // bg-warning
                    FontColor("#343A40"); // text-dark
                    break;
                case 6:
                    BackgroundColor("#17A2B8"); // bg-info
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 7:
                    BackgroundColor("#F8F9FA"); // bg-light
                    FontColor("#343A40"); // text-dark
                    break;
                case 8:
                    BackgroundColor("#343A40"); // bg-dark
                    FontColor("#FFFFFF"); // text-white
                    break;
                case 9:
                    BackgroundColor("#FFFFFF"); // bg-white
                    FontColor("#343A40"); // text-dark
                    break;
                default:
                    break;
            }

            void BackgroundColor(string htmlColor)
            {
                rng.Interior.PatternColorIndex = Constants.xlAutomatic;
                rng.Interior.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Interior.TintAndShade = 0;
                rng.Interior.PatternTintAndShade = 0;
            }

            void FontColor(string htmlColor)
            {
                rng.Font.Bold = false;
                rng.Font.Italic = false;
                rng.Font.Underline = false;
                rng.Font.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Font.TintAndShade = 0;
            }
        }

        public static void Styles_Emphasis(Range rng, int color)
        {
            switch (color)
            {
                case 1:
                    BackgroundColor("#999999");
                    FontColor("#FFFFFF");
                    break;
                case 2:
                    BackgroundColor("#727272");
                    FontColor("#FFFFFF");
                    break;
                case 3:
                    BackgroundColor("#4C4C4C");
                    FontColor("#FFFFFF");
                    break;
                case 4:
                    BackgroundColor("#262626");
                    FontColor("#FFFFFF");
                    break;
                case 5:
                    BackgroundColor("#000000");
                    FontColor("#FFFFFF");
                    break;
                default:
                    break;
            }
            void BackgroundColor(string htmlColor)
            {
                rng.Interior.PatternColorIndex = Constants.xlAutomatic;
                rng.Interior.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Interior.TintAndShade = 0;
                rng.Interior.PatternTintAndShade = 0;
            }

            void FontColor(string htmlColor)
            {
                rng.Font.Bold = true;
                rng.Font.Italic = false;
                rng.Font.Underline = false;
                rng.Font.Color = ColorTranslator.FromHtml(htmlColor);
                rng.Font.TintAndShade = 0;
            }
        }

        public static void DeleteRowsThatContainSpecificTextInColumn(Worksheet ws, string nameColumn, string criterion1 = "<>", string criterion2 = "<>")
        {
            int ColumnNumber = GetNumberColumnByName(ws, nameColumn);
            int usedRows = ws.UsedRange.Rows.Count;

            Range rng = ws.Range[ws.Cells[1, 1], ws.Cells[usedRows, ColumnNumber]];

            rng.AutoFilter(ColumnNumber, criterion1, XlAutoFilterOperator.xlAnd, criterion2, true);

            rng.Offset[1, 0].SpecialCells(XlCellType.xlCellTypeVisible).EntireRow.Delete();

            app.ActiveSheet.AutoFilterMode = false;
        }

        public static Worksheet SearchWorksheet(gcsApplication gcsApp, string sheetName)
        {
            foreach (Worksheet sheet in gcsApp.Worksheets)
            {
                if (sheet.Name.ToLower().Trim() == sheetName.ToLower())
                {
                    return sheet;
                }
            }


            return null;
        }

        public static bool CheckIfColumnsExist(Worksheet workSheet, List<string> columnsName)
        {
            foreach (var columnName in columnsName)
            {
                int usedColumns = workSheet.UsedRange.Columns.Count;
                Range rng = workSheet.Range[app.Cells[1, 1], app.Cells[1, usedColumns]].Find(What: columnName.Trim().ToLower(), LookAt: XlLookAt.xlWhole, MatchCase: false);
                if (rng == null)
                {
                    MessageBox.Show($"A coluna {columnName} não foi encontrada!", "ATENÇÃO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
            }
            return true;
        }

        public static void RemoveCF(Worksheet ws)
        {
            if (cl_Settings.More_SelectionType == 1)
            {
                RemoveConditionalFormatting(ws.Cells);
                app.Goto(ws.Range["A1"], true);
            }
            else if (cl_Settings.More_SelectionType == 2)
            {
                foreach (Worksheet sheet in app.ActiveWorkbook.Worksheets)
                {
                    RemoveConditionalFormatting(sheet.Cells);
                    app.Goto(ws.Range["A1"], true);
                }
                ws.Select();
            }
        }

        public static void ResetApp(gcsApplication xlapp)
        {
            Worksheet ws = xlapp.ActiveSheet;

            xlapp.Goto(ws.Range["A1"], true);
            xlapp.ScreenUpdating = true;
            //ws.Cells[1, 1].Select();
            //cl_ExcelFunctions.AdjustScroll();
        }

        public static void ApplyRemove(Worksheet ws)
        {

            Range rng = ws.Cells;
            var lstEmptySheets = new List<Worksheet>();
            var lstHiddenSheets = new List<Worksheet>();

            if (cl_Settings.ApplyRemove_Apply_AllSheets == false)
            {
                if (cl_Settings.ApplyRemove_Remove_Images == true) { RemoveImages(ws); }
                if (cl_Settings.ApplyRemove_Remove_Filter == true) { RemoveFilter(ws); }
                if (cl_Settings.ApplyRemove_Remove_Formula == true) { RemoveFormula(rng); }

                if (cl_Settings.ApplyRemove_Apply_Zoom == true) { PageZoom(ws, cl_Settings.ApplyRemove_Apply_ZoomValue); }

                if (cl_Settings.ApplyRemove_Remove_ConditionalFormatting == true) { RemoveConditionalFormatting(rng); }

                if (cl_Settings.ApplyRemove_Apply_FontName == true) { FontName(rng, cl_Settings.ApplyRemove_Apply_FontNameText); }
                if (cl_Settings.ApplyRemove_Apply_FontSize == true) { FontSize(rng, int.Parse(cl_Settings.ApplyRemove_Apply_FontSizeText)); }

                if (cl_Settings.ApplyRemove_Remove_FontBold == true) { FontBold(rng, false); }
                if (cl_Settings.ApplyRemove_Remove_FontItalic == true) { FontItalic(rng, false); }
                if (cl_Settings.ApplyRemove_Remove_FontUnderline == true) { FontUnderline(rng, false); }
                if (cl_Settings.ApplyRemove_Remove_Borders == true) { RemoveBorders(rng); }
                if (cl_Settings.ApplyRemove_Remove_Fill == true) { RemoveFill(rng); }
                if (cl_Settings.ApplyRemove_Remove_FontColor == true) { RemoveFontColor(rng); }
                if (cl_Settings.ApplyRemove_Remove_WrapText == true) { WrapText(rng, false); }
                if (cl_Settings.ApplyRemove_Remove_MergeCells == true) { MergeCells(rng, false); }

                if (cl_Settings.ApplyRemove_Apply_Align_Vertical == true) { VerticalAlignment(rng, cl_Settings.ApplyRemove_Apply_Align_VerticalValue); }
                if (cl_Settings.ApplyRemove_Apply_Align_Horizontal == true) { HorizontalAlignment(rng, cl_Settings.ApplyRemove_Apply_Align_HorizontalValue); }

                if (cl_Settings.ApplyRemove_Apply_RowHeight == true) { RowHeight(rng, cl_Settings.ApplyRemove_Apply_RowHeightValue); }
                if (cl_Settings.ApplyRemove_Apply_ColumnWidth == true) { ColumnWidth(rng, cl_Settings.ApplyRemove_Apply_ColumnWidthValue); }

                app.Goto(ws.Range["A1"], true);
            }
            else
            {
                foreach (Worksheet sheet in app.ActiveWorkbook.Worksheets)
                {
                    if (sheet.Visible == XlSheetVisibility.xlSheetHidden) { lstHiddenSheets.Add(sheet); continue; } // VERIFY IF SHEET IS HIDE
                    if (sheet.UsedRange.Count < 2) { lstEmptySheets.Add(sheet); continue; } // VERIFY IF SHEET IS EMPTY

                    sheet.Select();
                    rng = sheet.Cells;

                    if (cl_Settings.ApplyRemove_Remove_Images) { RemoveImages(sheet); }
                    if (cl_Settings.ApplyRemove_Remove_Filter) { RemoveFilter(sheet); }

                    if (cl_Settings.ApplyRemove_Remove_Formula) { RemoveFormula(rng); }

                    if (cl_Settings.ApplyRemove_Apply_Zoom) { PageZoom(sheet, cl_Settings.ApplyRemove_Apply_ZoomValue); }

                    if (cl_Settings.ApplyRemove_Remove_ConditionalFormatting) { RemoveConditionalFormatting(rng); }

                    if (cl_Settings.ApplyRemove_Apply_FontName) { FontName(rng, cl_Settings.ApplyRemove_Apply_FontNameText); }
                    if (cl_Settings.ApplyRemove_Apply_FontSize) { FontSize(rng, int.Parse(cl_Settings.ApplyRemove_Apply_FontSizeText)); }

                    if (cl_Settings.ApplyRemove_Remove_FontBold) { FontBold(rng, false); }
                    if (cl_Settings.ApplyRemove_Remove_FontItalic) { FontItalic(rng, false); }
                    if (cl_Settings.ApplyRemove_Remove_FontUnderline) { FontUnderline(rng, false); }
                    if (cl_Settings.ApplyRemove_Remove_Borders) { RemoveBorders(rng); }
                    if (cl_Settings.ApplyRemove_Remove_Fill) { RemoveFill(rng); }
                    if (cl_Settings.ApplyRemove_Remove_FontColor) { RemoveFontColor(rng); }
                    if (cl_Settings.ApplyRemove_Remove_WrapText) { WrapText(rng, false); }
                    if (cl_Settings.ApplyRemove_Remove_MergeCells) { MergeCells(rng, false); }

                    if (cl_Settings.ApplyRemove_Apply_Align_Vertical) { VerticalAlignment(rng, cl_Settings.ApplyRemove_Apply_Align_VerticalValue); }
                    if (cl_Settings.ApplyRemove_Apply_Align_Horizontal) { HorizontalAlignment(rng, cl_Settings.ApplyRemove_Apply_Align_HorizontalValue); }

                    if (cl_Settings.ApplyRemove_Apply_RowHeight) { RowHeight(rng, cl_Settings.ApplyRemove_Apply_RowHeightValue); }
                    if (cl_Settings.ApplyRemove_Apply_ColumnWidth) { ColumnWidth(rng, cl_Settings.ApplyRemove_Apply_ColumnWidthValue); }


                    app.Goto(sheet.Range["A1"], true);
                }
            }

            app.ActiveWorkbook.Sheets[1].Select();

            if (cl_Settings.ApplyRemove_RemoveAllSheets_HiddenSheets)
            {
                foreach (var item in lstHiddenSheets)
                {
                    item.Delete();
                }
            }

            if (cl_Settings.ApplyRemove_RemoveAllSheets_EmptySheets)
            {
                foreach (var item in lstEmptySheets)
                {
                    item.Delete();
                }
            }
        }

        public static void FontName(Range rng, string name = "Consolas")
        {
            rng.Font.Name = name;
        }

        public static void FontSize(Range rng, int size = 10)
        {
            rng.Font.Size = size;
        }

        public static void FontBold(Range rng, bool act = true)
        {
            rng.Font.Bold = act;
        }

        public static void FontItalic(Range rng, bool act = false)
        {
            rng.Font.Italic = act;
        }

        public static void FontUnderline(Range rng, bool act = false)
        {
            rng.Font.Underline = act;
        }

        public static void RemoveBorders(Range rng)
        {
            rng.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = Constants.xlNone;
            rng.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = Constants.xlNone;

            rng.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Constants.xlNone;
            rng.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = Constants.xlNone;
            rng.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Constants.xlNone;
            rng.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Constants.xlNone;

            rng.Borders[XlBordersIndex.xlInsideVertical].LineStyle = Constants.xlNone;
            rng.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = Constants.xlNone;
        }

        public static void RemoveFill(Range rng)
        {
            rng.Interior.Pattern = Constants.xlNone;
            rng.Interior.TintAndShade = 0;
            rng.Interior.PatternTintAndShade = 0;
        }

        public static void RemoveFontColor(Range rng)
        {
            rng.Font.ColorIndex = Constants.xlAutomatic;
            rng.Font.TintAndShade = 0;
        }

        public static void WrapText(Range rng, bool act = true)
        {
            rng.WrapText = act;
        }

        public static void MergeCells(Range rng, bool act = true)
        {
            rng.MergeCells = act;
        }

        public static void RemoveFormula(Range rng)
        {
            rng.Copy();
            rng.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        public static void RemoveFilter(Worksheet ws)
        {
            if (ws.AutoFilter != null)
                ws.AutoFilterMode = false;
        }

        public static void RemoveImages(Worksheet ws)
        {
            foreach (Shape sh in ws.Shapes)
                sh.Delete();
        }

        public static void RemoveConditionalFormatting(Range rng)
        {
            rng.FormatConditions.Delete();
        }

        public static void VerticalAlignment(Range rng, int option)
        {
            switch (option)
            {
                case 0:
                    rng.VerticalAlignment = Constants.xlTop;
                    break;
                case 1:
                    rng.VerticalAlignment = Constants.xlCenter;
                    break;
                case 2:
                    rng.VerticalAlignment = Constants.xlBottom;
                    break;
            }
        }

        public static void HorizontalAlignment(Range rng, int option)
        {
            switch (option)
            {
                case 0:
                    rng.HorizontalAlignment = Constants.xlLeft;
                    break;
                case 1:
                    rng.HorizontalAlignment = Constants.xlCenter;
                    break;
                case 2:
                    rng.HorizontalAlignment = Constants.xlRight;
                    break;
            }
        }

        public static void RowHeight(Range rng, decimal height)
        {
            // 0 = Auto
            if (height == 0)
            {
                rng.EntireRow.AutoFit();
            }
            else
            {
                rng.RowHeight = height;
            }
        }

        public static void ColumnWidth(Range rng, decimal width)
        {
            // 0 = Auto
            if (width == 0)
            {
                rng.EntireColumn.AutoFit();
            }
            else
            {
                rng.ColumnWidth = width;
            }
        }

        public static void PageZoom(Worksheet ws, decimal zoom)
        {
            ws.Application.ActiveWindow.Zoom = zoom;
            //app.ActiveWindow.Zoom = zoom;
        }

        public static void MoveColumns(Worksheet workSheet, List<string> nameColumnsInOrder)
        {
            int position = 1;

            foreach (string columnName in nameColumnsInOrder)
            {
                Move(columnName);
            }

            void Move(string name)
            {
                Range Column_Range = cl_ExcelFunctions.GetRangeColumnByName(workSheet, name);

                if (Column_Range != null)
                {
                    if (Column_Range.Cells.Column != position)
                    {
                        workSheet.Columns[Column_Range.Cells.Column].Cut();
                        workSheet.Columns[position].Insert(XlInsertShiftDirection.xlShiftToRight);
                    }
                    position++;
                }
            }
        }

        public static void SortDataByColumn(Worksheet workSheet, List<string> nameColumnsInOrder)
        {
            List<Range> list_ColumnsRange = new List<Range>();

            foreach (string column in nameColumnsInOrder)
            {
                list_ColumnsRange.Add(cl_ExcelFunctions.GetRangeColumnByName(workSheet, column));
            }

            workSheet.Sort.SortFields.Clear();

            foreach (Range columnRange in list_ColumnsRange)
            {
                if (columnRange != null)
                {
                    workSheet.Sort.SortFields.Add(Key: columnRange, SortOn: XlSortOn.xlSortOnValues, Order: XlSortOrder.xlAscending, DataOption: XlSortDataOption.xlSortNormal);
                }
            }

            workSheet.Sort.SetRange(workSheet.Cells);
            workSheet.Sort.Header = XlYesNoGuess.xlYes;
            workSheet.Sort.MatchCase = false;
            workSheet.Sort.Orientation = (XlSortOrientation)Constants.xlTopToBottom;
            workSheet.Sort.SortMethod = XlSortMethod.xlPinYin;
            workSheet.Sort.Apply();
        }

        public static void RemoveColumns(Worksheet workSheet, List<string> nameColumns)
        {
            foreach (string nameColumn in nameColumns)
            {
                while (true)
                {
                    Range rng = cl_ExcelFunctions.GetRangeColumnByName(workSheet, nameColumn);

                    if (rng != null)
                    {
                        rng.EntireColumn.Delete();
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }

        public static void MoveSheetOrder(string name, int position)
        {
            name = name.ToLower().Trim();
            foreach (Worksheet workSheet in app.Worksheets)
            {
                if (workSheet.Name.ToLower().Trim() == name)
                {
                    workSheet.Move(app.ActiveWorkbook.Sheets[position]);
                }
            }
        }

        public static void DeleteSheetContainsName(string name)
        {
            string sheetToBeDeleted = "";
            name = name.ToLower().Trim();

            while (true)
            {
                bool sheetExist = false;

                foreach (Worksheet workSheet in app.Worksheets)
                {
                    if (workSheet.Name.ToLower().Trim().Contains(name))
                    {
                        sheetToBeDeleted = workSheet.Name;
                        sheetExist = true;
                        app.Worksheets[sheetToBeDeleted].Delete();
                        continue;
                    }
                }

                if (!sheetExist)
                {
                    break;
                }
            }
        }

        public static void DeleteSheetEqualName(string name)
        {
            name = name.ToLower().Trim();

            foreach (Worksheet workSheet in app.Worksheets)
            {
                if (workSheet.Name.ToLower().Trim() == name)
                {
                    name = workSheet.Name;
                    app.Worksheets[name].Delete();
                }
            }

        }

        public static void SetBZPA(Worksheet workSheet, Range rng)
        {
            try
            {
                app.ActiveWindow.View = XlWindowView.xlPageBreakPreview;
                app.ActiveWindow.Zoom = 100;

                workSheet.PageSetup.PrintArea = rng.Address;

                // REMOVENDO BORDAS
                rng.Borders[XlBordersIndex.xlDiagonalUp].LineStyle = Constants.xlNone;
                rng.Borders[XlBordersIndex.xlDiagonalDown].LineStyle = Constants.xlNone;

                AplicarBorda(XlBordersIndex.xlEdgeLeft); // ESQUERDA
                AplicarBorda(XlBordersIndex.xlEdgeTop); // SUPERIOR
                AplicarBorda(XlBordersIndex.xlEdgeBottom); // INFERIOR
                AplicarBorda(XlBordersIndex.xlEdgeRight); // DIREITA
                AplicarBorda(XlBordersIndex.xlInsideVertical); // VERTICAL INTERNA
                AplicarBorda(XlBordersIndex.xlInsideHorizontal); // HORIZONTAL INTERNA

                void AplicarBorda(XlBordersIndex borda)
                {
                    rng.Borders[borda].LineStyle = XlLineStyle.xlDot;
                    rng.Borders[borda].ColorIndex = Constants.xlAutomatic;
                    rng.Borders[borda].TintAndShade = 0;
                    rng.Borders[borda].Weight = XlBorderWeight.xlThin;
                }
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 552929", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void SheetsOrderBy(bool descending = false)
        {
            if (app.Worksheets.Count <= 1) { return; }

            List<string> lst_Sheets = new List<string>();

            foreach (Worksheet sheet in app.Worksheets) { sheet.Name = cl_Settings.RemoverAcentos(sheet.Name).Trim(); }

            foreach (Worksheet sheet in app.Worksheets) { lst_Sheets.Add(sheet.Name); }

            lst_Sheets.Sort();

            if (descending) { lst_Sheets.Reverse(); }

            int position = 1;

            foreach (string sheet in lst_Sheets) { MoveSheetOrder(sheet, position); position++; }

            app.Worksheets[1].Select();
        }

        public static void RenameSheet(string oldName, string newName)
        {
            foreach (Worksheet sheet in app.Worksheets)
            {
                if (sheet.Name.ToLower().Trim() == oldName.ToLower().Trim())
                {
                    sheet.Name = newName;
                }
            }
        }

        public static int RemoveHiddenSheets()
        {
            int count = 0;
            foreach (Worksheet sheet in app.Worksheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetHidden)
                {
                    sheet.Delete();
                    count++;
                }
            }
            return count;
        }

        public static int ShowHiddenSheets()
        {
            int count = 0;
            foreach (Worksheet sheet in app.Worksheets)
            {
                if (sheet.Visible == XlSheetVisibility.xlSheetHidden)
                {
                    sheet.Visible = XlSheetVisibility.xlSheetVisible;
                    count++;
                }
            }
            return count;
        }

        public static void SetColumnWidthByName(Worksheet workSheet, string nameColumn, decimal width = 0)
        {
            Range rng = GetRangeColumnByName(workSheet, nameColumn);

            if (rng != null)
            {
                if (width == 0)
                {
                    rng.EntireColumn.AutoFit();
                }
                else
                {
                    rng.ColumnWidth = width;
                }
            }
        }

        public static void RoundDecimal(Range rng, int decimals = 2)
        {


        }

        public static void CreateBackup()
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Workbook wb = app.ActiveWorkbook;

            string FilePath = Path.GetDirectoryName(wb.FullName);
            string FileExt = Path.GetExtension(wb.FullName);

            if (!Directory.Exists(Path.Combine(FilePath, "_BACKUPS")))
            {
                Directory.CreateDirectory(Path.Combine(FilePath, "_BACKUPS"));
            }

            wb.SaveCopyAs(Path.Combine(FilePath, "_BACKUPS", cl_Tools.GetDateTime()) + FileExt);
        }

        public static void FileToSend()
        {
            gcsApplication app = Globals.ThisAddIn.Application;
            Workbook wb = app.ActiveWorkbook;

            string FilePath = Path.GetDirectoryName(wb.FullName);
            string FileName = "_" + Path.GetFileNameWithoutExtension(wb.FullName).Trim() + " - Enviar";
            string FileExt = Path.GetExtension(wb.FullName);
            string FileFullPath = Path.Combine(FilePath, FileName + FileExt);

            wb.SaveAs(FileFullPath);
        }

    }
}