using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using gcsApplication = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class Tools
    {
        static gcsApplication gcsApp = Globals.ThisAddIn.Application;

        public static void ObterTipoSelecao(Range selecao)
        {
            try
            {
                MessageBox.Show(selecao.Value.GetType().ToString());
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 553757", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void CopiarColarValor(Range rng)
        {
            rng.Copy();
            rng.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        public static void SelecionarRange()
        {
            Worksheet ws = gcsApp.ActiveSheet;
            ws.UsedRange.Select();
        }

        public static void SelecionarTudo()
        {
            Worksheet ws = gcsApp.ActiveSheet;

            Range teste = ws.Cells;

            teste.Select();
        }

        public static void LinhasUsadas()
        {
            Worksheet ws = gcsApp.ActiveSheet;
            MessageBox.Show("Linhas usadas: " + ws.UsedRange.Rows.Count.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ColunasUsadas()
        {
            Worksheet ws = gcsApp.ActiveSheet;
            MessageBox.Show("Colunas usadas: " + ws.UsedRange.Columns.Count.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void DefinirDados(string empresa, string mes, string ano, string compra)
        {
            Worksheet ws = gcsApp.ActiveSheet;
            ws.PageSetup.CenterHeader = "&\"Arial\"&B&14" + empresa + "\n&A - " + mes + "/" + ano + " (" + compra + ")";
        }

        public static bool ContainsWithRegex(string text, string pattern)
        {
            Regex regex = new Regex(pattern);
            return regex.IsMatch(text);
        }

        public static int ConverterCPF(Worksheet ws, Range rng)
        {
            int count = 0;
            foreach (Range item in rng.Cells)
            {
                if (item.Value == null)
                {
                    continue;
                }
                else
                {
                    string text = item.Value.ToString();
                    bool addZero = cl_Settings.CPF_ZeroAEsquerda;

                    if (cl_Settings.CPF_Opcao == 0) // 00000000000
                    {
                        text = TreatCpf(text, true, addZero);

                    }
                    else // 000.000.000-00
                    {
                        text = TreatCpf(text, false, addZero);
                    }

                    Range selection = ws.Cells[item.Row, item.Column];
                    selection.NumberFormat = "@";
                    selection.Value = text;
                    count++;
                }
            }
            return count;
        }

        public static string TreatCpf(string cpf, bool onlyNumbers = true, bool addZero = true)
        {
            string newCpf = cpf.Trim();
            newCpf = Regex.Replace(newCpf, @"[^\d]", "");

            if (addZero)
                newCpf = newCpf.Trim().PadLeft(11, '0');

            if (newCpf.Length == 11)
            {
                if (onlyNumbers)
                {
                    return newCpf;
                }
                else
                {
                    return Regex.Replace(newCpf, "([0-9]{3})([0-9]{3})([0-9]{3})([0-9]{2})", "$1.$2.$3-$4");
                }
            }
            else
            {
                return newCpf;
            }
        }

        public static string TreatWorkSchedule(string text)
        {
            text = text.ToUpper().Trim();
            string newText = Regex.Replace(text, @"\s", "");

            if (newText.Contains("6X1")) { return "6X1"; }
            else if (newText.Contains("06X01")) { return "6X1"; }
            else if (newText.Contains("6X2")) { return "6X1"; }
            else if (newText.Contains("60X01")) { return "6X1"; }
            else if (newText.Contains("45H")) { return "6X1"; }
            else if (newText.Contains("05X02")) { return "5X2"; }
            else if (newText.Contains("5X2")) { return "5X2"; }
            else if (newText.Contains("SX2")) { return "5X2"; }
            else if (newText.Contains("5X1")) { return "5X2"; }
            else if (newText.Contains("44H")) { return "5X2"; }
            else if (newText.Contains("12X36")) { return "12X36"; }
            else if (newText.Contains("13X36")) { return "12X36"; }
            else if (newText.Contains("24X48")) { return "24X48"; }
            else if (newText.Contains("04X03")) { return "4X3"; }
            else if (newText.Contains("4X3")) { return "4X3"; }
            else if (newText.Contains("03X04")) { return "3X4"; }
            else if (newText.Contains("3X4")) { return "3X4"; }
            else if (newText.Contains("02X05")) { return "2X5"; }
            else if (newText.Contains("2X5")) { return "2X5"; }
            else if (newText.Contains("01X06")) { return "1X6"; }
            else if (newText.Contains("1X6")) { return "1X6"; }
            else { return text; }
        }

        public static void ConverterDecimalPlaces(Worksheet ws, Range rng, int places)
        {
            int count = 0;
            foreach (Range item in rng.Cells)
            {
                if (item.Value == null)
                {
                    continue;
                }
                else
                {
                    string text = item.Value.ToString().Trim();

                    if (decimal.TryParse(text, out decimal out_text))
                    {
                        text = Math.Round(out_text, places).ToString();
                        count++;
                        Range selection = ws.Cells[item.Row, item.Column];

                        switch (places)
                        {
                            case 0:
                                selection.NumberFormat = @"_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)";
                                break;
                            case 1:
                                selection.NumberFormat = @"_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""?_-;_-@_-";
                                break;
                            case 2:
                                selection.NumberFormat = @"_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)";
                                break;
                            default:
                                break;
                        }

                        selection.Value = decimal.Parse(text);
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            MessageBox.Show("Números convertidos: " + count.ToString());
        }

        public static void Generator_CPF(Worksheet ws, Range rng)
        {
            int count = 0;
            foreach (Range item in rng.Cells)
            {
                string text = "";
                Thread.Sleep(1);
                text = GerarCpf();
                Range selection = ws.Cells[item.Row, item.Column];
                selection.NumberFormat = "@";
                selection.Value = text.ToString();
                count++;
            }
            MessageBox.Show("CPFs gerados: " + count.ToString());
        }

        public static string GerarCpf()
        {
            var random = new Random();

            int soma = 0;
            int resto = 0;
            int[] multiplicadores = new int[10] { 11, 10, 9, 8, 7, 6, 5, 4, 3, 2 };
            string semente;

            do
            {
                semente = random.Next(1, 999999999).ToString().PadLeft(9, '0');
            }
            while (
                semente == "000000000"
                || semente == "111111111"
                || semente == "222222222"
                || semente == "333333333"
                || semente == "444444444"
                || semente == "555555555"
                || semente == "666666666"
                || semente == "777777777"
                || semente == "888888888"
                || semente == "999999999"
            );

            for (int i = 1; i < multiplicadores.Count(); i++)
                soma += int.Parse(semente[i - 1].ToString()) * multiplicadores[i];

            resto = soma % 11;

            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;

            semente += resto;
            soma = 0;

            for (int i = 0; i < multiplicadores.Count(); i++)
                soma += int.Parse(semente[i].ToString()) * multiplicadores[i];

            resto = soma % 11;

            if (resto < 2)
                resto = 0;
            else
                resto = 11 - resto;

            semente = semente + resto;

            return semente;

        }

        public static string GetDateTime(bool getDate = true, bool getTime = true)
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH-mm-ss");

            if (getDate && getTime)
                return $"{date}_{time}";
            else if (getDate && getTime == false)
                return $"{date}";
            else
                return $"{time}";
        }

        public static string TreatText(string text, bool trim = true, bool toUpper = true, bool removeAccents = true, bool removeDuplicateSpaces = true)
        {
            if (trim)
                text = text.Trim();
            if (toUpper)
                text = text.ToUpper();
            if (removeAccents)
                text = RemoveAccents(text);
            if (removeDuplicateSpaces)
                text = RemoveDuplicateSpaces(text);
            return text;
        }

        public static string RemoveAccents(string texto)
        {
            var stringBuilder = new StringBuilder();
            StringBuilder sbReturn = stringBuilder;
            var arrayText = texto.Normalize(NormalizationForm.FormD).ToCharArray();
            foreach (char letter in arrayText)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(letter) != UnicodeCategory.NonSpacingMark)
                    sbReturn.Append(letter);
            }
            return sbReturn.ToString();
        }

        public static string RemoveDuplicateSpaces(string texto)
        {
            texto = Regex.Replace(texto, @"\s{2,}", " ");
            texto = texto.Trim();

            return texto;
        }
    }
}

