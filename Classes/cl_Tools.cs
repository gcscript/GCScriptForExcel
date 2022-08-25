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
using Appl = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class cl_Tools
    {
        static Appl app = Globals.ThisAddIn.Application;

        public static void T1(Worksheet ws, Range rng)
        {

        }

        public static void T2(Worksheet ws, Range rng)
        {

        }

        public static void T3()
        {

        }

        public static void T4()
        {

        }

        public static void T5()
        {

        }

        public static void T6()
        {
            try
            {
                MessageBox.Show(app.ActiveCell.Value.GetType().ToString());
            }
            catch (Exception erro)
            {
                MessageBox.Show(erro.Message, "ERRO: 553757", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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
            Worksheet ws = app.ActiveSheet;
            ws.UsedRange.Select();
        }

        public static void SelecionarTudo()
        {
            Worksheet ws = app.ActiveSheet;

            Range teste = ws.Cells;

            teste.Select();
        }

        public static void LinhasUsadas()
        {
            Worksheet ws = app.ActiveSheet;
            MessageBox.Show("Linhas usadas: " + ws.UsedRange.Rows.Count.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void ColunasUsadas()
        {
            Worksheet ws = app.ActiveSheet;
            MessageBox.Show("Colunas usadas: " + ws.UsedRange.Columns.Count.ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void DefinirDados(string empresa, string mes, string ano, string compra)
        {
            Worksheet ws = app.ActiveSheet;
            ws.PageSetup.CenterHeader = "&\"Arial\"&B&14" + empresa + "\n&A - " + mes + "/" + ano + " (" + compra + ")";
        }

        public static bool ContainsWithRegex(string text, string pattern)
        {
            Regex regex = new Regex(pattern);
            return regex.IsMatch(text);
        }

        public static void ConverterCPF(Worksheet ws, Range rng)
        {
            int contador = 0;
            foreach (Range item in rng.Cells)
            {
                if (item.Value == null)
                {
                    continue;
                }
                else
                {
                    string texto = item.Value.ToString();
                    bool addZero = cl_Settings.CPF_ZeroAEsquerda;

                    if (cl_Settings.CPF_Opcao == 0) // 00000000000
                    {
                        texto = TratarCPF_0(texto, addZero);

                    }
                    else if (cl_Settings.CPF_Opcao == 1) // 000.000.000-00
                    {
                        texto = TratarCPF_1(texto, addZero);
                    }
                    else
                    {
                        MessageBox.Show("Option de conversão inválida!", "ERRO: 871174", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    Range selecao = ws.Cells[item.Row, item.Column];
                    selecao.NumberFormat = "@";
                    selecao.Value = texto;
                    contador++;
                }
            }
            MessageBox.Show("CPFs alterados: " + contador.ToString());
        }

        public static string TratarCPF_0(string NumeroCPF, bool AddZero = false)
        {
            // FORMATO: 00000000000

            string NovoNumeroCPF = NumeroCPF.Trim();
            NovoNumeroCPF = Regex.Replace(NovoNumeroCPF, @"[^\d]", "");

            if (AddZero)
            {
                NovoNumeroCPF = NovoNumeroCPF.Trim().PadLeft(11, '0');
            }

            if (NovoNumeroCPF.Length == 11)
            {
                return NovoNumeroCPF;
            }
            else
            {
                return NumeroCPF;
            }
        }

        public static string TratarCPF_1(string NumeroCPF, bool AddZero = false)
        {
            // FORMATO: 000.000.000-00
            string NovoNumeroCPF = NumeroCPF.Trim();
            NovoNumeroCPF = Regex.Replace(NovoNumeroCPF, @"[^\d]", "");

            if (AddZero)
            {
                NovoNumeroCPF = NovoNumeroCPF.Trim().PadLeft(11, '0');
            }

            if (NovoNumeroCPF.Length == 11)
            {
                NovoNumeroCPF = Regex.Replace(NovoNumeroCPF, "([0-9][0-9][0-9])([0-9][0-9][0-9])([0-9][0-9][0-9])([0-9][0-9])", "$1.$2.$3-$4");
            }
            else
            {
                return NumeroCPF;
            }

            return NovoNumeroCPF;
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

        public static string GetDateTime()
        {
            string date = DateTime.Now.ToString("yyyy-MM-dd");
            string time = DateTime.Now.ToString("HH-mm-ss");

            return string.Format("{0}_{1}", date, time);
        }

        public static string GetDate()
        {
            return DateTime.Now.ToString("yyyy-MM-dd"); ;
        }

        public static string GetTime()
        {
            return DateTime.Now.ToString("HH-mm-ss");
        }
    }
}

