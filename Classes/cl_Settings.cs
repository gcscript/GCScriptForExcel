using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using GCScript_for_Excel;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using Appl = Microsoft.Office.Interop.Excel.Application;

namespace GCScript_for_Excel.Classes
{
    public static class cl_Settings
    {
        static Appl app = Globals.ThisAddIn.Application;

        public static bool Text_Trim = true;
        public static bool Text_RemoverEspacoDuplicado = true;
        public static bool Text_RemoverAcentos = true;
        public static int Text_Option = 0;
        // Option 0 = MAIUSCULO
        // Option 1 = minusculo
        // Option 2 = Modo Titulo
        // Option 3 = Original
        // Option 4 = Substituir
        // Option 5 = Alinhamento

        public static int converter_Text_Opcao_Alinhamento = 0;
        // Option 0 = Esquerda
        // Option 1 = Centralizado
        // Option 2 = Direita

        public static string converter_Text_Substituir_Origem = "";
        public static string converter_Text_Substituir_Destino = "";

        public static int converter_Text_Alinhamento_Comprimento = 50;
        public static string converter_Text_Alinhamento_Preenchimento = "-";

        public static bool CPF_ZeroAEsquerda = true;
        public static int CPF_Opcao = 0;
        // Option 0 = 00000000000
        // Option 1 = 000.000.000-00

        public static bool CNPJ_ZeroAEsquerda = true;
        public static int CNPJ_Opcao = 1;
        // Option 0 = 00000000000000
        // Option 1 = 00.000.000/0000-00

        public static int More_SelectionType = 1;
        // Option 0 = Seletion
        // Option 1 = Active Sheet
        // Option 2 = All Sheets

        #region APPLY & REMOVE
        public static bool ApplyRemove_Apply_AllSheets = true;
        public static bool ApplyRemove_Apply_FontName = true;
        public static string ApplyRemove_Apply_FontNameText = "Consolas";
        public static bool ApplyRemove_Apply_FontSize = true;
        public static string ApplyRemove_Apply_FontSizeText = "10";

        public static bool ApplyRemove_Apply_Align_Vertical = true;
        public static int ApplyRemove_Apply_Align_VerticalValue = 1;
        // Option 0 = Top
        // Option 1 = Middle
        // Option 2 = Bottom
        public static bool ApplyRemove_Apply_Align_Horizontal = true;
        public static int ApplyRemove_Apply_Align_HorizontalValue = 0;
        // Option 0 = Left
        // Option 1 = Center
        // Option 2 = Right
        public static bool ApplyRemove_Apply_RowHeight = true;
        public static decimal ApplyRemove_Apply_RowHeightValue = 0;
        public static bool ApplyRemove_Apply_ColumnWidth = true;
        public static decimal ApplyRemove_Apply_ColumnWidthValue = 30;
        public static bool ApplyRemove_Apply_Zoom = true;
        public static decimal ApplyRemove_Apply_ZoomValue = 100;

        public static bool ApplyRemove_Remove_FontBold = true;
        public static bool ApplyRemove_Remove_FontItalic = true;
        public static bool ApplyRemove_Remove_FontUnderline = true;
        public static bool ApplyRemove_Remove_Borders = true;
        public static bool ApplyRemove_Remove_Fill = true;
        public static bool ApplyRemove_Remove_FontColor = true;
        public static bool ApplyRemove_Remove_WrapText = true;
        public static bool ApplyRemove_Remove_MergeCells = true;
        public static bool ApplyRemove_Remove_Formula = true;
        public static bool ApplyRemove_Remove_ConditionalFormatting = true;
        #endregion

        public static void ConverterText(Worksheet ws, Range rng)
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
                    if (Text_Trim)
                    {
                        texto = texto.Trim();
                    }

                    if (Text_RemoverEspacoDuplicado)
                    {
                        texto = RemoverEspacosDuplicados(texto);
                    }

                    if (Text_RemoverAcentos)
                    {
                        texto = RemoverAcentos(texto);
                    }

                    if (Text_Option == 0) // Maiúsculo
                    {
                        texto = texto.ToUpper();
                    }
                    else if (Text_Option == 1) // minúsculo
                    {
                        texto = texto.ToLower();
                    }
                    else if (Text_Option == 2) // Titulo
                    {
                        texto = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(texto);
                    }
                    else if (Text_Option == 3) // Original
                    {
                        // Fazer nada
                    }
                    else if (Text_Option == 4) // Substituir
                    {
                        string origem = converter_Text_Substituir_Origem;
                        string destino = converter_Text_Substituir_Destino;

                        texto = texto.Replace(origem, destino);
                    }
                    else if (Text_Option == 5) // Alinhamento
                    {
                        if (converter_Text_Opcao_Alinhamento == 0) // Esquerda
                        {
                            texto = TextoAEsquerda(texto, converter_Text_Alinhamento_Comprimento, char.Parse(converter_Text_Alinhamento_Preenchimento));
                        }
                        else if (converter_Text_Opcao_Alinhamento == 2) // Direita
                        {
                            texto = TextoADireita(texto, converter_Text_Alinhamento_Comprimento, char.Parse(converter_Text_Alinhamento_Preenchimento));
                        }
                        else if (converter_Text_Opcao_Alinhamento == 1) // Centro
                        {
                            texto = TextoAoCentro(texto, converter_Text_Alinhamento_Comprimento, char.Parse(converter_Text_Alinhamento_Preenchimento));
                        }
                    }
                    else
                    {
                        MessageBox.Show("Option de conversão inválida!", "ERRO: 508027", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    Range selecao = ws.Cells[item.Row, item.Column];
                    selecao.NumberFormat = "@";
                    selecao.Value = texto;
                    contador++;
                }
            }
            MessageBox.Show("Valores alterados: " + contador.ToString());
        }

        public static void ConverterCNPJ(Worksheet ws, Range rng)
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
                    bool addZero = CNPJ_ZeroAEsquerda;

                    if (CNPJ_Opcao == 0) // 00000000000000
                    {
                        texto = TratarCNPJ_0(texto, addZero);

                    }
                    else if (CNPJ_Opcao == 1) // 00.000.000/0000-00
                    {
                        texto = TratarCNPJ_1(texto, addZero);
                    }
                    else
                    {
                        MessageBox.Show("Option de conversão inválida!", "ERRO: 672219", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    Range selecao = ws.Cells[item.Row, item.Column];
                    selecao.NumberFormat = "@";
                    selecao.Value = texto;
                    contador++;
                }
            }
            MessageBox.Show("CNPJs alterados: " + contador.ToString());
        }

        public static void ConverterWorkSchedule(Worksheet ws, Range rng)
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
                    texto = WorkSchedule(texto);
                    Range selecao = ws.Cells[item.Row, item.Column];
                    selecao.NumberFormat = "@";
                    selecao.Value = texto;
                    contador++;
                }
            }
            MessageBox.Show("Valores alterados: " + contador.ToString());
        }

        public static string RemoverAcentos(string texto)
        {
            StringBuilder sbReturn = new StringBuilder();
            var arrayText = texto.Normalize(NormalizationForm.FormD).ToCharArray();
            foreach (char letter in arrayText)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(letter) != UnicodeCategory.NonSpacingMark)
                    sbReturn.Append(letter);
            }
            return sbReturn.ToString();
        }

        public static string RemoverEspacosDuplicados(string texto)
        {
            texto = Regex.Replace(texto, @"\s{2,}", " ");
            texto = texto.Trim();

            return texto;
        }

        public static string WorkSchedule(string texto)
        {
            texto = texto.ToUpper().Trim();
            texto = Regex.Replace(texto, @"\s", "");
            if (texto.Contains("6X1")) { return "6X1"; }
            else if (texto.Contains("06X01")) { return "6X1"; }
            else if (texto.Contains("6X2")) { return "6X1"; }
            else if (texto.Contains("60X01")) { return "6X1"; }
            else if (texto.Contains("05X02")) { return "5X2"; }
            else if (texto.Contains("5X2")) { return "5X2"; }
            else if (texto.Contains("SX2")) { return "5X2"; }
            else if (texto.Contains("5X1")) { return "5X2"; }
            else if (texto.Contains("44H")) { return "5X2"; }
            else if (texto.Contains("12X36")) { return "12X36"; }
            else if (texto.Contains("13X36")) { return "12X36"; }
            else if (texto.Contains("24X48")) { return "24X48"; }
            else if (texto.Contains("04X03")) { return "4X3"; }
            else if (texto.Contains("4X3")) { return "4X3"; }
            else if (texto.Contains("03X04")) { return "3X4"; }
            else if (texto.Contains("3X4")) { return "3X4"; }
            else if (texto.Contains("02X05")) { return "2X5"; }
            else if (texto.Contains("2X5")) { return "2X5"; }
            else if (texto.Contains("01X06")) { return "1X6"; }
            else if (texto.Contains("1X6")) { return "1X6"; }
            else { return texto; }

        }

        public static string TextoAEsquerda(string texto, int comprimento, char preenchimento)
        {
            return texto.PadRight(comprimento, preenchimento);
        }

        public static string TextoADireita(string texto, int comprimento, char preenchimento)
        {
            return texto.PadLeft(comprimento, preenchimento);
        }

        public static string TextoAoCentro(string texto, int comprimento, char caractere)
        {
            int spaces = comprimento - texto.Length;
            int padLeft = spaces / 2 + texto.Length;
            return texto.PadLeft(padLeft, caractere).PadRight(comprimento, caractere);
        }
        
        public static string TratarCNPJ_0(string NumeroCNPJ, bool AddZero = false)
        {
            // FORMATO: 00000000000000

            string NovoNumeroCNPJ = NumeroCNPJ.Trim();
            NovoNumeroCNPJ = Regex.Replace(NovoNumeroCNPJ, @"[^\d]", "");

            if (AddZero)
            {
                NovoNumeroCNPJ = NovoNumeroCNPJ.Trim().PadLeft(14, '0');
            }

            if (NovoNumeroCNPJ.Length == 14)
            {
                return NovoNumeroCNPJ;
            }
            else
            {
                return NumeroCNPJ;
            }
        }

        public static string TratarCNPJ_1(string NumeroCNPJ, bool AddZero = false)
        {
            // FORMATO: 00.000.000/0000-00
            string NovoNumeroCNPJ = NumeroCNPJ.Trim();
            NovoNumeroCNPJ = Regex.Replace(NovoNumeroCNPJ, @"[^\d]", "");

            if (AddZero)
            {
                NovoNumeroCNPJ = NovoNumeroCNPJ.Trim().PadLeft(14, '0');
            }

            if (NovoNumeroCNPJ.Length == 14)
            {
                NovoNumeroCNPJ = Regex.Replace(NovoNumeroCNPJ, "([0-9]{2})([0-9]{3})([0-9]{3})([0-9]{4})([0-9]{2})", "$1.$2.$3/$4-$5");
            }
            else
            {
                return NumeroCNPJ;
            }

            return NovoNumeroCNPJ;
        }

    }
}
