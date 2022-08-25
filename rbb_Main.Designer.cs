namespace GCScript_for_Excel
{
    partial class rbb_Main : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variável de designer necessária.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rbb_Main()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Limpar os recursos que estão sendo usados.
        /// </summary>
        /// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código gerado pelo Designer de Componentes

        /// <summary>
        /// Método necessário para suporte ao Designer - não modifique 
        /// o conteúdo deste método com o editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(rbb_Main));
            this.rbt_Main = this.Factory.CreateRibbonTab();
            this.grp_Tools = this.Factory.CreateRibbonGroup();
            this.grp_Styles = this.Factory.CreateRibbonGroup();
            this.grp_Number = this.Factory.CreateRibbonGroup();
            this.grp_Info = this.Factory.CreateRibbonGroup();
            this.grp_Others = this.Factory.CreateRibbonGroup();
            this.rbt_Testes = this.Factory.CreateRibbonTab();
            this.grp_Tests = this.Factory.CreateRibbonGroup();
            this.grp_Beta = this.Factory.CreateRibbonGroup();
            this.glr_ApplyRemove = this.Factory.CreateRibbonGallery();
            this.ApplyRemove_btn_Start = this.Factory.CreateRibbonButton();
            this.ApplyRemove_btn_Settings = this.Factory.CreateRibbonButton();
            this.glr_Generate = this.Factory.CreateRibbonGallery();
            this.Generate_btn_Apportionment = this.Factory.CreateRibbonButton();
            this.Generate_btn_Purchase = this.Factory.CreateRibbonButton();
            this.Generate_btn_FileToSend = this.Factory.CreateRibbonButton();
            this.Generate_btn_Separator1 = this.Factory.CreateRibbonButton();
            this.Generate_btn_RandomCPF = this.Factory.CreateRibbonButton();
            this.glr_Converter = this.Factory.CreateRibbonGallery();
            this.Converter_btn_Text = this.Factory.CreateRibbonButton();
            this.Converter_btn_CPF = this.Factory.CreateRibbonButton();
            this.Converter_btn_CNPJ = this.Factory.CreateRibbonButton();
            this.Converter_btn_WorkSchedule = this.Factory.CreateRibbonButton();
            this.Converter_btn_Separator1 = this.Factory.CreateRibbonButton();
            this.Converter_btn_Round_0DecimalPlaces = this.Factory.CreateRibbonButton();
            this.Converter_btn_Round_1DecimalPlaces = this.Factory.CreateRibbonButton();
            this.Converter_btn_Round_2DecimalPlaces = this.Factory.CreateRibbonButton();
            this.Converter_btn_Separator2 = this.Factory.CreateRibbonButton();
            this.Converter_btn_Settings = this.Factory.CreateRibbonButton();
            this.Tools_btn_BZA = this.Factory.CreateRibbonButton();
            this.glr_Sort = this.Factory.CreateRibbonGallery();
            this.Sort_btn_Preset01 = this.Factory.CreateRibbonButton();
            this.Sort_btn_Preset02 = this.Factory.CreateRibbonButton();
            this.Sort_btn_Preset03 = this.Factory.CreateRibbonButton();
            this.Sort_btn_Preset04 = this.Factory.CreateRibbonButton();
            this.glr_Sheets = this.Factory.CreateRibbonGallery();
            this.Sheets_btn_SortSheetsASC = this.Factory.CreateRibbonButton();
            this.Sheets_btn_SortSheetsDESC = this.Factory.CreateRibbonButton();
            this.Sheets_btn_ShowHiddenSheets = this.Factory.CreateRibbonButton();
            this.Sheets_btn_RemoveHiddenSheets = this.Factory.CreateRibbonButton();
            this.glr_More = this.Factory.CreateRibbonGallery();
            this.More_btn_OnlyValues = this.Factory.CreateRibbonButton();
            this.More_btn_RemoveConditionalFormatting = this.Factory.CreateRibbonButton();
            this.More_btn_Separator1 = this.Factory.CreateRibbonButton();
            this.More_btn_CheckSelection = this.Factory.CreateRibbonButton();
            this.More_btn_CheckActiveSheet = this.Factory.CreateRibbonButton();
            this.More_btn_CheckAllSheets = this.Factory.CreateRibbonButton();
            this.Styles_btn_Primary = this.Factory.CreateRibbonButton();
            this.Styles_btn_Secondary = this.Factory.CreateRibbonButton();
            this.Styles_btn_Success = this.Factory.CreateRibbonButton();
            this.Styles_btn_Danger = this.Factory.CreateRibbonButton();
            this.Styles_btn_Warning = this.Factory.CreateRibbonButton();
            this.Styles_btn_Info = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap = this.Factory.CreateRibbonGallery();
            this.Styles_glr_Bootstrap_Primary = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Secondary = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Success = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Danger = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Warning = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Info = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Light = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_Dark = this.Factory.CreateRibbonButton();
            this.Styles_glr_Bootstrap_White = this.Factory.CreateRibbonButton();
            this.Styles_glr_Emphasis = this.Factory.CreateRibbonGallery();
            this.Styles_glr_Emphasis1 = this.Factory.CreateRibbonButton();
            this.Styles_glr_Emphasis2 = this.Factory.CreateRibbonButton();
            this.Styles_glr_Emphasis3 = this.Factory.CreateRibbonButton();
            this.Styles_glr_Emphasis4 = this.Factory.CreateRibbonButton();
            this.Styles_glr_Emphasis5 = this.Factory.CreateRibbonButton();
            this.Styles_btn_Default = this.Factory.CreateRibbonButton();
            this.Number_btn_General = this.Factory.CreateRibbonButton();
            this.Number_btn_Text = this.Factory.CreateRibbonButton();
            this.Number_btn_Accounting = this.Factory.CreateRibbonButton();
            this.Info_btn_LinhasUsadas = this.Factory.CreateRibbonButton();
            this.Info_btn_ColunasUsadas = this.Factory.CreateRibbonButton();
            this.Info_btn_ObterTipoCell = this.Factory.CreateRibbonButton();
            this.Info_btn_SelecionarRange = this.Factory.CreateRibbonButton();
            this.Info_btn_SelecionarTudo = this.Factory.CreateRibbonButton();
            this.Tools_btn_SetTitle = this.Factory.CreateRibbonButton();
            this.btn_Settings = this.Factory.CreateRibbonButton();
            this.btn_T1 = this.Factory.CreateRibbonButton();
            this.btn_T2 = this.Factory.CreateRibbonButton();
            this.btn_T3 = this.Factory.CreateRibbonButton();
            this.btn_T4 = this.Factory.CreateRibbonButton();
            this.btn_T5 = this.Factory.CreateRibbonButton();
            this.btn_T6 = this.Factory.CreateRibbonButton();
            this.btn_RemoverFC = this.Factory.CreateRibbonButton();
            this.rbt_Main.SuspendLayout();
            this.grp_Tools.SuspendLayout();
            this.grp_Styles.SuspendLayout();
            this.grp_Number.SuspendLayout();
            this.grp_Info.SuspendLayout();
            this.grp_Others.SuspendLayout();
            this.rbt_Testes.SuspendLayout();
            this.grp_Tests.SuspendLayout();
            this.grp_Beta.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbt_Main
            // 
            this.rbt_Main.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.rbt_Main.Groups.Add(this.grp_Tools);
            this.rbt_Main.Groups.Add(this.grp_Styles);
            this.rbt_Main.Groups.Add(this.grp_Number);
            this.rbt_Main.Groups.Add(this.grp_Info);
            this.rbt_Main.Groups.Add(this.grp_Others);
            this.rbt_Main.KeyTip = "GCS";
            this.rbt_Main.Label = "GCScript";
            this.rbt_Main.Name = "rbt_Main";
            // 
            // grp_Tools
            // 
            this.grp_Tools.Items.Add(this.glr_ApplyRemove);
            this.grp_Tools.Items.Add(this.glr_Generate);
            this.grp_Tools.Items.Add(this.glr_Converter);
            this.grp_Tools.Items.Add(this.Tools_btn_BZA);
            this.grp_Tools.Items.Add(this.glr_Sort);
            this.grp_Tools.Items.Add(this.glr_Sheets);
            this.grp_Tools.Items.Add(this.glr_More);
            this.grp_Tools.KeyTip = "1";
            this.grp_Tools.Label = "Tools";
            this.grp_Tools.Name = "grp_Tools";
            // 
            // grp_Styles
            // 
            this.grp_Styles.Items.Add(this.Styles_btn_Primary);
            this.grp_Styles.Items.Add(this.Styles_btn_Secondary);
            this.grp_Styles.Items.Add(this.Styles_btn_Success);
            this.grp_Styles.Items.Add(this.Styles_btn_Danger);
            this.grp_Styles.Items.Add(this.Styles_btn_Warning);
            this.grp_Styles.Items.Add(this.Styles_btn_Info);
            this.grp_Styles.Items.Add(this.Styles_glr_Bootstrap);
            this.grp_Styles.Items.Add(this.Styles_glr_Emphasis);
            this.grp_Styles.Items.Add(this.Styles_btn_Default);
            this.grp_Styles.KeyTip = "2";
            this.grp_Styles.Label = "Styles";
            this.grp_Styles.Name = "grp_Styles";
            // 
            // grp_Number
            // 
            this.grp_Number.Items.Add(this.Number_btn_General);
            this.grp_Number.Items.Add(this.Number_btn_Text);
            this.grp_Number.Items.Add(this.Number_btn_Accounting);
            this.grp_Number.KeyTip = "3";
            this.grp_Number.Label = "Number";
            this.grp_Number.Name = "grp_Number";
            // 
            // grp_Info
            // 
            this.grp_Info.Items.Add(this.Info_btn_LinhasUsadas);
            this.grp_Info.Items.Add(this.Info_btn_ColunasUsadas);
            this.grp_Info.Items.Add(this.Info_btn_ObterTipoCell);
            this.grp_Info.Items.Add(this.Info_btn_SelecionarRange);
            this.grp_Info.Items.Add(this.Info_btn_SelecionarTudo);
            this.grp_Info.KeyTip = "4";
            this.grp_Info.Label = "Info";
            this.grp_Info.Name = "grp_Info";
            // 
            // grp_Others
            // 
            this.grp_Others.Items.Add(this.Tools_btn_SetTitle);
            this.grp_Others.Items.Add(this.btn_Settings);
            this.grp_Others.KeyTip = "5";
            this.grp_Others.Label = "Others";
            this.grp_Others.Name = "grp_Others";
            // 
            // rbt_Testes
            // 
            this.rbt_Testes.Groups.Add(this.grp_Tests);
            this.rbt_Testes.Groups.Add(this.grp_Beta);
            this.rbt_Testes.Label = "GCScript (Testes)";
            this.rbt_Testes.Name = "rbt_Testes";
            // 
            // grp_Tests
            // 
            this.grp_Tests.Items.Add(this.btn_T1);
            this.grp_Tests.Items.Add(this.btn_T2);
            this.grp_Tests.Items.Add(this.btn_T3);
            this.grp_Tests.Items.Add(this.btn_T4);
            this.grp_Tests.Items.Add(this.btn_T5);
            this.grp_Tests.Items.Add(this.btn_T6);
            this.grp_Tests.Label = "Testes";
            this.grp_Tests.Name = "grp_Tests";
            // 
            // grp_Beta
            // 
            this.grp_Beta.Items.Add(this.btn_RemoverFC);
            this.grp_Beta.Label = "Beta";
            this.grp_Beta.Name = "grp_Beta";
            // 
            // glr_ApplyRemove
            // 
            this.glr_ApplyRemove.Buttons.Add(this.ApplyRemove_btn_Start);
            this.glr_ApplyRemove.Buttons.Add(this.ApplyRemove_btn_Settings);
            this.glr_ApplyRemove.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_ApplyRemove.Image = global::GCScript_for_Excel.Properties.Resources.apply_remove;
            this.glr_ApplyRemove.KeyTip = "1";
            this.glr_ApplyRemove.Label = "Apply Remove";
            this.glr_ApplyRemove.Name = "glr_ApplyRemove";
            this.glr_ApplyRemove.ShowImage = true;
            // 
            // ApplyRemove_btn_Start
            // 
            this.ApplyRemove_btn_Start.Image = global::GCScript_for_Excel.Properties.Resources.play;
            this.ApplyRemove_btn_Start.Label = "Start";
            this.ApplyRemove_btn_Start.Name = "ApplyRemove_btn_Start";
            this.ApplyRemove_btn_Start.ScreenTip = "Start";
            this.ApplyRemove_btn_Start.ShowImage = true;
            this.ApplyRemove_btn_Start.SuperTip = "Start";
            this.ApplyRemove_btn_Start.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ApplyRemove_btn_Start_Click);
            // 
            // ApplyRemove_btn_Settings
            // 
            this.ApplyRemove_btn_Settings.Image = global::GCScript_for_Excel.Properties.Resources.settings;
            this.ApplyRemove_btn_Settings.Label = "Settings";
            this.ApplyRemove_btn_Settings.Name = "ApplyRemove_btn_Settings";
            this.ApplyRemove_btn_Settings.ScreenTip = "Settings";
            this.ApplyRemove_btn_Settings.ShowImage = true;
            this.ApplyRemove_btn_Settings.SuperTip = "Settings";
            this.ApplyRemove_btn_Settings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ApplyRemove_btn_Settings_Click);
            // 
            // glr_Generate
            // 
            this.glr_Generate.Buttons.Add(this.Generate_btn_Apportionment);
            this.glr_Generate.Buttons.Add(this.Generate_btn_Purchase);
            this.glr_Generate.Buttons.Add(this.Generate_btn_FileToSend);
            this.glr_Generate.Buttons.Add(this.Generate_btn_Separator1);
            this.glr_Generate.Buttons.Add(this.Generate_btn_RandomCPF);
            this.glr_Generate.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_Generate.Image = global::GCScript_for_Excel.Properties.Resources.create;
            this.glr_Generate.KeyTip = "2";
            this.glr_Generate.Label = "Generate";
            this.glr_Generate.Name = "glr_Generate";
            this.glr_Generate.ScreenTip = "Generate";
            this.glr_Generate.ShowImage = true;
            // 
            // Generate_btn_Apportionment
            // 
            this.Generate_btn_Apportionment.Image = global::GCScript_for_Excel.Properties.Resources.rateio;
            this.Generate_btn_Apportionment.Label = "Apportionment";
            this.Generate_btn_Apportionment.Name = "Generate_btn_Apportionment";
            this.Generate_btn_Apportionment.ScreenTip = "Apportionment";
            this.Generate_btn_Apportionment.ShowImage = true;
            this.Generate_btn_Apportionment.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Generate_btn_Apportionment_Click);
            // 
            // Generate_btn_Purchase
            // 
            this.Generate_btn_Purchase.Image = global::GCScript_for_Excel.Properties.Resources.shop;
            this.Generate_btn_Purchase.Label = "Purchase";
            this.Generate_btn_Purchase.Name = "Generate_btn_Purchase";
            this.Generate_btn_Purchase.ScreenTip = "Purchase";
            this.Generate_btn_Purchase.ShowImage = true;
            this.Generate_btn_Purchase.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Generate_btn_Purchase_Click);
            // 
            // Generate_btn_FileToSend
            // 
            this.Generate_btn_FileToSend.Image = global::GCScript_for_Excel.Properties.Resources.file_to_send;
            this.Generate_btn_FileToSend.Label = "File To Send";
            this.Generate_btn_FileToSend.Name = "Generate_btn_FileToSend";
            this.Generate_btn_FileToSend.ScreenTip = "File To Send";
            this.Generate_btn_FileToSend.ShowImage = true;
            this.Generate_btn_FileToSend.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Generate_btn_FileToSend_Click);
            // 
            // Generate_btn_Separator1
            // 
            this.Generate_btn_Separator1.Enabled = false;
            this.Generate_btn_Separator1.Label = "---------------------[OTHERS]---------------------";
            this.Generate_btn_Separator1.Name = "Generate_btn_Separator1";
            // 
            // Generate_btn_RandomCPF
            // 
            this.Generate_btn_RandomCPF.Image = global::GCScript_for_Excel.Properties.Resources.cpf;
            this.Generate_btn_RandomCPF.Label = "Random CPF";
            this.Generate_btn_RandomCPF.Name = "Generate_btn_RandomCPF";
            this.Generate_btn_RandomCPF.ScreenTip = "Random CPF";
            this.Generate_btn_RandomCPF.ShowImage = true;
            this.Generate_btn_RandomCPF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Generate_btn_RandomCPF_Click);
            // 
            // glr_Converter
            // 
            this.glr_Converter.Buttons.Add(this.Converter_btn_Text);
            this.glr_Converter.Buttons.Add(this.Converter_btn_CPF);
            this.glr_Converter.Buttons.Add(this.Converter_btn_CNPJ);
            this.glr_Converter.Buttons.Add(this.Converter_btn_WorkSchedule);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Separator1);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Round_0DecimalPlaces);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Round_1DecimalPlaces);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Round_2DecimalPlaces);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Separator2);
            this.glr_Converter.Buttons.Add(this.Converter_btn_Settings);
            this.glr_Converter.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_Converter.Image = global::GCScript_for_Excel.Properties.Resources.change;
            this.glr_Converter.KeyTip = "3";
            this.glr_Converter.Label = "Converter";
            this.glr_Converter.Name = "glr_Converter";
            this.glr_Converter.ShowImage = true;
            // 
            // Converter_btn_Text
            // 
            this.Converter_btn_Text.Image = global::GCScript_for_Excel.Properties.Resources.text;
            this.Converter_btn_Text.Label = "Text";
            this.Converter_btn_Text.Name = "Converter_btn_Text";
            this.Converter_btn_Text.ShowImage = true;
            this.Converter_btn_Text.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_Text_Click);
            // 
            // Converter_btn_CPF
            // 
            this.Converter_btn_CPF.Image = global::GCScript_for_Excel.Properties.Resources.cpf;
            this.Converter_btn_CPF.Label = "CPF";
            this.Converter_btn_CPF.Name = "Converter_btn_CPF";
            this.Converter_btn_CPF.ScreenTip = "CPF";
            this.Converter_btn_CPF.ShowImage = true;
            this.Converter_btn_CPF.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_CPF_Click);
            // 
            // Converter_btn_CNPJ
            // 
            this.Converter_btn_CNPJ.Image = global::GCScript_for_Excel.Properties.Resources.cnpj;
            this.Converter_btn_CNPJ.Label = "CNPJ";
            this.Converter_btn_CNPJ.Name = "Converter_btn_CNPJ";
            this.Converter_btn_CNPJ.ScreenTip = "CNPJ";
            this.Converter_btn_CNPJ.ShowImage = true;
            this.Converter_btn_CNPJ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_CNPJ_Click);
            // 
            // Converter_btn_WorkSchedule
            // 
            this.Converter_btn_WorkSchedule.Image = global::GCScript_for_Excel.Properties.Resources.clock;
            this.Converter_btn_WorkSchedule.Label = "Work Schedule";
            this.Converter_btn_WorkSchedule.Name = "Converter_btn_WorkSchedule";
            this.Converter_btn_WorkSchedule.ShowImage = true;
            this.Converter_btn_WorkSchedule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_WorkSchedule_Click);
            // 
            // Converter_btn_Separator1
            // 
            this.Converter_btn_Separator1.Enabled = false;
            this.Converter_btn_Separator1.Label = "---------------------[ROUND]----------------------";
            this.Converter_btn_Separator1.Name = "Converter_btn_Separator1";
            // 
            // Converter_btn_Round_0DecimalPlaces
            // 
            this.Converter_btn_Round_0DecimalPlaces.Image = global::GCScript_for_Excel.Properties.Resources.decimal_place;
            this.Converter_btn_Round_0DecimalPlaces.Label = "0 Decimal Places";
            this.Converter_btn_Round_0DecimalPlaces.Name = "Converter_btn_Round_0DecimalPlaces";
            this.Converter_btn_Round_0DecimalPlaces.ScreenTip = "0 Decimal Place";
            this.Converter_btn_Round_0DecimalPlaces.ShowImage = true;
            this.Converter_btn_Round_0DecimalPlaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_Round_0DecimalPlaces_Click);
            // 
            // Converter_btn_Round_1DecimalPlaces
            // 
            this.Converter_btn_Round_1DecimalPlaces.Image = global::GCScript_for_Excel.Properties.Resources.decimal_place;
            this.Converter_btn_Round_1DecimalPlaces.Label = "1 Decimal Places";
            this.Converter_btn_Round_1DecimalPlaces.Name = "Converter_btn_Round_1DecimalPlaces";
            this.Converter_btn_Round_1DecimalPlaces.ScreenTip = "1 Decimal Place";
            this.Converter_btn_Round_1DecimalPlaces.ShowImage = true;
            this.Converter_btn_Round_1DecimalPlaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_Round_1DecimalPlaces_Click);
            // 
            // Converter_btn_Round_2DecimalPlaces
            // 
            this.Converter_btn_Round_2DecimalPlaces.Image = global::GCScript_for_Excel.Properties.Resources.decimal_place;
            this.Converter_btn_Round_2DecimalPlaces.Label = "2 Decimal Places";
            this.Converter_btn_Round_2DecimalPlaces.Name = "Converter_btn_Round_2DecimalPlaces";
            this.Converter_btn_Round_2DecimalPlaces.ScreenTip = "2 Decimal Places";
            this.Converter_btn_Round_2DecimalPlaces.ShowImage = true;
            this.Converter_btn_Round_2DecimalPlaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_Round_2DecimalPlaces_Click);
            // 
            // Converter_btn_Separator2
            // 
            this.Converter_btn_Separator2.Enabled = false;
            this.Converter_btn_Separator2.Label = "--------------------[SETTINGS]--------------------";
            this.Converter_btn_Separator2.Name = "Converter_btn_Separator2";
            // 
            // Converter_btn_Settings
            // 
            this.Converter_btn_Settings.Image = global::GCScript_for_Excel.Properties.Resources.settings;
            this.Converter_btn_Settings.Label = "Settings";
            this.Converter_btn_Settings.Name = "Converter_btn_Settings";
            this.Converter_btn_Settings.ShowImage = true;
            this.Converter_btn_Settings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Converter_btn_Settings_Click);
            // 
            // Tools_btn_BZA
            // 
            this.Tools_btn_BZA.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Tools_btn_BZA.Image = global::GCScript_for_Excel.Properties.Resources.bzpa;
            this.Tools_btn_BZA.KeyTip = "4";
            this.Tools_btn_BZA.Label = "BZPA";
            this.Tools_btn_BZA.Name = "Tools_btn_BZA";
            this.Tools_btn_BZA.ScreenTip = "Border + Zoom + Print Area";
            this.Tools_btn_BZA.ShowImage = true;
            this.Tools_btn_BZA.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Tools_DefinirAreaMex_Click);
            // 
            // glr_Sort
            // 
            this.glr_Sort.Buttons.Add(this.Sort_btn_Preset01);
            this.glr_Sort.Buttons.Add(this.Sort_btn_Preset02);
            this.glr_Sort.Buttons.Add(this.Sort_btn_Preset03);
            this.glr_Sort.Buttons.Add(this.Sort_btn_Preset04);
            this.glr_Sort.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_Sort.Image = global::GCScript_for_Excel.Properties.Resources.sort;
            this.glr_Sort.KeyTip = "5";
            this.glr_Sort.Label = "Sort";
            this.glr_Sort.Name = "glr_Sort";
            this.glr_Sort.ShowImage = true;
            // 
            // Sort_btn_Preset01
            // 
            this.Sort_btn_Preset01.Image = global::GCScript_for_Excel.Properties.Resources.sort;
            this.Sort_btn_Preset01.Label = "UF > Operadora > Nome";
            this.Sort_btn_Preset01.Name = "Sort_btn_Preset01";
            this.Sort_btn_Preset01.ScreenTip = "UF > Operadora > Nome";
            this.Sort_btn_Preset01.ShowImage = true;
            this.Sort_btn_Preset01.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sort_btn_Preset01_Click);
            // 
            // Sort_btn_Preset02
            // 
            this.Sort_btn_Preset02.Image = global::GCScript_for_Excel.Properties.Resources.sort;
            this.Sort_btn_Preset02.Label = "UF > Operadora > Empresa> C.Unid > Nome";
            this.Sort_btn_Preset02.Name = "Sort_btn_Preset02";
            this.Sort_btn_Preset02.ScreenTip = "UF > Operadora > Empresa> C.Unid > Nome";
            this.Sort_btn_Preset02.ShowImage = true;
            this.Sort_btn_Preset02.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sort_btn_Preset02_Click);
            // 
            // Sort_btn_Preset03
            // 
            this.Sort_btn_Preset03.Image = global::GCScript_for_Excel.Properties.Resources.sort;
            this.Sort_btn_Preset03.Label = "UF > Operadora > Empresa > C.Unid > C.Depto > Nome";
            this.Sort_btn_Preset03.Name = "Sort_btn_Preset03";
            this.Sort_btn_Preset03.ScreenTip = "UF > Operadora > Empresa > C.Unid > C.Depto > Nome";
            this.Sort_btn_Preset03.ShowImage = true;
            this.Sort_btn_Preset03.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sort_btn_Preset03_Click);
            // 
            // Sort_btn_Preset04
            // 
            this.Sort_btn_Preset04.Image = global::GCScript_for_Excel.Properties.Resources.sort;
            this.Sort_btn_Preset04.Label = "UF > Operadora > Empresa > C.Unid > C.Depto > Depto > Nome";
            this.Sort_btn_Preset04.Name = "Sort_btn_Preset04";
            this.Sort_btn_Preset04.ScreenTip = "UF > Operadora > Empresa > C.Unid > C.Depto > Depto > Nome";
            this.Sort_btn_Preset04.ShowImage = true;
            this.Sort_btn_Preset04.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sort_btn_Preset04_Click);
            // 
            // glr_Sheets
            // 
            this.glr_Sheets.Buttons.Add(this.Sheets_btn_SortSheetsASC);
            this.glr_Sheets.Buttons.Add(this.Sheets_btn_SortSheetsDESC);
            this.glr_Sheets.Buttons.Add(this.Sheets_btn_ShowHiddenSheets);
            this.glr_Sheets.Buttons.Add(this.Sheets_btn_RemoveHiddenSheets);
            this.glr_Sheets.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_Sheets.Image = global::GCScript_for_Excel.Properties.Resources.sheet;
            this.glr_Sheets.KeyTip = "6";
            this.glr_Sheets.Label = "Sheets";
            this.glr_Sheets.Name = "glr_Sheets";
            this.glr_Sheets.ShowImage = true;
            // 
            // Sheets_btn_SortSheetsASC
            // 
            this.Sheets_btn_SortSheetsASC.Image = global::GCScript_for_Excel.Properties.Resources.sheet;
            this.Sheets_btn_SortSheetsASC.Label = "Sort Sheets [ASC]";
            this.Sheets_btn_SortSheetsASC.Name = "Sheets_btn_SortSheetsASC";
            this.Sheets_btn_SortSheetsASC.ScreenTip = "Sort Sheets [ASC]";
            this.Sheets_btn_SortSheetsASC.ShowImage = true;
            this.Sheets_btn_SortSheetsASC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheets_btn_SortSheetsASC_Click);
            // 
            // Sheets_btn_SortSheetsDESC
            // 
            this.Sheets_btn_SortSheetsDESC.Image = global::GCScript_for_Excel.Properties.Resources.sheet;
            this.Sheets_btn_SortSheetsDESC.Label = "Sort Sheets [DESC]";
            this.Sheets_btn_SortSheetsDESC.Name = "Sheets_btn_SortSheetsDESC";
            this.Sheets_btn_SortSheetsDESC.ScreenTip = "Sort Sheets [DESC]";
            this.Sheets_btn_SortSheetsDESC.ShowImage = true;
            this.Sheets_btn_SortSheetsDESC.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheets_btn_SortSheetsDESC_Click);
            // 
            // Sheets_btn_ShowHiddenSheets
            // 
            this.Sheets_btn_ShowHiddenSheets.Image = global::GCScript_for_Excel.Properties.Resources.show;
            this.Sheets_btn_ShowHiddenSheets.Label = "Show Hidden Sheets";
            this.Sheets_btn_ShowHiddenSheets.Name = "Sheets_btn_ShowHiddenSheets";
            this.Sheets_btn_ShowHiddenSheets.ScreenTip = "Show Hidden Sheets";
            this.Sheets_btn_ShowHiddenSheets.ShowImage = true;
            this.Sheets_btn_ShowHiddenSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheets_btn_ShowHiddenSheets_Click);
            // 
            // Sheets_btn_RemoveHiddenSheets
            // 
            this.Sheets_btn_RemoveHiddenSheets.Image = global::GCScript_for_Excel.Properties.Resources.remove;
            this.Sheets_btn_RemoveHiddenSheets.Label = "Remove Hidden Sheets";
            this.Sheets_btn_RemoveHiddenSheets.Name = "Sheets_btn_RemoveHiddenSheets";
            this.Sheets_btn_RemoveHiddenSheets.ScreenTip = "Remove Hidden Sheets";
            this.Sheets_btn_RemoveHiddenSheets.ShowImage = true;
            this.Sheets_btn_RemoveHiddenSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Sheets_btn_RemoveHiddenSheets_Click);
            // 
            // glr_More
            // 
            this.glr_More.Buttons.Add(this.More_btn_OnlyValues);
            this.glr_More.Buttons.Add(this.More_btn_RemoveConditionalFormatting);
            this.glr_More.Buttons.Add(this.More_btn_Separator1);
            this.glr_More.Buttons.Add(this.More_btn_CheckSelection);
            this.glr_More.Buttons.Add(this.More_btn_CheckActiveSheet);
            this.glr_More.Buttons.Add(this.More_btn_CheckAllSheets);
            this.glr_More.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.glr_More.Image = global::GCScript_for_Excel.Properties.Resources.more;
            this.glr_More.KeyTip = "7";
            this.glr_More.Label = "More";
            this.glr_More.Name = "glr_More";
            this.glr_More.ShowImage = true;
            this.glr_More.ShowItemImage = false;
            this.glr_More.ItemsLoading += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.glr_More_ItemsLoading);
            // 
            // More_btn_OnlyValues
            // 
            this.More_btn_OnlyValues.Image = global::GCScript_for_Excel.Properties.Resources.value;
            this.More_btn_OnlyValues.Label = "Only Value (Selection)";
            this.More_btn_OnlyValues.Name = "More_btn_OnlyValues";
            this.More_btn_OnlyValues.ShowImage = true;
            this.More_btn_OnlyValues.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.More_btn_OnlyValues_Click);
            // 
            // More_btn_RemoveConditionalFormatting
            // 
            this.More_btn_RemoveConditionalFormatting.Image = global::GCScript_for_Excel.Properties.Resources.conditional_formatting;
            this.More_btn_RemoveConditionalFormatting.Label = "Remove Conditional Formatting (Selection)";
            this.More_btn_RemoveConditionalFormatting.Name = "More_btn_RemoveConditionalFormatting";
            this.More_btn_RemoveConditionalFormatting.ShowImage = true;
            this.More_btn_RemoveConditionalFormatting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.More_btn_RemoveConditionalFormatting_Click);
            // 
            // More_btn_Separator1
            // 
            this.More_btn_Separator1.Enabled = false;
            this.More_btn_Separator1.Label = "----------------------------------------";
            this.More_btn_Separator1.Name = "More_btn_Separator1";
            // 
            // More_btn_CheckSelection
            // 
            this.More_btn_CheckSelection.Image = global::GCScript_for_Excel.Properties.Resources.uncheck;
            this.More_btn_CheckSelection.Label = "Selection";
            this.More_btn_CheckSelection.Name = "More_btn_CheckSelection";
            this.More_btn_CheckSelection.ScreenTip = "Selection";
            this.More_btn_CheckSelection.ShowImage = true;
            this.More_btn_CheckSelection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.More_btn_CheckSelection_Click);
            // 
            // More_btn_CheckActiveSheet
            // 
            this.More_btn_CheckActiveSheet.Image = global::GCScript_for_Excel.Properties.Resources.check;
            this.More_btn_CheckActiveSheet.Label = "Active Sheet";
            this.More_btn_CheckActiveSheet.Name = "More_btn_CheckActiveSheet";
            this.More_btn_CheckActiveSheet.ScreenTip = "Active Sheet";
            this.More_btn_CheckActiveSheet.ShowImage = true;
            this.More_btn_CheckActiveSheet.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.More_btn_CheckActiveSheet_Click);
            // 
            // More_btn_CheckAllSheets
            // 
            this.More_btn_CheckAllSheets.Image = global::GCScript_for_Excel.Properties.Resources.uncheck;
            this.More_btn_CheckAllSheets.Label = "All Sheets";
            this.More_btn_CheckAllSheets.Name = "More_btn_CheckAllSheets";
            this.More_btn_CheckAllSheets.ScreenTip = "All Sheets";
            this.More_btn_CheckAllSheets.ShowImage = true;
            this.More_btn_CheckAllSheets.Tag = "";
            this.More_btn_CheckAllSheets.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.More_btn_CheckAllSheets_Click);
            // 
            // Styles_btn_Primary
            // 
            this.Styles_btn_Primary.Image = global::GCScript_for_Excel.Properties.Resources.styles_primary;
            this.Styles_btn_Primary.KeyTip = "1";
            this.Styles_btn_Primary.Label = "Primary";
            this.Styles_btn_Primary.Name = "Styles_btn_Primary";
            this.Styles_btn_Primary.ScreenTip = "Primary";
            this.Styles_btn_Primary.ShowImage = true;
            this.Styles_btn_Primary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Primary_Click);
            // 
            // Styles_btn_Secondary
            // 
            this.Styles_btn_Secondary.Image = global::GCScript_for_Excel.Properties.Resources.styles_secondary;
            this.Styles_btn_Secondary.KeyTip = "2";
            this.Styles_btn_Secondary.Label = "Secondary";
            this.Styles_btn_Secondary.Name = "Styles_btn_Secondary";
            this.Styles_btn_Secondary.ScreenTip = "Secondary";
            this.Styles_btn_Secondary.ShowImage = true;
            this.Styles_btn_Secondary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Secondary_Click);
            // 
            // Styles_btn_Success
            // 
            this.Styles_btn_Success.Image = global::GCScript_for_Excel.Properties.Resources.styles_success;
            this.Styles_btn_Success.KeyTip = "3";
            this.Styles_btn_Success.Label = "Success";
            this.Styles_btn_Success.Name = "Styles_btn_Success";
            this.Styles_btn_Success.ScreenTip = "Success";
            this.Styles_btn_Success.ShowImage = true;
            this.Styles_btn_Success.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Success_Click);
            // 
            // Styles_btn_Danger
            // 
            this.Styles_btn_Danger.Image = global::GCScript_for_Excel.Properties.Resources.styles_danger;
            this.Styles_btn_Danger.KeyTip = "4";
            this.Styles_btn_Danger.Label = "Danger";
            this.Styles_btn_Danger.Name = "Styles_btn_Danger";
            this.Styles_btn_Danger.ScreenTip = "Danger";
            this.Styles_btn_Danger.ShowImage = true;
            this.Styles_btn_Danger.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Danger_Click);
            // 
            // Styles_btn_Warning
            // 
            this.Styles_btn_Warning.Image = global::GCScript_for_Excel.Properties.Resources.styles_warning;
            this.Styles_btn_Warning.KeyTip = "5";
            this.Styles_btn_Warning.Label = "Warning";
            this.Styles_btn_Warning.Name = "Styles_btn_Warning";
            this.Styles_btn_Warning.ScreenTip = "Warning";
            this.Styles_btn_Warning.ShowImage = true;
            this.Styles_btn_Warning.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Warning_Click);
            // 
            // Styles_btn_Info
            // 
            this.Styles_btn_Info.Image = global::GCScript_for_Excel.Properties.Resources.styles_info;
            this.Styles_btn_Info.KeyTip = "6";
            this.Styles_btn_Info.Label = "Info";
            this.Styles_btn_Info.Name = "Styles_btn_Info";
            this.Styles_btn_Info.ScreenTip = "Info";
            this.Styles_btn_Info.ShowImage = true;
            this.Styles_btn_Info.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Info_Click);
            // 
            // Styles_glr_Bootstrap
            // 
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Primary);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Secondary);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Success);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Danger);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Warning);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Info);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Light);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_Dark);
            this.Styles_glr_Bootstrap.Buttons.Add(this.Styles_glr_Bootstrap_White);
            this.Styles_glr_Bootstrap.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap.KeyTip = "7";
            this.Styles_glr_Bootstrap.Label = "Bootstrap";
            this.Styles_glr_Bootstrap.Name = "Styles_glr_Bootstrap";
            this.Styles_glr_Bootstrap.ScreenTip = "Bootstrap";
            this.Styles_glr_Bootstrap.ShowImage = true;
            this.Styles_glr_Bootstrap.ShowItemImage = false;
            // 
            // Styles_glr_Bootstrap_Primary
            // 
            this.Styles_glr_Bootstrap_Primary.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Primary.Label = "Primary";
            this.Styles_glr_Bootstrap_Primary.Name = "Styles_glr_Bootstrap_Primary";
            this.Styles_glr_Bootstrap_Primary.ScreenTip = "Primary";
            this.Styles_glr_Bootstrap_Primary.ShowImage = true;
            this.Styles_glr_Bootstrap_Primary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Primary_Click);
            // 
            // Styles_glr_Bootstrap_Secondary
            // 
            this.Styles_glr_Bootstrap_Secondary.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Secondary.Label = "Secondary";
            this.Styles_glr_Bootstrap_Secondary.Name = "Styles_glr_Bootstrap_Secondary";
            this.Styles_glr_Bootstrap_Secondary.ScreenTip = "Secondary";
            this.Styles_glr_Bootstrap_Secondary.ShowImage = true;
            this.Styles_glr_Bootstrap_Secondary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Secondary_Click);
            // 
            // Styles_glr_Bootstrap_Success
            // 
            this.Styles_glr_Bootstrap_Success.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Success.Label = "Success";
            this.Styles_glr_Bootstrap_Success.Name = "Styles_glr_Bootstrap_Success";
            this.Styles_glr_Bootstrap_Success.ScreenTip = "Success";
            this.Styles_glr_Bootstrap_Success.ShowImage = true;
            this.Styles_glr_Bootstrap_Success.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Success_Click);
            // 
            // Styles_glr_Bootstrap_Danger
            // 
            this.Styles_glr_Bootstrap_Danger.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Danger.Label = "Danger";
            this.Styles_glr_Bootstrap_Danger.Name = "Styles_glr_Bootstrap_Danger";
            this.Styles_glr_Bootstrap_Danger.ScreenTip = "Danger";
            this.Styles_glr_Bootstrap_Danger.ShowImage = true;
            this.Styles_glr_Bootstrap_Danger.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Danger_Click);
            // 
            // Styles_glr_Bootstrap_Warning
            // 
            this.Styles_glr_Bootstrap_Warning.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Warning.Label = "Warning";
            this.Styles_glr_Bootstrap_Warning.Name = "Styles_glr_Bootstrap_Warning";
            this.Styles_glr_Bootstrap_Warning.ScreenTip = "Warning";
            this.Styles_glr_Bootstrap_Warning.ShowImage = true;
            this.Styles_glr_Bootstrap_Warning.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Warning_Click);
            // 
            // Styles_glr_Bootstrap_Info
            // 
            this.Styles_glr_Bootstrap_Info.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Info.Label = "Info";
            this.Styles_glr_Bootstrap_Info.Name = "Styles_glr_Bootstrap_Info";
            this.Styles_glr_Bootstrap_Info.ScreenTip = "Info";
            this.Styles_glr_Bootstrap_Info.ShowImage = true;
            this.Styles_glr_Bootstrap_Info.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Info_Click);
            // 
            // Styles_glr_Bootstrap_Light
            // 
            this.Styles_glr_Bootstrap_Light.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Light.Label = "Light";
            this.Styles_glr_Bootstrap_Light.Name = "Styles_glr_Bootstrap_Light";
            this.Styles_glr_Bootstrap_Light.ScreenTip = "Light";
            this.Styles_glr_Bootstrap_Light.ShowImage = true;
            this.Styles_glr_Bootstrap_Light.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Light_Click);
            // 
            // Styles_glr_Bootstrap_Dark
            // 
            this.Styles_glr_Bootstrap_Dark.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_Dark.Label = "Dark";
            this.Styles_glr_Bootstrap_Dark.Name = "Styles_glr_Bootstrap_Dark";
            this.Styles_glr_Bootstrap_Dark.ScreenTip = "Dark";
            this.Styles_glr_Bootstrap_Dark.ShowImage = true;
            this.Styles_glr_Bootstrap_Dark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_Dark_Click);
            // 
            // Styles_glr_Bootstrap_White
            // 
            this.Styles_glr_Bootstrap_White.Image = global::GCScript_for_Excel.Properties.Resources.styles_bootstrap;
            this.Styles_glr_Bootstrap_White.Label = "White";
            this.Styles_glr_Bootstrap_White.Name = "Styles_glr_Bootstrap_White";
            this.Styles_glr_Bootstrap_White.ScreenTip = "White";
            this.Styles_glr_Bootstrap_White.ShowImage = true;
            this.Styles_glr_Bootstrap_White.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Bootstrap_White_Click);
            // 
            // Styles_glr_Emphasis
            // 
            this.Styles_glr_Emphasis.Buttons.Add(this.Styles_glr_Emphasis1);
            this.Styles_glr_Emphasis.Buttons.Add(this.Styles_glr_Emphasis2);
            this.Styles_glr_Emphasis.Buttons.Add(this.Styles_glr_Emphasis3);
            this.Styles_glr_Emphasis.Buttons.Add(this.Styles_glr_Emphasis4);
            this.Styles_glr_Emphasis.Buttons.Add(this.Styles_glr_Emphasis5);
            this.Styles_glr_Emphasis.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis.KeyTip = "8";
            this.Styles_glr_Emphasis.Label = "Emphasis";
            this.Styles_glr_Emphasis.Name = "Styles_glr_Emphasis";
            this.Styles_glr_Emphasis.ScreenTip = "Emphasis";
            this.Styles_glr_Emphasis.ShowImage = true;
            this.Styles_glr_Emphasis.ShowItemImage = false;
            // 
            // Styles_glr_Emphasis1
            // 
            this.Styles_glr_Emphasis1.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis1.Label = "1 (C.UNID)";
            this.Styles_glr_Emphasis1.Name = "Styles_glr_Emphasis1";
            this.Styles_glr_Emphasis1.ScreenTip = "1 (C.UNID)";
            this.Styles_glr_Emphasis1.ShowImage = true;
            this.Styles_glr_Emphasis1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Emphasis1_Click);
            // 
            // Styles_glr_Emphasis2
            // 
            this.Styles_glr_Emphasis2.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis2.Label = "2 (EMPRESA)";
            this.Styles_glr_Emphasis2.Name = "Styles_glr_Emphasis2";
            this.Styles_glr_Emphasis2.ScreenTip = "2 (EMPRESA)";
            this.Styles_glr_Emphasis2.ShowImage = true;
            this.Styles_glr_Emphasis2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Emphasis2_Click);
            // 
            // Styles_glr_Emphasis3
            // 
            this.Styles_glr_Emphasis3.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis3.Label = "3 (OPERADORA)";
            this.Styles_glr_Emphasis3.Name = "Styles_glr_Emphasis3";
            this.Styles_glr_Emphasis3.ScreenTip = "3 (OPERADORA)";
            this.Styles_glr_Emphasis3.ShowImage = true;
            this.Styles_glr_Emphasis3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Emphasis3_Click);
            // 
            // Styles_glr_Emphasis4
            // 
            this.Styles_glr_Emphasis4.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis4.Label = "4 (UF)";
            this.Styles_glr_Emphasis4.Name = "Styles_glr_Emphasis4";
            this.Styles_glr_Emphasis4.ScreenTip = "4 (UF)";
            this.Styles_glr_Emphasis4.ShowImage = true;
            this.Styles_glr_Emphasis4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Emphasis4_Click);
            // 
            // Styles_glr_Emphasis5
            // 
            this.Styles_glr_Emphasis5.Image = global::GCScript_for_Excel.Properties.Resources.styles_emphasis;
            this.Styles_glr_Emphasis5.Label = "5 (TOTAL GERAL)";
            this.Styles_glr_Emphasis5.Name = "Styles_glr_Emphasis5";
            this.Styles_glr_Emphasis5.ScreenTip = "5 (TOTAL GERAL)";
            this.Styles_glr_Emphasis5.ShowImage = true;
            this.Styles_glr_Emphasis5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_glr_Emphasis5_Click);
            // 
            // Styles_btn_Default
            // 
            this.Styles_btn_Default.Image = global::GCScript_for_Excel.Properties.Resources.styles_default;
            this.Styles_btn_Default.KeyTip = "9";
            this.Styles_btn_Default.Label = "Default";
            this.Styles_btn_Default.Name = "Styles_btn_Default";
            this.Styles_btn_Default.ScreenTip = "Default";
            this.Styles_btn_Default.ShowImage = true;
            this.Styles_btn_Default.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Styles_btn_Default_Click);
            // 
            // Number_btn_General
            // 
            this.Number_btn_General.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Number_btn_General.Image = ((System.Drawing.Image)(resources.GetObject("Number_btn_General.Image")));
            this.Number_btn_General.Label = "General";
            this.Number_btn_General.Name = "Number_btn_General";
            this.Number_btn_General.ShowImage = true;
            this.Number_btn_General.SuperTip = "Gustavo";
            this.Number_btn_General.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Number_btn_General_Click);
            // 
            // Number_btn_Text
            // 
            this.Number_btn_Text.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Number_btn_Text.Image = ((System.Drawing.Image)(resources.GetObject("Number_btn_Text.Image")));
            this.Number_btn_Text.Label = "Text";
            this.Number_btn_Text.Name = "Number_btn_Text";
            this.Number_btn_Text.ShowImage = true;
            this.Number_btn_Text.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Number_btn_Text_Click);
            // 
            // Number_btn_Accounting
            // 
            this.Number_btn_Accounting.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Number_btn_Accounting.Image = ((System.Drawing.Image)(resources.GetObject("Number_btn_Accounting.Image")));
            this.Number_btn_Accounting.Label = "Accounting";
            this.Number_btn_Accounting.Name = "Number_btn_Accounting";
            this.Number_btn_Accounting.ShowImage = true;
            this.Number_btn_Accounting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Number_btn_Accounting_Click);
            // 
            // Info_btn_LinhasUsadas
            // 
            this.Info_btn_LinhasUsadas.Image = global::GCScript_for_Excel.Properties.Resources.rows;
            this.Info_btn_LinhasUsadas.Label = "Linhas Usadas";
            this.Info_btn_LinhasUsadas.Name = "Info_btn_LinhasUsadas";
            this.Info_btn_LinhasUsadas.ShowImage = true;
            this.Info_btn_LinhasUsadas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Info_btn_LinhasUsadas_Click);
            // 
            // Info_btn_ColunasUsadas
            // 
            this.Info_btn_ColunasUsadas.Image = global::GCScript_for_Excel.Properties.Resources.columns;
            this.Info_btn_ColunasUsadas.Label = "Colunas Usadas";
            this.Info_btn_ColunasUsadas.Name = "Info_btn_ColunasUsadas";
            this.Info_btn_ColunasUsadas.ShowImage = true;
            this.Info_btn_ColunasUsadas.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Info_btn_ColunasUsadas_Click);
            // 
            // Info_btn_ObterTipoCell
            // 
            this.Info_btn_ObterTipoCell.Image = ((System.Drawing.Image)(resources.GetObject("Info_btn_ObterTipoCell.Image")));
            this.Info_btn_ObterTipoCell.Label = "Tipo da Célula";
            this.Info_btn_ObterTipoCell.Name = "Info_btn_ObterTipoCell";
            this.Info_btn_ObterTipoCell.ShowImage = true;
            this.Info_btn_ObterTipoCell.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Info_btn_ObterTipoSelecao_Click);
            // 
            // Info_btn_SelecionarRange
            // 
            this.Info_btn_SelecionarRange.Image = global::GCScript_for_Excel.Properties.Resources.select_range;
            this.Info_btn_SelecionarRange.Label = "Selecionar Range";
            this.Info_btn_SelecionarRange.Name = "Info_btn_SelecionarRange";
            this.Info_btn_SelecionarRange.ShowImage = true;
            this.Info_btn_SelecionarRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Info_btn_SelecionarRange_Click);
            // 
            // Info_btn_SelecionarTudo
            // 
            this.Info_btn_SelecionarTudo.Image = global::GCScript_for_Excel.Properties.Resources.select_all;
            this.Info_btn_SelecionarTudo.Label = "Selecionar Tudo";
            this.Info_btn_SelecionarTudo.Name = "Info_btn_SelecionarTudo";
            this.Info_btn_SelecionarTudo.ShowImage = true;
            this.Info_btn_SelecionarTudo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Info_btn_SelecionarTudo_Click);
            // 
            // Tools_btn_SetTitle
            // 
            this.Tools_btn_SetTitle.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Tools_btn_SetTitle.Image = global::GCScript_for_Excel.Properties.Resources.title;
            this.Tools_btn_SetTitle.Label = "Title";
            this.Tools_btn_SetTitle.Name = "Tools_btn_SetTitle";
            this.Tools_btn_SetTitle.ScreenTip = "Title";
            this.Tools_btn_SetTitle.ShowImage = true;
            this.Tools_btn_SetTitle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Tools_btn_SetTitle_Click);
            // 
            // btn_Settings
            // 
            this.btn_Settings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_Settings.Image = global::GCScript_for_Excel.Properties.Resources.settings;
            this.btn_Settings.Label = "Configurações";
            this.btn_Settings.Name = "btn_Settings";
            this.btn_Settings.ShowImage = true;
            this.btn_Settings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_Settings_Click);
            // 
            // btn_T1
            // 
            this.btn_T1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T1.Image = ((System.Drawing.Image)(resources.GetObject("btn_T1.Image")));
            this.btn_T1.Label = "T1";
            this.btn_T1.Name = "btn_T1";
            this.btn_T1.ShowImage = true;
            this.btn_T1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T1_Click);
            // 
            // btn_T2
            // 
            this.btn_T2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T2.Image = ((System.Drawing.Image)(resources.GetObject("btn_T2.Image")));
            this.btn_T2.Label = "T2";
            this.btn_T2.Name = "btn_T2";
            this.btn_T2.ShowImage = true;
            this.btn_T2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T2_Click);
            // 
            // btn_T3
            // 
            this.btn_T3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T3.Image = ((System.Drawing.Image)(resources.GetObject("btn_T3.Image")));
            this.btn_T3.Label = "T3";
            this.btn_T3.Name = "btn_T3";
            this.btn_T3.ShowImage = true;
            this.btn_T3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T3_Click);
            // 
            // btn_T4
            // 
            this.btn_T4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T4.Image = ((System.Drawing.Image)(resources.GetObject("btn_T4.Image")));
            this.btn_T4.Label = "T4";
            this.btn_T4.Name = "btn_T4";
            this.btn_T4.ShowImage = true;
            this.btn_T4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T4_Click);
            // 
            // btn_T5
            // 
            this.btn_T5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T5.Image = ((System.Drawing.Image)(resources.GetObject("btn_T5.Image")));
            this.btn_T5.Label = "T5";
            this.btn_T5.Name = "btn_T5";
            this.btn_T5.ShowImage = true;
            this.btn_T5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T5_Click);
            // 
            // btn_T6
            // 
            this.btn_T6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btn_T6.Image = ((System.Drawing.Image)(resources.GetObject("btn_T6.Image")));
            this.btn_T6.Label = "T6";
            this.btn_T6.Name = "btn_T6";
            this.btn_T6.ShowImage = true;
            this.btn_T6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_T6_Click);
            // 
            // btn_RemoverFC
            // 
            this.btn_RemoverFC.Image = ((System.Drawing.Image)(resources.GetObject("btn_RemoverFC.Image")));
            this.btn_RemoverFC.Label = "Remover FC";
            this.btn_RemoverFC.Name = "btn_RemoverFC";
            this.btn_RemoverFC.ScreenTip = "Remove todas as formatações condicional da página";
            this.btn_RemoverFC.ShowImage = true;
            // 
            // rbb_Main
            // 
            this.Name = "rbb_Main";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.rbt_Main);
            this.Tabs.Add(this.rbt_Testes);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.rbb_Main_Load);
            this.rbt_Main.ResumeLayout(false);
            this.rbt_Main.PerformLayout();
            this.grp_Tools.ResumeLayout(false);
            this.grp_Tools.PerformLayout();
            this.grp_Styles.ResumeLayout(false);
            this.grp_Styles.PerformLayout();
            this.grp_Number.ResumeLayout(false);
            this.grp_Number.PerformLayout();
            this.grp_Info.ResumeLayout(false);
            this.grp_Info.PerformLayout();
            this.grp_Others.ResumeLayout(false);
            this.grp_Others.PerformLayout();
            this.rbt_Testes.ResumeLayout(false);
            this.rbt_Testes.PerformLayout();
            this.grp_Tests.ResumeLayout(false);
            this.grp_Tests.PerformLayout();
            this.grp_Beta.ResumeLayout(false);
            this.grp_Beta.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab rbt_Main;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Tools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Tools_btn_BZA;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Info;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Info_btn_ColunasUsadas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Info_btn_LinhasUsadas;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Info_btn_SelecionarRange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Info_btn_SelecionarTudo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Styles;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Default;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery Styles_glr_Emphasis;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Emphasis1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Emphasis2;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Emphasis3;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Emphasis4;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Emphasis5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Tools_btn_SetTitle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Info_btn_ObterTipoCell;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_Generate;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Generate_btn_Apportionment;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Generate_btn_Purchase;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_Converter;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Text;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_CPF;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_CNPJ;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Settings;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Settings;
        private Microsoft.Office.Tools.Ribbon.RibbonTab rbt_Testes;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Tests;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_T6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Beta;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_RemoverFC;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Primary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Secondary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Success;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Danger;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Warning;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Styles_btn_Info;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery Styles_glr_Bootstrap;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Primary;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Secondary;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Success;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Danger;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Warning;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Info;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Light;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_Dark;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Styles_glr_Bootstrap_White;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_More;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Others;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_RemoveConditionalFormatting;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_CheckAllSheets;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_CheckActiveSheet;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_OnlyValues;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_Separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_ApplyRemove;
        private Microsoft.Office.Tools.Ribbon.RibbonButton ApplyRemove_btn_Start;
        private Microsoft.Office.Tools.Ribbon.RibbonButton ApplyRemove_btn_Settings;
        private Microsoft.Office.Tools.Ribbon.RibbonButton More_btn_CheckSelection;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_WorkSchedule;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Number;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Number_btn_General;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Number_btn_Text;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Number_btn_Accounting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_Sort;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sort_btn_Preset02;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Generate_btn_FileToSend;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sort_btn_Preset03;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sort_btn_Preset04;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery glr_Sheets;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sheets_btn_SortSheetsASC;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sheets_btn_SortSheetsDESC;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Separator1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Round_2DecimalPlaces;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Separator2;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Round_0DecimalPlaces;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Converter_btn_Round_1DecimalPlaces;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sheets_btn_RemoveHiddenSheets;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sheets_btn_ShowHiddenSheets;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Generate_btn_Separator1;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Generate_btn_RandomCPF;
        private Microsoft.Office.Tools.Ribbon.RibbonButton Sort_btn_Preset01;
    }

    partial class ThisRibbonCollection
    {
        internal rbb_Main Ribbon1
        {
            get { return this.GetRibbon<rbb_Main>(); }
        }
    }
}
