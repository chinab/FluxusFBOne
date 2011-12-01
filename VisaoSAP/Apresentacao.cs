using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Apresentacao
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;
        private SAPbouiCOM.ComboBox oComboItem;
        private SAPbouiCOM.Button oButton;
        private SAPbouiCOM.LinkedButton oLink;
        private SAPbouiCOM.DBDataSource DBDSflx;

        public Apresentacao(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("Grade");

            oNewItem = oForm.Items.Add("Apr_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 65;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Análise Crítica";

            oNewItem = oForm.Items.Add("Apr_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Aprovação";

            oNewItem = oForm.Items.Add("Apr_Text2a", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem = oForm.Items.Item("Grade");
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Data";

            oNewItem = oForm.Items.Add("Apr_Text2b", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 120;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Hora";

            oNewItem = oForm.Items.Add("Apr_Text2c", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem = oForm.Items.Item("Grade");
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 175;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Projetista";

            oNewItem = oForm.Items.Add("Apr_Data", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_APS_DATE");

            oNewItem = oForm.Items.Add("Apr_Hora", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 120;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            //oEditItem.DataBind.SetBound(true, "", "EditSource"); 
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_APS_HORA");

            oNewItem = oForm.Items.Add("Apr_Proj", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 140;
            oNewItem.Left = 175;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.DisplayDesc = true;
            //oComboItem.DataBind.SetBound(true, "", "CombSource"); 
            oComboItem = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));
            //oComboItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_APS_PROJT");
            //LoadResponsavelComboVals(oComboItem);

            oNewItem = oForm.Items.Add("Apr_Ped", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 235;
            oNewItem.Width = 100;
            oNewItem.Left = 725;
            oNewItem.Height = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Pedido (fechamento)";

            oNewItem = oForm.Items.Add("Apr_Text3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 19;
            oNewItem.Width = 50;
            oNewItem.Left = 840;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Pedido";

            oNewItem = oForm.Items.Add("Apr_Pedido", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 240;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 840;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Enabled = false;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Apr_LinkPd", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
            oNewItem.Top = oItem.Top + 240;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 810;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oNewItem.Enabled = true;
            oLink = ((SAPbouiCOM.LinkedButton)(oNewItem.Specific));
            //oLink.LinkedObject = "Apr_Pedido";

            oNewItem = oForm.Items.Add("Apr_Ctr", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 300;
            oNewItem.Width = 100;
            oNewItem.Left = 725;
            oNewItem.Height = 25;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Gerar Contrato(s)";

            oNewItem = oForm.Items.Add("Apr_Text4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 19;
            oNewItem.Width = 55;
            oNewItem.Left = 840;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Contrato(s)";

            oNewItem = oForm.Items.Add("Apr_Contr", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 305;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 840;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));



            oNewItem = oForm.Items.Add("Apr_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 420;
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Height = 110;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Apr_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Apr_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Detalhamento";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Apr_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Id";
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Apr_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");
            oColumn = oColumns.Item("Apr_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ANC_DETALHA");
            oColumn = oColumns.Item("Apr_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "Line");

            oNewItem = oForm.Items.Add("NvAnalise", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Width = 140;
            oNewItem.Left = oForm.Width - 270;
            oNewItem.Height = 21;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Nova análise crítica";


            oNewItem = oForm.Items.Add("Ans_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            SAPbouiCOM.Item teste = oForm.Items.Item("Apr_Amb");
            oNewItem.Left = oForm.Width - 270;
            oNewItem.Width = 200;
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Height = 110;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ans_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Analise Crítica";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ans_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Id";
            oColumn.Editable = true;
            oColumn.Visible = false;

            DBDSflx = oForm.DataSources.DBDataSources.Add("@FLX_FB_CONFMED");
            oForm.DataSources.DataTables.Add("oDataTableAnalise");
            oForm.DataSources.DataTables.Item("oDataTableAnalise").ExecuteQuery("SELECT * FROM [@FLX_FB_ANLCRI] where U_FLX_FB_ANLCRI_ID = '1' and U_FLX_FB_ANLCRI_AMBI = '1'");

            oColumn = oColumns.Item("Ans_Amb_C0");
            oColumn.DataBind.Bind("oDataTableAnalise", "U_FLX_FB_ANLCRI_ANEX");

            oColumn = oColumns.Item("Ans_Amb_C1");
            oColumn.DataBind.Bind("oDataTableAnalise", "Code");









            oNewItem = oForm.Items.Add("Apv_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 700;
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 110;
            oNewItem.FromPane = 12;
            oNewItem.ToPane = 12;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Apv_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Apv_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Aprovado por";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Apv_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Data Aprovação";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Apv_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "PDFs visto com o cliente";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Apv_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Prancha de Imagem";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Apv_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Memorial Descritivo";
            oColumn.Width = 130 ;
            oColumn.Editable = true;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Apv_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Apv_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_APR_APROVAD");

            oColumn = oColumns.Item("Apv_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_APR_DATAAPR");

            oColumn = oColumns.Item("Apv_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_APR_PDFCLIE");

            oColumn = oColumns.Item("Apv_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_APR_PRANIMG");

            oColumn = oColumns.Item("Apv_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_APR_MEMDESC");
        }
    }
}
