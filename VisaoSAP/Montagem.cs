using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Montagem
    {   
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;
        private SAPbouiCOM.Button oButton;

        public Montagem(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Laudo_Text", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Laudo de vistoria inicial: ";

            oNewItem = oForm.Items.Add("Ini_An", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 17;
            oNewItem.Width = 250;
            oNewItem.Left = 170;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_MTG_INICIAL");

            oNewItem = oForm.Items.Add("Laudo_Ini", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Width = 200;
            oNewItem.Left = 430;
            oNewItem.Height = 20;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Novo laudo de vistoria inicial";

            oNewItem = oForm.Items.Add("Mon_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 80;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Pendências";

            oNewItem = oForm.Items.Add("Mon_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 740;
            oNewItem.Top = oItem.Top + 100;
            oNewItem.Height = 120;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Mon_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Mon_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Mon_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Responsável";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Mon_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descrição";
            oColumn.Width = 320;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Mon_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Anexo Vistoria Int. 1";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Mon_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Anexo Vistoria Int. 2";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Mon_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Anexo Vistoria Int. 3";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Mon_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Mon_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MTG_RESPONS");

            oColumn = oColumns.Item("Mon_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MTG_DESCRIC");

            oColumn = oColumns.Item("Mon_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MTG_VSTINT1");

            oColumn = oColumns.Item("Mon_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MTG_VSTINT2");

            oColumn = oColumns.Item("Mon_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MTG_VSTINT3");

            //Botao Nova vistoria intermediaria
            oNewItem = oForm.Items.Add("Laudo_Int", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 230;
            oNewItem.Width = 200;
            oNewItem.Left = 25;
            oNewItem.Height = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Novo laudo de vistoria intermediária";

            oNewItem = oForm.Items.Add("Mon_OS", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Width = 100;
            oNewItem.Left = 780;
            oNewItem.Height = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Gerar OS";

            oNewItem = oForm.Items.Add("Mon_Text2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 80;
            oNewItem.Height = 19;
            oNewItem.Width = 60;
            oNewItem.Left = 810;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "No. OS";

            oNewItem = oForm.Items.Add("Mon_NoOS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 100;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 810;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Enabled = false;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Mon_Planej", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 150;
            oNewItem.Width = 100;
            oNewItem.Left = 780;
            oNewItem.Height = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Planejamento";

            oNewItem = oForm.Items.Add("Mon_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Itens Complementares";

            oNewItem = oForm.Items.Add("Mon_Itc", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 305;
            oNewItem.Height = 120;
            oNewItem.FromPane = 16;
            oNewItem.ToPane = 16;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Itc_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item";
            oColumn.Width = 40;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descrição";
            oColumn.Width = 200;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C8", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Estoque";
            oColumn.Width = 40;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Qtd";
            oColumn.Width = 40;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Fornecedor";
            oColumn.Width = 150;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Observação";
            oColumn.Width = 200;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Itc_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Prz Entrega";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Itc_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Solicitante";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Itc_Amb_C7", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Recebido";
            oColumn.Width = 80;
            oColumn.Editable = true;
            oColumn.DisplayDesc = true;
            oColumn.ValOn = "1";
            oColumn.ValOff = "0";

            oColumn = oColumns.Add("Itc_Amb_C9", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "idFornecedor";
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oColumns.Add("Itc_Amb_10", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "idItensComp";
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Item("oDataTableItc").ExecuteQuery("select T1.ItemCode, T1.ItemName, T1.OnHand, T0.U_FLX_FB_ITC_QTD, T2.CardCode, T2.CardName, T0.U_FLX_FB_ITC_OBS, T0.DocEntry, T0.U_FLX_FB_ITC_PRZETG, T0.U_FLX_FB_ITC_SOLICI, T0.U_FLX_FB_ITC_RECEB from [@FLX_FB_ITC] T0 inner join OITM T1 on T1.ItemCode = T0.U_FLX_FB_ITC_IDOITM inner join OCRD T2 on T2.CardCode = T0.U_FLX_FB_ITC_IDOCRD where T0.U_FLX_FB_ITC_IDOOPR = '1'");

            oColumn = oColumns.Item("Itc_Amb_C0");
            oColumn.DataBind.Bind("oDataTableItc", "ItemCode");

            oColumn = oColumns.Item("Itc_Amb_C1");
            oColumn.DataBind.Bind("oDataTableItc", "ItemName");

            oColumn = oColumns.Item("Itc_Amb_C8");
            oColumn.DataBind.Bind("oDataTableItc", "OnHand");

            oColumn = oColumns.Item("Itc_Amb_C2");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_QTD");

            oColumn = oColumns.Item("Itc_Amb_C9");
            oColumn.DataBind.Bind("oDataTableItc", "CardCode");

            oColumn = oColumns.Item("Itc_Amb_C3");
            oColumn.DataBind.Bind("oDataTableItc", "CardName");

            oColumn = oColumns.Item("Itc_Amb_C4");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_OBS");

            oColumn = oColumns.Item("Itc_Amb_10");
            oColumn.DataBind.Bind("oDataTableItc", "DocEntry");

            oColumn = oColumns.Item("Itc_Amb_C5");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_PRZETG");

            oColumn = oColumns.Item("Itc_Amb_C6");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_SOLICI");

            oColumn = oColumns.Item("Itc_Amb_C7");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_RECEB");
        }
    }
}
