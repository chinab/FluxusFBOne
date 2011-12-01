using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Detalhamento
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;

        public Detalhamento(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("Grade");

            oNewItem = oForm.Items.Add("Det_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 30;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 14;
            oNewItem.ToPane = 14;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Detalhamento do Projeto";

            oNewItem = oForm.Items.Add("Det_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 185;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 14;
            oNewItem.ToPane = 14;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Itens Complementares";

            oNewItem = oForm.Items.Add("Det_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 860;
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Height = 120;
            oNewItem.FromPane = 14;
            oNewItem.ToPane = 14;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Det_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Det_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Det_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Projetista";
            oColumn.Width = 120;
            oColumn.Editable = true;
            oColumn.DisplayDesc = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_ENT_PENDENC");

            oColumn = oColumns.Add("Det_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Início Previsto";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Det_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Término Previsto";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Det_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Início Realizado";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Det_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Término Realzado";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Det_Amb_C7", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "PDFs Detalhamento";
            oColumn.Width = 200;
            oColumn.Editable = true;
            oColumn.DataBind.SetBound(true, "@FLX_FB_CONFMED", "U_FLX_FB_CONFMED_PRJ");

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Det_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Det_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_PROJETI");

            oColumn = oColumns.Item("Det_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_INICIOP");

            oColumn = oColumns.Item("Det_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_TERMINP");

            oColumn = oColumns.Item("Det_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_INICIRE");

            oColumn = oColumns.Item("Det_Amb_C6");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_TERMINO");

            oColumn = oColumns.Item("Det_Amb_C7");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_DET_PDF");



            oNewItem = oForm.Items.Add("Det_Cmp", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 205;
            oNewItem.Height = 120;
            oNewItem.FromPane = 14;
            oNewItem.ToPane = 14;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Cmp_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Cmp_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item";
            oColumn.Width = 40;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Cmp_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descrição";
            oColumn.Width = 200;
            oColumn.Editable = true;
            oColumn.DataBind.SetBound(true, "", "dt");
            oColumn.ChooseFromListUID = "CFL1";
            oColumn.ChooseFromListAlias = "ItemName";

            oColumn = oColumns.Add("Cmp_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Estoque";
            oColumn.Width = 70;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Cmp_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Qtd";
            oColumn.Width = 40;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Cmp_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "IdFornecedor";
            oColumn.Width = 40;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oColumns.Add("Cmp_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Fornecedor";
            oColumn.Width = 150;
            oColumn.Editable = true;
            oColumn.DataBind.SetBound(true, "", "dt");
            oColumn.ChooseFromListUID = "CFL2";
            oColumn.ChooseFromListAlias = "CardName";

            oColumn = oColumns.Add("Cmp_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Observação";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Cmp_Amb_C7", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "id";
            oColumn.Width = 120;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Add("oDataTableItc");
            oForm.DataSources.DataTables.Item("oDataTableItc").ExecuteQuery("select T1.ItemCode, T1.ItemName, T1.OnHand, T0.U_FLX_FB_ITC_QTD, T2.CardCode, T2.CardName, T0.U_FLX_FB_ITC_OBS, T0.Code from [@FLX_FB_ITC] T0 inner join OITM T1 on T1.ItemCode = T0.U_FLX_FB_ITC_IDOITM inner join OCRD T2 on T2.CardCode = T0.U_FLX_FB_ITC_IDOCRD where T0.U_FLX_FB_ITC_IDOOPR = '1'");

            oColumn = oColumns.Item("Cmp_Amb_C0");
            oColumn.DataBind.Bind("oDataTableItc", "ItemCode");
            oColumn = oColumns.Item("Cmp_Amb_C1");
            oColumn.DataBind.Bind("oDataTableItc", "ItemName");
            oColumn = oColumns.Item("Cmp_Amb_C4");
            oColumn.DataBind.Bind("oDataTableItc", "OnHand");
            oColumn = oColumns.Item("Cmp_Amb_C2");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_QTD");
            oColumn = oColumns.Item("Cmp_Amb_C6");
            oColumn.DataBind.Bind("oDataTableItc", "CardCode");
            oColumn = oColumns.Item("Cmp_Amb_C3");
            oColumn.DataBind.Bind("oDataTableItc", "CardName");
            oColumn = oColumns.Item("Cmp_Amb_C5");
            oColumn.DataBind.Bind("oDataTableItc", "U_FLX_FB_ITC_OBS");
            oColumn = oColumns.Item("Cmp_Amb_C7");
            oColumn.DataBind.Bind("oDataTableItc", "Code");
        }
    }
}
