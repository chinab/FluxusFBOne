using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Pedido
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;

        public Pedido(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Ped_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Height = 120;
            oNewItem.FromPane = 13;
            oNewItem.ToPane = 13;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Ped_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ped_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ped_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Data Pedido";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "No.Pedido";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ordem de Compra";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Valor";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Solicitante";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Prazo Entrega";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C7", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Anexo (pedido impresso)";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ped_Amb_C8", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "URL";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Ped_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Ped_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_DATE");

            oColumn = oColumns.Item("Ped_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_NPEDIDO");

            oColumn = oColumns.Item("Ped_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_ORDENDE");

            oColumn = oColumns.Item("Ped_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_VALOR");

            oColumn = oColumns.Item("Ped_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_SOLICIT");

            oColumn = oColumns.Item("Ped_Amb_C6");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_PRAZOEN");

            oColumn = oColumns.Item("Ped_Amb_C7");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_ANEXOPE");

            oColumn = oColumns.Item("Ped_Amb_C8");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_PED_URL");
        }
    }
}
