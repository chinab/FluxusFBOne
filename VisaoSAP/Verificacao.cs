using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Verificacao
    {
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;

        public Verificacao(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Ver_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 205;
            oNewItem.Height = 120;
            oNewItem.FromPane = 11;
            oNewItem.ToPane = 11;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Ver_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ver_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ver_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Data Verificação";
            oColumn.Width = 80;
            oColumn.Editable = true;

            /*oColumn = oColumns.Add("Ver_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Verificado por";
            oColumn.Width = 100;
            oColumn.Editable = true;
            oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_ENT_PENDENC");*/

            oColumn = oColumns.Add("Ver_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Verificado por";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ver_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Observações";
            oColumn.Width = 470;
            oColumn.Editable = true;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Ver_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Ver_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_VRF_DATEVER");

            oColumn = oColumns.Item("Ver_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_VRF_VERIFPO");

            oColumn = oColumns.Item("Ver_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_VRF_OBS");

        }
    }
}
