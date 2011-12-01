using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Fabrica
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;

        public Fabrica(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("Grade");

            oNewItem = oForm.Items.Add("Fab_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 30;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 15;
            oNewItem.ToPane = 15;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Recebimento dos Ambientes";

            oNewItem = oForm.Items.Add("Fab_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 185;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 15;
            oNewItem.ToPane = 15;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Registro de Avarias";

            oNewItem = oForm.Items.Add("Fab_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 480;
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Height = 120;
            oNewItem.FromPane = 15;
            oNewItem.ToPane = 15;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Fab_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Fab_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            /*oColumn = oColumns.Add("Fab_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Expedição";
            oColumn.Width = 80;
            oColumn.Editable = true;*/

            oColumn = oColumns.Add("Fab_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Expedição";
            oColumn.Width = 80;
            oColumn.Editable = true;

            /*oColumn = oColumns.Add("Fab_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Recebimento";
            oColumn.Width = 80;
            oColumn.Editable = true;*/

            oColumn = oColumns.Add("Fab_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Recebimento";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Fab_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Conferente";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Fab_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "IdAmbiente";
            oColumn.Width = 80;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(
                "SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 " +
                "INNER JOIN OOIN T1 ON T1.Num = T0.IntId " +
                "WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Fab_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Fab_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_FAB_EXPEDIC");

            oColumn = oColumns.Item("Fab_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_FAB_RECEBIM");

            oColumn = oColumns.Item("Fab_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_FAB_CONFERE");

            oColumn = oColumns.Item("Fab_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "Line");


            oNewItem = oForm.Items.Add("Ava_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 680;
            oNewItem.Top = oItem.Top + 205;
            oNewItem.Height = 120;
            oNewItem.FromPane = 15;
            oNewItem.ToPane = 15;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;            
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Ava_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ava_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descrição";
            oColumn.Width = 600;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ava_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "id";
            oColumn.Width = 600;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Add("oDataTableAvr");
            oForm.DataSources.DataTables.Item("oDataTableAvr").ExecuteQuery("select * from [@FLX_FB_AVR] where U_FLX_FB_AVR_IDOOPR = '1' and U_FLX_FB_AVR_IDAMBI = '1'");

            oColumn = oColumns.Item("Ava_Amb_C0");
            oColumn.DataBind.Bind("oDataTableAvr", "U_FLX_FB_AVR_DESC");

            oColumn = oColumns.Item("Ava_Amb_C1");
            oColumn.DataBind.Bind("oDataTableAvr", "Code");

        }
    }
}
