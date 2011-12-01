using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Medicao
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;
        private SAPbouiCOM.DBDataSource DBDSflx;
        private SAPbouiCOM.ComboBox oComboItem;
        private SAPbouiCOM.Button oButton;

        public Medicao(SAPbouiCOM.Form oForm)
        {
           this.oForm = oForm;
           desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Med_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 110;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Levantamento";

            oNewItem = oForm.Items.Add("Med_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 270;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Conferências de Medições";

            oNewItem = oForm.Items.Add("Med_Text2a", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem = oForm.Items.Item("Grade");
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 25;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Data";

            oNewItem = oForm.Items.Add("Med_Text2b", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 120;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Hora";

            oNewItem = oForm.Items.Add("Med_Text2c", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 17;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 175;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Projetista";

            // Comentado campos de ligação com a tabela a partir daqui
                DBDSflx = oForm.DataSources.DBDataSources.Add("@FLX_FB_MED");

            oNewItem = oForm.Items.Add("Med_Data", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = 25;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_MED_DATEMED");

            oNewItem = oForm.Items.Add("Med_Hora", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 120;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_MED_HORAMED");

            oNewItem = oForm.Items.Add("Med_Proj", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Height = 17;
            oNewItem.Width = 140;
            oNewItem.Left = 175;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.DisplayDesc = true;
            oComboItem = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));


            DBDSflx = oForm.DataSources.DBDataSources.Add("@FLX_FB_CONFMED");
            oNewItem = oForm.Items.Add("Med_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 500;
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Height = 110;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;

            //SBO_Application.MessageBox("PASSO 14", 1, "Ok", "", "");


            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Med_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 60;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Med_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Levantamento (anexos)";
            oColumn.Width = 140;
            oColumn.Editable = true;
            oColumn.DataBind.SetBound(true, "@FLX_FB_CONFMED", "U_FLX_FB_CONFMED_PRJ");

            oColumn = oColumns.Add("Med_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "IdAmbiente";
            oColumn.Width = 80;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery(
                "SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 " +
                "INNER JOIN OOIN T1 ON T1.Num = T0.IntId " +
                "WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Med_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");
            oColumn = oColumns.Item("Med_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_MED_LEVANTA");
            oColumn = oColumns.Item("Med_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "Line");

            oMatrix.LoadFromDataSource();

            oNewItem = oForm.Items.Add("NvLev", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 210;
            oNewItem.Width = 180;
            oNewItem.Left = 400;
            oNewItem.Height = 19;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Novo levantamento";


            oNewItem = oForm.Items.Add("Med_Cnf", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 300;
            oNewItem.Top = oItem.Top + 250;
            oNewItem.Height = 110;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Med_Cnf_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Data";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("med_Cnf_C1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Conferente";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Med_Cnf_C2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Id";
            oColumn.Width = 100;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oNewItem = oForm.Items.Add("Med_Age", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 35;
            oNewItem.Width = 80;
            oNewItem.Left = 430;
            oNewItem.Height = 19;
            oNewItem.FromPane = 10;
            oNewItem.ToPane = 10;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Agendar";

            oForm.DataSources.DataTables.Add("oDataTable");
            oForm.DataSources.DataTables.Item("oDataTable").ExecuteQuery("select * from [@FLX_FB_CONFMED] where U_FLX_FB_CONFMED_ID = '1'");

            oColumn = oColumns.Item("Med_Cnf_C0");
            oColumn.DataBind.Bind("oDataTable", "U_FLX_FB_CONFMED_DAT");
            oColumn = oColumns.Item("med_Cnf_C1");
            oColumn.DataBind.Bind("oDataTable", "U_FLX_FB_CONFMED_PRJ");
            oColumn = oColumns.Item("Med_Cnf_C2");
            oColumn.DataBind.Bind("oDataTable", "Code");
        }
    }
}
