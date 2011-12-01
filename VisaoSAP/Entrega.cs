using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Entrega
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

        public Entrega(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Etg_TextY", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 25;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Declaração de conformidade: ";

            oNewItem = oForm.Items.Add("Etg_Dec_An", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 17;
            oNewItem.Width = 250;
            oNewItem.Left = 170;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ENT_DECLARA");

            oNewItem = oForm.Items.Add("Etg_Decl", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Width = 180;
            oNewItem.Left = 430;
            oNewItem.Height = 20;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Nova declaração de conformidade";

          ///////////////////////////////////////////

            oNewItem = oForm.Items.Add("Etg_Text2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 25;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Pesquisa de satisfação: ";

            oNewItem = oForm.Items.Add("Etg_Pesqu", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Height = 17;
            oNewItem.Width = 250;
            oNewItem.Left = 170;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ENT_PESQUIS");

            oNewItem = oForm.Items.Add("Etg_Pq", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 85;
            oNewItem.Width = 180;
            oNewItem.Left = 430;
            oNewItem.Height = 20;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Nova pesquisa de satisfação";

            /////////////////////////////////////////////////////
            oNewItem = oForm.Items.Add("Etg_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 105;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Entrega";
            
            oNewItem = oForm.Items.Add("Etg_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 580;
            oNewItem.Top = oItem.Top + 125;
            oNewItem.Height = 120;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Etg_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Etg_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Etg_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Entrega";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Etg_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Responsável";
            oColumn.Width = 120;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Etg_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Laudo de Entrega";
            oColumn.Width = 80;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Etg_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Data p/Solução";
            oColumn.Width = 100;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Etg_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Resolvido";
            oColumn.Width = 80;
            oColumn.Editable = true;
            oColumn.DisplayDesc = true;
            oColumn.ValOn = "1";
            oColumn.ValOff = "0";

            oColumn = oColumns.Add("Etg_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Id Ambiente";
            oColumn.Width = 30;
            oColumn.Editable = true;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Etg_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Etg_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ENT_ENTREGA");

            oColumn = oColumns.Item("Etg_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ENT_RESPONS");

            oColumn = oColumns.Item("Etg_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ENT_LAUDO");

            oColumn = oColumns.Item("Etg_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ENT_DATASOL");

            oColumn = oColumns.Item("Etg_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ENT_RESOLVI");

            oColumn = oColumns.Item("Etg_Amb_C6");
            oColumn.DataBind.Bind("oMatrixDT", "Line");
             
            /////////////////////////////

            //Botao Nova vistoria de entrega
            oNewItem = oForm.Items.Add("Laudo_Ent", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 125;
            oNewItem.Width = 200;
            oNewItem.Left = oForm.Width - 250;
            oNewItem.Height = 25;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Novo laudo de vistoria de entrega";

            oNewItem = oForm.Items.Add("Etg_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 265;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Registro de Pendências";

            oNewItem = oForm.Items.Add("Pen_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 120;
            oNewItem.FromPane = 17;
            oNewItem.ToPane = 17;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Pen_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Pen_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Descrição";
            oColumn.Width = 600;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Pen_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Code";
            oColumn.Width = 50;
            oColumn.Editable = false;
            oColumn.Visible = false;

            oForm.DataSources.DataTables.Add("oDataTablePend");
            oForm.DataSources.DataTables.Item("oDataTablePend").ExecuteQuery("select * from [@FLX_FB_PEN] where U_FLX_FB_PEN_IDOOPR = '1' and U_FLX_FB_PEN_IDAMB = '1'");

            oColumn = oColumns.Item("Pen_Amb_C0");
            oColumn.DataBind.Bind("oDataTablePend", "U_FLX_FB_PEN_DESC");

            oColumn = oColumns.Item("Pen_Amb_C1");
            oColumn.DataBind.Bind("oDataTablePend", "Code");
        }
    }
}
