using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Elaboracao
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;
        private SAPbouiCOM.Button oButton;

        public Elaboracao(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("Grade");

            oNewItem = oForm.Items.Add("Ela_Cot", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 30;
            oNewItem.Width = 150;
            oNewItem.Left = 730;
            oNewItem.Height = 19;
            oNewItem.FromPane = 11;
            oNewItem.ToPane = 11;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Gerar Orçamento...";

            oNewItem = oForm.Items.Add("Ela_Text1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 30;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 11;
            oNewItem.ToPane = 11;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Elaboração do projeto";

            oNewItem = oForm.Items.Add("Ela_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 185;
            oNewItem.Height = 19;
            oNewItem.Width = 350;
            oNewItem.Left = 25;
            oNewItem.FromPane = 11;
            oNewItem.ToPane = 11;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente / Verificação do projeto";

            oNewItem = oForm.Items.Add("Ela_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = 25;
            oNewItem.Width = 880;
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Height = 120;
            oNewItem.FromPane = 11;
            oNewItem.ToPane = 11;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Ela_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 30;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Ela_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 80;
            oColumn.Editable = false;

            /*oColumn = oColumns.Add("Ela_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX);
            oColumn.TitleObject.Caption = "Revisão";
            oColumn.Width = 60;
            oColumn.Editable = true;*/

            oColumn = oColumns.Add("Ela_Amb_C1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Revisão";
            oColumn.Width = 60;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ela_Amb_C2", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oColumn.TitleObject.Caption = "Projetista";
            oColumn.Width = 100;
            oColumn.Editable = true;
            oColumn.DisplayDesc = true;


            oColumn = oColumns.Add("Ela_Amb_C3", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Início Prev";
            oColumn.Width = 90;
            oColumn.Editable = true;

            oColumn = oColumns.Add("Ela_Amb_C4", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Término Prev";
            oColumn.Width = 90;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oColumn = oColumns.Add("Ela_Amb_C5", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Início Realiz";
            oColumn.Width = 90;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oColumn = oColumns.Add("Ela_Amb_C6", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Término Realiz";
            oColumn.Width = 90;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oColumn = oColumns.Add("Ela_Amb_C7", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Arquivos CAD";
            oColumn.Width = 120;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oColumn = oColumns.Add("Ela_Amb_C8", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Arquivos PRJ";
            oColumn.Width = 120;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oColumn = oColumns.Add("Ela_Amb_C9", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Arquivos JPG";
            oColumn.Width = 120;
            oColumn.Editable = true;
            //oColumn.DataBind.SetBound(true, "OPR4", "U_FLX_FB_VRF_OBS");

            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Ela_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oColumn = oColumns.Item("Ela_Amb_C1");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_REVISAO");

            oColumn = oColumns.Item("Ela_Amb_C2");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_PROJETI");

            oColumn = oColumns.Item("Ela_Amb_C3");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_INICIOP");

            oColumn = oColumns.Item("Ela_Amb_C4");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_TERMINP");

            oColumn = oColumns.Item("Ela_Amb_C5");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_INICIOR");

            oColumn = oColumns.Item("Ela_Amb_C6");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_TERMINR");

            oColumn = oColumns.Item("Ela_Amb_C7");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_ARQCAD");

            oColumn = oColumns.Item("Ela_Amb_C8");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_ARQPRJ");

            oColumn = oColumns.Item("Ela_Amb_C9");
            oColumn.DataBind.Bind("oMatrixDT", "U_FLX_FB_ELB_ARQJPG");
        }
    }
}
