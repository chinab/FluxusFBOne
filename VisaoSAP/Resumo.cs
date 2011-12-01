using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    public sealed class Resumo
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Folder oFolderItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Column oColumn;
        private SAPbouiCOM.DBDataSource DBDSflx;

        public Resumo(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharAba();
            desenharCampos();
        }

        private void desenharAba()
        {
            oNewItem = oForm.Items.Add("Projeto1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oItem = oForm.Items.Item("7");
            oNewItem.Top = oItem.Top;
            oNewItem.Height = oItem.Height;
            oNewItem.Width = oItem.Width;
            oNewItem.Left = oItem.Left + oItem.Width;
            oFolderItem = ((SAPbouiCOM.Folder)(oNewItem.Specific));
            oFolderItem.Caption = "Móveis (Resumo)";
            oFolderItem.GroupWith("7");
        }

        private void desenharCampos()
        {

            oForm.PaneLevel = 1;
            oForm.Height = 835;
            oForm.Width = 900;

            //Botao atualizar
            oItem = oForm.Items.Item("1");
            oItem.Top = oForm.Height - 60;

            //Botao cancelar
            oItem = oForm.Items.Item("2");
            oItem.Top = oForm.Height - 60;

            //Botao Documentos relacionados
            oItem = oForm.Items.Item("77");
            oItem.Top = oForm.Height - 60;
            oItem.Left = oForm.Width - 260;

            //Botao Atividades relacionadas
            oItem = oForm.Items.Item("78");
            oItem.Top = oForm.Height - 60;
            oItem.Left = oForm.Width - 130;

            oItem = oForm.Items.Item("52");
            oItem.Width = oForm.Width - 25;
            oItem.Top = oForm.Height - 80;

            oItem = oForm.Items.Item("55");
            oItem.Width = oForm.Items.Item("52").Width;
            
            oItem = oForm.Items.Item("54");
            oItem.Left = oForm.Width - 20;
            oItem.Height = oForm.Height - 234;

            oItem = oForm.Items.Item("53");
            oItem.Height = oForm.Height - 234;

            oNewItem = oForm.Items.Add("Grade", SAPbouiCOM.BoFormItemTypes.it_RECTANGLE);
            oNewItem.Top = oForm.Items.Item("55").Top + 30;
            oNewItem.Height = oForm.Items.Item("53").Height - 40;
            oNewItem.Width = 900;
            oNewItem.Left = oForm.Left + 15;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 17;

            oItem = oForm.Items.Item("55");
            int leftFases = oForm.Width - 455;
            int leftPeriodo = oForm.Width - 365;
            int leftPeriodo2 = oForm.Width - 265;
            int leftResponsaveis = oForm.Width - 165;

            //Grid Ambientes
            oNewItem = oForm.Items.Add("Res_Text0", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 30;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = oForm.Left + 25;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambientes Contratados";

            oNewItem = oForm.Items.Add("Res_Amb", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
            oNewItem.Left = oForm.Left + 25;
            oNewItem.Width = 100;
            oNewItem.Top = oItem.Top + 50;
            oNewItem.Height = 150;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;

            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            oColumns = oMatrix.Columns;

            oColumn = oColumns.Add("Res_#", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "#";
            oColumn.Width = 15;
            oColumn.Editable = false;

            oColumn = oColumns.Add("Res_Amb_C0", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Ambiente";
            oColumn.Width = 85;
            oColumn.Editable = false;

            oForm.DataSources.DataTables.Add("oMatrixDT");
            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '1'");

            oColumn = oColumns.Item("Res_Amb_C0");
            oColumn.DataBind.Bind("oMatrixDT", "Descript");

            oMatrix.LoadFromDataSource();
            //Fim Grid Ambientes

            //Coluna Fases

            oItem = oForm.Items.Item("Res_Text0");

            oNewItem = oForm.Items.Add("Res_Text11", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top;
            oNewItem.Height = 19;
            oNewItem.Width = 50;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Fase";

            oNewItem = oForm.Items.Add("Res_Text2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 20;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "1.Entrevista";

            oNewItem = oForm.Items.Add("Res_Text3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 40;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "2.Medição";

            oNewItem = oForm.Items.Add("Res_Text4", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "3.Elaboração";

            oNewItem = oForm.Items.Add("Res_Text5", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 80;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "4.Verificação";

            oNewItem = oForm.Items.Add("Res_Text6", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 100;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "5.Apresentação";

            oNewItem = oForm.Items.Add("Res_Text7", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 120;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "6.Aprovação";

            oNewItem = oForm.Items.Add("Res_Text8", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 140;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "7.Pedido";

            oNewItem = oForm.Items.Add("Res_Text9", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 160;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "8.Detalhamento";

            oNewItem = oForm.Items.Add("Res_TextA", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 180;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "9.Fábrica";

            oNewItem = oForm.Items.Add("Res_TextB", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "10.Montagem";

            oNewItem = oForm.Items.Add("Res_TextC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = leftFases;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "11.Entrega";
            //Fim da coluna Fases

            //Coluna Período
            oNewItem = oForm.Items.Add("Res_Text12", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top;
            oNewItem.Height = 19;
            oNewItem.Width = 50;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Período";

            oNewItem = oForm.Items.Add("Ent_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 20;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ETV_DATE");

            oNewItem = oForm.Items.Add("Med_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 40;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_MED_DATEMED");

            oNewItem = oForm.Items.Add("Ela_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ela_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo2;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ver_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 80;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            //Aba Resumo, Data Apresentação
            oNewItem = oForm.Items.Add("Aps_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 100;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_APS_DATE");

            oNewItem = oForm.Items.Add("Apv_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 120;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ped_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 140;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Det_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 160;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Det_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 160;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo2;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Fab_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 180;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Fab_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 180;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo2;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Mon_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Mon_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo2;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Eng_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = leftPeriodo;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //Fim da coluna período

            //Coluna dos responsáveis
            oNewItem = oForm.Items.Add("Res_Text13", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top;
            oNewItem.Height = 19;
            oNewItem.Width = 100;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Responsável";

            oNewItem = oForm.Items.Add("Ent_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 20;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Med_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 40;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ela_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 60;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ver_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 80;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Aps_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 100;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Apv_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 120;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Ped_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 140;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Det_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 160;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Fab_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 180;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Mon_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Eng_Res", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Enabled = false;
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = leftResponsaveis;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //Fim da coluna dos responsáveis

            //Previsao e Realizacao
            oItem = oForm.Items.Item("Res_Amb");
          
            oNewItem = oForm.Items.Add("Res_TextD", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = oItem.Left;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Previsão:";

            oNewItem = oForm.Items.Add("Res_TextE", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = oItem.Left;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Realização:";

            oNewItem = oForm.Items.Add("Pre_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = oItem.Left + 60;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_PREVINI");

            oNewItem = oForm.Items.Add("Pre_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 200;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = oItem.Left + 160;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_PREVFIM");

            oNewItem = oForm.Items.Add("Rea_Dat1", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = oItem.Left + 60;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_INICIOR");

            oNewItem = oForm.Items.Add("Rea_Dat2", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 220;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = oItem.Left + 160;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_FIMR");
            //Fim da Previsao e Realizacao

            //Endereço
            oNewItem = oForm.Items.Add("Res_TextF", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 250;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 25;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Logradouro";

            DBDSflx = oForm.DataSources.DBDataSources.Add("@FLX_FB_MED");
            oNewItem = oForm.Items.Add("End_Log", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 265;
            oNewItem.Height = 17;
            oNewItem.Width = 300;
            oNewItem.Left = 25;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_LOGR");

            oNewItem = oForm.Items.Add("Res_TextG", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 250;
            oNewItem.Height = 19;
            oNewItem.Width = 50;
            oNewItem.Left = 335;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "No.";

            oNewItem = oForm.Items.Add("End_Num", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 265;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 335;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_NUM");

            oNewItem = oForm.Items.Add("Res_TextH", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 250;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 385;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Complemento";

            oNewItem = oForm.Items.Add("End_Com", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 265;
            oNewItem.Height = 17;
            oNewItem.Width = 150;
            oNewItem.Left = 385;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_COMP");

            oNewItem = oForm.Items.Add("Res_TextI", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 25;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Bairro";

            oNewItem = oForm.Items.Add("End_Bai", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 300;
            oNewItem.Height = 17;
            oNewItem.Width = 160;
            oNewItem.Left = 25;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_BAIRRO");

            oNewItem = oForm.Items.Add("Res_TextJ", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 19;
            oNewItem.Width = 70;
            oNewItem.Left = 205;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Cidade";

            oNewItem = oForm.Items.Add("End_Cid", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 300;
            oNewItem.Height = 17;
            oNewItem.Width = 160;
            oNewItem.Left = 205;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_CIDADE");

            oNewItem = oForm.Items.Add("Res_TextK", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 285;
            oNewItem.Height = 19;
            oNewItem.Width = 20;
            oNewItem.Left = 385;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "UF";

            oNewItem = oForm.Items.Add("End_UF", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 300;
            oNewItem.Height = 17;
            oNewItem.Width = 20;
            oNewItem.Left = 385;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_UF");

            oNewItem = oForm.Items.Add("Res_TextL", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 250;
            oNewItem.Height = 19;
            oNewItem.Width = 150;
            oNewItem.Left = 550;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ponto de Referência";

            oNewItem = oForm.Items.Add("End_Ref", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            oNewItem.Top = oItem.Top + 265;
            oNewItem.Height = 45;
            oNewItem.Width = 360;
            oNewItem.Left = 550;
            oNewItem.FromPane = 8;
            oNewItem.ToPane = 8;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_RES_PONREF");
            //Fim do Endereço
        }

        public void disableCampos()
        {
            oForm.Items.Item("Ent_Dat1").Enabled = false;
            oForm.Items.Item("Med_Dat1").Enabled = false;
            oForm.Items.Item("Ela_Dat1").Enabled = false;
            oForm.Items.Item("Ela_Dat2").Enabled = false;
            oForm.Items.Item("Ver_Dat1").Enabled = false;
            oForm.Items.Item("Aps_Dat1").Enabled = false;
            oForm.Items.Item("Apv_Dat1").Enabled = false;
            oForm.Items.Item("Ped_Dat1").Enabled = false;
            oForm.Items.Item("Det_Dat1").Enabled = false;
            oForm.Items.Item("Det_Dat2").Enabled = false;
            oForm.Items.Item("Fab_Dat1").Enabled = false;
            oForm.Items.Item("Fab_Dat2").Enabled = false;
            oForm.Items.Item("Mon_Dat1").Enabled = false;
            oForm.Items.Item("Mon_Dat2").Enabled = false;
            oForm.Items.Item("Eng_Dat1").Enabled = false;
            oForm.Items.Item("Ent_Res").Enabled = false;
            oForm.Items.Item("Med_Res").Enabled = false;
            oForm.Items.Item("Ela_Res").Enabled = false;
            oForm.Items.Item("Ver_Res").Enabled = false;
            oForm.Items.Item("Aps_Res").Enabled = false;
            oForm.Items.Item("Apv_Res").Enabled = false;
            oForm.Items.Item("Ped_Res").Enabled = false;
            oForm.Items.Item("Det_Res").Enabled = false;
            oForm.Items.Item("Fab_Res").Enabled = false;
            oForm.Items.Item("Mon_Res").Enabled = false;
            oForm.Items.Item("Eng_Res").Enabled = false;
        }
    }
}
