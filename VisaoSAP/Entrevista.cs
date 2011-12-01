using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    class Entrevista
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.StaticText oTextItem;
        private SAPbouiCOM.DBDataSource DBDSflx;
        private SAPbouiCOM.Button oButton;
        private SAPbouiCOM.ComboBox oComboItem;

        public Entrevista(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharCampos();
        }

        private void desenharCampos()
        {
            oItem = oForm.Items.Item("Grade");

            oNewItem = oForm.Items.Add("Ent_Text1a", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 27;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 25;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Data";

            oNewItem = oForm.Items.Add("Ent_Text1b", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 27;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 120;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Hora";

            oNewItem = oForm.Items.Add("Ent_Text1c", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 27;
            oNewItem.Height = 19;
            oNewItem.Width = 175;
            oNewItem.Left = 175;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Responsável";

            oNewItem = oForm.Items.Add("Ent_Text1d", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 27;
            oNewItem.Height = 19;
            oNewItem.Width = 400;
            oNewItem.Left = 330;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Prev. Apresent.";

            oNewItem = oForm.Items.Add("Ent_Data", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 45;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = 25;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //oEditItem.DataBind.SetBound(true, "OOPR", "U_DataEntrevista");
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ETV_DATE");

            oNewItem = oForm.Items.Add("Ent_Hora", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 45;
            oNewItem.Height = 17;
            oNewItem.Width = 40;
            oNewItem.Left = 120;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //oEditItem.DataBind.SetBound(true, "OOPR", "U_HoraEntrevista");
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ETV_HORA");

            //SBO_Application.MessageBox("PASSO 11", 1, "Ok", "", "");

            DBDSflx = oForm.DataSources.DBDataSources.Add("@FLX_FB_PRJ");
            oNewItem = oForm.Items.Add("Ent_Proj", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oNewItem.Top = oItem.Top + 45;
            oNewItem.Height = 17;
            oNewItem.Width = 140;
            oNewItem.Left = 175;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.DisplayDesc = true;
            oComboItem = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));
            //LoadResponsavelComboVals(oComboItem);

            oNewItem = oForm.Items.Add("Ent_Prev", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oNewItem.Top = oItem.Top + 45;
            oNewItem.Height = 17;
            oNewItem.Width = 80;
            oNewItem.Left = 330;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            oEditItem.DataBind.SetBound(true, "OOPR", "U_FLX_FB_ETV_PREVAPR");

            oNewItem = oForm.Items.Add("Text2", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 105;
            oNewItem.Height = 17;
            oNewItem.Width = 120;
            oNewItem.Left = 25;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "Ambiente";

            oNewItem = oForm.Items.Add("Ent_Amb", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
            oNewItem.Top = oItem.Top + 125;
            oNewItem.Height = 17;
            oNewItem.Width = 140;
            oNewItem.Left = 25;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.DisplayDesc = true;
            //oComboItem.DataBind.SetBound(true, "", "CombSource"); 
            oComboItem = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));

            oNewItem = oForm.Items.Add("Text3", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oNewItem.Top = oItem.Top + 105;
            oNewItem.Height = 17;
            oNewItem.Width = 600;
            oNewItem.Left = 300;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oTextItem = ((SAPbouiCOM.StaticText)(oNewItem.Specific));
            oTextItem.Caption = "DESCRIÇÃO DO AMBIENTE:   Móveis Planejados (novos)";

            oNewItem = oForm.Items.Add("Ent_Det", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
            oNewItem.Top = oItem.Top + 125;
            oNewItem.Height = 200;
            oNewItem.Width = 600;
            oNewItem.Left = 300;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            //oEditItem.DataBind.SetBound(true, "", "EditSource"); 
            oEditItem = ((SAPbouiCOM.EditText)(oNewItem.Specific));
            //oEditItem.DataBind.SetBound(true, "OPR4", "U_FLX_FB_ETV_PREVAPR");

            oNewItem = oForm.Items.Add("Ent_Age", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 43;
            oNewItem.Width = 80;
            oNewItem.Left = 430;
            oNewItem.Height = 19;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Agendar";

            oNewItem = oForm.Items.Add("Ent_Imp", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oItem.Top + 327;
            oNewItem.Width = 65;
            oNewItem.Left = oForm.Width-100;
            oNewItem.Height = 19;
            oNewItem.FromPane = 9;
            oNewItem.ToPane = 9;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Imprimir";
        
        }
    }
}
