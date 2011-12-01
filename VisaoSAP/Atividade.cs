using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    public class Atividade
    {
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Button oButton;

        public Atividade(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharBotao();
        }

        private void desenharBotao(){
            oNewItem = oForm.Items.Add("Ata_Ativ", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oNewItem.Top = oForm.Height - 30;
            oNewItem.Width = 100;
            oNewItem.Left = oForm.Width - 125;
            oNewItem.Height = 20;
            oNewItem.Visible = true;
            oButton = ((SAPbouiCOM.Button)(oNewItem.Specific));
            oButton.Caption = "Ata de reunião";

            desabilitaBotaoAta();
        }

        public void desabilitaBotaoAta()
        {
            oNewItem = oForm.Items.Item("Ata_Ativ");
            oNewItem.Enabled = false;
        }
        public void habilitaBotaoAta()
        {
            oNewItem = oForm.Items.Item("Ata_Ativ");
            oNewItem.Enabled = true;
        }

    }
}
