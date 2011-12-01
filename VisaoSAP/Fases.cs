using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VisaoSAP
{
    public class Fases
    {
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Folder oFolderItem;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Folder oFolderInicial;

        public Fases(SAPbouiCOM.Form oForm)
        {
            this.oForm = oForm;
            desenharAba();
            desenharConteudo();
        }

        private void desenharAba()
        {
            oItem = oForm.Items.Item("7");

            oNewItem = oForm.Items.Add("Projeto2", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
            oNewItem.Top = oItem.Top;
            oNewItem.Height = oItem.Height;
            oNewItem.Width = oItem.Width;
            oNewItem.Left = oItem.Left + oItem.Width;
            oFolderItem = ((SAPbouiCOM.Folder)(oNewItem.Specific));
            oFolderItem.Caption = "Móveis (Fases)";
            oFolderItem.GroupWith("7");

        }

        private void desenharConteudo()
        {
            for (int i = 1; i <= 9; i++)
            {
                oNewItem = oForm.Items.Add("Folder" + i, SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oNewItem.Top = oForm.Items.Item("55").Top + 10;
                oNewItem.Height = 20;
                oNewItem.Width = 100;
                oNewItem.Left = 15 + ((i - 1) * 100);
                oNewItem.FromPane = 9;
                oNewItem.ToPane = 17;
                oNewItem.Visible = true;
                oFolderItem = ((SAPbouiCOM.Folder)(oNewItem.Specific));
                if (i == 1) { oFolderItem.Caption = "Entrevista"; }
                if (i == 2) { oFolderItem.Caption = "Medição"; }
                if (i == 3) { oFolderItem.Caption = "Elaboração/Verificação"; }
                if (i == 4) { oFolderItem.Caption = "Apresentação/Aprovação"; }
                if (i == 5) { oFolderItem.Caption = "Pedido"; }
                if (i == 6) { oFolderItem.Caption = "Detalhamento"; }
                if (i == 7) { oFolderItem.Caption = "Fábrica"; }
                if (i == 8) { oFolderItem.Caption = "Montagem"; }
                if (i == 9) { oFolderItem.Caption = "Entrega"; }
                oFolderItem.DataBind.SetBound(true, "", "FolderDS");
                if (i == 1)
                {
                    oFolderItem.Select();
                    oFolderInicial = oFolderItem;
                }
                else
                {
                    oFolderItem.GroupWith("Folder" + (i - 1));
                }
            }

            Entrevista entrevista = new Entrevista(oForm);
            Medicao medicao = new Medicao(oForm);
            Elaboracao elaboracao = new Elaboracao(oForm);
            Verificacao verificacao = new Verificacao(oForm);
            Apresentacao apresentacao = new Apresentacao(oForm);
            Pedido pedido = new Pedido(oForm);
            Detalhamento detalhamento = new Detalhamento(oForm);
            Fabrica fabrica = new Fabrica(oForm);
            Montagem montagem = new Montagem(oForm);
            Entrega entrega = new Entrega(oForm);
                
        }

    }
}
