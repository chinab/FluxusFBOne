using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace DataModuloSAP
{
    public class Ambiente
    {
        private SAPbobsCOM.Recordset ambientes = null;
        private long RecCount = 0;

        public Ambiente(string idOOPR)
        {
            ambientes = LoadAmbientesCadastrados(idOOPR);
            RecCount = ambientes.RecordCount;
        }

        private SAPbobsCOM.Recordset LoadAmbientesCadastrados(string idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;

            RecSet = ((SAPbobsCOM.Recordset)(ConexaoSAP.Instance.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "select T0.Line, T1.Descript, T0.U_FLX_FB_ETV_DESCAMB from OPR4 T0 inner join OOIN T1 on T1.Num = T0.IntId where T0.OprId =" + idOOPR + "";
            
            RecSet.DoQuery(QryStr);
            return RecSet;
        }

        public bool possuiAmbientesCadastrados()
        {
            return (RecCount > 0);
        }

        public ArrayList getIds()
        {
            ArrayList idsAmbientes = new ArrayList();
            ambientes.MoveFirst();

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                idsAmbientes.Add(Convert.ToInt32(ambientes.Fields.Item(0).Value));
                ambientes.MoveNext();
            }
            System.GC.Collect();
            return idsAmbientes;
        }

        public string getDescricaoEntrevista(string id)
        {
            ambientes.MoveFirst();

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                if (ambientes.Fields.Item(0).Value.ToString().Equals(id))
                {
                    return ambientes.Fields.Item(2).Value.ToString();
                }
                ambientes.MoveNext();
            }
            return "";
        }

        /*
        public void loadCombo(SAPbouiCOM.ComboBox oCombo)
        {
            ambientes.MoveFirst();

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                oCombo.ValidValues.Add(Convert.ToString(ambientes.Fields.Item(0).Value),
                                       Convert.ToString(ambientes.Fields.Item(1).Value));
                if (RecIndex == 0)
                {
                    string selectedDesc = ambientes.Fields.Item(1).Value.ToString();
                    oCombo.Select(selectedDesc, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }

                ambientes.MoveNext();
            }

        }
        */
        public SAPbobsCOM.Recordset getAmbientes()
        {
            return ambientes;
        }

    }
}
