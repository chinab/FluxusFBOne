using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataModuloSAP
{
    public class Projetista
    {
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbobsCOM.Recordset projetistas;
        private long RecCount = 0;
        private string idOOPR;

        public Projetista(string idOOPR)
        {
            LoadProjetistasCadastrados(idOOPR);
            RecCount = projetistas.RecordCount;
        }

        public void LoadProjetistasCadastrados(string idOOPR)
        {
            string QryStr = null;

            QryStr = "Select Code, Name from [@FLX_FB_PRJ]";
            projetistas = ((SAPbobsCOM.Recordset)(ConexaoSAP.Instance.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            projetistas.DoQuery(QryStr);
            RecCount = projetistas.RecordCount;
        }

        public SAPbobsCOM.Recordset getProjetistas()
        {
            return this.projetistas;
        }

        public SAPbobsCOM.Recordset trazerProjetistasOportunidade(string idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;

            RecSet = ((SAPbobsCOM.Recordset)(ConexaoSAP.Instance.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "select T1.Name, T2.Name, T3.Name from OOPR T0 left outer join [@FLX_FB_PRJ] T1 on T1.Code = T0.U_FLX_FB_ETV_RESP left outer join [@FLX_FB_PRJ] T2 on T2.Code = T0.U_FLX_FB_APS_PROJT left outer join [@FLX_FB_PRJ] T3 on T3.Code = T0.U_FLX_FB_MED_PROJT where OpprId = " + idOOPR + "";
            RecSet.DoQuery(QryStr);

            return RecSet;
        }

        public SAPbobsCOM.Recordset trazerProjetistasElaboracao(string idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;

            RecSet = ((SAPbobsCOM.Recordset)(ConexaoSAP.Instance.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "select T1.Name from OPR4 T0 left outer join [@FLX_FB_PRJ] T1 on T1.Code = T0.U_FLX_FB_ELB_PROJETI where T0.OprId = " + idOOPR + "";
            RecSet.DoQuery(QryStr);

            return RecSet;
        }
    }
}
