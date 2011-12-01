using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbobsCOM;


namespace DataModuloSAP
{
    public class ConexaoSAP
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        public static ConexaoSAP conexao;
        private bool isConnectedContext = false;
        private bool isConnectionToCompany = false;
        private ConexaoSAP()
        {
        }

        public static ConexaoSAP Instance
        {
            get
            {
                if (conexao == null)
                {
                    conexao = new ConexaoSAP();
                }
                return conexao;
            }
        }

        public int SetConnectionContext()
        {
            int setConnectionContextReturn = 0;
            string sCookie = null;
            string sConnectionContext = null;

            oCompany = new SAPbobsCOM.Company();
            sCookie = oCompany.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);

            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            setConnectionContextReturn = oCompany.SetSboLoginContext(sConnectionContext);

            return setConnectionContextReturn;
        }

        public int ConnectToCompany()
        {
            int connectToCompanyReturn = 0;
            connectToCompanyReturn = oCompany.Connect();
            return connectToCompanyReturn;
        }

        public void Conectar()
        {
            if (SetConnectionContext() == 0)
            {
                isConnectedContext = true;
            }

            if (ConnectToCompany() == 0)
            {
                isConnectionToCompany = true;
            }
        }

        public SAPbobsCOM.Company getOCompany()
        {
            return oCompany;
        }

        public bool getIsConnectedContext()
        {
            return isConnectedContext;
        }

        public bool getIsConnectionToCompany()
        {
            return isConnectionToCompany;
        }

        public void setSBOApplication(SAPbouiCOM.Application SBO_Application){
            this.SBO_Application = SBO_Application;
        }
    }
}
