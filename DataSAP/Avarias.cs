using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataSAP
{
    public class Avarias
    {

        public static void AddAvarias(int idOOPR, string descricao, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodeAvarias(oCompany);

            try
            {
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_AVR");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_AVR_DESC", descricao);

                oGeneralParams = oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void UpdateAvarias(string code, string name, int idOOPR, string descricao, SAPbobsCOM.Company oCompany, SAPbouiCOM.Application SBO_Application)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_AVR");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", code);
                oGeneralData.SetProperty("Name", name);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_AVR_DESC", descricao);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static string GetProxCodeAvarias(SAPbobsCOM.Company oCompany)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_AVR] order by Code desc) if @Numero is null begin set @Numero = 0000000 + 1 end SELECT case len(CAST(@Numero AS varchar(7))) WHEN 1 THEN '000000' + CAST(@Numero AS varchar(7)) WHEN 2 THEN '00000' + CAST(@Numero AS varchar(7)) WHEN 3 THEN '0000' + CAST(@Numero AS varchar(7)) WHEN 4 THEN '000' + CAST(@Numero AS varchar(7)) WHEN 5 THEN '00' + CAST(@Numero AS varchar(7)) WHEN 6 THEN '0' + CAST(@Numero AS varchar(7)) WHEN 7 THEN CAST(@Numero AS varchar(7)) END";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }
    }
}
