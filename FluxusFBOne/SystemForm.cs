using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using VisaoSAP;
using DataModuloSAP;

namespace FluxusFBOne
{

    public class SystemForm
    {
        private SAPbouiCOM.Application SBO_Application;
        private SAPbouiCOM.Form oForm;
        private SAPbouiCOM.Form oFormPai;//
        private SAPbouiCOM.Form oFormAtual;
        private SAPbouiCOM.Item oNewItem;
        private SAPbouiCOM.Item oItem;
        private SAPbouiCOM.Folder oFolderInicial;
        private SAPbouiCOM.EditText oEditItem;
        private SAPbouiCOM.ComboBox oComboItem;
        private SAPbouiCOM.ComboBox oComboItenPrjEntrevista;
        private SAPbouiCOM.ComboBox oComboItenPrjMed;
        private SAPbouiCOM.ComboBox oComboItenPrjApr;
        private SAPbouiCOM.ComboBox comboElbProjetista = null;
        private SAPbouiCOM.ComboBox comboDetProjetista = null;
        private SAPbouiCOM.Matrix oMatrix;
        private SAPbouiCOM.Matrix matrixApresentacao;
        private SAPbouiCOM.Matrix matrixAprovacao;
        private SAPbouiCOM.Columns oColumns;
        private SAPbouiCOM.Columns oColumnsAnaliseCritica;
        private SAPbouiCOM.Column oColumnAnaliseCritica;
        private SAPbouiCOM.DBDataSource DBDSflx;
        private bool upProjEnt = false;
        private bool upProjAps = false;
        private bool upProjMed = false;
        private bool upEtvAmb = false;
        private Process newProcess;
        private ProcessStartInfo info;
        private int countMatrixConfMedAntes = 0;
        private int countMatrixAvariasAntes = 0;
        private int countMatrixAnaliseCriticaAntes = 0;
        private int countMatrixPendenciaAntes = 0;
        private int countMatrixItensComplementaresAntes = 0;
        private SAPbobsCOM.Recordset projetistas;
        private ArrayList ListDataConfMed = new ArrayList();
        private ArrayList ListConferenteCofMed = new ArrayList();
        private ArrayList ListNomeAvarias = new ArrayList();
        private ArrayList ListNomePendencias = new ArrayList();
        private SAPbouiCOM.UserDataSource oUserDataSource;
        private SAPbobsCOM.SalesOpportunities oSalesOpportunities = null;
        private string sBPCode = "";
        private int idAmbiente = 1000;
        private int idAmbientePendencia = 1000;
        private string sSalesOpportunities_Id = "";
        private int iUltimoFormTypeCount_SalesOpportunities = 0;
        private string sDescricaoOriginalAmbiente = "";       
        private int iIdAmbienteMedicao = 1000;                
		private int iRowAmbiente = 1000;        
        private bool bGravouAvarias = false;
        private bool modificouAnsCritica = false;
        private bool bGravouMedicoes = false;
        private int iRowAmbienteMedicao = 0;
        SAPbobsCOM.Recordset RecSet;
        Resumo resumo;
        Fases fases;
        ConexaoSAP conexao;
        Ambiente ambiente;
        Projetista projetista;
        Atividade atividade;
        private SAPbouiCOM.Column oColumnConferenciaMedicoes;
        private SAPbouiCOM.Columns oColumnsConferenciaMedicoes;
        private bool bBotaoAgendarFoiClicado;
        private bool modificouPendecia = false;
        private SAPbouiCOM.Columns oColumnsPendencia;
        private SAPbouiCOM.Column oColumnPendencia;

        private void SetApplication()
        {

            SAPbouiCOM.SboGuiApi SboGuiApi = null;
            string sConnectionString = null;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
            try
            {
                SboGuiApi.Connect(sConnectionString);
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                System.Environment.Exit(0);
            }

            SBO_Application = SboGuiApi.GetApplication(-1);
        }

        private void AddItemsToForm()
        {
            oForm.DataSources.UserDataSources.Add("OpBtnDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("CheckDS1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("CheckDS2", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("CheckDS3", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
            oForm.DataSources.UserDataSources.Add("EditSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            oForm.DataSources.UserDataSources.Add("CombSource", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);

            oUserDataSource = oForm.DataSources.UserDataSources.Add("dt", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 254);
            AddChooseFromList();
            AddChooseFromList2();
        }

        public SystemForm()
        {
            SetApplication();

            SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);

            SBO_Application.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(SBO_Application_DataEvent);

            SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);

            SBO_Application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);

            conexao = ConexaoSAP.Instance;
            conexao.setSBOApplication(SBO_Application);
            conexao.Conectar();


            if (!conexao.getIsConnectedContext())
            {
                SBO_Application.MessageBox("Failed setting a connection to DI API", 1, "Ok", "", "");
                System.Environment.Exit(0); //  Terminating the Add-On Application
            }
            if (!conexao.getIsConnectionToCompany())
            {
                SBO_Application.MessageBox("Failed connecting to the company's Data Base", 1, "Ok", "", "");
                System.Environment.Exit(0); //  Terminating the Add-On Application
            }
        }

        private void SBO_Application_DataEvent(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if ((pVal.FormTypeEx == "320" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) & (pVal.BeforeAction == false))
            {
                string idOOPR = ((SAPbouiCOM.EditText)oForm.Items.Item("74").Specific).Value;
                ambiente = new Ambiente(idOOPR);
                projetista = new Projetista(idOOPR);
                projetistas = projetista.getProjetistas();

                if (ambiente.possuiAmbientesCadastrados())
                {
                    //Seta todas a Matrizes para carregar os ambientes cadastrados.
                    LoadAmbientesInMatrix();
                    oMatrix.LoadFromDataSource();
                }

                //Grids com Tabela de Usuário - Inclui o select em um datatable e seta este datatable para a matriz específica.
                idAmbiente = 1000;
                iRowAmbiente = 1000;
                iIdAmbienteMedicao = 1000;
                idAmbientePendencia = 1000;
                LoadGridConferenciaMedicao();
                LoadGridAvarias();
                LoadGridPendencias();
                LoadGridItensComplementares();

                //Limpa a variável que pega o id do ambiente selecionado e carrega a grid de analise critica com parâmetro idAmbiente = 0
                LoadGridAnaliseCritica();


                /**** INÍCIO - CARREGAMENTO DE COMBOS *******/

                loadCombo("Ent_Amb", ambiente.getAmbientes());

                oEditItem = ((SAPbouiCOM.EditText)oForm.Items.Item("Ent_Det").Specific);
                string selectedValue = ((SAPbouiCOM.ComboBox)oForm.Items.Item("Ent_Amb").Specific).Value;
                oEditItem.Value = ambiente.getDescricaoEntrevista(selectedValue);

                //Apresentação - Pega o valor cadastrado na drop de projetistas, lista o nome e todos os projetistas.
                loadCombo("Apr_Proj", projetistas);
                //Entrevista - Pega o valor cadastrado na drop de projetistas, lista o nome e todos os projetistas.
                loadCombo("Ent_Proj", projetistas);
                //Medição - Pega o valor cadastrado na drop de projetistas, lista o nome e todos os projetistas.
                loadCombo("Med_Proj", projetistas);

                if (ambiente.possuiAmbientesCadastrados())
                {
                    //Projetistas - Grid Elaboração
                    loadComboEmGrid("Ela_Amb", "Ela_Amb_C2", projetistas);
                    //Projetistas - Grid Detalhamento
                    loadComboEmGrid("Det_Amb", "Det_Amb_C2", projetistas);
                    //Projetistas - Grid Verificação
                    loadComboEmGrid("Ver_Amb", "Ver_Amb_C2", projetistas);
                    //Projetistas - Grid Aprovação
                    loadComboEmGrid("Apv_Amb", "Apv_Amb_C1", projetistas);
                    //Projetistas - Grid Pedido
                    loadComboEmGrid("Ped_Amb", "Ped_Amb_C5", projetistas);
                    //Projetistas - Grid Fabrica
                    loadComboEmGrid("Fab_Amb", "Fab_Amb_C3", projetistas);
                    //Projetistas - Grid Montagem
                    loadComboEmGrid("Mon_Amb", "Mon_Amb_C1", projetistas);
                    //Projetistas - Grid Entrega
                    loadComboEmGrid("Etg_Amb", "Etg_Amb_C2", projetistas);
                }
                /**** FIM - CARREGAMENTO DE COMBOS *******/

                oNewItem = oForm.Items.Item("Apr_Proj");
                oComboItenPrjApr = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));

                //Entrevista - Pega o valor cadastrado na drop de projetistas, lista o nome e todos os projetistas.
                oNewItem = oForm.Items.Item("Ent_Proj");
                oComboItenPrjEntrevista = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));

                //Medição - Pega o valor cadastrado na drop de projetistas, lista o nome e todos os projetistas.
                oNewItem = oForm.Items.Item("Med_Proj");
                oComboItenPrjMed = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));

                //LoadComboProjetistaCadastradoElaboracao(idOOPR);
                LoadComboProjetistaCadastradoOOPR(oComboItenPrjEntrevista, oComboItenPrjMed, oComboItenPrjApr, idOOPR);

                ContarLinhasIniciaisMatrix();
                resumo.disableCampos();
                PreencherCamposResumo(idOOPR);
            }

            if (pVal.FormTypeEx == "651" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD & !pVal.BeforeAction)
            {
                if (oFormAtual.Items.Item("67").Specific.Value == "M")
                {
                    atividade.habilitaBotaoAta();
                }
                else
                {
                    atividade.desabilitaBotaoAta();
                }
            }

        }

        public void loadCombo(string combo, SAPbobsCOM.Recordset valores)
        {
            string sPrimeirovalorCombo = "";
            oNewItem = oForm.Items.Item(combo);
            SAPbouiCOM.ComboBox oCombo = ((SAPbouiCOM.ComboBox)(oNewItem.Specific));
            RemoveValoresDeCombo(ref oCombo);

            oCombo.ValidValues.Add("", "");

            int RecCount = valores.RecordCount;
            valores.MoveFirst();

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                //Se for combo de ambientes da aba de Entrevista
                if ((combo == "Ent_Amb") & (RecIndex == 0))
                    sPrimeirovalorCombo = valores.Fields.Item(0).Value.ToString();
                oCombo.ValidValues.Add(Convert.ToString(valores.Fields.Item(0).Value), Convert.ToString(valores.Fields.Item(1).Value));
                valores.MoveNext();
            }

            //Se for combo de ambientes da aba de Entrevista
            if (combo == "Ent_Amb")
            {
                oCombo.Select(sPrimeirovalorCombo, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
        }

        public void loadComboEmGrid(string matrix, string combo, SAPbobsCOM.Recordset valores)
        {
            oMatrix = checked((SAPbouiCOM.Matrix)oForm.Items.Item(matrix).Specific);
            SAPbouiCOM.ComboBox oCombo = checked((SAPbouiCOM.ComboBox)oMatrix.Columns.Item(combo).Cells.Item(1).Specific);
            RemoveValoresDeCombo(ref oCombo);

            oCombo.ValidValues.Add("", "");

            int RecCount = valores.RecordCount;
            valores.MoveFirst();

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                oCombo.ValidValues.Add(Convert.ToString(valores.Fields.Item(0).Value), Convert.ToString(valores.Fields.Item(1).Value));
                valores.MoveNext();
            }
        }

        public void RemoveValoresDeCombo(ref SAPbouiCOM.ComboBox oComboItem)
        {
            //Remove valores da drop ao mudar de registro.
            while (oComboItem.ValidValues.Count > 0)
            {
                oComboItem.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
        }

        public void PreencherCamposResumo(string idOOPR)
        {
            //Preenche Campos da Aba Resumo
            SAPbouiCOM.EditText etvResp = null;
            SAPbouiCOM.EditText medResp = null;
            SAPbouiCOM.EditText elbIni = null;
            SAPbouiCOM.EditText elbFim = null;
            SAPbouiCOM.EditText elbResp = null;
            SAPbouiCOM.EditText verificacao = null;
            SAPbouiCOM.EditText verResp = null;
            SAPbouiCOM.EditText apsResp = null;
            SAPbouiCOM.EditText aprovacao = null;
            SAPbouiCOM.EditText aprResp = null;
            SAPbouiCOM.EditText pedido = null;
            SAPbouiCOM.EditText pedResp = null;
            SAPbouiCOM.EditText detIni = null;
            SAPbouiCOM.EditText detFim = null;
            SAPbouiCOM.EditText detResp = null;
            SAPbouiCOM.EditText fabIni = null;
            SAPbouiCOM.EditText fabFim = null;
            SAPbouiCOM.EditText fabResp = null;
            SAPbouiCOM.EditText montagem = null;
            SAPbouiCOM.EditText montagemFim = null;
            SAPbouiCOM.EditText montResp = null;
            SAPbouiCOM.EditText entrega = null;
            SAPbouiCOM.EditText etgResp = null;

            etvResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Ent_Res").Specific);
            medResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Med_Res").Specific);
            elbIni = ((SAPbouiCOM.EditText)oForm.Items.Item("Ela_Dat1").Specific);
            elbFim = ((SAPbouiCOM.EditText)oForm.Items.Item("Ela_Dat2").Specific);
            elbResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Ela_Res").Specific);
            verificacao = ((SAPbouiCOM.EditText)oForm.Items.Item("Ver_Dat1").Specific);
            verResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Ver_Res").Specific);
            apsResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Aps_Res").Specific);
            aprovacao = ((SAPbouiCOM.EditText)oForm.Items.Item("Apv_Dat1").Specific);
            aprResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Apv_Res").Specific);
            pedido = ((SAPbouiCOM.EditText)oForm.Items.Item("Ped_Dat1").Specific);
            pedResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Ped_Res").Specific);
            detIni = ((SAPbouiCOM.EditText)oForm.Items.Item("Det_Dat1").Specific);
            detFim = ((SAPbouiCOM.EditText)oForm.Items.Item("Det_Dat2").Specific);
            detResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Det_Res").Specific);
            fabIni = ((SAPbouiCOM.EditText)oForm.Items.Item("Fab_Dat1").Specific);
            fabFim = ((SAPbouiCOM.EditText)oForm.Items.Item("Fab_Dat2").Specific);
            fabResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Fab_Res").Specific);
            montagem = ((SAPbouiCOM.EditText)oForm.Items.Item("Mon_Dat1").Specific);
            montagemFim = ((SAPbouiCOM.EditText)oForm.Items.Item("Mon_Dat2").Specific);
            montResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Mon_Res").Specific);
            entrega = ((SAPbouiCOM.EditText)oForm.Items.Item("Eng_Dat1").Specific);
            etgResp = ((SAPbouiCOM.EditText)oForm.Items.Item("Eng_Res").Specific);

            LoadResumo(etvResp, medResp, elbIni, elbFim, elbResp, verificacao, verResp, apsResp, aprovacao, aprResp, pedido, pedResp, detIni, detFim, detResp, fabIni, fabFim, fabResp, montagem, montagemFim, montResp, entrega, etgResp, idOOPR);
        }

        public void ContarLinhasIniciaisMatrix()
        {
            ContarLinhasIniciaisMatrixMedicoes();

            ContarLinhasIniciaisMatrixAvarias();

            ContarLinhasIniciaisMatrixPendencias();

            ContarLinhasIniciaisMatrixItensComplementares();
        }

        private void ContarLinhasIniciaisMatrixPendencias()
        {
            //Pendencias
            oNewItem = oForm.Items.Item("Pen_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            countMatrixPendenciaAntes = oMatrix.RowCount;
            ListNomePendencias.Clear();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                ListNomePendencias.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C0").Cells.Item(i).Specific).String);
            }
        }

        private void ContarLinhasIniciaisMatrixAvarias()
        {
            //Avarias
            oNewItem = oForm.Items.Item("Ava_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            countMatrixAvariasAntes = oMatrix.RowCount;
            ListNomeAvarias.Clear();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                ListNomeAvarias.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ava_Amb_C0").Cells.Item(i).Specific).String);
            }
        }

        private void ContarLinhasIniciaisMatrixMedicoes()
        {
            //Conferência de Medições
            oNewItem = oForm.Items.Item("Med_Cnf");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            countMatrixConfMedAntes = oMatrix.RowCount;
            ListDataConfMed.Clear();
            ListConferenteCofMed.Clear();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                ListDataConfMed.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Cnf_C0").Cells.Item(i).Specific).String);
                ListConferenteCofMed.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("med_Cnf_C1").Cells.Item(i).Specific).String);
            }
        }

        private void ContarLinhasIniciaisMatrixItensComplementares()
        {
            //Pendencias
            oNewItem = oForm.Items.Item("Det_Cmp");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            countMatrixItensComplementaresAntes = oMatrix.RowCount;
            /*ListNomePendencias.Clear();
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                ListNomePendencias.Add(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C0").Cells.Item(i).Specific).String);
            }*/
        }

        public void AddPendencias(int idOOPR, string descricao, int idAmbientePend)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodePendencias();

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_PEN");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_PEN_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_PEN_IDAMB", idAmbientePend);
                oGeneralData.SetProperty("U_FLX_FB_PEN_DESC", descricao);

                oGeneralParams = oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdatePendencias(string descricao, string pkPendencia)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;   

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_PEN");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", pkPendencia);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_FLX_FB_PEN_DESC", descricao);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void AddAvarias(int idOOPR, string descricao, int identificadoAmbiente)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodeAvarias();

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_AVR");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_AVR_DESC", descricao);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDAMBI", identificadoAmbiente);

                oGeneralParams = oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdateAvarias(string code, string name, int idOOPR, string descricao, int identificadorAmbiente)
        {
            if (code == "")
                return;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_AVR");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", code);
                oGeneralData.SetProperty("Name", name);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_AVR_DESC", descricao);
                oGeneralData.SetProperty("U_FLX_FB_AVR_IDAMBI", identificadorAmbiente);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void AddConferenciaDeMedicao(int idOOPR, string dataConfMedicao, string nomeConferente, int idAmbiente)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodeConferenciaMedicao();

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_CONFMED");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_ID", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_DAT", dataConfMedicao);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_PRJ", Convert.ToInt32(nomeConferente));
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_IDA", iIdAmbienteMedicao);

                oGeneralParams = oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdateConferenciaDeMedicao(string code, string name, int idOOPR, string dataConfMedicao, string nomeConferente, int idAmbiente)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_CONFMED");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", code);
                oGeneralData.SetProperty("Name", name);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_ID", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_DAT", dataConfMedicao);
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_PRJ", Convert.ToInt32(nomeConferente));
                oGeneralData.SetProperty("U_FLX_FB_CONFMED_IDA", iIdAmbienteMedicao);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void AddOportunidadeVendas(int idOOPR, string etvProjetistaResp, string medProjetista, string apsProjetista, int etvAmbiente, string descAmb)
        {
            try
            {
                oSalesOpportunities = checked((SAPbobsCOM.SalesOpportunities)conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities));

                if (oSalesOpportunities.GetByKey(idOOPR) == true)
                {
                    //Aba Fases Entrevista
                    if (etvProjetistaResp != "0")
                    {
                        oSalesOpportunities.UserFields.Fields.Item("U_FLX_FB_ETV_RESP").Value = Convert.ToInt32(etvProjetistaResp);
                    }

                    if (medProjetista != "0")
                    {
                        oSalesOpportunities.UserFields.Fields.Item("U_FLX_FB_MED_PROJT").Value = Convert.ToInt32(medProjetista);
                    }

                    //Abas Fase Apresentação
                    if (apsProjetista != "0")
                    {
                        oSalesOpportunities.UserFields.Fields.Item("U_FLX_FB_APS_PROJT").Value = Convert.ToInt32(apsProjetista);
                    }

                    for (int i = 0; i < oSalesOpportunities.Interests.Count; i++)
                    {
                        oSalesOpportunities.Interests.SetCurrentLine(i);
                        /*if (oSalesOpportunities.Interests.InterestId == etvAmbiente)
                        {
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ETV_DESCAMB").Value = descAmb;
                        }*/
                        if (oSalesOpportunities.Interests.RowNo == etvAmbiente)
                        {
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ETV_DESCAMB").Value = descAmb;
                        }
                    }

                    oSalesOpportunities.Update();
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void AddAnexoMedicao(int idOOPR)
        {
            try
            {
                SAPbobsCOM.SalesOpportunities oSalesOpportunities = null;
                oSalesOpportunities = checked((SAPbobsCOM.SalesOpportunities)conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.oSalesOpportunities));
                string medAnexoLevantamento = null;
                string elbIniPrev = null;
                string elbTermPrev = null;
                string elbIniRealizado = null;
                string elbTermRealizado = null;
                string elbArqCad = null;
                string elbArqPRJ = null;
                string elbArqJPG = null;
                int elbProjetista = 0;
                string elbDatRevisao = null;
                string apsDetalhamento = null;
                string apvAprovadoPor = null;
                string apvDataApv = null;
                string apvAnexoPdf = null;
                string apvPranchaImagem = null;
                string apvMemorialDescritivo = null;
                string verData = null;
                string verObs = null;
                string verVerificadoPor = null;
                string pedData = null;
                string pedUrl = null;
                string pedNumero = null;
                string pedOrdemCompra = null;
                string pedValor = null;
                string pedSolicitante = null;
                string pedPrazEntrega = null;
                string pedAnexo = null;
                string detIniPrev = null;
                string detTermPrev = null;
                string detIniRealzidado = null;
                string detTermRealizado = null;
                string detAnexo = null;
                int detProjetista = 0;
                string fabExpedicao = null;
                string fabConferente = null;
                string fabRecebimento = null;
                string montResponsavel = null;
                string montDescricao = null;
                string montVstInt1 = null;
                string montVstInt2 = null;
                string montVstInt3 = null;
                string etgDatEntrega = null;
                string etgLaudoEntrega = null;
                string etgResponsavel = null;
                string etgDatSolucao = null;
                bool etgResolvido = false;

                oNewItem = oForm.Items.Item("Med_Amb");
                oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Ela_Amb");
                SAPbouiCOM.Matrix matrixElaboracao;
                matrixElaboracao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Apr_Amb");
                SAPbouiCOM.Matrix matrixApresentacao;
                matrixApresentacao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Apv_Amb");
                SAPbouiCOM.Matrix matrixAprovacao;
                matrixAprovacao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Ver_Amb");
                SAPbouiCOM.Matrix matrixVerificacao;
                matrixVerificacao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Ped_Amb");
                SAPbouiCOM.Matrix matrixPedido;
                matrixPedido = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Det_Amb");
                SAPbouiCOM.Matrix matrixDetalhamento;
                matrixDetalhamento = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Fab_Amb");
                SAPbouiCOM.Matrix matrixFabrica;
                matrixFabrica = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Mon_Amb");
                SAPbouiCOM.Matrix matrixMontagem;
                matrixMontagem = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                oNewItem = oForm.Items.Item("Etg_Amb");
                SAPbouiCOM.Matrix matrixEntrega;
                matrixEntrega = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                ArrayList idsAmbientes = ambiente.getIds();

                if (oSalesOpportunities.GetByKey(idOOPR) == true)
                {
                    for (int i = 0; i < oSalesOpportunities.Interests.Count; i++)
                    {
                        oSalesOpportunities.Interests.SetCurrentLine(i);
                        int id = Convert.ToInt32(idsAmbientes[i]);
                        if (oSalesOpportunities.Interests.RowNo == id)
                        {
                            //Instanciar a Grid de Ambientes
                            medAnexoLevantamento = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Amb_C1").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MED_LEVANTA").Value = medAnexoLevantamento;

                            //Instanciar a Grid de Elaboração

                            elbDatRevisao = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C1").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_REVISAO").Value = elbDatRevisao;

                            SAPbouiCOM.ComboBox combo;
                            //oItem = oForm.Items.Item("Ela_Amb_C2");
                            combo = (SAPbouiCOM.ComboBox)matrixElaboracao.Columns.Item("Ela_Amb_C2").Cells.Item(i + 1).Specific;
                            if (combo.Value != "")
                            {
                                elbProjetista = Convert.ToInt32(((SAPbouiCOM.ComboBox)matrixElaboracao.Columns.Item("Ela_Amb_C2").Cells.Item(i + 1).Specific).Value);
                                oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_PROJETI").Value = elbProjetista;
                            }

                            elbIniPrev = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C3").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_INICIOP").Value = elbIniPrev;

                            elbTermPrev = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C4").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_TERMINP").Value = elbTermPrev;

                            elbIniRealizado = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C5").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_INICIOR").Value = elbIniRealizado;

                            elbTermRealizado = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C6").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_TERMINR").Value = elbTermRealizado;

                            elbArqCad = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C7").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_ARQCAD").Value = elbArqCad;

                            elbArqPRJ = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C8").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_ARQPRJ").Value = elbArqPRJ;

                            elbArqJPG = ((SAPbouiCOM.EditText)matrixElaboracao.Columns.Item("Ela_Amb_C9").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ELB_ARQJPG").Value = elbArqJPG;


                            //Instanciar a Grid de Verificação
                            verData = ((SAPbouiCOM.EditText)matrixVerificacao.Columns.Item("Ver_Amb_C1").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_VRF_DATEVER").Value = verData;

                            verVerificadoPor = ((SAPbouiCOM.ComboBox)matrixVerificacao.Columns.Item("Ver_Amb_C2").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_VRF_VERIFPO").Value = Convert.ToInt32(verVerificadoPor);

                            verObs = ((SAPbouiCOM.EditText)matrixVerificacao.Columns.Item("Ver_Amb_C3").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_VRF_OBS").Value = verObs;


                            //Instanciar a Grid de Apresentação
                            apsDetalhamento = ((SAPbouiCOM.EditText)matrixApresentacao.Columns.Item("Apr_Amb_C1").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ANC_DETALHA").Value = apsDetalhamento;


                            //Instanciar a Grid de Aprovação
                            apvAprovadoPor = ((SAPbouiCOM.ComboBox)matrixAprovacao.Columns.Item("Apv_Amb_C1").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_APR_APROVAD").Value = Convert.ToInt32(apvAprovadoPor);

                            apvDataApv = ((SAPbouiCOM.EditText)matrixAprovacao.Columns.Item("Apv_Amb_C2").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_APR_DATAAPR").Value = apvDataApv;

                            apvAnexoPdf = ((SAPbouiCOM.EditText)matrixAprovacao.Columns.Item("Apv_Amb_C3").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_APR_PDFCLIE").Value = apvAnexoPdf;

                            apvPranchaImagem = ((SAPbouiCOM.EditText)matrixAprovacao.Columns.Item("Apv_Amb_C4").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_APR_PRANIMG").Value = apvPranchaImagem;

                            apvMemorialDescritivo = ((SAPbouiCOM.EditText)matrixAprovacao.Columns.Item("Apv_Amb_C5").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_APR_MEMDESC").Value = apvMemorialDescritivo;


                            //Instanciar a Grid de Pedidos
                            pedData = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C1").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_DATE").Value = pedData;

                            pedNumero = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C2").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_NPEDIDO").Value = pedNumero;

                            pedOrdemCompra = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C3").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_ORDENDE").Value = pedOrdemCompra;

                            pedValor = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C4").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_VALOR").Value = pedValor;

                            pedSolicitante = ((SAPbouiCOM.ComboBox)matrixPedido.Columns.Item("Ped_Amb_C5").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_SOLICIT").Value = Convert.ToInt32(pedSolicitante);

                            pedPrazEntrega = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C6").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_PRAZOEN").Value = pedPrazEntrega;

                            pedAnexo = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C7").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_ANEXOPE").Value = pedAnexo;

                            pedUrl = ((SAPbouiCOM.EditText)matrixPedido.Columns.Item("Ped_Amb_C8").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_PED_URL").Value = pedUrl;


                            //Instanciar a Grid de Detalhamento

                            combo = (SAPbouiCOM.ComboBox)matrixDetalhamento.Columns.Item("Det_Amb_C2").Cells.Item(i + 1).Specific;
                            if (combo.Value != "")
                            {
                                detProjetista = Convert.ToInt32(((SAPbouiCOM.ComboBox)matrixDetalhamento.Columns.Item("Det_Amb_C2").Cells.Item(i + 1).Specific).Value);
                                oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_PROJETI").Value = detProjetista;
                            }

                            detIniPrev = ((SAPbouiCOM.EditText)matrixDetalhamento.Columns.Item("Det_Amb_C3").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_INICIOP").Value = detIniPrev;

                            detTermPrev = ((SAPbouiCOM.EditText)matrixDetalhamento.Columns.Item("Det_Amb_C4").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_TERMINP").Value = detTermPrev;

                            detIniRealzidado = ((SAPbouiCOM.EditText)matrixDetalhamento.Columns.Item("Det_Amb_C5").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_INICIRE").Value = detIniRealzidado;

                            detTermRealizado = ((SAPbouiCOM.EditText)matrixDetalhamento.Columns.Item("Det_Amb_C6").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_TERMINO").Value = detTermRealizado;

                            detAnexo = ((SAPbouiCOM.EditText)matrixDetalhamento.Columns.Item("Det_Amb_C7").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_DET_PDF").Value = detAnexo;


                            //Instanciar a Grid de Fábrica
                            fabExpedicao = ((SAPbouiCOM.EditText)matrixFabrica.Columns.Item("Fab_Amb_C1").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_FAB_EXPEDIC").Value = fabExpedicao;

                            fabRecebimento = ((SAPbouiCOM.EditText)matrixFabrica.Columns.Item("Fab_Amb_C2").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_FAB_RECEBIM").Value = fabRecebimento;

                            fabConferente = ((SAPbouiCOM.ComboBox)matrixFabrica.Columns.Item("Fab_Amb_C3").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_FAB_CONFERE").Value = Convert.ToInt32(fabConferente);


                            //Instanciar a Grid de Montagem
                            montResponsavel = ((SAPbouiCOM.ComboBox)matrixMontagem.Columns.Item("Mon_Amb_C1").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MTG_RESPONS").Value = Convert.ToInt32(montResponsavel);

                            montDescricao = ((SAPbouiCOM.EditText)matrixMontagem.Columns.Item("Mon_Amb_C2").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MTG_DESCRIC").Value = montDescricao;

                            montVstInt1 = ((SAPbouiCOM.EditText)matrixMontagem.Columns.Item("Mon_Amb_C3").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MTG_VSTINT1").Value = montVstInt1;

                            montVstInt2 = ((SAPbouiCOM.EditText)matrixMontagem.Columns.Item("Mon_Amb_C4").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MTG_VSTINT2").Value = montVstInt2;

                            montVstInt3 = ((SAPbouiCOM.EditText)matrixMontagem.Columns.Item("Mon_Amb_C5").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_MTG_VSTINT3").Value = montVstInt3;

                            //Instanciar a Grid de Entrega
                            etgDatEntrega = ((SAPbouiCOM.EditText)matrixEntrega.Columns.Item("Etg_Amb_C1").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_ENTREGA").Value = etgDatEntrega;

                            etgResponsavel = ((SAPbouiCOM.ComboBox)matrixEntrega.Columns.Item("Etg_Amb_C2").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_RESPONS").Value = Convert.ToInt32(etgResponsavel);

                            etgLaudoEntrega = ((SAPbouiCOM.EditText)matrixEntrega.Columns.Item("Etg_Amb_C3").Cells.Item(i + 1).Specific).Value;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_LAUDO").Value = etgLaudoEntrega;

                            etgDatSolucao = ((SAPbouiCOM.EditText)matrixEntrega.Columns.Item("Etg_Amb_C4").Cells.Item(i + 1).Specific).String;
                            oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_DATASOL").Value = etgDatSolucao;

                            etgResolvido = ((SAPbouiCOM.CheckBox)matrixEntrega.Columns.Item("Etg_Amb_C5").Cells.Item(i + 1).Specific).Checked;
                            if (etgResolvido)
                            {
                                oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_RESOLVI").Value = 1;
                            }
                            else
                            {
                                oSalesOpportunities.Interests.UserFields.Fields.Item("U_FLX_FB_ENT_RESOLVI").Value = 0;
                            }


                            oSalesOpportunities.Update();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            SAPbouiCOM.BoEventTypes EventEnum = 0;
            EventEnum = pVal.EventType;
            BubbleEvent = true;

            if (pVal.FormType == 320)
            {
                if ((pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (!pVal.Before_Action))
                {
                    //Laudo_Ini
                    if (pVal.ItemUID == "Laudo_Ini")
                    {
                        abrirRelatorio("Laudo inicial", oForm.Items.Item("74").Specific.Value);
                    }

                    //Ent_Imp
                    if (pVal.ItemUID == "Ent_Imp")
                    {
                        abrirRelatorio("Descricao dos ambientes", oForm.Items.Item("74").Specific.Value);
                    }

                    //Laudo_Int
                    if (pVal.ItemUID == "Laudo_Int")
                    {
                        abrirRelatorio("Laudo intermediario", oForm.Items.Item("74").Specific.Value);
                    }
                    //NvAnalise
                    if (pVal.ItemUID == "NvAnalise")
                    {
                        abrirRelatorio("Analise critica", oForm.Items.Item("74").Specific.Value);
                    }
                    //Laudo_ent
                    if (pVal.ItemUID == "Laudo_Ent")
                    {
                        abrirRelatorio("Laudo de entrega", oForm.Items.Item("74").Specific.Value);
                    }
					//Pesquisa de satisfacao
                    if (pVal.ItemUID == "Etg_Pq")
                    {
                        abrirRelatorio("Pesquisa de satisfacao", oForm.Items.Item("74").Specific.Value);
                    }					
                    //NvLev
                    if (pVal.ItemUID == "NvLev")
                    {
                        abrirRelatorio("Levantamento", "");
                    }
                    //Etg_Decl
                    if (pVal.ItemUID == "Etg_Decl")
                    {
                        abrirRelatorio("Declaracao de conformidade", oForm.Items.Item("74").Specific.Value);
                    }                
                
                }

                // Ao mudar o ambiente
                if (pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) && pVal.ItemUID == "Ent_Amb" && pVal.ItemChanged)
                {
                    // Verifica a descricao de Ent_Det se mudou e captura
                    oEditItem = ((SAPbouiCOM.EditText)oForm.Items.Item("Ent_Det").Specific);
                    string sEnt_Det = oEditItem.String;
                    try
                    {
                        if (sEnt_Det != sDescricaoOriginalAmbiente)
                        {
                            int idOOPR = int.Parse(((SAPbouiCOM.EditText)oForm.Items.Item("74").Specific).Value);
                            string selectedValue = ((SAPbouiCOM.ComboBox)oForm.Items.Item("Ent_Amb").Specific).Value;
                            int iSelectedValue;
                            if (selectedValue != "")
                            {
                                iSelectedValue = int.Parse(selectedValue);
                                // Atualiza a Descrição na Oportunidade de Vendas
                                AddOportunidadeVendas(idOOPR, "0", "0", "0", iSelectedValue, sEnt_Det);
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) && pVal.ItemUID == "Ent_Amb" && pVal.ItemChanged)
                {
                    oEditItem = ((SAPbouiCOM.EditText)oForm.Items.Item("Ent_Det").Specific);

                    string idOOPR = ((SAPbouiCOM.EditText)oForm.Items.Item("74").Specific).Value;
                    string selectedValue = ((SAPbouiCOM.ComboBox)oForm.Items.Item("Ent_Amb").Specific).Value;

                    ambiente = new Ambiente(idOOPR);
                    oEditItem.Value = ambiente.getDescricaoEntrevista(selectedValue);
                }

                //Evento da Drop.

                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "Ela_Amb"))
                {
                    string coluna2 = pVal.ColUID;
                    if (coluna2 == "Ela_Amb_C2")
                    {
                        string linha = pVal.Row.ToString();
                    }
                }

                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "Ent_Proj"))
                {
                    upProjEnt = true;
                    //SBO_Application.MessageBox("Mudou Proj Entrevista.");
                }
                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "Med_Proj"))
                {
                    upProjMed = true;
                    //SBO_Application.MessageBox("Mudou Proj Medição.");
                }
                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "Apr_Proj"))
                {
                    upProjAps = true;
                    //SBO_Application.MessageBox("Mudou Proj Apresentação.");
                }
                if (!pVal.Before_Action && (EventEnum == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "Ent_Amb"))
                {
                    upEtvAmb = true;
                    //SBO_Application.MessageBox("Mudou Combo de Ambiente");
                }

                //Abre tela de Atividades.
                if (((pVal.ItemUID == "Ent_Age") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (pVal.Before_Action == false)))
                {
                    bBotaoAgendarFoiClicado = true;
                    SBO_Application.ActivateMenuItem("2563");
                }

                //Abre tela de Atividades.
                if (((pVal.ItemUID == "Med_Age") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (pVal.Before_Action == false)))
                {
                    bBotaoAgendarFoiClicado = true;
                    SBO_Application.ActivateMenuItem("2563");
                }

                //Clique do Botão Atualizar
                if (((pVal.ItemUID == "1") & (pVal.FormMode == 1) & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED) & (pVal.Before_Action == false)))
                {
                    if (ambiente.possuiAmbientesCadastrados())
                    {
                       Atualizar();
                    }
                }

                //Criar os campos do formulario.
                if (pVal.Before_Action && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    oForm = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    AddItemsToForm();

                    oForm.Resize(300, 130);// (132, 100);

                    resumo = new Resumo(oForm);
                    fases = new Fases(oForm);
                }

                //Evento do Clique da aba Resumo.
                if (pVal.ItemUID == "Projeto1" & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) & pVal.Before_Action)
                {
                    resumo.disableCampos();
                    oForm.PaneLevel = 8;
                }

                //Evento do Clique da aba Fases.
                if (pVal.ItemUID == "Projeto2" & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) & pVal.Before_Action)
                {
                    oForm.PaneLevel = 9;
                }

                int panel = 9;
                if (pVal.ItemUID.StartsWith("Folder") & (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED || pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK) & pVal.Before_Action)
                {
                    switch (pVal.ItemUID)
                    {
                        case "Folder1": panel = 9;
                            break;
                        case "Folder2": panel = 10;
                            break;
                        case "Folder3": panel = 11;
                            break;
                        case "Folder4": panel = 12;
                            break;
                        case "Folder5": panel = 13;
                            break;
                        case "Folder6": panel = 14;
                            break;
                        case "Folder7": panel = 15;
                            break;
                        case "Folder8": panel = 16;
                            break;
                        case "Folder9": panel = 17;
                            break;
                    }

                    oForm.PaneLevel = panel;
                }


                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = ((SAPbouiCOM.IChooseFromListEvent)(pVal));
                    string sCFL_ID = oCFLEvento.ChooseFromListUID;
                    SAPbouiCOM.Form oForm = SBO_Application.Forms.Item(FormUID);
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item(sCFL_ID);


                    if (oCFLEvento.BeforeAction == false && sCFL_ID == "CFL1")
                    {
                        SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                        string valItemName = null;
                        string valItemCode = null;
                        try
                        {
                            valItemCode = System.Convert.ToString(oDataTable.GetValue(0, 0));
                            valItemName = System.Convert.ToString(oDataTable.GetValue(1, 0));

                            string qtdEstoque = GetQtdEmEstoque(valItemCode);
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C4").Cells.Item(pVal.Row).Specific).Value = qtdEstoque;

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C0").Cells.Item(pVal.Row).Specific).Value = valItemCode;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C1").Cells.Item(pVal.Row).Specific).Value = valItemName;
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    else if (oCFLEvento.BeforeAction == false && sCFL_ID == "CFL2")
                    {
                        SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                        string valCarName = null;
                        string idFornecedor = null;
                        try
                        {
                            idFornecedor = System.Convert.ToString(oDataTable.GetValue(0, 0));
                            valCarName = System.Convert.ToString(oDataTable.GetValue(1, 0));

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C6").Cells.Item(pVal.Row).Specific).Value = idFornecedor;
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C3").Cells.Item(pVal.Row).Specific).Value = valCarName;
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }

                string coluna = pVal.ColUID;
                
                if (EventEnum == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && !pVal.BeforeAction)
                {
                    //Anexo de arquivo
                    if (coluna == "Ela_Amb_C7" || coluna == "Ela_Amb_C8" || coluna == "Ela_Amb_C9" || coluna == "Med_Amb_C1"
                                               || coluna == "Apv_Amb_C3" || coluna == "Ped_Amb_C7" || coluna == "Det_Amb_C7"
                                               || coluna == "Etg_Amb_C3" || coluna == "Mon_Amb_C3" || coluna == "Mon_Amb_C4" 
                                               || coluna == "Mon_Amb_C5" || coluna == "Apv_Amb_C4" || coluna == "Ans_Amb_C0")
                    {
                        oNewItem = oForm.Items.Item(pVal.ItemUID);
                        oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
                        oEditItem = (SAPbouiCOM.EditText)oMatrix.Columns.Item(coluna).Cells.Item(pVal.Row).Specific;

                        GridComAnexo(oEditItem);
                    }
                    
                    //Url
                    if (coluna == "Ped_Amb_C8" && ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ped_Amb_C8").Cells.Item(pVal.Row).Specific).Value != "")
                    {
                        newProcess = new Process();
                        string valor = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ped_Amb_C8").Cells.Item(pVal.Row).Specific).Value;
                        info = new ProcessStartInfo(valor);
                        newProcess.StartInfo = info;
                        newProcess.Start();

                    }
                }

                if (EventEnum == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS && !pVal.BeforeAction)
                {
                    if (coluna == "Cmp_Amb_C2")
                    {
                        oNewItem = oForm.Items.Item("Det_Cmp");
                        oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                        string qtd = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C2").Cells.Item(pVal.Row).Specific).String;
                        decimal teste = Convert.ToDecimal(qtd);
                        string estoque = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C4").Cells.Item(pVal.Row).Specific).String;
                        decimal teste2 = Convert.ToDecimal(estoque);
                        if (qtd != "" && teste > teste2)
                        {
                            SBO_Application.MessageBox("Sem ítens sufucintes no estoque");
                        }
                    }
                }

                //Evento da grid de ambiente/análise crítica.
                if (!pVal.BeforeAction && pVal.ItemUID == "Apr_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ColUID == "#" && pVal.Row > 0)
                {

                    if (modificouAnsCritica)
                    {
                        SBO_Application.MessageBox("Vai atualizar");
                        Atualizar();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        modificouAnsCritica = false;
                    }

                    //Instancia a matriz de ambiente da aba apresentação.
                    oNewItem = oForm.Items.Item("Apr_Amb");
                    SAPbouiCOM.Matrix matrixApresentacao;
                    matrixApresentacao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Instancia a matriz de análise crítica da aba apresentação.
                    oNewItem = oForm.Items.Item("Ans_Amb");
                    SAPbouiCOM.Matrix matrixAnaliseCritica;
                    matrixAnaliseCritica = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Pega a coluna onde vai setar os valores pra o ambiente na matriz de análise crítica.
                    oColumnsAnaliseCritica = matrixAnaliseCritica.Columns;
                    oColumnAnaliseCritica = oColumnsAnaliseCritica.Item("Ans_Amb_C0");

                    //Pega o id do ambiente e a descrição do ambiente.
                    oEditItem = (SAPbouiCOM.EditText)matrixApresentacao.Columns.Item("Apr_Amb_C2").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText oItemGrid = (SAPbouiCOM.EditText)matrixApresentacao.Columns.Item("Apr_Amb_C0").Cells.Item(pVal.Row).Specific;
                    idAmbiente = int.Parse(oEditItem.String);
                    string nomeGrid = oItemGrid.String;

                    //Mostra na matriz de análise crítica qual ambiente selecionado.
                    oColumnAnaliseCritica.TitleObject.Caption = "Analise Crítica (" + nomeGrid + ")";
                    LoadGridAnaliseCritica();
                    countMatrixAnaliseCriticaAntes = matrixAnaliseCritica.RowCount;

                    if (matrixAnaliseCritica.RowCount == 0)
                    {
                        matrixAnaliseCritica.AddRow(1, 1);
                    }
                }
                //Evento da grid de ambiente/análise crítica.
                if (!pVal.BeforeAction && pVal.ItemUID == "Ans_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ColUID == "Ans_Amb_C0" && pVal.CharPressed == 9)
                {
                    oNewItem = oForm.Items.Item("Ans_Amb");
                    SAPbouiCOM.Matrix matrixAnaliseCritica;
                    matrixAnaliseCritica = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    oEditItem = (SAPbouiCOM.EditText)matrixAnaliseCritica.Columns.Item("Ans_Amb_C0").Cells.Item(matrixAnaliseCritica.RowCount).Specific;
                    string nome = oEditItem.String;

                    if (matrixAnaliseCritica.RowCount > 0 && nome != "")
                    {
                        matrixAnaliseCritica.AddRow(1, matrixAnaliseCritica.RowCount + 1);
                        ((SAPbouiCOM.EditText)matrixAnaliseCritica.Columns.Item("Ans_Amb_C0").Cells.Item(matrixAnaliseCritica.RowCount).Specific).Value = "";
                        ((SAPbouiCOM.EditText)matrixAnaliseCritica.Columns.Item("Ans_Amb_C1").Cells.Item(matrixAnaliseCritica.RowCount).Specific).Value = "";
                    }
                }

                if (pVal.ItemUID == "Ans_Amb" && pVal.ColUID == "Ans_Amb_C0" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    SBO_Application.MessageBox("Teste");
                    modificouAnsCritica = true;
                }

                if (!pVal.BeforeAction && pVal.ItemUID == "Fab_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ColUID == "Fab_#" && pVal.Row > 0)
                {
                    if (bGravouAvarias)
                    {
                        SBO_Application.MessageBox("Vai atualizar Avarias");
                        Atualizar();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        bGravouAvarias = false;
                    }

                    oNewItem = oForm.Items.Item("Fab_Amb");
                    SAPbouiCOM.Matrix matrixFabrica;
                    matrixFabrica = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    oNewItem = oForm.Items.Item("Ava_Amb");
                    SAPbouiCOM.Matrix matrixAvarias;
                    matrixAvarias = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    SAPbouiCOM.Columns oColumnsAvarias = null;
                    SAPbouiCOM.Column oColumnAvarias = null;
                    oColumnsAvarias = matrixAvarias.Columns;
                    oColumnAvarias = oColumnsAvarias.Item("Ava_Amb_C0");

                    //Pega o id do ambiente e a descrição do ambiente.
                    oEditItem = (SAPbouiCOM.EditText)matrixFabrica.Columns.Item("Fab_Amb_C4").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText oItemGrid = (SAPbouiCOM.EditText)matrixFabrica.Columns.Item("Fab_Amb_C0").Cells.Item(pVal.Row).Specific;
                    iRowAmbiente = int.Parse(oEditItem.String);
                    string nomeGrid = oItemGrid.String;

                    oColumnAvarias.TitleObject.Caption = "Descrição (" + nomeGrid + ")";
                    LoadGridAvarias();
                    countMatrixAvariasAntes = matrixAvarias.RowCount;

                    if (matrixAvarias.RowCount == 0)
                    {
                        matrixAvarias.AddRow(1, 1);
                    }
                }

                if (!pVal.BeforeAction && pVal.ItemUID == "Ava_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ColUID == "Ava_Amb_C0" && pVal.CharPressed == 9)
                {
                    oNewItem = oForm.Items.Item("Ava_Amb");
                    SAPbouiCOM.Matrix matrixAvarias;
                    matrixAvarias = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    oEditItem = (SAPbouiCOM.EditText)matrixAvarias.Columns.Item("Ava_Amb_C0").Cells.Item(matrixAvarias.RowCount).Specific;
                    string nome = oEditItem.String;

                    if (matrixAvarias.RowCount > 0 && nome != "")
                    {
                        matrixAvarias.AddRow(1, matrixAvarias.RowCount + 1);
                        ((SAPbouiCOM.EditText)matrixAvarias.Columns.Item("Ava_Amb_C0").Cells.Item(matrixAvarias.RowCount).Specific).Value = "";
                        ((SAPbouiCOM.EditText)matrixAvarias.Columns.Item("Ava_Amb_C1").Cells.Item(matrixAvarias.RowCount).Specific).Value = "";
                    }
                }

                if (pVal.ItemUID == "Ava_Amb" && pVal.ColUID == "Ava_Amb_C0" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    SBO_Application.MessageBox("Teste Avarias");
                    bGravouAvarias = true;
                }

                //Evento da grid de ambiente/conferência medições.
                if (!pVal.BeforeAction && pVal.ItemUID == "Med_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ColUID == "#" && pVal.Row > 0)
                {
                    //Instancia a matriz de ambiente da aba medições.
                    oNewItem = oForm.Items.Item("Med_Amb");
                    SAPbouiCOM.Matrix matrixMedicoes;
                    matrixMedicoes = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Instancia a matriz de conferência medições da aba medições.
                    oNewItem = oForm.Items.Item("Med_Cnf");
                    SAPbouiCOM.Matrix matrixConferenciaMedicoes;
                    matrixConferenciaMedicoes = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Pega a coluna onde vai setar os valores para o ambiente na matriz de conferência medições.
                    oColumnsConferenciaMedicoes = matrixConferenciaMedicoes.Columns;
                    oColumnConferenciaMedicoes = oColumnsConferenciaMedicoes.Item("med_Cnf_C1");

                    //Pega o id do ambiente e a descrição do ambiente.
                    oEditItem = (SAPbouiCOM.EditText)matrixMedicoes.Columns.Item("Med_Amb_C2").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText oItemGrid = (SAPbouiCOM.EditText)matrixMedicoes.Columns.Item("Med_Amb_C0").Cells.Item(pVal.Row).Specific;
                    iIdAmbienteMedicao = int.Parse(oEditItem.String);
                    string nomeGrid = oItemGrid.String;

                    //Mostra na matriz de conferência medições qual ambiente selecionado.
                    oColumnConferenciaMedicoes.TitleObject.Caption = "Conferente (" + nomeGrid + ")";
                    LoadGridConferenciaMedicao();
                    countMatrixConfMedAntes = matrixConferenciaMedicoes.RowCount;

                    if (matrixConferenciaMedicoes.RowCount == 0)
                    {
                        matrixConferenciaMedicoes.AddRow(1, 1);
                        //Projetistas - Grid Conferencia de Medicao
                        loadComboEmGrid("Med_Cnf", "med_Cnf_C1", projetistas);
                    }
                }
                //Evento da grid de ambiente/conferência medições.
                if (pVal.CharPressed == 9 && !pVal.BeforeAction && pVal.ItemUID == "Med_Cnf" && EventEnum == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ColUID == "med_Cnf_C1")
                {
                    oNewItem = oForm.Items.Item("Med_Cnf");
                    SAPbouiCOM.Matrix matrixConferenciaMedicao;
                    matrixConferenciaMedicao = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    oEditItem = (SAPbouiCOM.EditText)matrixConferenciaMedicao.Columns.Item("Med_Cnf_C0").Cells.Item(matrixConferenciaMedicao.RowCount).Specific;
                    string data = oEditItem.String;
                    SAPbouiCOM.ComboBox combo = (SAPbouiCOM.ComboBox)matrixConferenciaMedicao.Columns.Item("med_Cnf_C1").Cells.Item(matrixConferenciaMedicao.RowCount).Specific;
                    string nome = combo.Value;

                    if (matrixConferenciaMedicao.RowCount > 0 && data != "" && nome != "")
                    {
                        matrixConferenciaMedicao.AddRow(1, matrixConferenciaMedicao.RowCount + 1);
                        ((SAPbouiCOM.EditText)matrixConferenciaMedicao.Columns.Item("Med_Cnf_C0").Cells.Item(matrixConferenciaMedicao.RowCount).Specific).Value = "";
                        //Projetistas - Grid Conferencia de Medicao
                        RemoveValoresDeCombo(ref combo);
                        loadComboEmGrid("Med_Cnf", "med_Cnf_C1", projetistas);
                        ((SAPbouiCOM.ComboBox)matrixConferenciaMedicao.Columns.Item("med_Cnf_C1").Cells.Item(matrixConferenciaMedicao.RowCount).Specific).Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                        ((SAPbouiCOM.EditText)matrixConferenciaMedicao.Columns.Item("Med_Cnf_C2").Cells.Item(matrixConferenciaMedicao.RowCount).Specific).Value = "";

                    }
                }

                //Evento da grid de Entrega
                if (!pVal.BeforeAction && pVal.ItemUID == "Etg_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.ColUID == "Etg_#" && pVal.Row > 0)
                {

                    if (modificouPendecia)
                    {
                        SBO_Application.MessageBox("Vai atualizar Pendencia");
                        Atualizar();
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        modificouPendecia = false;
                    }

                    //Instancia a matriz de ambiente da aba apresentação.
                    oNewItem = oForm.Items.Item("Etg_Amb");
                    SAPbouiCOM.Matrix matrixEntrega;
                    matrixEntrega = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Instancia a matriz de análise crítica da aba apresentação.
                    oNewItem = oForm.Items.Item("Pen_Amb");
                    SAPbouiCOM.Matrix matrixPendencia;
                    matrixPendencia = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    //Pega a coluna onde vai setar os valores pra o ambiente na matriz de Pendencia.
                    oColumnsPendencia = matrixPendencia.Columns;
                    oColumnPendencia = oColumnsPendencia.Item("Pen_Amb_C0");

                    //Pega o id do ambiente e a descrição do ambiente.
                    oEditItem = (SAPbouiCOM.EditText)matrixEntrega.Columns.Item("Etg_Amb_C6").Cells.Item(pVal.Row).Specific;
                    SAPbouiCOM.EditText oItemGrid = (SAPbouiCOM.EditText)matrixEntrega.Columns.Item("Etg_Amb_C0").Cells.Item(pVal.Row).Specific;
                    idAmbientePendencia = int.Parse(oEditItem.String);
                    string nomeAmbiente = oItemGrid.String;

                    //Mostra na matriz de Pendencia qual ambiente selecionado.
                    oColumnPendencia.TitleObject.Caption = "Ambiente (" + nomeAmbiente + ")";
                    LoadGridPendencias();
                    countMatrixPendenciaAntes = matrixPendencia.RowCount;

                    if (matrixPendencia.RowCount == 0)
                    {
                        matrixPendencia.AddRow(1, 1);
                    }
                }
                //Evento da grid de ambiente/análise crítica.
                if (!pVal.BeforeAction && pVal.ItemUID == "Pen_Amb" && EventEnum == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.ColUID == "Pen_Amb_C0" && pVal.CharPressed == 9)
                {
                    oNewItem = oForm.Items.Item("Pen_Amb");
                    SAPbouiCOM.Matrix matrixPendencia;
                    matrixPendencia = ((SAPbouiCOM.Matrix)(oNewItem.Specific));

                    oEditItem = (SAPbouiCOM.EditText)matrixPendencia.Columns.Item("Pen_Amb_C0").Cells.Item(matrixPendencia.RowCount).Specific;
                    string nome = oEditItem.String;

                    if (matrixPendencia.RowCount > 0 && nome != "")
                    {
                        matrixPendencia.AddRow(1, matrixPendencia.RowCount + 1);
                        ((SAPbouiCOM.EditText)matrixPendencia.Columns.Item("Pen_Amb_C0").Cells.Item(matrixPendencia.RowCount).Specific).Value = "";
                        ((SAPbouiCOM.EditText)matrixPendencia.Columns.Item("Pen_Amb_C1").Cells.Item(matrixPendencia.RowCount).Specific).Value = "";
                    }
                }

                if (pVal.ItemUID == "Pen_Amb" && pVal.ColUID == "Pen_Amb_C0" && pVal.ItemChanged && !pVal.BeforeAction)
                {
                    SBO_Application.MessageBox("Teste Pendencia");
                    modificouPendecia = true;
                }
            }

            if (pVal.FormType == 651)
            {
                if (pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    oFormAtual = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    if (pVal.ItemUID == "Ata_Ativ" & pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED & !pVal.Before_Action)
                    {
                        abrirRelatorio("Ata de reuniao", oFormAtual.Items.Item("5").Specific.Value);
                    }

                    if (!pVal.Before_Action && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT) & pVal.ItemChanged & (pVal.ItemUID == "67"))
                    {
                        if (oFormAtual.Items.Item("67").Specific.Value == "M")
                        {
                            atividade.habilitaBotaoAta();
                        }
                        else
                        {
                            atividade.desabilitaBotaoAta();
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    {

                        if (pVal.Before_Action)
                        {
                            atividade = new Atividade(oFormAtual);
                        }

                        if (bBotaoAgendarFoiClicado)
                        {
                            oFormPai = SBO_Application.Forms.GetFormByTypeAndCount(320, iUltimoFormTypeCount_SalesOpportunities);

                            sSalesOpportunities_Id = ((SAPbouiCOM.EditText)oFormPai.Items.Item("74").Specific).Value;
                            sBPCode = ((SAPbouiCOM.EditText)oFormPai.Items.Item("9").Specific).Value;

                            oFormAtual = SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                            ((SAPbouiCOM.EditText)oFormAtual.Items.Item("9").Specific).Value = sBPCode;

                            bBotaoAgendarFoiClicado = false;
                        }
                    }
                }
            }


        }

        private void abrirRelatorio(string id, string param)
        {
            try
            {
                SAPbobsCOM.Recordset oRS = conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("SELECT MenuUID FROM OCMN WHERE Name = '" + id + "' AND Type = 'C'");

                SBO_Application.ActivateMenuItem(oRS.Fields.Item(0).Value.ToString());

                SAPbouiCOM.Form oForm1;
                oForm1 = SBO_Application.Forms.ActiveForm;
                
                if (!param.Equals(""))
                {
                    oForm1.Items.Item("1000003").Specific.String = param;
                    oForm1.Items.Item("1").Click();
                }

            }catch (Exception e)
            {
                SBO_Application.MessageBox(e.Message);
            }
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {

            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:

                    // Take care of terminating your AddOn application

                    SBO_Application.MessageBox("A Shut Down Event has been caught" + Constants.vbNewLine + "Terminating 'Order Form Manipulation' Add On...", 1, "Ok", "", "");

                    System.Environment.Exit(0);

                    break;
            }
        }

        private void LoadDataHoraMedicao(SAPbouiCOM.EditText data, SAPbouiCOM.EditText hora, SAPbouiCOM.EditText dataResumo, string idOOPR)
        {
            RecSet = null;
            string QryStr = null;

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "Select CONVERT(char,U_FLX_FB_MED_DATAMED,103), Convert(varchar(05),U_FLX_FB_MED_HORAMED,108) from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_MED] where U_FLX_FB_MED_IDOOPR =" + idOOPR;
            RecSet.DoQuery(QryStr);
            data.Value = RecSet.Fields.Item(0).Value.ToString();
            dataResumo.Value = RecSet.Fields.Item(0).Value.ToString();
            hora.Value = RecSet.Fields.Item(1).Value.ToString();
        }

        private void LoadResumo(SAPbouiCOM.EditText etvResp, SAPbouiCOM.EditText medResp, SAPbouiCOM.EditText elbIni, SAPbouiCOM.EditText elbFim, SAPbouiCOM.EditText elbResp, SAPbouiCOM.EditText verificacao, SAPbouiCOM.EditText verResp, SAPbouiCOM.EditText apsResp, SAPbouiCOM.EditText aprovacao, SAPbouiCOM.EditText aprResp, SAPbouiCOM.EditText pedido, SAPbouiCOM.EditText pedResp, SAPbouiCOM.EditText detIni, SAPbouiCOM.EditText detFim, SAPbouiCOM.EditText detResp, SAPbouiCOM.EditText fabIni, SAPbouiCOM.EditText fabFim, SAPbouiCOM.EditText fabResp, SAPbouiCOM.EditText montagem, SAPbouiCOM.EditText montagemFim, SAPbouiCOM.EditText montResp, SAPbouiCOM.EditText entrega, SAPbouiCOM.EditText etgResp, string idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "SELECT TOP 1 T2.Name as etvResp, T4.Name as medResp,CONVERT(char,T0.U_FLX_FB_ELB_INICIOR,103) as elaboracaoIni, CONVERT(char,T0.U_FLX_FB_ELB_TERMINR,103) as elaboracaoFim, T5.Name as elbResp, CONVERT(char,T0.U_FLX_FB_VRF_DATEVER,103) as verificacao, T0.U_FLX_FB_VRF_VERIFPO as verResp,T6.Name as apsResp,CONVERT(char,T0.U_FLX_FB_APR_DATAAPR,103)as aprovacao, T0.U_FLX_FB_APR_APROVAD as aprovPor,CONVERT(char,T0.U_FLX_FB_PED_DATE,103)as pedido, T0.U_FLX_FB_PED_SOLICIT as pedSolicitante,CONVERT(char,T0.U_FLX_FB_DET_INICIRE,103)as detIni,CONVERT(char,T0.U_FLX_FB_DET_TERMINO,103) as detFim, T7.Name as detResp,CONVERT(char,T0.U_FLX_FB_FAB_EXPEDIC,103)as fabIni,CONVERT(char,T0.U_FLX_FB_FAB_RECEBIM,103)as fabFim,T0.U_FLX_FB_FAB_CONFERE as fabConf,T0.U_FLX_FB_MTG_RESPONS as montResp,CONVERT(char,T0.U_FLX_FB_ENT_ENTREGA,103)as entrega, T0.U_FLX_FB_ENT_RESPONS as etgResp FROM OPR4 T0 left join OOPR T1 on T1.OpprId = T0.OprId left join [@FLX_FB_PRJ] T2 on T2.Code = T1.U_FLX_FB_ETV_RESP left join [@FLX_FB_MED] T3 on T3.U_FLX_FB_MED_IDOOPR = T0.OprId left join [@FLX_FB_PRJ] T4 on T4.Code = T3.U_FLX_FB_MED_PROJT left join [@FLX_FB_PRJ] T5 on T5.Code = T0.U_FLX_FB_ELB_PROJETI left join [@FLX_FB_PRJ] T6 on T6.Code = T1.U_FLX_FB_APS_PROJT left join [@FLX_FB_PRJ] T7 on T7.Code = T0.U_FLX_FB_DET_PROJETI WHERE OprId = " + idOOPR + "order by OprId asc";
            RecSet.DoQuery(QryStr);
            etvResp.Value = RecSet.Fields.Item(0).Value.ToString();
            medResp.Value = RecSet.Fields.Item(1).Value.ToString();
            elbIni.Value = RecSet.Fields.Item(2).Value.ToString();
            elbFim.Value = RecSet.Fields.Item(3).Value.ToString();
            elbResp.Value = RecSet.Fields.Item(4).Value.ToString();
            verificacao.Value = RecSet.Fields.Item(5).Value.ToString();
            verResp.Value = RecSet.Fields.Item(6).Value.ToString();
            apsResp.Value = RecSet.Fields.Item(7).Value.ToString();
            aprovacao.Value = RecSet.Fields.Item(8).Value.ToString();
            aprResp.Value = RecSet.Fields.Item(9).Value.ToString();
            pedido.Value = RecSet.Fields.Item(10).Value.ToString();
            pedResp.Value = RecSet.Fields.Item(11).Value.ToString();
            detIni.Value = RecSet.Fields.Item(12).Value.ToString();
            detFim.Value = RecSet.Fields.Item(13).Value.ToString();
            detResp.Value = RecSet.Fields.Item(14).Value.ToString();
            fabIni.Value = RecSet.Fields.Item(15).Value.ToString();
            fabFim.Value = RecSet.Fields.Item(16).Value.ToString();
            fabResp.Value = RecSet.Fields.Item(17).Value.ToString();
            //montagem.Value = RecSet.Fields.Item(18).Value.ToString();
            //montagemFim.Value = RecSet.Fields.Item(18).Value.ToString();
            montResp.Value = RecSet.Fields.Item(18).Value.ToString();
            entrega.Value = RecSet.Fields.Item(19).Value.ToString();
            etgResp.Value = RecSet.Fields.Item(20).Value.ToString();
        }

        private void LoadComboProjetistaCadastradoOOPR(SAPbouiCOM.ComboBox oComboEntrevista, SAPbouiCOM.ComboBox oComboMedicao, SAPbouiCOM.ComboBox oComboApresentacao, string idOOPR)
        {
            RecSet = projetista.trazerProjetistasOportunidade(idOOPR);
            RecSet.MoveFirst();
            int RecCount = RecSet.RecordCount;

            if (RecCount == 0)
            {
                oComboEntrevista.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                oComboApresentacao.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }

            for (int RecIndex = 0; RecIndex <= RecCount - 1; RecIndex++)
            {
                string valorComboEntrevista = RecSet.Fields.Item(0).Value.ToString();
                string valorComboApresentacao = RecSet.Fields.Item(1).Value.ToString();
                string valorComboMedicao = RecSet.Fields.Item(2).Value.ToString();

                oComboEntrevista.Select(valorComboEntrevista, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oComboApresentacao.Select(valorComboApresentacao, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oComboMedicao.Select(valorComboMedicao, SAPbouiCOM.BoSearchKey.psk_ByValue);

                RecSet.MoveNext();
            }

            RecSet = null;
            System.GC.Collect();
        }


        private void LoadGridConferenciaMedicao()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oDataTable").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oDataTable").ExecuteQuery("select * from [@FLX_FB_CONFMED] where U_FLX_FB_CONFMED_ID = '" + OpID + "' and U_FLX_FB_CONFMED_IDA = '" + iIdAmbienteMedicao.ToString() + "'");

            oItem = oForm.Items.Item("Med_Cnf");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void LoadGridAnaliseCritica()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oDataTableAnalise").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oDataTableAnalise").ExecuteQuery("SELECT * FROM [@FLX_FB_ANLCRI] where U_FLX_FB_ANLCRI_ID = '" + OpID + "'" + "and U_FLX_FB_ANLCRI_AMBI = '" + idAmbiente + "'");

            oItem = oForm.Items.Item("Ans_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void LoadGridAvarias()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oDataTableAvr").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oDataTableAvr").ExecuteQuery("select * from [@FLX_FB_AVR] where U_FLX_FB_AVR_IDOOPR = '" + OpID + "' and U_FLX_FB_AVR_IDAMBI = '" + iRowAmbiente.ToString() + "'");

            oItem = oForm.Items.Item("Ava_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void LoadGridItensComplementares()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oDataTableItc").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oDataTableItc").ExecuteQuery("select T1.ItemCode, T1.ItemName, T1.OnHand, T0.U_FLX_FB_ITC_QTD, T2.CardCode, T2.CardName, T0.U_FLX_FB_ITC_OBS, T0.Code, T0.U_FLX_FB_ITC_PRZETG, T0.U_FLX_FB_ITC_SOLICI, T0.U_FLX_FB_ITC_RECEB from [@FLX_FB_ITC] T0 inner join OITM T1 on T1.ItemCode = T0.U_FLX_FB_ITC_IDOITM inner join OCRD T2 on T2.CardCode = T0.U_FLX_FB_ITC_IDOCRD where T0.U_FLX_FB_ITC_IDOOPR = '" + OpID + "'");

            oItem = oForm.Items.Item("Det_Cmp");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Mon_Itc");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void LoadGridPendencias()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oDataTablePend").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oDataTablePend").ExecuteQuery("select * from [@FLX_FB_PEN] where U_FLX_FB_PEN_IDOOPR = '" + OpID + "'" + "and U_FLX_FB_PEN_IDAMB = '" + idAmbientePendencia + "'");

            oItem = oForm.Items.Item("Pen_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void LoadAmbientesInMatrix()
        {
            string OpID = null;

            oForm.DataSources.DataTables.Item("oMatrixDT").Clear();
            oItem = oForm.Items.Item("74");
            oEditItem = ((SAPbouiCOM.EditText)(oItem.Specific));
            OpID = oEditItem.Value;
            oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId WHERE T0.[OprId] = '" + OpID + "'");
            //oForm.DataSources.DataTables.Item("oMatrixDT").ExecuteQuery("SELECT T2.Name as elbProj, T3.Name as detProj, T1.[Num], T1.[Descript], T0.* FROM OPR4 T0 INNER JOIN OOIN T1 ON T1.Num = T0.IntId inner join [@FLX_FB_PRJ] T2 on T2.Code = T0.U_FLX_FB_ELB_PROJETI inner join [@FLX_FB_PRJ] T3 on T3.Code = T0.U_FLX_FB_DET_PROJETI WHERE T0.[OprId] = '" + OpID + "'");

            oItem = oForm.Items.Item("Med_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
            iRowAmbienteMedicao = 1;
            iIdAmbienteMedicao = int.Parse(((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Amb_C2").Cells.Item(iRowAmbienteMedicao).Specific).Value);

            oItem = oForm.Items.Item("Res_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Ela_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Ver_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Apr_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Apv_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Ped_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Det_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Fab_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Mon_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();

            oItem = oForm.Items.Item("Etg_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oItem.Specific));
            oMatrix.LoadFromDataSource();
        }

        private void GridComAnexo(SAPbouiCOM.EditText oEditItem)
        {
            try
            {
                if (oEditItem.Value.Equals("") || oEditItem.Value.Equals(null))
                {
                    ThreadStart starter = delegate { Anexar(oEditItem); };
                    Thread t = new Thread(starter);
                    t.SetApartmentState(ApartmentState.STA);
                    t.Start();
                    t.Join();
                }
                else
                {
                    newProcess = new Process();
                    info = new ProcessStartInfo(oEditItem.Value);
                    newProcess.StartInfo = info;
                    newProcess.Start();
                }
            }
            catch (Exception e)
            {
                SBO_Application.MessageBox(e.Message);
            }
        }

        public void Anexar(SAPbouiCOM.EditText oEditItem)
        {
            try
            {
                OpenFileDialog fDialog = new OpenFileDialog();
                fDialog.Title = "Anexar arquivo";
                fDialog.Filter = "(*.*)|*.*";
                fDialog.InitialDirectory = @"C:\";

                if (fDialog.ShowDialog() == DialogResult.OK)
                {
                    oEditItem.Value = fDialog.FileName.ToString();
                }
            }
            catch (Exception e)
            {
                SBO_Application.MessageBox(e.Message);
            }
        }

        public void Atualizar()
        {
            SAPbouiCOM.EditText oEdit = null;
            SAPbouiCOM.ComboBox oCombo = null;

            string etvProjetista = "";
            string apsProjetista = "";
            string medProjetista = "";
            string logradouro, numero, complemento, bairro, cidade, uf, pontoRef, etvData, etvHora, etvPrevisao, apsData, apsHora,
                medData, medHora, idOOPR = null;

            idOOPR = ((SAPbouiCOM.EditText)oForm.Items.Item("74").Specific).Value;

            oItem = oForm.Items.Item("End_Log");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            logradouro = oEdit.String;

            oItem = oForm.Items.Item("End_Num");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            numero = oEdit.String;

            oItem = oForm.Items.Item("End_Com");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            complemento = oEdit.String;

            oItem = oForm.Items.Item("End_Bai");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            bairro = oEdit.String;

            oItem = oForm.Items.Item("End_Cid");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            cidade = oEdit.String;

            oItem = oForm.Items.Item("End_UF");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            uf = oEdit.String;

            oItem = oForm.Items.Item("End_Ref");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            pontoRef = oEdit.String;

            //Abas Fases Entrevista
            oItem = oForm.Items.Item("Ent_Data");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            etvData = oEdit.String;

            oItem = oForm.Items.Item("Ent_Hora");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            etvHora = oEdit.String;

            if (upProjEnt)
            {
                oItem = oForm.Items.Item("Ent_Proj");
                oCombo = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                etvProjetista = oCombo.Value;
                upProjEnt = false;
            }
            else
                etvProjetista = "0";

            oItem = oForm.Items.Item("Ent_Prev");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            etvPrevisao = oEdit.String;

            //Abas - Fase Apresentação

            oItem = oForm.Items.Item("Apr_Data");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            apsData = oEdit.String;

            oItem = oForm.Items.Item("Apr_Hora");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            apsHora = oEdit.String;

            if (upProjAps)
            {
                oItem = oForm.Items.Item("Apr_Proj");
                oCombo = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                apsProjetista = oCombo.Value;
                upProjAps = false;
            }
            else
                apsProjetista = "0";

            //Aba - Fase Medição

            oItem = oForm.Items.Item("Med_Data");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            medData = oEdit.String;

            oItem = oForm.Items.Item("Med_Hora");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            medHora = oEdit.String;

            if (upProjMed)
            {
                oItem = oForm.Items.Item("Med_Proj");
                oCombo = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                medProjetista = oCombo.Value;
                upProjMed = false;
            }
            else
                medProjetista = "0";

            int etvAmbiente = 0;
            if (upEtvAmb)
            {
                oItem = oForm.Items.Item("Ent_Amb");
                oCombo = ((SAPbouiCOM.ComboBox)(oItem.Specific));
                //if (oCombo.ValidValues.Count == 0)
                if (oCombo.Value != "")
                {
                    try
                    {
                        etvAmbiente = Convert.ToInt32(oCombo.Value);
                    }
                    catch
                    {
                    }
                }
            }

            string descAmb = null;
            oItem = oForm.Items.Item("Ent_Det");
            oEdit = ((SAPbouiCOM.EditText)(oItem.Specific));
            descAmb = oEdit.String;


            //Conferencia de Medições
            oNewItem = oForm.Items.Item("Med_Cnf");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string data = null;
            string conferente = null;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (countMatrixConfMedAntes >= i)
                {
                    data = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Cnf_C0").Cells.Item(i).Specific).String;
                    conferente = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("med_Cnf_C1").Cells.Item(i).Specific).Value;

                    //string codeMed = GetIdConferenciaMedicaoParaUpdate(ListDataConfMed[i - 1].ToString(), Convert.ToInt32(idOOPR), ListConferenteCofMed[i - 1].ToString(), iIdAmbienteMedicao);
                    string codeMed = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Cnf_C2").Cells.Item(i).Specific).String; ;

                    if (data != "" && conferente != "")
                        UpdateConferenciaDeMedicao(codeMed, codeMed, int.Parse(idOOPR), data, conferente, iIdAmbienteMedicao);
                }
                else
                {
                    data = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Cnf_C0").Cells.Item(i).Specific).String;
                    conferente = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("med_Cnf_C1").Cells.Item(i).Specific).Value;
                    if (data != "" && conferente != "")
                        AddConferenciaDeMedicao(int.Parse(idOOPR), data, conferente, iIdAmbienteMedicao);
                }
            }

            //Adiciona o anexo da medição.
            AddAnexoMedicao(Convert.ToInt32(idOOPR));



            //Adiciona Avarias
            oNewItem = oForm.Items.Item("Ava_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string descricao = null;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                descricao = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ava_Amb_C0").Cells.Item(i).Specific).Value;
                if (countMatrixAvariasAntes >= i)
                {
                    string codeAvarias = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ava_Amb_C1").Cells.Item(i).Specific).Value; ; //GetIdAvariasParaUpdate(ListNomeAvarias[i - 1].ToString(), Convert.ToInt32(idOOPR), iIdAmbiente);

                    UpdateAvarias(codeAvarias, codeAvarias, int.Parse(idOOPR), descricao, iRowAmbiente);
                }
                else
                {
                    AddAvarias(int.Parse(idOOPR), descricao, iRowAmbiente);
                }
            }

            //Adiciona Pendencias
            oNewItem = oForm.Items.Item("Pen_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string descricaoPend = null;
            string idPendencia = null;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (countMatrixPendenciaAntes >= i)
                {
                    descricaoPend = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C0").Cells.Item(i).Specific).Value;
                    idPendencia = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C1").Cells.Item(i).Specific).Value;
                    UpdatePendencias(descricaoPend, idPendencia);
                }
                else
                {
                    descricaoPend = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C0").Cells.Item(i).Specific).Value;
                    AddPendencias(Convert.ToInt32(idOOPR), descricaoPend, idAmbientePendencia);
                }
            }

            //Adiciona Analise Critica
            oNewItem = oForm.Items.Item("Ans_Amb");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string txtAnexo = null;
            string idAnaliseCritica = null;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (countMatrixAnaliseCriticaAntes >= i)
                {
                    txtAnexo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ans_Amb_C0").Cells.Item(i).Specific).Value;
                    idAnaliseCritica = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ans_Amb_C1").Cells.Item(i).Specific).Value;
                    UpdateAnaliseCritica(txtAnexo, idAnaliseCritica);
                }
                else
                {
                    txtAnexo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ans_Amb_C0").Cells.Item(i).Specific).Value;
                    AddAnaliseCritica(Convert.ToInt32(idOOPR), idAmbiente, txtAnexo);
                }

                if (i == oMatrix.RowCount)
                {
                    LoadGridAnaliseCritica();
                }
            }

            //Adiciona Itens Complementares
            oNewItem = oForm.Items.Item("Det_Cmp");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string idOITM = null;
            string idOCRD = null;
            string qtd = null;
            string observacao = null;
            string idItensComplementares = null;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                if (countMatrixItensComplementaresAntes >= i)
                {
                    idOITM = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C0").Cells.Item(i).Specific).String;
                    idOCRD = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C6").Cells.Item(i).Specific).Value;
                    qtd = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C2").Cells.Item(i).Specific).Value;
                    observacao = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C5").Cells.Item(i).Specific).Value;
                    idItensComplementares = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C7").Cells.Item(i).Specific).String;
                    UpdateItensComplementares(idOITM, idOCRD, qtd, observacao, idItensComplementares);
                }
                else
                {
                    idOITM = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C0").Cells.Item(i).Specific).String;
                    idOCRD = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C6").Cells.Item(i).Specific).Value;
                    qtd = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C2").Cells.Item(i).Specific).Value;
                    observacao = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C5").Cells.Item(i).Specific).Value;

                    AddItensComplementares(Convert.ToInt32(idOOPR), idOITM, idOCRD, qtd, observacao);
                }

                if (i == oMatrix.RowCount)
                {
                    //LoadGridAnaliseCritica();
                }
            }


            //Update Itens Complementares (Aba Montagem)
            oNewItem = oForm.Items.Item("Mon_Itc");
            oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
            string prazoEntrega = null;
            string solicitante = null;
            bool recebido = false;
            int check = 0;

            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                prazoEntrega = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Itc_Amb_C5").Cells.Item(i).Specific).String;
                solicitante = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Itc_Amb_C6").Cells.Item(i).Specific).String;
                recebido = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Itc_Amb_C7").Cells.Item(i).Specific).Checked;
                idItensComplementares = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Itc_Amb_10").Cells.Item(i).Specific).String;

                if (recebido)
                {
                    check = 1;
                }
                else
                {
                    check = 0;
                }

                UpdateItensComplementares(idItensComplementares, prazoEntrega, solicitante, check);
                LoadGridItensComplementares();
            }


            AddOportunidadeVendas(Convert.ToInt32(idOOPR), etvProjetista, medProjetista, apsProjetista, etvAmbiente, descAmb);
        }

        public string[] VerificarSeExisteCadastroMedicao(int idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string retorno = "";
            string code = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = " declare @retorno INT declare @code INT select @retorno = (select U_FLX_FB_MED_IDOOPR from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_MED] where U_FLX_FB_MED_IDOOPR =" + idOOPR + ") if @retorno is null begin set @retorno = 0 select @retorno, @code end else begin select @retorno, Code from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_MED] where U_FLX_FB_MED_IDOOPR =" + idOOPR + " end";
            RecSet.DoQuery(QryStr);
            retorno = Convert.ToString(RecSet.Fields.Item(0).Value);
            code = Convert.ToString(RecSet.Fields.Item(1).Value);

            string[] valores = { retorno, code };

            return valores;
        }

        public string GetProxCodeMedicao()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_MED] order by Code desc) if @Numero is null begin set @Numero = 0000000 + 1 end SELECT case len(CAST(@Numero AS varchar(7))) WHEN 1 THEN '000000' + CAST(@Numero AS varchar(7)) WHEN 2 THEN '00000' + CAST(@Numero AS varchar(7)) WHEN 3 THEN '0000' + CAST(@Numero AS varchar(7)) WHEN 4 THEN '000' + CAST(@Numero AS varchar(7)) WHEN 5 THEN '00' + CAST(@Numero AS varchar(7)) WHEN 6 THEN '0' + CAST(@Numero AS varchar(7)) WHEN 7 THEN CAST(@Numero AS varchar(7)) END";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }

        public string GetProxCodeConferenciaMedicao()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_CONFMED] order by Code desc) if @Numero is null begin set @Numero = 0000000 + 1 end SELECT case len(CAST(@Numero AS varchar(7))) WHEN 1 THEN '000000' + CAST(@Numero AS varchar(7)) WHEN 2 THEN '00000' + CAST(@Numero AS varchar(7)) WHEN 3 THEN '0000' + CAST(@Numero AS varchar(7)) WHEN 4 THEN '000' + CAST(@Numero AS varchar(7)) WHEN 5 THEN '00' + CAST(@Numero AS varchar(7)) WHEN 6 THEN '0' + CAST(@Numero AS varchar(7)) WHEN 7 THEN CAST(@Numero AS varchar(7)) END";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }

        public string GetProxCodeAvarias()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_AVR] order by Code desc) if @Numero is null begin set @Numero = 0000000 + 1 end SELECT case len(CAST(@Numero AS varchar(7))) WHEN 1 THEN '000000' + CAST(@Numero AS varchar(7)) WHEN 2 THEN '00000' + CAST(@Numero AS varchar(7)) WHEN 3 THEN '0000' + CAST(@Numero AS varchar(7)) WHEN 4 THEN '000' + CAST(@Numero AS varchar(7)) WHEN 5 THEN '00' + CAST(@Numero AS varchar(7)) WHEN 6 THEN '0' + CAST(@Numero AS varchar(7)) WHEN 7 THEN CAST(@Numero AS varchar(7)) END";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }

        public string GetQtdEmEstoque(string itemCode)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string qtdEstoque = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "SELECT T0.[OnHand] FROM OITM T0 WHERE T0.ItemCode = '" + itemCode + "'";
            RecSet.DoQuery(QryStr);
            qtdEstoque = Convert.ToString(RecSet.Fields.Item(0).Value);

            return qtdEstoque;
        }

        public string GetProxCodePendencias()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_PEN] order by Code desc) if @Numero is null begin set @Numero = 0000000 + 1 end SELECT case len(CAST(@Numero AS varchar(7))) WHEN 1 THEN '000000' + CAST(@Numero AS varchar(7)) WHEN 2 THEN '00000' + CAST(@Numero AS varchar(7)) WHEN 3 THEN '0000' + CAST(@Numero AS varchar(7)) WHEN 4 THEN '000' + CAST(@Numero AS varchar(7)) WHEN 5 THEN '00' + CAST(@Numero AS varchar(7)) WHEN 6 THEN '0' + CAST(@Numero AS varchar(7)) WHEN 7 THEN CAST(@Numero AS varchar(7)) END";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }

        public string GetIdConferenciaMedicaoParaUpdate(string data, int idOOPR, string conferente, int idAmbiente)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string code = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "SELECT top 1 [Code] FROM [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_CONFMED] where U_FLX_FB_CONFMED_DAT = '" + data + "' and U_FLX_FB_CONFMED_PRJ = '" + conferente +
                "' and U_FLX_FB_CONFMED_ID ='" + idOOPR + "' and U_FLX_FB_CONFMED_IDA ='" + iIdAmbienteMedicao + "'  order by CreateDate desc";
            RecSet.DoQuery(QryStr);
            code = Convert.ToString(RecSet.Fields.Item(0).Value);

            return code;
        }

        public string GetIdAvariasParaUpdate(string descricao, int idOOPR, int idAmbiente)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string code = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "SELECT top 1 [Code] FROM [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_AVR] where U_FLX_FB_AVR_DESC = '" + descricao + "' and U_FLX_FB_AVR_IDOOPR ='" + idOOPR +
                "' and U_FLX_FB_AVR_IDAMBI ='" + idAmbiente + "'  order by CreateDate desc";
            RecSet.DoQuery(QryStr);
            code = Convert.ToString(RecSet.Fields.Item(0).Value);

            return code;
        }

        public string GetIdPendenciasParaUpdate(string descricao, int idOOPR)
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string code = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "SELECT top 1 [Code] FROM [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_PEN] where U_FLX_FB_PEN_DESC = '" + descricao + "' and U_FLX_FB_PEN_IDOOPR ='" + idOOPR + "'  order by CreateDate desc";
            RecSet.DoQuery(QryStr);
            code = Convert.ToString(RecSet.Fields.Item(0).Value);

            return code;
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            //Matrix de Conferencia de Medicao
            if (pVal.MenuUID == "AddRowMatrixConf" && pVal.BeforeAction == true)
            {
                oNewItem = oForm.Items.Item("Med_Cnf");
                oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
                oColumns = oMatrix.Columns;
                int numeroLinhas = oMatrix.RowCount;
                oMatrix.AddRow(1, numeroLinhas + 1);
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Med_Cnf_C0").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                //((SAPbouiCOM.EditText)oMatrix.Columns.Item("med_Cnf_C1").Cells.Item(oMatrix.RowCount).Specific).Value = "";
            }

            //Matrix de Conferencia de Avarias
            if (pVal.MenuUID == "AddRowMatrixAvr" && pVal.BeforeAction == true)
            {
                oNewItem = oForm.Items.Item("Ava_Amb");
                oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
                oColumns = oMatrix.Columns;
                int numeroLinhas = oMatrix.RowCount;
                oMatrix.AddRow(1, numeroLinhas + 1);
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Ava_Amb_C0").Cells.Item(oMatrix.RowCount).Specific).Value = "";
            }

            //Matrix de Conferencia de Avarias
            if (pVal.MenuUID == "AddRowMatrixPend" && pVal.BeforeAction == true)
            {
                oNewItem = oForm.Items.Item("Pen_Amb");
                oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
                oColumns = oMatrix.Columns;
                int numeroLinhas = oMatrix.RowCount;
                oMatrix.AddRow(1, numeroLinhas + 1);
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Pen_Amb_C0").Cells.Item(oMatrix.RowCount).Specific).Value = "";
            }

            //Matrix de Itens Complementares
            if (pVal.MenuUID == "AddRowMatrixIt" && pVal.BeforeAction == true)
            {
                oNewItem = oForm.Items.Item("Det_Cmp");
                oMatrix = ((SAPbouiCOM.Matrix)(oNewItem.Specific));
                oColumns = oMatrix.Columns;
                int numeroLinhas = oMatrix.RowCount;
                oMatrix.AddRow(1, numeroLinhas + 1);
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C0").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C1").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C4").Cells.Item(oMatrix.RowCount).Specific).Value = "0.000000";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C2").Cells.Item(oMatrix.RowCount).Specific).Value = "0.000000";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C6").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C3").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C5").Cells.Item(oMatrix.RowCount).Specific).Value = "";
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("Cmp_Amb_C7").Cells.Item(oMatrix.RowCount).Specific).Value = "";

                oMatrix.FlushToDataSource();
            }
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;


            if (eventInfo.ItemUID == "Med_Cnf")
            {
                if ((eventInfo.BeforeAction == true))
                {
                    SAPbouiCOM.MenuItem oMenuItem = null;
                    SAPbouiCOM.Menus oMenus = null;

                    try
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackage = null;
                        oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "AddRowMatrixConf";
                        oCreationPackage.String = "Adicionar";
                        oCreationPackage.Enabled = true;


                        oMenuItem = SBO_Application.Menus.Item("1280"); // Data'
                        oMenus = oMenuItem.SubMenus;
                        oMenus.AddEx(oCreationPackage);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    SAPbouiCOM.MenuItem oMenuItem = null;
                    SAPbouiCOM.Menus oMenus = null;


                    try
                    {
                        SBO_Application.Menus.RemoveEx("AddRowMatrixConf");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }

            //Matrix de Avarias

            if (eventInfo.ItemUID == "Ava_Amb")
            {
                if ((eventInfo.BeforeAction == true))
                {
                    SAPbouiCOM.MenuItem oMenuItemAvr = null;
                    SAPbouiCOM.Menus oMenusAvr = null;

                    try
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackageAvr = null;
                        oCreationPackageAvr = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        oCreationPackageAvr.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackageAvr.UniqueID = "AddRowMatrixAvr";
                        oCreationPackageAvr.String = "Adicionar";
                        oCreationPackageAvr.Enabled = true;


                        oMenuItemAvr = SBO_Application.Menus.Item("1280"); // Data'
                        oMenusAvr = oMenuItemAvr.SubMenus;
                        oMenusAvr.AddEx(oCreationPackageAvr);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        SBO_Application.Menus.RemoveEx("AddRowMatrixAvr");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }


            //Matrix de Pendencias

            if (eventInfo.ItemUID == "Pen_Amb")
            {
                if ((eventInfo.BeforeAction == true))
                {
                    SAPbouiCOM.MenuItem oMenuItemAvr = null;
                    SAPbouiCOM.Menus oMenusAvr = null;

                    try
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackageAvr = null;
                        oCreationPackageAvr = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        oCreationPackageAvr.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackageAvr.UniqueID = "AddRowMatrixPend";
                        oCreationPackageAvr.String = "Adicionar";
                        oCreationPackageAvr.Enabled = true;


                        oMenuItemAvr = SBO_Application.Menus.Item("1280"); // Data'
                        oMenusAvr = oMenuItemAvr.SubMenus;
                        oMenusAvr.AddEx(oCreationPackageAvr);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        SBO_Application.Menus.RemoveEx("AddRowMatrixPend");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }

            //Matrix de Itens Complementares

            if (eventInfo.ItemUID == "Det_Cmp")
            {
                if ((eventInfo.BeforeAction == true))
                {
                    SAPbouiCOM.MenuItem oMenuItemIt = null;
                    SAPbouiCOM.Menus oMenusIt = null;

                    try
                    {
                        SAPbouiCOM.MenuCreationParams oCreationPackageIt = null;
                        oCreationPackageIt = ((SAPbouiCOM.MenuCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));

                        oCreationPackageIt.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackageIt.UniqueID = "AddRowMatrixIt";
                        oCreationPackageIt.String = "Adicionar";
                        oCreationPackageIt.Enabled = true;


                        oMenuItemIt = SBO_Application.Menus.Item("1280"); // Data'
                        oMenusIt = oMenuItemIt.SubMenus;
                        oMenusIt.AddEx(oCreationPackageIt);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                else
                {
                    try
                    {
                        SBO_Application.Menus.RemoveEx("AddRowMatrixIt");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
            }
        }

        private void AddChooseFromList()
        {
            try
            {
                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                // SAPbouiCOM.Conditions oConditions;
                //SAPbouiCOM.Condition oCondition;

                //  Filtro pelo cadastro de itens.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFL1";

                // oConditions = oCFL.GetConditions();
                //oCondition = oConditions.Add();

                oCFL = oCFLs.Add(oCFLCreationParams);

            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        private void AddChooseFromList2()
        {
            try
            {
                SAPbouiCOM.Conditions oConditions;
                SAPbouiCOM.Condition oCondition;

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //  Filtro pelo cadastro de itens.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFL2";

                oConditions = new SAPbouiCOM.Conditions();
                oCondition = oConditions.Add();
                oCondition.BracketOpenNum = 1;
                oCondition.Alias = "CardType";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "S";
                oCondition.BracketCloseNum = 1;

                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oConditions);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void AddAnaliseCritica(int idOOPR, int idAmbiente, string descAnexo)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodeAnaliseCritica();

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_ANLCRI");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_ANLCRI_ID", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_ANLCRI_AMBI", idAmbiente);
                oGeneralData.SetProperty("U_FLX_FB_ANLCRI_ANEX", descAnexo);

                oGeneralParams = oGeneralService.Add(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdateAnaliseCritica(string descAnexo, string pkAnaliseCritica)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_ANLCRI");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", pkAnaliseCritica);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_FLX_FB_ANLCRI_ANEX", descAnexo);

                oGeneralService.Update(oGeneralData);
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public string GetProxCodeAnaliseCritica()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_ANLCRI] order by DocEntry desc) if @Numero is null begin set @Numero = 1 end select @Numero";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }

        public void AddItensComplementares(int idOOPR, string idOITM, string idOCRD, string qtd, string observacao)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;
            string proxCode = GetProxCodeItensComplementares();

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_ITC_SOLICI");
                oGeneralData = ((SAPbobsCOM.GeneralData)(oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)));
                oGeneralData.SetProperty("Code", proxCode);
                oGeneralData.SetProperty("Name", proxCode);
                oGeneralData.SetProperty("U_FLX_FB_ITC_IDOOPR", idOOPR);
                oGeneralData.SetProperty("U_FLX_FB_ITC_IDOITM", idOITM);
                oGeneralData.SetProperty("U_FLX_FB_ITC_IDOCRD", idOCRD);
                oGeneralData.SetProperty("U_FLX_FB_ITC_QTD", qtd);
                oGeneralData.SetProperty("U_FLX_FB_ITC_OBS", observacao);

                oGeneralParams = oGeneralService.Add(oGeneralData);
                //LoadGridItensComplementares();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdateItensComplementares(string idOITM, string idOCRD, string qtd, string observacao, string pkItensComplementares)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_ITC_SOLICI");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", pkItensComplementares);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_FLX_FB_ITC_IDOITM", idOITM);
                oGeneralData.SetProperty("U_FLX_FB_ITC_IDOCRD", idOCRD);
                oGeneralData.SetProperty("U_FLX_FB_ITC_QTD", qtd);
                oGeneralData.SetProperty("U_FLX_FB_ITC_OBS", observacao);

                oGeneralService.Update(oGeneralData);
                //LoadGridItensComplementares();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public void UpdateItensComplementares(string pkItensComplementares, string prazoEntreg, string solicitante, int recebido)
        {
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralData oGeneralData = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.CompanyService oCompanyService = null;

            try
            {
                oCompanyService = conexao.getOCompany().GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("FLX_FB_ITC_SOLICI");
                oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("Code", pkItensComplementares);
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                oGeneralData.SetProperty("U_FLX_FB_ITC_PRZETG", prazoEntreg);
                oGeneralData.SetProperty("U_FLX_FB_ITC_SOLICI", solicitante);
                oGeneralData.SetProperty("U_FLX_FB_ITC_RECEB", recebido);

                oGeneralService.Update(oGeneralData);
                LoadGridItensComplementares();
            }
            catch (Exception ex)
            {
                SBO_Application.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public string GetProxCodeItensComplementares()
        {
            SAPbobsCOM.Recordset RecSet = null;
            string QryStr = null;
            string proxCod = "";

            RecSet = ((SAPbobsCOM.Recordset)(conexao.getOCompany().GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)));
            QryStr = "DECLARE @Numero AS INT SELECT @Numero = (select top 1 cast (Code as INT) + 1 from [SBO_SEA_Design_Prod].[dbo].[@FLX_FB_ITC] order by Code desc) if @Numero is null begin set @Numero = 1 end select @Numero";
            RecSet.DoQuery(QryStr);
            proxCod = Convert.ToString(RecSet.Fields.Item(0).Value);

            return proxCod;
        }
    }
}
