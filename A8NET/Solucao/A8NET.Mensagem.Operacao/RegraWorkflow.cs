using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace A8NET.Mensagem.Operacao
{
    class RegraWorkflow : IDisposable
    {

       protected Data.DsParametrizacoes _DsCache;

        #region <<< Construtores >>>
        public RegraWorkflow(Data.DsParametrizacoes dataSetCache)
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            _DsCache = dataSetCache;
            #endregion
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~RegraWorkflow()
        {
            this.Dispose();
        }
        #endregion

        #region <<< enumeradores >>>
        public enum enumMacroFuncao
        {
            Confirmacao = 1,
            Conciliacao = 2,
            Liberacao = 3
        }
        public enum enumFuncaoSistema
        {
            Confirmar = 1,
            Conciliar = 2,
            Liberar = 3,
            Pagar = 4,
            Concordar = 5,
            PagarContingencia = 6,
            PagarSTR = 7,
            PagarBACEN = 8,
            Discordar = 9,
            LberarCancelamentoEspecificacaoCompromissada = 10,
            LiberarLiquidacaoLeilaoBMA = 11,
            LiberarPagamento = 12,
            Receber = 13,
            ConfirmarReativacao = 14,
            LiberarReativacao = 15,
            IntegracaoCC = 16,
            LiberarCAM0054 = 17
        }
        #endregion

        public bool VerificarRegraAutomatica(udtOperacao udtOperacaoRecebida, enumFuncaoSistema enumFuncaoSistema, ref int codigoRetornoVerificacao)
        {
            codigoRetornoVerificacao = 0;
            string UsuarioOperacao = string.Empty;

            try
            {
                //'Obter o grupo Usuario
                //Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso") -- lê de um metodo dessa classe q retorna string
                //Set xmlControleAcesso = CreateObject("MSXML2.DOMDocument.4.0") -- carrega string nesse domdocument

                //strUsuarioOperacao = UCase(Mid$(xmlOperacao.documentElement.selectSingleNode("//CO_USUA_CADR_OPER").Text, 1, 8))
                if (udtOperacaoRecebida.RowOperacao["CO_USUA_CADR_OPER"].ToString().Length <= 8) UsuarioOperacao = udtOperacaoRecebida.RowOperacao["CO_USUA_CADR_OPER"].ToString();
                else UsuarioOperacao = udtOperacaoRecebida.RowOperacao["CO_USUA_CADR_OPER"].ToString().Substring(0, 8);
    
                #region <<< ObterGruposAcessoDadosPorUsuario >>>
                //strControleAcesso = objControleAcesso.ObterGruposAcessoDadosPorUsuario(strUsuarioOperacao, plngCodigoRetornoVerificacao)

                //CRIAR UMA FUNCAO ObterGruposAcessoDadosPorUsuario QUE PADRONIZE RETORNO DATASET/STRINGXML INDEPENDENTE DE LER DO MBS/CACHE OU DA TABELA MBSGRUPO
                _DsCache.CaseSensitive = false;
                _DsCache.MBS_GRUPO.DefaultView.RowFilter = string.Format("cd_usr ='{0}'", UsuarioOperacao);
                #endregion

                //If strControleAcesso = vbNullString Then
                if (_DsCache.MBS_GRUPO.DefaultView.Count == 0)
                {
                    //Sai sem erro, pois irá informar que o grupo não está cadastrado para o usuário
                    if (codigoRetornoVerificacao == 0)
                    {
                        codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.CadastroGrupoUsuario;
                    }
                    return false;
                }

                _DsCache.TB_PARM_FCAO_SIST.DefaultView.RowFilter = string.Format(@"TP_OPER      ={0} 
                                                                               AND CO_FCAO_SIST ={1} 
                                                                               AND CO_EMPR      ={2} 
                                                                               AND TP_BKOF      ={3}", 
                                                                               udtOperacaoRecebida.RowOperacao["TP_OPER"].ToString(), 
                                                                               (int)enumFuncaoSistema, 
                                                                               udtOperacaoRecebida.RowOperacao["CO_EMPR"].ToString(),
                                                                               udtOperacaoRecebida.RowOperacao["TP_BKOF"].ToString());
                
                if (_DsCache.TB_PARM_FCAO_SIST.DefaultView.Count == 0) 
                {   // não há parametrização cadastrada
                    codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.RegraWorkflow;
                    return false;
                }

                if (_DsCache.TB_PARM_FCAO_SIST.DefaultView[0]["IN_FCAO_SIST_AUTM"].ToString() == ((int)Comum.Comum.EnumInidicador.Sim).ToString())
                {   // Automatica = Sim
                    return true;
                }
                else
                {   // Automático = Não
                    codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.RegraWorkflow;
                    return false;
                }

            }

            catch
            {
                throw;
            }

        }

        public void ObterFuncaoSistemaStatus(udtOperacao udtOperacaoRecebida, 
                                             enumMacroFuncao macroFuncao, 
                                         ref enumFuncaoSistema funcaoSistema, 
                                         ref Comum.Comum.EnumStatusOperacao? statusOperacao)
        {
            try
            {

                #region >>> Default >>>
                switch (macroFuncao)
                {
                    case enumMacroFuncao.Confirmacao: 
                        funcaoSistema = enumFuncaoSistema.Confirmar;
                        statusOperacao = Comum.Comum.EnumStatusOperacao.ConcordanciaAutomatica;
                        break;
                    case enumMacroFuncao.Conciliacao: 
                        funcaoSistema = enumFuncaoSistema.Conciliar;
                        //statusOperacao da conciliacao é definido dentro da funcao de conciliacao
                        break;
                    case enumMacroFuncao.Liberacao:   
                        funcaoSistema = enumFuncaoSistema.Liberar;
                        statusOperacao = Comum.Comum.EnumStatusOperacao.LiberadaAutomatica;
                        break;
                }
                #endregion

                #region >>> Regras específicas >>>
                
                // Regras se TipoOperacao = InformaConfirmacaoContrCamaraTelaCega então muda para funcaosistema específica LiberarCAM0054
                if (udtOperacaoRecebida.TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancarioEletronico)
                {
                    funcaoSistema = RegraWorkflow.enumFuncaoSistema.LiberarCAM0054;
                    return;
                }

                // Regras se TipoSolicitacao = Reativacao
                if (int.Parse(udtOperacaoRecebida.RowOperacao["TP_SOLI"].ToString()) == (int)Comum.Comum.enumTipoSolicitacao.Reativacao)
                {
                    if (macroFuncao == enumMacroFuncao.Confirmacao)
                    {
                        funcaoSistema = RegraWorkflow.enumFuncaoSistema.ConfirmarReativacao;
                        statusOperacao = Comum.Comum.EnumStatusOperacao.ConcordanciaReativacaoAutomatica;
                        return;
                    }
                    else if (macroFuncao == enumMacroFuncao.Liberacao)
                    {
                        funcaoSistema = RegraWorkflow.enumFuncaoSistema.LiberarReativacao;
                        statusOperacao = Comum.Comum.EnumStatusOperacao.LiberadaReativacaoAutomatica;
                        return;
                    }
                }
                #endregion

            }

            catch (Exception ex)
            {
                throw new Exception("Operacao.VerificarRegraAutomatica() - " + ex.ToString());
            }

        }



    }
}
