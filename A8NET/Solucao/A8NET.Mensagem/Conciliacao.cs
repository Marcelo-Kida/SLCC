using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using A8NET.Comum;
using A8NET.Data.DAO;
using A8NET.Data;

namespace A8NET.Mensagem
{
    public class Conciliacao : Mensagem
    {
        #region <<< Variáveis >>>
        MensagemSpbDAO.EstruturaMensagemSPB _EstruturaMensagemSpbR0 = new MensagemSpbDAO.EstruturaMensagemSPB();
        DsTB_MESG_RECB_ENVI_SPB _DsMensagemSPB;
        DateTime _DataOperacao = new DateTime();
        string _StatusOperacaoSeConciliacaoOK = string.Empty;
        string _StatusMensagemSeConciliacaoOK = string.Empty;
        Comum.Comum.EnumStatusOperacao[] _ListaStatusOperacao = null;
        string _NumeroComando = string.Empty;
        bool _ConciliacaoOK = false;
        long _NumeroSequenciaConciliacao = 0;
        long _NumeroSequenciaOperacao = 0;
        long _NumeraOperacaoCambial2 = 0;
        int _TipoOperacao = 0;
        int _TipoAcao = 0;
        int _CodigoTextXML = 0;
        int _RetornoConciliacao = 0;
        #endregion

        #region <<< Construtores >>>
        public Conciliacao()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            _MensagemSpbDATA = new MensagemSpbDAO();
            _OperacaoDATA = new OperacaoDAO();
            _ConciliacaoDATA = new ConciliacaoDAO();
            #endregion
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~Conciliacao()
        {
            this.Dispose();
        }
        #endregion

        #region <<< ProcessaMensagem >>>
        public override void ProcessaMensagem(string nomeFila, string mensagemRecebida)
        {
        }
        #endregion

        #region <<< VerificaConciliacao - OVERLOAD Operacao com MensagemSPB >>>
        public bool VerificaConciliacao(OperacaoDAO.EstruturaOperacao entidadeOperacao,
                                    ref XmlDocument xmlOperacao,  
                                    ref int codigoRetornoVerificacao)
        {
            bool AtualizaStatusMensagemSPB = false;
            bool AppendaMensagemSPBConciliada = false;
            string CodigoMensagemSPBAConciliar;

            try
            {
                // Carrega variáveis
                _TipoOperacao = int.Parse(entidadeOperacao.TP_OPER.ToString());
                _NumeroSequenciaOperacao = int.Parse(entidadeOperacao.NU_SEQU_OPER_ATIV.ToString());
                _NumeroComando = entidadeOperacao.NU_COMD_OPER.ToString();
                long.TryParse(entidadeOperacao.NR_OPER_CAMB_2.ToString(), out _NumeraOperacaoCambial2);
                _DataOperacao = Comum.Comum.ConvertDtToDateTime(entidadeOperacao.DT_OPER_ATIV.ToString());
                CodigoMensagemSPBAConciliar = ObterCodigoMensagemSPBAConciliar((Comum.Comum.EnumTipoOperacao)_TipoOperacao);

                #region >>> Conciliação de Operação de Registro Interbancario com BMC0015 >>>
                if (_TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.RegistroOperacaoInterbancarioEletronico)
                {
                    if (_ConciliacaoDATA.ConciliarOperacaoBMC0015(_NumeroSequenciaOperacao
                                                               ,Comum.Comum.EnumStatusOperacao.AConciliarRegistro
                                                               ,Comum.Comum.EnumStatusMensagem.AConciliar
                                                               ,ref _NumeroSequenciaConciliacao
                                                               ,ref _RetornoConciliacao) == true)
                    {
                        _ConciliacaoOK = true;
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.LiberadaAutomatica).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.Conciliada).ToString();
                        _TipoAcao = base.ObterTipoAcaoEnvioMensagemSPB(_TipoOperacao);
                        AtualizaStatusMensagemSPB = true;
                        AppendaMensagemSPBConciliada = true;
                    }
                    else
                    { 
                        if (_RetornoConciliacao == 0) codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.SemMensagemBMC0015;
                        else if (_RetornoConciliacao >= 2) codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.VariasMensagensBMC0015;
                    }
                }
                #endregion
  
                #region >>> Conciliação de diversos TipoOperacao com MensagensSPB, por RegistroOperaçãoCambial e RegistroOperaçãoCambial2 >>>
                else if (_TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemTelaCega
                      || _TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemCamara
                      || _TipoOperacao == (int)Comum.Comum.EnumTipoOperacao.InformaConfirmacaoOperArbitragemParceiroPais) 
                {
                    // Verifica Conciliacao
                    if (_ConciliacaoDATA.ConciliarComMensagemSPB(_NumeroComando,
                                                                 _NumeraOperacaoCambial2,
                                                                 _DataOperacao,
                                                                 Comum.Comum.EnumStatusMensagem.EnviadaLegado,
                                                                 CodigoMensagemSPBAConciliar,
                                                             ref _DsMensagemSPB,
                                                             ref _RetornoConciliacao) == true)
                    {
                        _ConciliacaoOK = true;
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.LiberadaAutomatica).ToString();
                    }

                }
                #endregion

                // Se Conciliacao OK então atualiza o status da Operaçao e MensagemSPB conciliadas
                if (_ConciliacaoOK == true)
                {
                    // Atualiza status Operacao
                    entidadeOperacao.CO_ULTI_SITU_PROC = int.Parse(_StatusOperacaoSeConciliacaoOK);
                    base.AlterarStatusOperacao((int)_NumeroSequenciaOperacao, int.Parse(_StatusOperacaoSeConciliacaoOK), 0, _TipoAcao);

                    // Obter Conciliacao para ler dados da MensagemSPB conciliada
                    _ConciliacaoDATA.SelecionarConciliacaoOperacao(_NumeroSequenciaConciliacao);
                    if (_ConciliacaoDATA.Itens.Length == 0)
                    {
                        codigoRetornoVerificacao = (int)Comum.Comum.enumJustificativa.SemMensagemBMC0015;
                        return false;
                    }

                    // Atualiza status Mensagem
                    if (AtualizaStatusMensagemSPB == true)
                    {
                        _MensagemSpbDATA.SelecionarMensagensPorControleIF(_ConciliacaoDATA.TB_CNCL_OPER_ATIV.NU_CTRL_IF.ToString());
                        _EstruturaMensagemSpbR0 = _MensagemSpbDATA.ObterMensagemLidaUnica(DateTime.Parse(_ConciliacaoDATA.TB_CNCL_OPER_ATIV.DH_REGT_MESG_SPB.ToString()), int.Parse(_ConciliacaoDATA.TB_CNCL_OPER_ATIV.NU_SEQU_CNTR_REPE.ToString()));
                        _EstruturaMensagemSpbR0.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                        _EstruturaMensagemSpbR0.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                        _EstruturaMensagemSpbR0.CO_ULTI_SITU_PROC = int.Parse(_StatusMensagemSeConciliacaoOK);
                        base.AlterarStatusMensagemSPB(ref _EstruturaMensagemSpbR0,
                                                          (Comum.Comum.EnumStatusMensagem)int.Parse(_StatusMensagemSeConciliacaoOK));
                    }

                    // Appenda na Operacao o XML da Mensagem conciliada, necessário devido Tags de Entrada da Regra de Transporte da MensagemSPB no A7NET
                    if (AppendaMensagemSPBConciliada == true)
                    {
                        _CodigoTextXML = int.Parse(_EstruturaMensagemSpbR0.CO_TEXT_XML.ToString());
                        Comum.Comum.AppendNode(ref xmlOperacao, "MESG", "MESG", base.SelecionarTextoBase64(_CodigoTextXML));
                    }

                    return true;
                }
                else
                {
                    return false;
                }

            }
	        
            catch (Exception ex)
	        {
                throw new Exception("Conciliacao.VerificaConciliacao() - OVERLOAD Operacao com MensagemSPB" + ex.ToString());
	        }
        }
        #endregion

        #region <<< VerificaConciliacao - OVERLOAD MensagemSPB com Operacao >>>
        public bool VerificaConciliacao(ref MensagemSpbDAO.EstruturaMensagemSPB entidadeMensagem,
                                            string indicadorAceite,
                                        ref long numeroSequenciaOperacao,
                                            bool somenteConsulta)
        {
            
            try
            {
                // se numeroSequenciaOperacao != 0 significa que esta conciliacao já foi verificada antes, portanto sinaliza que já está ok
                if (numeroSequenciaOperacao != 0) _ConciliacaoOK = true;

                // Carrega variáveis
                _NumeroComando = entidadeMensagem.NU_COMD_OPER.ToString();
                long.TryParse(entidadeMensagem.NR_OPER_CAMB_2.ToString(), out _NumeraOperacaoCambial2);
                _DataOperacao = DateTime.Parse(DateTime.Parse(entidadeMensagem.DH_REGT_MESG_SPB.ToString()).ToShortDateString());

                // o Status Operacao a ser procurado varia de acordo com a MensagemSPB que está sendo processada
                switch (entidadeMensagem.CO_MESG_SPB.ToString().Trim())
                {
                    case "CAM0007R2":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.Respondida, Comum.Comum.EnumStatusOperacao.Reativada, Comum.Comum.EnumStatusOperacao.RegistradaAutomatica };
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.Confirmada).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.EnviadaLegado).ToString();
                        break;
                    case "CAM0008R2":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.Confirmada, Comum.Comum.EnumStatusOperacao.RegistradaAutomatica };
                        if (indicadorAceite == "S") _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.AConciliarAceite).ToString();
                        else _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.Rejeitada).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.EnviadaLegado).ToString();
                        break;
                    case "CAM0010R2": case "CAM0014R2":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.Registrada, Comum.Comum.EnumStatusOperacao.Reativada };
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.Confirmada).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.EnviadaLegado).ToString();
                        break;
                    case "CAM0055":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.Respondida, Comum.Comum.EnumStatusOperacao.Registrada, Comum.Comum.EnumStatusOperacao.Confirmada, Comum.Comum.EnumStatusOperacao.Liberada, Comum.Comum.EnumStatusOperacao.LiberadaAutomatica };
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.CanceladaCamara).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.Registrada).ToString();
                        break;
                    case "CAM0005R2":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.CanceladaCamara };
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.Reativada).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.Reativada).ToString();
                        break;
                    case "BMC0005":
                        _ListaStatusOperacao = new Comum.Comum.EnumStatusOperacao[] { Comum.Comum.EnumStatusOperacao.AConciliarAceite };
                        _StatusOperacaoSeConciliacaoOK = ((int)Comum.Comum.EnumStatusOperacao.ConcordanciaAceiteAutomatica).ToString();
                        _StatusMensagemSeConciliacaoOK = ((int)Comum.Comum.EnumStatusMensagem.Conciliada).ToString();
                        break;
                    default:
                        return false; // se não for nenhuma das MensagensSPB listadas acima, então encerra o processamento
                };

                // Verifica Conciliacao (se _ConciliacaoOK = true significa que a conciliacao não precisa ser verificada novamente)
                if (_ConciliacaoOK == false)
                {
                    if (_ConciliacaoDATA.ConciliarComOperacao(_NumeroComando,
                                                              _NumeraOperacaoCambial2,
                                                              _DataOperacao,
                                                              _ListaStatusOperacao,
                                                          ref _NumeroSequenciaOperacao,
                                                          ref _RetornoConciliacao) == true)
                    {
                        _ConciliacaoOK = true;
                        numeroSequenciaOperacao = _NumeroSequenciaOperacao;
                    }
                    else
                    {
                        if (_RetornoConciliacao == 0) _RetornoConciliacao = (int)Comum.Comum.enumJustificativa.OperacaoNaoEncontrada;
                        else if (_RetornoConciliacao >= 2) _RetornoConciliacao = 0; // deixar = 0 significa que não haverá Justificativa
                    }
                }

                // Se Conciliacao OK e somenteConsulta=false então atualiza o status da Operaçao e MensagemSPB conciliadas, e associa ambas
                if (_ConciliacaoOK == true)
                {
                    if (somenteConsulta == false)
                    {
                        // Atualiza status Operacao
                        if (base.AtualizaStatusOperacao(entidadeMensagem.CO_MESG_SPB.ToString().Trim(), long.Parse(_NumeroSequenciaOperacao.ToString())) == true)
                        {
                            base.AlterarStatusOperacao((int)_NumeroSequenciaOperacao, int.Parse(_StatusOperacaoSeConciliacaoOK), 0, _TipoAcao);
                        }

                        // Atualiza Status da MensagemSPB e a associa com a Operação Conciliada
                        _MensagemSpbDATA.SelecionarMensagensPorControleIF(entidadeMensagem.NU_CTRL_IF.ToString());
                        entidadeMensagem = _MensagemSpbDATA.ObterMensagemLida();
                        entidadeMensagem.CO_USUA_ULTI_ATLZ = Comum.Comum.UsuarioSistema;
                        entidadeMensagem.CO_ETCA_TRAB_ULTI_ATLZ = Comum.Comum.NomeMaquina;
                        entidadeMensagem.CO_ULTI_SITU_PROC = int.Parse(_StatusMensagemSeConciliacaoOK);
                        //if (entidadeMensagem.CO_MESG_SPB.ToString().Trim() != "BMC0005")
                        //{
                        entidadeMensagem.NU_SEQU_OPER_ATIV = (int)_NumeroSequenciaOperacao; // para associar a MensagemSPB à Operação
                        //}
                        base.AlterarStatusMensagemSPB(ref entidadeMensagem, (Comum.Comum.EnumStatusMensagem)int.Parse(_StatusMensagemSeConciliacaoOK));
                    }
                    return true;
                }
                else
                {
                    return false;
                }

            }

            catch (Exception ex)
            {
                throw new Exception("Conciliacao.VerificaConciliacao() - OVERLOAD MensagemSPB com Operacao" + ex.ToString());
            }
        }
        #endregion

        #region <<< VerificaConciliacao - OVERLOAD MensagemSPB com MensagemSPB >>>
        public bool VerificaConciliacao(MensagemSpbDAO.EstruturaMensagemSPB entidadeMensagem,
                                    ref DsTB_MESG_RECB_ENVI_SPB dsMensagemSPBConciliada)
        {

            string CodigoMensagemSPBAConciliar = string.Empty;
            Comum.Comum.EnumStatusMensagem StatusMensagemAConciliar = new A8NET.Comum.Comum.EnumStatusMensagem();

            try
            {
                // Carrega variáveis
                _NumeroComando = entidadeMensagem.NU_COMD_OPER.ToString();
                long.TryParse(entidadeMensagem.NR_OPER_CAMB_2.ToString(), out _NumeraOperacaoCambial2);
                _DataOperacao = DateTime.Parse(DateTime.Parse(entidadeMensagem.DH_REGT_MESG_SPB.ToString()).ToShortDateString());

                switch (entidadeMensagem.CO_MESG_SPB.ToString().Trim())
                {
                    case "CAM0013R1":
                        CodigoMensagemSPBAConciliar = "CAM0014R2";
                        StatusMensagemAConciliar = A8NET.Comum.Comum.EnumStatusMensagem.R2;
                        break;
                    case "CAM0009R1":
                        CodigoMensagemSPBAConciliar = "CAM0010R2";
                        StatusMensagemAConciliar = A8NET.Comum.Comum.EnumStatusMensagem.R2;
                        break;
                    case "CAM0054R1":
                        CodigoMensagemSPBAConciliar = "BMC0005";
                        StatusMensagemAConciliar = A8NET.Comum.Comum.EnumStatusMensagem.AConciliar;
                        break;
                    default:
                        return true;
                }

                // Verifica Conciliacao
                if (_ConciliacaoDATA.ConciliarComMensagemSPB(_NumeroComando,
                                                             _NumeraOperacaoCambial2,
                                                             _DataOperacao,
                                                             StatusMensagemAConciliar,
                                                             CodigoMensagemSPBAConciliar,
                                                         ref dsMensagemSPBConciliada,
                                                         ref _RetornoConciliacao) == true)
                {
                    return true;
                }

                return false;

            }
            catch (Exception ex)
            {
                throw new Exception("Conciliacao.VerificaConciliacao() - OVERLOAD MensagemSPB com MensagemSPB" + ex.ToString());
            }
        }
        #endregion

        #region >>> ObterCodigoMensagemSPBAConciliar >>>
        public string ObterCodigoMensagemSPBAConciliar(Comum.Comum.EnumTipoOperacao enumTipoOperacao)
        {
            try
            {
                switch (enumTipoOperacao)
                {
                    case Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemTelaCega : return "CAM0006R2";
                    case Comum.Comum.EnumTipoOperacao.ConfirmacaoOperacaoInterbancariaSemCamara   : return "CAM0009R2";
                    case Comum.Comum.EnumTipoOperacao.InformaConfirmacaoOperArbitragemParceiroPais: return "CAM0013R2";
                    default                                                                       : return string.Empty;
                }

            }
            catch (Exception ex)
            {
                throw new Exception("Conciliacao.ObterCodigoMensagemSPBAConciliar() - OVERLOAD MensagemSPB" + ex.ToString());
            }
        }
        #endregion
    }
}
