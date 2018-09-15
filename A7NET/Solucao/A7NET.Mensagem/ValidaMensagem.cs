using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using A7NET.Data;

namespace A7NET.Mensagem
{
    public class ValidaMensagem
    {
        #region <<< Variaveis >>>
        private DsParametrizacoes _DataSetCache;
        #endregion

        #region >>> Construtor >>>
        public ValidaMensagem(DsParametrizacoes DataSetCache)
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion

            _DataSetCache = DataSetCache;

        }
        #endregion

        #region <<< ValidaTipoMensagem >>>
        public bool ValidaTipoMensagem(int tipoFormatoMensagemSaida, string tipoMensagem)
        {
            int OutN;

            try
            {
                if (tipoFormatoMensagemSaida != 0)
                {
                    if (_DataSetCache.TB_TIPO_MESG.Select("TRIM(TP_MESG)='" + (int.TryParse(tipoMensagem, out OutN) ? Convert.ToString(OutN) : tipoMensagem.Trim()) + "'" +
                                                          " AND TP_FORM_MESG_SAID=" + tipoFormatoMensagemSaida).Length == 0)
                    {
                        return false;
                    }
                }
                else
                {
                    if (_DataSetCache.TB_TIPO_MESG.Select("TRIM(TP_MESG)='" + (int.TryParse(tipoMensagem, out OutN) ? Convert.ToString(OutN) : tipoMensagem.Trim()) + "'").Length == 0)
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {

                return false;
            }
        }
        #endregion

        #region <<< ValidaEmpresa >>>
        public bool ValidaEmpresa(string codigoEmpresa)
        {
            try
            {
                if (_DataSetCache.TB_EMPRESA_HO.Select("CO_EMPR=" + int.Parse(codigoEmpresa)).Length == 0)
                {
                    return false;
                }

                return true;
            }
            catch
            {

                return false;
            }
        }
        #endregion

        #region <<< ValidaSistema >>>
        public bool ValidaSistema(string codigoEmpresa, string siglaSistema)
        {
            try
            {
                if (_DataSetCache.TB_SIST.Select("CO_EMPR=" + int.Parse(codigoEmpresa) +
                                                 " AND SG_SIST='" + siglaSistema.Trim().ToUpper() + "'").Length == 0)
                {
                    return false;
                }

                return true;
            }
            catch
            {

                return false;
            }
        }
        #endregion

        #region <<< ObtemRegraTransporte >>>
        public bool ObtemRegraTransporte(ref DataRow regra, string tipoMensagem, string codigoEmpresa,
                                         string siglaSistemaOrigem, string siglaSistemaDestino, ref string erro)
        {
            int OutN;

            try
            {
                if (_DataSetCache.TB_REGR_TRAP_MESG.Select("TRIM(TP_MESG)='" + (int.TryParse(tipoMensagem, out OutN) ? Convert.ToString(OutN) : tipoMensagem.Trim()) + "'" +
                                                          " AND CO_EMPR_ORIG=" + int.Parse(codigoEmpresa) +
                                                          " AND SG_SIST_ORIG='" + siglaSistemaOrigem.Trim().ToUpper() + "'" +
                                                          " AND SG_SIST_DEST='" + siglaSistemaDestino.Trim().ToUpper() + "'").Length == 0)
                {
                    return false;
                }

                regra = _DataSetCache.TB_REGR_TRAP_MESG.Select("TRIM(TP_MESG)='" + (int.TryParse(tipoMensagem, out OutN) ? Convert.ToString(OutN) : tipoMensagem.Trim()) + "'" +
                                                              " AND CO_EMPR_ORIG=" + int.Parse(codigoEmpresa) +
                                                              " AND SG_SIST_ORIG='" + siglaSistemaOrigem.Trim().ToUpper() + "'" +
                                                              " AND SG_SIST_DEST='" + siglaSistemaDestino.Trim().ToUpper() + "'")[0];


                return true;
            }
            catch (Exception ex)
            {

                erro = ex.ToString();
                return false;
            }
        }
        #endregion

    }
}
