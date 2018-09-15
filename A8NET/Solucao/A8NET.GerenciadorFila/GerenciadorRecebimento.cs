using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Transactions;
using System.Threading;
using System.Xml;
using A8NET.Data;
using A8NET.ConfiguracaoMQ;
using A8NET.Factory;

namespace A8NET.GerenciadorFila
{
    public class GerenciadorRecebimento : IDisposable
    {

        #region <<< Variables >>>
        long _TotalMensagem = 100;
        MQConnector _MqConnector = null;
        #endregion

        #region <<< Constructors Members >>>
        public GerenciadorRecebimento()
        {
            #region >>> Setar a Cultura >>>
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("pt-BR");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("pt-BR");
            #endregion
        }
        #endregion

        #region <<< IDisposable >>>
        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        ~GerenciadorRecebimento()
        {
            this.Dispose();
        }
        #endregion

        #region <<< ProcessaMensagemMQ >>>
        public void ProcessaMensagemMQ(string nomeFila, string quantidadePorThread)
        {
            object ObjRetorno = null;
            bool ExisteMensagem = true;
            long Contador = 0;
            long ContadorDominio = 0;
            A8NET.Data.DsParametrizacoes DataSetCache;

            _TotalMensagem = Convert.ToInt16(quantidadePorThread);

            try
            {
                A8NET.Data.Dominio DataDominio = new A8NET.Data.Dominio();
                DataDominio.CarregaDominio();

                DataSetCache = DataDominio.DataSetCache;

                while (true)
                {
                    if (DataSetCache.TB_CTRL_PROC_OPER_ATIV.Count != 0
                    && DataSetCache.TB_FCAO_SIST_TIPO_OPER.Count != 0
                    && DataSetCache.TB_PARM_FCAO_SIST.Count != 0
                    && DataSetCache.TB_TIPO_OPER.Count != 0
                    && DataSetCache.MBS_GRUPO.Count != 0
                    && DataSetCache.TB_VEIC_LEGA.Count != 0)
                    {
                        while (ExisteMensagem)
                        {
                            Contador++;
                            try
                            {
                                ObjRetorno = CriaMensagemMQ(nomeFila, DataSetCache);
                            }
                            catch { }

                            ExisteMensagem = Convert.ToBoolean(ObjRetorno);
                            ObjRetorno = null;
                            if (Contador > _TotalMensagem) { ExisteMensagem = false; }
                        }
                        break;
                    }
                    else
                    {
                        if (ContadorDominio < 10)
                        {
                            Thread.Sleep(5000);
                            ContadorDominio++;
                            DataDominio.CarregaDominio();
                            DataSetCache = DataDominio.DataSetCache;
                        }
                        else
                        {
                            throw new Exception("AS TABELAS DE DOMINIO NAO FORAM CARREGADAS CORRETAMENTE!");
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                GravaLogErro("GerenciadorRecebimento.ProcessaMensagemMQ:ERR", ex);
                throw ex;
            }
        }
        #endregion

        #region <<< CriaMensagemMQ >>>
        private bool CriaMensagemMQ(string nomeFila, A8NET.Data.DsParametrizacoes dataSetCache)
        {
            string SistemaOrigem = string.Empty;
            string Mensagem = string.Empty;
            bool Retorno = false;
            int backOutCount = 0;
            XmlDocument xmlRemessa = new XmlDocument();

            try
            {
                using (TransactionScope ts = new TransactionScope(TransactionScopeOption.RequiresNew))
                {
                    using (_MqConnector = new MQConnector())
                    {
                        _MqConnector.MQConnect();
                        _MqConnector.MQQueueOpen(nomeFila, MQConnector.enumMQOpenOptions.GET);

                        if (_MqConnector.MQGetMessage())
                        {
                            backOutCount = _MqConnector.BackoutCount;
                            Mensagem = _MqConnector.Message;

                            //Busca o sistema origem
                            xmlRemessa.LoadXml(Mensagem);

                            if (xmlRemessa.SelectSingleNode("//SG_SIST_ORIG") != null)
                            {
                                SistemaOrigem = xmlRemessa.SelectSingleNode("//SG_SIST_ORIG").InnerXml;

                                if (backOutCount > 2)
                                {
                                    // Put na Fila de Erro
                                    _MqConnector.MQQueueOpen("A7Q.E.ERRO", MQConnector.enumMQOpenOptions.PUT);
                                    _MqConnector.Message = Mensagem;
                                    _MqConnector.MQPutMessage();
                                    _MqConnector.MQQueueClose();
                                }
                                else
                                {
                                    MensagemFactory.CriaMensagem(SistemaOrigem, dataSetCache).ProcessaMensagem(nomeFila, Mensagem);
                                }
                            }
                            Retorno = true;
                        }
                        _MqConnector.MQQueueClose();
                        _MqConnector.MQEnd();
                    }

                    ts.Complete();

                }

            }
            catch (TransactionAbortedException ex)
            {
                GravaLogErro("Erro de transação - CriaMensagemMQ(" + nomeFila + ")", ex);
                PutMensagemMQ("A7Q.E.ERRO", Mensagem);
                throw ex;
            }

            catch (Exception ex)
            {
                GravaLogErro("CriaMensagemMQ(" + nomeFila + ")", ex);
                PutMensagemMQ("A7Q.E.ERRO", Mensagem);
                throw ex;
            }

            return Retorno;

        }
        #endregion

        #region <<< GravaLogErro >>>
        private void GravaLogErro(string strErro, Exception ex)
        {

            if (strErro == null || strErro.Trim() == "")
            {
                return;
            }

            string pathRel = "";
            string nomeArquivo = "Erro_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + ".txt";

            pathRel = AppDomain.CurrentDomain.BaseDirectory + "LogErro\\";

            StringBuilder stringBuilder1 = new StringBuilder();
            stringBuilder1.AppendFormat(strErro);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append("MESSAGE: ");
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(ex.ToString().ToUpper());
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(ex.Message.ToUpper());
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append("SOURCE: ");
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(ex.Source.ToUpper());
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append("STACK TRACE: ");
            stringBuilder1.Append(Environment.NewLine);
            stringBuilder1.Append(ex.StackTrace.ToUpper());
            stringBuilder1.Append(Environment.NewLine);

            try
            {

                if (strErro.Length > 0)
                {
                    #region Tratamento dos arquivos no diretório
                    // Verifica se existe o Path.
                    if (Directory.Exists(pathRel))
                    {
                        // Máximo de arquivos permitido. Expurgo parametrizável.
                        if (Directory.GetFiles(pathRel).Length >= 2)
                        {
                            foreach (string sArquivo in Directory.GetFiles(pathRel))
                            {
                                TimeSpan dataDiff = DateTime.Now.Subtract(File.GetLastWriteTime(sArquivo).Date);

                                int diasDiff = dataDiff.Days;

                                if (diasDiff >= 1)
                                {
                                    File.Delete(sArquivo);
                                }
                            }
                        }
                    }
                    else
                    {
                        Directory.CreateDirectory(@pathRel);
                    }
                    #endregion

                    #region Grava ou adiciona registro no log

                    StringBuilder sb = new StringBuilder();

                    sb.Append("###############");
                    sb.Append(DateTime.Now.ToLongDateString());
                    sb.Append(" - ");
                    sb.Append(DateTime.Now.ToLongTimeString());
                    sb.Append("###############");
                    sb.Append(Environment.NewLine);
                    sb.Append(stringBuilder1.ToString());
                    sb.Append(Environment.NewLine);
                    sb.Append("###############");
                    sb.Append("###############");

                    if (File.Exists(pathRel + nomeArquivo))
                    {
                        using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Append, FileAccess.Write, FileShare.Write))
                        {
                            using (StreamWriter sw = new StreamWriter(fs))
                            {

                                sw.Write(sb.ToString());
                            }
                        }
                    }
                    else
                    {
                        using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Create, FileAccess.Write, FileShare.Write))
                        {
                            using (StreamWriter sw = new StreamWriter(fs))
                            {
                                sw.Write(sb.ToString());
                            }
                        }
                    }
                    #endregion
                }

            }
            catch (Exception)
            {

            }

        }
        #endregion

        #region <<< PutMensagemMQ >>>
        private void PutMensagemMQ(string nomeFila, string mensagem)
        {
            try
            {

                using (_MqConnector = new MQConnector())
                {
                    _MqConnector.MQConnect();
                    _MqConnector.MQQueueOpen(nomeFila, MQConnector.enumMQOpenOptions.PUT);
                    _MqConnector.Message = mensagem;
                    _MqConnector.MQPutMessage();
                    _MqConnector.MQQueueClose();
                    _MqConnector.MQEnd();
                }

            }
            catch (Exception ex) { }
        }
        #endregion

    }
}
