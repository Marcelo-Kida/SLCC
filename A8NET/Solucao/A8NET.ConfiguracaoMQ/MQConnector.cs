using System;
using System.IO;
using System.Text;
using System.Net;
using System.Configuration;
using IBM.WMQ;

namespace A8NET.ConfiguracaoMQ
{
    public class MQConnector : IDisposable
    {

        #region <<< Variables >>>
        // Variaveis MQ
        protected MQQueueManager queueManager = null;
        protected MQQueue queue = null;
        protected int OpenOptions = 0;
        protected enumMQOpenOptions mqOpenOption;

        // Variaveis Apoio PUT/GET
        private string message = null;  // mensagem PUT/GET

        private string messageIdHex = null; // Usado no  PUT/GET
        private string correlationIdHex = null;// Usado no  PUT/GET

        private byte[] messageID = null; // Usado no  PUT/GET
        private byte[] correlationID = null;// Usado no  PUT/GET
        private int waitInterval;// Usado no  GET
        private int backoutCount = 0;
        private string replyToQueueName = null;

        private int prioridade = MQC.MQPRI_PRIORITY_AS_Q_DEF;
        #endregion

        #region <<< Constructors Members >>>
        public MQConnector()
        {
        }
        #endregion

        #region <<< IDisposable Members >>>
        public void Dispose()
        {
            if (this.queueManager != null && this.queueManager.IsConnected)
                this.queueManager.Disconnect();

            GC.SuppressFinalize(this);
        }
        ~MQConnector()
        {
            this.Dispose();
        }
        #endregion

        #region <<< Get / Set Members >>>
        public int BackoutCount
        {
            get { return backoutCount; }
            set { backoutCount = value; }
        }

        public string ReplyToQueueName
        {
            get { return replyToQueueName; }
            set { replyToQueueName = value; }
        }

        public int Prioridade
        {
            get { return prioridade; }
            set { prioridade = value; }
        }

        public int WaitInterval
        {
            get { return waitInterval; }
            set { waitInterval = value; }
        }

        public byte[] CorrelationID
        {
            get { return correlationID; }
            set { correlationID = value; }
        }

        public byte[] MessageID
        {
            get { return messageID; }
            set { messageID = value; }
        }
        public string MessageIdHex
        {
            get { return messageIdHex; }
            set { messageIdHex = value; }
        }

        public string CorrelationIdHex
        {
            get { return correlationIdHex; }
            set { correlationIdHex = value; }
        }

        public string Message
        {
            get { return message; }
            set { message = value; }
        }
        #endregion

        #region <<< Enum Members >>>
        public enum enumMQOpenOptions
        {
            GET, 
            PUT, 
            BROWSE, 
            GET_WAIT
        }

        public enum PrioridadeMensagem
        {
            Minima = 0,
            Maxima = 9
        }
        #endregion

        #region <<< MQConnect >>>
        public void MQConnect()
        {
            try
            {
                // TESTES MQ AMBIENTE MVAR
                //string SPort = "1415";//ConfigurationManager.AppSettings["Port"].ToString().Trim();
                //Int64 Port = Convert.ToInt64("0" + SPort);
                //string Channel = "CHANNEL1";//ConfigurationManager.AppSettings["Channel"].Trim();
                //string QueueManagerName = "QM.SLCC.01";//ConfigurationManager.AppSettings["QueueManagerName"].Trim();
                //String HostName = "Orion";//ConfigurationManager.AppSettings["Hostname"].Trim();
                //string connString = HostName + "(" + Port + ")";
                //queueManager = new MQQueueManager(QueueManagerName, Channel, connString);

                string QueueManagerName = "QM.SLCC.01";
                queueManager = new MQQueueManager(QueueManagerName);

            }
            catch (MQException exMQ)
            {

                throw new Exception("Erro Acesso MQ - Reason :" + exMQ.Reason, exMQ);
            }
            catch (Exception ex) { throw ex; }

        }
        #endregion

        #region <<< MQQueueOpen >>>
        public void MQQueueOpen(string queueName, enumMQOpenOptions openOption)
        {

            mqOpenOption = openOption;

            switch (openOption)
            {
                case enumMQOpenOptions.BROWSE:
                    OpenOptions = MQC.MQOO_INQUIRE + MQC.MQOO_BROWSE + MQC.MQOO_FAIL_IF_QUIESCING;
                    break;
                case enumMQOpenOptions.GET:
                    OpenOptions = MQC.MQOO_INPUT_SHARED + MQC.MQOO_INQUIRE;
                    break;
                case enumMQOpenOptions.PUT:
                    OpenOptions = MQC.MQOO_OUTPUT + MQC.MQOO_FAIL_IF_QUIESCING;
                    break;
                case enumMQOpenOptions.GET_WAIT:
                    OpenOptions = MQC.MQOO_INPUT_SHARED + MQC.MQOO_INQUIRE;
                    break;

            }
            try
            {

                queue = queueManager.AccessQueue(queueName, OpenOptions, null, null, null);
            }
            catch (Exception MQe)
            {
                throw MQe;
            }
        }
        #endregion

        #region <<< MQQueueClose >>>
        public void MQQueueClose()
        {
            queue.Close();
            message = null;
            messageID = null;
            correlationID = null;
            backoutCount = 0;
            replyToQueueName = null;
            messageIdHex = null;
            correlationIdHex = null;

        }
        #endregion

        #region <<< MQEnd >>>
        public void MQEnd()
        {
            this.queue.Close();
            this.queueManager.Disconnect();

            this.queue = null;
            this.queueManager = null;

        }
        #endregion

        #region <<< MQGetMessage >>>
        public Boolean MQGetMessage()
        {

            MQGetMessageOptions getOptions = new MQGetMessageOptions();

            switch (mqOpenOption)
            {
                case enumMQOpenOptions.BROWSE:
                    getOptions.Options = MQC.MQGMO_BROWSE_NEXT + MQC.MQGMO_NO_WAIT;
                    getOptions.WaitInterval = MQC.MQWI_UNLIMITED;
                    break;
                case enumMQOpenOptions.GET_WAIT:
                    getOptions.Options = MQC.MQGMO_SYNCPOINT + MQC.MQGMO_WAIT;
                    getOptions.WaitInterval = WaitInterval;
                    break;
                default:
                    getOptions.Options = MQC.MQGMO_SYNCPOINT + MQC.MQGMO_NO_WAIT;
                    getOptions.WaitInterval = MQC.MQWI_UNLIMITED;
                    break;

            }

            getOptions.MatchOptions = MQC.MQMO_MATCH_CORREL_ID + MQC.MQMO_MATCH_MSG_ID;

            MQMessage getMessage = new MQMessage();

            if (correlationID != null)
            {
                getMessage.CorrelationId = correlationID;
            }
            else
            {
                getMessage.CorrelationId = MQC.MQCI_NONE;
            }

            if (messageID != null)
            {
                getMessage.MessageId = messageID;
            }
            else
            {
                getMessage.MessageId = MQC.MQMI_NONE;
            }


            try
            {
                
                queue.Get(getMessage, getOptions);
                string receiveMessage = getMessage.ReadString(getMessage.MessageLength);
                message = this.RetiraCaracteresEspeciais(receiveMessage);

                messageID = getMessage.MessageId;
                correlationID = getMessage.CorrelationId;

                prioridade = getMessage.Priority;

                messageIdHex = ByteArrayToStringHex(messageID);
                correlationIdHex = ByteArrayToStringHex(correlationID);


                backoutCount = getMessage.BackoutCount;

                getMessage.ClearMessage();
                return true;

            }
            catch (MQException e)
            {

                if (e.ReasonCode == MQC.MQRC_NO_MSG_AVAILABLE)
                {
                    return false;
                }
                else
                {
                    throw e;
                }
            }
        }
        #endregion

        #region <<< MQPutMessage >>>
        public void MQPutMessage()
        {

            MQMessage sendmsg = new MQMessage();

            sendmsg.ClearMessage();
            sendmsg.Format = MQC.MQFMT_STRING;
            sendmsg.Feedback = MQC.MQFB_NONE;
            sendmsg.Feedback = MQC.MQFB_EXPIRATION;
            sendmsg.MessageType = MQC.MQMT_DATAGRAM;
            sendmsg.Persistence = MQC.MQPER_PERSISTENT;
            sendmsg.MessageId = MQC.MQMI_NONE;
            sendmsg.CorrelationId = MQC.MQCI_NONE;

            if (replyToQueueName != null)
            {
                sendmsg.ReplyToQueueManagerName = queueManager.Name;
                sendmsg.ReplyToQueueName = replyToQueueName;
            }


            if (messageID != null)
            {
                sendmsg.MessageId = messageID;
            }


            if (messageIdHex != null)
            {
                sendmsg.MessageId = this.StringHexToByteArray(messageIdHex);
            }


            if (correlationID != null)
            {
                sendmsg.CorrelationId = correlationID;
            }

            if (correlationIdHex != null)
            {
                sendmsg.MessageId = this.StringHexToByteArray(correlationIdHex);
            }

            sendmsg.Priority = prioridade;

            MQPutMessageOptions putOptions = new MQPutMessageOptions();  // accept the defaults, same
            putOptions.Options = MQC.MQPMO_SYNCPOINT;
            sendmsg.WriteBytes(Message.ToString());
            queue.Put(sendmsg, putOptions);

            messageID = sendmsg.MessageId;
            correlationID = sendmsg.CorrelationId;

            messageIdHex = ByteArrayToStringHex(messageID);
            correlationIdHex = ByteArrayToStringHex(correlationID);

            sendmsg.ClearMessage();

        }
        #endregion

        #region <<< RetiraCaracteresEspeciais >>>
        private string RetiraCaracteresEspeciais(string mensagem)
        {
            try
            {
                string Retorno = System.Text.Encoding.Default.GetString(System.Text.Encoding.Default.GetBytes(mensagem));
                byte[] lbAC = new byte[1];
                string lsC = String.Empty;

                Retorno = Retorno.Replace("&amp;", "E");
                Retorno = Retorno.Replace("&gt;", " ");
                Retorno = Retorno.Replace("&lt;", " ");
                Retorno = Retorno.Replace("&quot;", " ");
                Retorno = Retorno.Replace("&", "E");
                Retorno = Retorno.Replace("'", " ");


                //Remove caracteres < 32
                for (short i = 0; i < 32; i++)
                {
                    lbAC[0] = Convert.ToByte(i);
                    lsC = System.Text.Encoding.GetEncoding(0).GetString(lbAC);
                    if (Retorno.IndexOf(lsC) > -1)
                        Retorno = Retorno.Replace(lsC, " ");
                }

                //Remove acentuaзгo
                string[] lcAAcentos = { "З", "з", "Ь", "ь", "А", "а", "Б", "Й", "Н", "У", "Ъ", "б", "й", "н", "у", "ъ", "Г", "Х", "г", "х", "В", "К", "Ф", "в", "к", "ф" };
                string[] lcASemacentos = { "C", "c", "U", "u", "A", "a", "A", "E", "I", "O", "U", "a", "e", "i", "o", "u", "A", "O", "a", "o", "A", "E", "O", "a", "e", "o" };

                for (short liA = 0; liA < lcAAcentos.Length; liA++)
                {
                    if (Retorno.IndexOf(lcAAcentos[liA]) > -1)
                        Retorno = Retorno.Replace(lcAAcentos[liA], lcASemacentos[liA]);
                }

                //Remove caracteres > 128
                for (short liI = 123; liI < 256; liI++)
                {
                    lbAC[0] = Convert.ToByte(liI);
                    lsC = System.Text.Encoding.GetEncoding(0).GetString(lbAC);
                    if (Retorno.IndexOf(lsC) > -1)
                        Retorno = Retorno.Replace(lsC, " ");
                }
                return Retorno;
            }
            catch
            {
                throw;
            }
        }
        #endregion

        #region <<< StringHexToByteArray >>>
        private byte[] StringHexToByteArray(String hex)
        {
            int NumberChars = hex.Length;
            byte[] bytes = new byte[NumberChars / 2];
            for (int i = 0; i < NumberChars; i += 2)
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            return bytes;
        }
        #endregion

        #region <<< StrToByteArray >>>
        private static byte[] StrToByteArray(string str)
        {
            System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();

            return encoding.GetBytes(str);
        }
        #endregion

        #region <<< ByteArrayToStringHex >>>
        private string ByteArrayToStringHex(byte[] ba)
        {
            string hex = BitConverter.ToString(ba);
            return hex.Replace("-", "");
        }
        #endregion

        #region <<< ByteArrayToStr >>>
        private static string ByteArrayToStr(byte[] sBytes)
        {
            System.Text.ASCIIEncoding encoding = new System.Text.ASCIIEncoding();

            return encoding.GetString(sBytes);
        }
        #endregion

        #region <<< getQueueDepthMessage >>>
        public int getQueueDepthMessage()
        {
            return queue.CurrentDepth;
        }
        #endregion

        #region <<< isConnected >>>
        public Boolean isConnected()
        {
            return queueManager.IsConnected;
        }
        #endregion

        #region <<< isOpen >>>
        public Boolean isOpen()
        {
            return queueManager.IsOpen;
        }
        #endregion

        #region <<< MQCommit >>>
        public void MQCommit()
        {
            this.queueManager.Commit();
        }
        #endregion

        #region <<< MQBackOut >>>
        public void MQBackOut()
        {
            this.queueManager.Backout();
        }
        #endregion

        #region <<< PutMensagem >>>
        public static void PutMensagem(string nomeFila, string mensagem)
        {
            using (MQConnector MqConnector = new MQConnector())
            {
                MqConnector.MQQueueOpen(nomeFila, MQConnector.enumMQOpenOptions.PUT);
                MqConnector.Message = mensagem;
                MqConnector.MQPutMessage();
                MqConnector.MQQueueClose();
                MqConnector.MQEnd();
            }
        }
        #endregion
    }
}