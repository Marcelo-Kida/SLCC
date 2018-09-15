using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Text;

namespace br.santander.SLCC
{

    public delegate bool AsyncDelegate(SLCCFila fila, out SLCCFila fila1);

    public class SLCCExecutor : IDisposable
    {
        private object objRetorno = null;
        private object objHandle = null;
        private Type TypeComp = null;

        #region Consttutores Members
            public SLCCExecutor(){

            }
        #endregion
        
        #region IDisposable Members
            
            public void Dispose()
            {
                GC.SuppressFinalize(this);
            }

            ~SLCCExecutor()
            {
                this.Dispose();
                
                GC.Collect();
                GC.WaitForPendingFinalizers();

                
            }

        #endregion

        #region COMTaskExecutor Members
            public bool COMTaskExecutor(SLCCFila fila, out SLCCFila fila1)
            {
                try{

                    int availableThreads = 0;
                    int outComp = 0;
                    bool processar = true;

                    ThreadPool.GetAvailableThreads(out availableThreads, out outComp);
                    Thread.CurrentThread.Name = fila.nomeFila;
                    
                    if (availableThreads == 0){
                        processar = false;
                    }else{

                        if (fila.quantidadeAtualThread <= fila.quantidadeMaxThreads){
                            processar = true;
                        }else {
                            processar = false;
                        }
                    }

                    if (processar){

                        if (Thread.CurrentThread.TrySetApartmentState(ApartmentState.MTA)){
                            Thread.CurrentThread.SetApartmentState(ApartmentState.MTA);
                        }

                        string nomeObjeto = fila.nomeObjeto.Trim();
                        string nomeMetodo = fila.Metodo.Trim();

                        TypeComp = Type.GetTypeFromProgID(nomeObjeto);

                        objHandle = Activator.CreateInstance(TypeComp);

                        string[] arrParams = { fila.xmlNodeFila, fila.DataHoraUltiExec.ToString("yyyyMMddHHmmss") };

                        objRetorno = TypeComp.InvokeMember(nomeMetodo,
                                                           BindingFlags.InvokeMethod,
                                                           null,
                                                           objHandle,
                                                           arrParams);

                        nomeObjeto = null;
                        nomeMetodo = null;

                    }

                }catch (Exception objErr){
                    Console.WriteLine("---->>ERROR...{0}", objErr.Message);
                }finally{
                    
                    if (objHandle != null){
                        int ret = Marshal.ReleaseComObject(objHandle);
                    }

                    TypeComp = null;
                    objHandle = null;
                }

                fila1 = fila;

                return true;
        }   
        #endregion

        #region LogErro Members
        
        private void GravaLogErro(string strErro)
        {
            
            if (strErro == null)
            {
                return;
            }

            string pathRel = "";
            string nomeArquivo = "Trace_" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + ".txt";

            pathRel = AppDomain.CurrentDomain.BaseDirectory + "Trace\\";

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
                    if (File.Exists(pathRel + nomeArquivo)){
                        using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Append, FileAccess.Write, FileShare.Write))
                        {
                            using (StreamWriter sw = new StreamWriter(fs)){
                                sw.Write(strErro);
                            }
                        }
                    }else{
                        using (FileStream fs = File.Open(pathRel + nomeArquivo, FileMode.Create, FileAccess.Write, FileShare.Write))
                        {
                            using (StreamWriter sw = new StreamWriter(fs)){
                                sw.Write(strErro);
                            }
                        }
                    }
                    #endregion
                }


            }
            catch (Exception ex)
            {
                Console.WriteLine("---->>ERROR...{0}", ex.Message);
            }

        }
        #endregion

    }
}