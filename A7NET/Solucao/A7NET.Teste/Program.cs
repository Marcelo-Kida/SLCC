using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Threading;
using A7NET.GerenciadorFila;

namespace A7NET.Teste
{
    class Program
    {
        static void Main(string[] args)
        {
            #region >>> Setar a Cultura >>>
            Thread.CurrentThread.CurrentCulture = new CultureInfo("pt-BR");
            Thread.CurrentThread.CurrentUICulture = new CultureInfo("pt-BR");
            #endregion

            A7NET.GerenciadorFila.GerenciadorRecebimento msg = new GerenciadorRecebimento();
            msg.ProcessaMensagemMQ("A7Q.E.ENTRADA_NET", "1");

            Console.ReadLine();
        }
    }
}
