Attribute VB_Name = "basVerificaServer"

'Verificar existecia de mensagens nas Filas do gerenciador de Filas MQSeries

Option Explicit

Dim xmlConfiguracaoEntrada                  As MSXML2.DOMDocument40

Private llTempoEspera                       As Long 'em segundos
Private llTempoVerifSTR0008R2               As Long 'em minutos
Private llTempoVerifRodaDolar               As Long 'em segundos

Private datUltiVerifSTR0008R2               As Date
Private datUltiVerifUsuarioInativo          As Date
Private datUltiVerifRodaDolar               As Date

Private lngQuantidadeTotalThreads           As Long
Private lngQuantidadeFilas                  As Long
Private strQMName                           As String

Private objMQSession                        As MQAX200.MQSession
Private objMQQueueManager                   As MQAX200.MQQueueManager

Private Enum enumMSMQAcess
    MQ_RECEIVE_ACCESS = 1
End Enum

Private Enum enumMSMQShareMode
    MQ_DENY_NONE = 0
End Enum

Private Enum enumSourceAlerta
    Oracle = 1
    MQSeries = 2
    MSMQ = 3
    Todos = 4
End Enum

Public Enum enumPutOptions
    Binario = 1
    MainFrame = 2
End Enum


Public Enum enumFilasEntrada
    enumA7QEENTRADA = 1
    enumA7QEMENSAGEMRECEBIDA = 2
    enumA7QEREPORT = 3
    enumA8QEENTRADA = 4
    enumA6QEREMESSASUBRESERVA = 5
    enumA6QEREMESSAFUTURO = 6
End Enum

Private strHoraInicioVerificacao           As String
Private strHoraFimVerificacao              As String

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Rotina Inicial de verificação

Public Sub Main()
    
Dim strErro                                     As String
Dim strErroMQ                                   As String
    
On Error GoTo ErrorHandler
    
    If App.PrevInstance Then End
        
    Call flInicializar
    
    Call flVerificaMensagemMQ

    Exit Sub

ErrorHandler:
                
    strErro = Err.Description
    
    Set xmlConfiguracaoEntrada = Nothing
               
    strErroMQ = fgGetError(strErro, strQMName, "")
        
    If Trim(strErroMQ) <> "" Then
        strErro = strErroMQ
    End If
               
    flGravaArquivo strErro
    
End Sub

'Inicialização do Processamento
' - Verificação do Banco de Dados
' - Verificação do Gerenciador de Filas MSMQ
' - Leitura do filas.xml
' - Carregar as variaveis

Private Function flInicializar() As Boolean


Dim strErro                                 As String
Dim strErroMQ                               As String
Dim vntTempo                                As Variant
Dim objFuncao                               As Object 'A6A7A8.clsA6A7A8Funcoes

On Error GoTo ErrorHandler
            
    datUltiVerifSTR0008R2 = "00:00:00"
    datUltiVerifUsuarioInativo = "00:00:00"
    datUltiVerifRodaDolar = "00:00:00"
            
    Set objFuncao = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    llTempoVerifSTR0008R2 = Val(objFuncao.ObterValorParametrosGerais("CONSULTA_STR0008/PERIODICIDADE_EM_MINUTOS"))
    Set objFuncao = Nothing
                    
    If llTempoVerifSTR0008R2 = 0 Then llTempoVerifSTR0008R2 = 15
                
    If Trim$(Command$) = vbNullString Then
        llTempoEspera = 10
    Else
        llTempoEspera = Command$
    End If
    
    llTempoVerifRodaDolar = 30000
    
    'Verificar se a conexão com Oracle está OK
    flVerificarBancoDados
    
    'Verificar se a conexão com MSMQ esta OK
    flVerificaMSMQ
    
    Set xmlConfiguracaoEntrada = CreateObject("MSXML2.DOMDocument.4.0")

    With xmlConfiguracaoEntrada
        .validateOnParse = False
        .resolveExternals = False
        .setProperty "SelectionLanguage", "XPath"
        
        If Len(Dir(App.Path & "\Filas.xml")) = 0 Then
            Err.Raise Err.Number + 513, "flInicializar", App.EXEName & " - Arquivo: " & App.Path & "\Filas.xml não existe!"""
        Else
            If Not .Load(App.Path & "\Filas.xml") Then
                Call fgErroLoadXML(xmlConfiguracaoEntrada, App.EXEName, "basVerificaServer", "flInicializar")
                Set xmlConfiguracaoEntrada = Nothing
                Exit Function
            End If
        End If
    End With
    
    strQMName = Trim(xmlConfiguracaoEntrada.documentElement.selectSingleNode("NomeQueueManager").Text)
    
    If xmlConfiguracaoEntrada.documentElement.selectSingleNode("Janela_ParadaVerificacao/@HoraInicio") Is Nothing Then
        strHoraInicioVerificacao = "0300"
    Else
        strHoraInicioVerificacao = Trim$(xmlConfiguracaoEntrada.documentElement.selectSingleNode("Janela_ParadaVerificacao/@HoraInicio").Text)
    End If
    
    If xmlConfiguracaoEntrada.documentElement.selectSingleNode("Janela_ParadaVerificacao/@HoraFim") Is Nothing Then
        strHoraFimVerificacao = "0400"
    Else
        strHoraFimVerificacao = Trim$(xmlConfiguracaoEntrada.documentElement.selectSingleNode("Janela_ParadaVerificacao/@HoraFim").Text)
    End If
    
    lngQuantidadeFilas = xmlConfiguracaoEntrada.documentElement.selectNodes("//Grupo_Parametros_Entrada").length
    lngQuantidadeTotalThreads = xmlConfiguracaoEntrada.documentElement.selectSingleNode("QuantidadeTotalThreads").Text
        
    'Se a quantidade de threads estiver configurada for menor que o numero de filas
    'A quantidade máxima de treads sera a quantidade de filas de entrada
    If lngQuantidadeTotalThreads < lngQuantidadeFilas Then
        lngQuantidadeTotalThreads = lngQuantidadeFilas
    End If
         
    flSetControleThread vbNullString
         
    flInicializar = True
    
    Exit Function

ErrorHandler:
    
    Set objFuncao = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Verificar se existe mensagens nas filas no gerencidor de filas MQSEries
'Controle de threads
'Chamada a rotinas de processamento das mensagens

Private Sub flVerificaMensagemMQ()

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objPublisher                            As Object
Dim strNomeObjeto                           As String
Dim lngNumeroMaxThreads                     As Long
Dim blnProcessouOperacoes                   As Boolean

On Error GoTo ErrorHandler
        
    blnProcessouOperacoes = False
    
    Call flInicializaValidaRemessa(True)
    
    Do
        If flVerificaJanelaVerificacao() Then
            
            flVerificarBancoDados
            flVerificaMSMQ
            flVerificaMQSeries
                    
            If Not flExisteInformacaoAlerta(enumSourceAlerta.Todos) Then
                    
                'Verifica o envio de mensagens agendadas
                For Each xmlNode In xmlConfiguracaoEntrada.selectNodes("//Repet_Parametros_Agendamento/*")
                    
                    If Not xmlNode.selectSingleNode("QuantidadeMaxThreads") Is Nothing Then
                       
                       If CLng("0" & xmlNode.selectSingleNode("QuantidadeMaxThreads").Text) > 0 Then
                            
                            If UCase$(xmlNode.selectSingleNode("NomeObjeto").Text) = "A8LQS.CLSPROCESSOMENSAGEMSTR" Then
                                
                                If datUltiVerifSTR0008R2 = "00:00:00" Then
                                    Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                    objPublisher.AcionaGerenciadores xmlNode.xml
                                    Set objPublisher = Nothing
                                    datUltiVerifSTR0008R2 = Now
                                Else
                                    If DateDiff("n", datUltiVerifSTR0008R2, Now) > llTempoVerifSTR0008R2 Then
                                        Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                        objPublisher.AcionaGerenciadores xmlNode.xml
                                        Set objPublisher = Nothing
                                        datUltiVerifSTR0008R2 = Now
                                    End If
                                End If
                            
                            ElseIf UCase$(xmlNode.selectSingleNode("NomeObjeto").Text) = "A7SERVER.CLSCONTROLEUSUARIOSISTEMA" Then
                            
                                If datUltiVerifUsuarioInativo = "00:00:00" Then
                                    Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                    objPublisher.AcionaGerenciadores xmlNode.xml
                                    Set objPublisher = Nothing
                                    datUltiVerifUsuarioInativo = Date$
                                Else
                                    If datUltiVerifUsuarioInativo <> Date$ Then
                                        Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                        objPublisher.AcionaGerenciadores xmlNode.xml
                                        Set objPublisher = Nothing
                                        datUltiVerifUsuarioInativo = Date$
                                    End If
                                End If
                                
                            ElseIf UCase$(xmlNode.selectSingleNode("NomeObjeto").Text) = "A8LQS.CLSLIQUIDACAOFUTURA" Then
                                                            
                                If DatePart("H", Now) > 5 Then
                                    If Not blnProcessouOperacoes Then
                                        blnProcessouOperacoes = True
                                        Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                        objPublisher.AcionaGerenciadores xmlNode.xml
                                        Set objPublisher = Nothing
                                    End If
                                Else
                                    blnProcessouOperacoes = False
                                End If
                            
                            ElseIf UCase$(xmlNode.selectSingleNode("NomeObjeto").Text) = "A6A7A8.CLSCONTROLEACESSO" Then
                                
                                If flVerificaCacheUsuario = "S" Then
                                    Call flRemoveUsuarioCache
                                End If
                                
                            Else
                                Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                objPublisher.AcionaGerenciadores xmlNode.xml
                                Set objPublisher = Nothing
                            End If
                                
                        End If
                    End If
                    
                Next
                
                'Processa as mensagens das filas que possuem mensagens,
                'mas não possuem threads.
                Call flProcessaMensagemFilaSemThreads(xmlConfiguracaoEntrada)
                
                For Each xmlNode In xmlConfiguracaoEntrada.selectNodes("//Repet_Parametros_Entrada/*")
                    
                    lngNumeroMaxThreads = CLng("0" & xmlNode.selectSingleNode("QuantidadeMaxThreads").Text)
                    strNomeObjeto = Trim(xmlNode.selectSingleNode("NomeObjeto").Text)
                    
                    'Se lngNumeroMaxThreads = 0 , leitura da fila não esta ativa
                    If lngNumeroMaxThreads > 0 Then
                        'Verifica se existem mensagens na fila
                        If flVerificaMensagemFila(xmlNode.selectSingleNode("NomeFila").Text) Then
                            'Controle Total de Treads Ativas
                            If flObterTotalThreadsAtivas < lngQuantidadeTotalThreads Then
                                'Verifica se o numero de treads ativas é menor que o numero maximo de Threads
                                If flObterNumeroThreadsAtivas(strNomeObjeto) < lngNumeroMaxThreads Then
                                    Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                                    objPublisher.AcionaGerenciadores xmlNode.xml
                                    Set objPublisher = Nothing
                                Else
                                    Call flVerificaThreadSemMensagem(strNomeObjeto)
                                End If
                            End If
                        Else
                            Call flVerificaThreadSemMensagem(strNomeObjeto)
                        End If
                    End If
                Next
            End If
                               
        Else
            Call flRenovaCache
        End If
        Sleep llTempoEspera * 1000
      Loop
    
    Set objMQSession = Nothing
    Set objMQQueueManager = Nothing
    
    Exit Sub

ErrorHandler:
        
        
    Set objMQSession = Nothing
    Set objMQQueueManager = Nothing
    Set objPublisher = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub

'Veirifica se existe mensagem em um fila

Private Function flVerificaMensagemFila(ByVal psNomeFila As String) As Boolean
    
Dim objMQQueue                              As MQAX200.MQQueue
Dim llNumeroMensagens                       As Long
Dim strErroMQ                               As String
Dim strErro                                 As String

On Error GoTo ErrorHandler
    
    Set objMQQueue = objMQQueueManager.AccessQueue(psNomeFila, MQOO_INQUIRE + MQQT_LOCAL)
    
    If Not objMQQueue Is Nothing Then
        llNumeroMensagens = objMQQueue.CurrentDepth
        objMQQueue.Close
        Set objMQQueue = Nothing
    End If
    
    If llNumeroMensagens > 0 Then
        flVerificaMensagemFila = True
    Else
        flVerificaMensagemFila = False
    End If
    
    Exit Function
    
ErrorHandler:
    
    strErro = Err.Description
        
    If flExisteInformacaoAlerta(enumSourceAlerta.MQSeries) Then Exit Function
    
    Call flGeraInformacaoAlerta("Verificar MQSeries (Queue Manager: " & strQMName & " ;Nome da Fila : " & psNomeFila & ")", _
                                Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                "MQSeries", _
                                strErro)
        
    flGravaArquivo "Verificar Gerenciador de Filas - MQSeries" & vbCrLf & _
                   "Queue Manager: " & strQMName & vbCrLf & _
                   "Nome da Fila : " & psNomeFila & vbCrLf & _
                   strErro
    
End Function

'Adicionar um Node a um xml

Private Function flAppendNode(ByRef objDomDocument As MSXML2.DOMDocument40, _
                              ByVal psNodeContext As String, _
                              ByVal psNodeNome As String, _
                              ByVal psNodeValor As String, _
                     Optional ByVal psNodeRepet As String = "") As Boolean

Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If psNodeContext = "" Then
        'Se a Tag for o root passar psNodeContext = Nome Tag Principal
        Set objDomNodeContext = objDomDocument
    Else
        'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
        'Passar o argumento pNomeRepet = Repet_Usuario
        If psNodeRepet <> "" Then
            Set objDomNodeContext = objDomDocument.selectSingleNode("//Repet_Origem").childNodes.Item(objDomDocument.selectSingleNode("//Repet_Origem").childNodes.length - 1)
        Else
            Set objDomNodeContext = objDomDocument.documentElement.selectSingleNode("//" & psNodeContext)
        End If
    End If
    
    Set objDomNodeAux = objDomDocument.createElement(psNodeNome)
    objDomNodeAux.Text = psNodeValor
    objDomNodeContext.appendChild objDomNodeAux

    Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.Description

End Function

'Gravar aqruivos de log

Private Sub flGravaArquivo(ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strMensagem                              As String

On Error GoTo ErrorHandler

    strMensagem = String(50, "*") & vbCrLf
    strMensagem = strMensagem & pstrErro & vbCrLf
    strMensagem = strMensagem & String(50, "*")

    strNomeArquivoLogErro = App.Path & "\log\A6A7A8VerificaServer_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strMensagem
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'Gerar informações de informações de alerta de erro no A7

Private Function flGeraInformacaoAlerta(ByVal pstrDescricaoOcorrencia As String, _
                                        ByVal pstrDataHoraOcorrencia As String, _
                                        ByVal pstrSorce As String, _
                                        ByVal pstrErro As String) As Boolean

Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlAlertaAux                            As MSXML2.DOMDocument40
Dim strxmlAlerta                            As String

On Error GoTo ErrorHandler
    
    strxmlAlerta = flObterInformacaoAlerta()
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAlertaAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call flAppendNode(xmlAlertaAux, "", "Grupo_Alerta", "")
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_OCOR", pstrDescricaoOcorrencia)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DH_OCOR", pstrDataHoraOcorrencia)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_FONT_ERRO", pstrSorce)
    Call flAppendNode(xmlAlertaAux, "Grupo_Alerta", "DE_ERRO", pstrErro)
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then
        Call flAppendNode(xmlAlerta, "", "Alerta", "")
        Call flAppendNode(xmlAlerta, "Alerta", "Repet_Alerta", "")
    Else
        xmlAlerta.loadXML strxmlAlerta
    End If
    
    Call flAppendXML(xmlAlerta, "Repet_Alerta", xmlAlertaAux.xml)
    
    Call flGerarPropriedadeAlerta(xmlAlerta.xml)
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlAlerta = Nothing
    Set xmlAlertaAux = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Obter informações de alerta de errro do A7 no Shared Property memory (COM+)

Private Function flObterInformacaoAlerta() As String

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    flObterInformacaoAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Verifica se existe informação de alerta de errro do A7

Private Function flExisteInformacaoAlerta(ByVal plngSouce As enumSourceAlerta) As Boolean

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String
Dim blnExisteAlerta                         As Boolean
Dim strAux                                  As String

On Error GoTo ErrorHandler
        
    blnExisteAlerta = False
    
    Exit Function
        
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    strxmlAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then Exit Function
    
    xmlAlerta.loadXML strxmlAlerta
        
    For Each xmlNode In xmlAlerta.selectNodes("Alerta/Repet_Alerta/*")
             
         Select Case plngSouce
            Case enumSourceAlerta.MQSeries
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MQSERIES" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.MSMQ
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MSMQ" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.Oracle
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "ORACLE" Then
                    blnExisteAlerta = True
                    Exit For
                End If
            Case enumSourceAlerta.Todos
            
                strAux = strAux & xmlNode.selectSingleNode("DE_FONT_ERRO").Text
         End Select
    Next
    
    If plngSouce = enumSourceAlerta.Todos Then
        If strAux <> vbNullString Then
            blnExisteAlerta = True
        End If
    End If
    
    flExisteInformacaoAlerta = blnExisteAlerta
    
    Set xmlAlerta = Nothing
    Set xmlNode = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlNode = Nothing
    Set xmlAlerta = Nothing
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Gravar as informações de alerta de erro do A7 no Shred Property Memory (COM+)

Private Function flGerarPropriedadeAlerta(ByVal pstrxmlAlerta As String)

Dim objSharedPropMem                        As Object

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    Call objSharedPropMem.SetSPMProperty("A7BUSALERTA", "ALERTA", pstrxmlAlerta)
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Verifica se o erro é de MQSeries e Oracle

Public Function fgGetError(ByVal pstrErro As String, _
                           ByVal pstrQMgrName As String, _
                           ByVal pstrQueueName As String) As String

Dim objDomError                             As MSXML2.DOMDocument40
Dim intErrPosicaoInicio                     As Integer
Dim intErrPosicaoFim                        As Integer
Dim strErrSource                            As String
Dim strErrDesciption                        As String
Dim strTxtPesquisa                          As String
Dim lngErrNumber                            As Long
Dim strCodOcorrencia                        As String
Dim strOcorrencia                           As String
Dim strComplemento                          As String
    
On Error GoTo ErrorHandler
    
    strErrDesciption = pstrErro
    
    If InStr(1, UCase(strErrDesciption), UCase("mqax200")) > 0 Then
        
        strTxtPesquisa = "ReasonCode = "
        
        intErrPosicaoInicio = InStr(1, pstrErro, strTxtPesquisa)
        strErrSource = "MQSeries"
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio + Len(strTxtPesquisa)
            intErrPosicaoFim = InStr(intErrPosicaoInicio, pstrErro, ",")
            lngErrNumber = Mid(pstrErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
            strCodOcorrencia = "MQ-" & Format(lngErrNumber, "00000")
        End If
       
        Select Case lngErrNumber
            Case 2009
                strComplemento = "Queue Manager name: " & pstrQMgrName
                strOcorrencia = "A conexão com o gerenciador de fila foi perdida(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2035
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
                strOcorrencia = "O usuário não está autorizado a executar a operação(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2058, 2059
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "O gerenciador de fila não está disponível para conexão(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2085, 2086
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "Fila não Criada(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case 2162
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = "o gerenciador de fila está sendo encerrando(" & strComplemento & ")."
                fgGetError = strOcorrencia
            Case Else
                strComplemento = "Queue Manager name: " & pstrQMgrName & vbCrLf & _
                                 "Queue Name        : " & pstrQueueName
            
                strOcorrencia = strErrDesciption & " - (" & strComplemento & ")"
                fgGetError = strOcorrencia
        End Select
        
    ElseIf InStr(1, UCase(strErrDesciption), UCase("Microsoft OLE DB Provider for Oracle")) > 0 Then
        
        strTxtPesquisa = "ORA"
        
        intErrPosicaoInicio = InStr(1, pstrErro, strTxtPesquisa)
        
        strErrSource = "Oracle"
        
        If intErrPosicaoInicio > 0 Then
            intErrPosicaoInicio = intErrPosicaoInicio
            intErrPosicaoFim = InStr(intErrPosicaoInicio, pstrErro, ":", vbBinaryCompare)
            strCodOcorrencia = Mid(pstrErro, intErrPosicaoInicio, intErrPosicaoFim - intErrPosicaoInicio)
        End If
        
        Select Case strCodOcorrencia
            Case "ORA-12545"
                strOcorrencia = "ORA-12545: A conexão falhou porque o objeto ou host de destino não existe."
                fgGetError = strOcorrencia
            
            Case "ORA-12154"
                strOcorrencia = "ORA-12154: TNS - Não foi possível resorver nome de serviço."
                fgGetError = strOcorrencia
            
            Case "ORA-12541"
                strOcorrencia = "ORA-12541: TNS - não há listener."
                fgGetError = strOcorrencia
            
            Case "ORA-12500"
                strOcorrencia = "ORA-12500: TNS - Listener falhou ao iniciar um processo de servidor dedicado."
                fgGetError = strOcorrencia
            
            Case "ORA-00018"
                strOcorrencia = "ORA-00020: Maximum number of sessions exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00019"
                strOcorrencia = "ORA-00019: Maximum number of session licenses exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00022"
                strOcorrencia = "ORA-00022: Invalid session ID; access denied."
                fgGetError = strOcorrencia
            
            Case "ORA-00028"
                strOcorrencia = "ORA-00028 your session has been killed."
                fgGetError = strOcorrencia
            
            Case "ORA-00029"
                strOcorrencia = "ORA-00029 session is not a user session."
                fgGetError = strOcorrencia

            Case "ORA-00035"
                strOcorrencia = "ORA-00035 LICENSE_MAX_USERS cannot be less than current number of users."
                fgGetError = strOcorrencia
            
            Case "ORA-00036"
                strOcorrencia = "ORA-00036 maximum number of recursive SQL levels (string) exceeded."
                fgGetError = strOcorrencia

            Case "ORA-00052"
                strOcorrencia = "ORA-00052 maximum number of enqueue resources (string) exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-00052"
                strOcorrencia = "ORA-00052 maximum number of enqueue resources (string) exceeded"
                fgGetError = strOcorrencia
            
            Case "ORA-00053"
                strOcorrencia = "ORA-00053 maximum number of enqueues exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00055"
                strOcorrencia = "ORA-00055 maximum number of DML locks exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00057"
                strOcorrencia = "ORA-00057 maximum number of temporary table locks exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00063"
                strOcorrencia = "ORA-00063 maximum number of LOG_FILES exceeded"
                fgGetError = strOcorrencia
                
            Case "ORA-00107"
                strOcorrencia = "ORA-00107 failed to connect to ORACLE listener process"
                fgGetError = strOcorrencia
                
            Case "ORA-00115"
                strOcorrencia = "ORA-00115 connection refused; dispatcher connection table is full"
                fgGetError = strOcorrencia
                
            Case "ORA-00152"
                strOcorrencia = "ORA-00152 current session does not match requested session"
                fgGetError = strOcorrencia
                
            Case "ORA-00603"
                strOcorrencia = "ORA-00603 ORACLE server session terminated by fatal error"
                fgGetError = strOcorrencia
                
            Case "ORA-00606"
                strOcorrencia = "ORA-00606 Internal error code"
                fgGetError = strOcorrencia
                
            Case "ORA-00816"
                strOcorrencia = "ORA-00816 error message translation failed"
                fgGetError = strOcorrencia

            Case "ORA-00020"
                strOcorrencia = "ORA-00020: maximum number of processes exceeded."
                fgGetError = strOcorrencia
            
            Case "ORA-01034", "ORA-12535", "ORA-12560", "ORA-12541"
                strOcorrencia = "ORA-01034: ORACLE não disponível."
                fgGetError = strOcorrencia
        
            Case Else
                strOcorrencia = "Verificar Banco de Dados Oracle - " & strErrDesciption
                fgGetError = strOcorrencia
        End Select
    
    End If
        
    If strOcorrencia <> "" Then
        Call flGeraInformacaoAlerta(strOcorrencia, _
                                    Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                    strErrSource, _
                                    strErrDesciption)
    End If
        
    Set objDomError = Nothing
    
    Exit Function

ErrorHandler:
       
    Set objDomError = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Adicional um estrutura de xml em um node de um xml

Private Sub flAppendXML(ByRef objMapaNavegacao As MSXML2.DOMDocument30, _
                        ByVal pstrNodeContext As String, _
                        ByVal pstrXMLFilho As String)

Dim objDomFilho                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    'Leitura do XML Filho
    Set objDomFilho = CreateObject("MSXML2.DOMDocument.4.0")
    If Not objDomFilho.loadXML(pstrXMLFilho) Then
        '100 - Documento XML Inválido.
        'llCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
       
    'Setar o nivel que deverá entrar o XML Filho
     Set objDomNode = objMapaNavegacao.documentElement.selectSingleNode("//" & pstrNodeContext)

    'Adicionar XML Filho na Saida
    objDomNode.appendChild objDomFilho.childNodes.Item(0)
    
    Set objDomFilho = Nothing
    Set objDomNode = Nothing
    Exit Sub

ErrorHandler:
    
    Set objDomFilho = Nothing
    Set objDomNode = Nothing
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd

End Sub

'Obter xml de controle de threads

Private Function flObterControleThread() As String

Dim objSharedPropMem                        As Object

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    flObterControleThread = objSharedPropMem.GetSPMProperty("A6A7A8", "CONTROLETHREAD")
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Obter o numero de threas ativas de um componente

Private Function flObterNumeroThreadsAtivas(ByVal pstrNomeObjeto As String) As Long

Dim objSharedPropMem                        As Object
Dim strControleThread                       As String
Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
        
    strControleThread = flObterControleThread
    
    If Trim(strControleThread) = vbNullString Then
        flObterNumeroThreadsAtivas = 0
        Exit Function
    End If
    
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlControleThread.loadXML strControleThread
    
    With xmlControleThread.documentElement
        
        Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & UCase(Trim(pstrNomeObjeto)) & "']")
        
        If xmlNode Is Nothing Then
            flObterNumeroThreadsAtivas = 0
        Else
            flObterNumeroThreadsAtivas = CLng("0" & xmlNode.attributes.getNamedItem("QuantidadeThreadsAtivas").Text)
        End If
        
    End With
    
    Set xmlControleThread = Nothing
    Set xmlNode = Nothing
    Exit Function
ErrorHandler:
    
    Set xmlNode = Nothing
    Set xmlControleThread = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Obter total de threas ativas

Private Function flObterTotalThreadsAtivas() As Long

Dim objSharedPropMem                        As Object
Dim strControleThread                       As String
Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim lngTotalThreads                         As Long

On Error GoTo ErrorHandler
        
    strControleThread = flObterControleThread
    
    If Trim(strControleThread) = vbNullString Then
        flObterTotalThreadsAtivas = 0
        Exit Function
    End If
    
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlControleThread.loadXML strControleThread
    
    lngTotalThreads = 0
    
    For Each xmlNode In xmlControleThread.selectNodes("//Grupo_Thread")
        lngTotalThreads = lngTotalThreads + xmlNode.selectSingleNode("@QuantidadeThreadsAtivas").Text
    Next
    
    flObterTotalThreadsAtivas = lngTotalThreads
    
    Set xmlControleThread = Nothing
    Set xmlNode = Nothing
    Exit Function
ErrorHandler:
    Set xmlNode = Nothing
    Set xmlControleThread = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Raise error do objeto MSXML2.Domdocument

Public Function fgErroLoadXML(ByRef objDomDocument As MSXML2.DOMDocument40, _
                              ByVal pstrComponente As String, _
                              ByVal pstrClasse As String, _
                              ByVal pstrMetodo As String)
    

    Err.Raise objDomDocument.parseError.errorCode, pstrComponente & " - " & pstrClasse & " - " & pstrMetodo, objDomDocument.parseError.reason
    
End Function

'Verificar se o bando de dados do SLCC está OK

Private Sub flVerificarBancoDados()

Dim objA6A7A8CA                             As Object
Dim strTextoIncialErroDB                    As String
Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strErro                                 As String
Dim strErroAux                              As String
Dim strErroMQ                               As String

On Error GoTo ErrorHandler
    
    Set objA6A7A8CA = CreateObject("A6A7A8CA.clsConsulta")
    strTextoIncialErroDB = "Verificar Banco de Dados do SLCC"
    Call objA6A7A8CA.QuerySQL("SELECT * FROM DUAL")
        
    Set objA6A7A8CA = Nothing
    
    'Se a conexão com Oracle estiver OK e existir alerta de Oracle, exlcuir Alerta
    If flExisteInformacaoAlerta(enumSourceAlerta.Oracle) Then
        'Excluir alerta de Oracle
        Call flExcluirInformacaoAlerta(enumSourceAlerta.Oracle)
    End If

    Exit Sub
    
ErrorHandler:

    strErroAux = Err.Description

    Set objA6A7A8CA = Nothing
    
    If flExisteInformacaoAlerta(enumSourceAlerta.Oracle) Then Exit Sub
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlErro.loadXML(strErroAux) Then
        
        Call flGeraInformacaoAlerta(strTextoIncialErroDB, _
                                    Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                    "Oracle", _
                                    strErro)
        
        flGravaArquivo strTextoIncialErroDB & vbCrLf & strErroAux
        
    Else
        
        strErro = xmlErro.selectSingleNode("//Description").Text
        
        strErroMQ = fgGetError(strErroAux, strQMName, "")
            
        If Trim(strErroMQ) <> "" Then
            strErro = strErroMQ
        End If
                   
        flGravaArquivo strTextoIncialErroDB & vbCrLf & _
                       strErroMQ & vbCrLf & _
                       String(50, "*") & vbCrLf & _
                       strErroAux
    
    End If
   
    Set xmlErro = Nothing

End Sub

'Graver o xml de controle de threads no Shared Property Memory

Private Function flSetControleThread(ByVal strControleThread As String) As Boolean

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    Call objSharedPropMem.SetSPMProperty("A6A7A8", "CONTROLETHREAD", strControleThread)
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Processa as mensagens que estão nas filas mas não existe thread ativa para o processamento das mensagens

Private Function flProcessaMensagemFilaSemThreads(ByRef xmlCofiguracaoEntrada As MSXML2.DOMDocument40) As String

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objPublisher                            As Object
Dim strNomeObjeto                           As String
Dim lngNumeroThreadsAtivas                  As Long
Dim blnTemMensagens                         As Boolean
Dim lngNumeroMaxThreads                     As Long

On Error GoTo ErrorHandler
    
    For Each xmlNode In xmlConfiguracaoEntrada.selectNodes("//Repet_Parametros_Entrada/*")
        
        strNomeObjeto = Trim(xmlNode.selectSingleNode("NomeObjeto").Text)
        
        blnTemMensagens = flVerificaMensagemFila(xmlNode.selectSingleNode("NomeFila").Text)
        
        lngNumeroThreadsAtivas = flObterNumeroThreadsAtivas(strNomeObjeto)
        lngNumeroMaxThreads = CLng(xmlNode.selectSingleNode("QuantidadeMaxThreads").Text)
            
        If blnTemMensagens = True And lngNumeroThreadsAtivas = 0 Then
            'Controle Total de Treads Ativas
            If flObterTotalThreadsAtivas <= lngQuantidadeTotalThreads Then
                Set objPublisher = CreateObject("A6A7A8Publisher.clsPublisher")
                objPublisher.AcionaGerenciadores xmlNode.xml
                Set objPublisher = Nothing
            End If
        End If
    Next
    
   Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Verifica se a conexão com gerenciador de filas MQseries está OK

Private Function flVerificaMQSeries() As Boolean

Dim strErro                                 As String
Dim strErroMQ                               As String

On Error GoTo ErrorHandler
        
    'Pikachu - 29/07/2005
    'Conectar QueueManager somente uma vez , na inicialização
    
    If objMQSession Is Nothing Then
        Set objMQSession = CreateObject("MQAX200.MQSession")
        Set objMQQueueManager = objMQSession.AccessQueueManager(strQMName)
    End If
    
    If Not objMQQueueManager.IsConnected Then
        Set objMQSession = Nothing
        Set objMQSession = CreateObject("MQAX200.MQSession")
        Set objMQQueueManager = objMQSession.AccessQueueManager(strQMName)
    End If
        
    If Not objMQQueueManager.IsOpen Then
        Set objMQSession = Nothing
        Set objMQSession = CreateObject("MQAX200.MQSession")
        Set objMQQueueManager = objMQSession.AccessQueueManager(strQMName)
    End If
        
    flVerificaMQSeries = True
            
    If flExisteInformacaoAlerta(enumSourceAlerta.MQSeries) Then
        'Excluir alerta de MSMQ
        Call flExcluirInformacaoAlerta(enumSourceAlerta.MQSeries)
    End If
            
    Exit Function
ErrorHandler:
    Set objMQSession = Nothing
    
    strErro = Err.Description
        
    If flExisteInformacaoAlerta(enumSourceAlerta.MQSeries) Then Exit Function
    
    Call flGeraInformacaoAlerta("Verificar MQSeries (Queue Manager: " & strQMName & ")", _
                                Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                "MQSeries", _
                                strErro)
        
    flGravaArquivo "Verificar Gerenciador de Filas - MQSeries" & vbCrLf & _
                   "Queue Manager: " & strQMName & vbCrLf & _
                   strErro
        
       
End Function

'Verifica se a conexão com gerenciador de filas MSMQ está OK

Private Function flVerificaMSMQ() As Boolean

Dim objQueueInfo                                As Object 'MSMQQueueInfo
Dim objQueue                                    As Object 'MSMQ.MSMQQueue
Dim strErro                                     As String

On Error GoTo ErrorHandler
    
    flVerificaMSMQ = True
    
    Exit Function
    
    Set objQueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    objQueueInfo.PathName = ".\private$\a6a7a8"
    
    Set objQueue = objQueueInfo.Open(enumMSMQAcess.MQ_RECEIVE_ACCESS, enumMSMQShareMode.MQ_DENY_NONE)
    
    objQueue.Close
    
    Set objQueue = Nothing
    Set objQueueInfo = Nothing
            
    flDeletaMensagemMSMQDeadQueue
    
    If flExisteInformacaoAlerta(enumSourceAlerta.MSMQ) Then
        'Excluir alerta de MSMQ
        Call flExcluirInformacaoAlerta(enumSourceAlerta.MSMQ)
    End If
    
    flVerificaMSMQ = True
        
    Set objQueue = Nothing
    Set objQueueInfo = Nothing
        
    Exit Function
ErrorHandler:
    
    strErro = Err.Description
    
    Set objQueue = Nothing
    Set objQueueInfo = Nothing
       
    If flExisteInformacaoAlerta(enumSourceAlerta.MSMQ) Then Exit Function
    
    Call flGeraInformacaoAlerta("Verificar Gerenciador de Filas - MSMQ : " & strErro, _
                                Format(Now, "dd/mm/yyy HH:mm:ss"), _
                                "MSMQ", _
                                strErro)
        
    flGravaArquivo "Verificar Gerenciador de Filas - MSMQ" & vbCrLf & strErro
    
    Err.Clear
    
End Function

'Deleta as mensagens da fila .\private$\a6a7a8_deadqueue do MSMQ

Private Sub flDeletaMensagemMSMQDeadQueue()

Dim objQueueInfo                            As Object 'MSMQQueueInfo
Dim objQueue                                As Object 'MSMQ.MSMQQueue
Dim objMessage                              As Object 'MSMQ.MSMQMessage
    
On Error GoTo ErrorHandler
    
    Set objQueueInfo = CreateObject("MSMQ.MSMQQueueInfo")
    
    objQueueInfo.PathName = ".\private$\a6a7a8_deadqueue"
    
    Set objQueue = objQueueInfo.Open(enumMSMQAcess.MQ_RECEIVE_ACCESS, _
                                     enumMSMQShareMode.MQ_DENY_NONE)
    
    Set objMessage = CreateObject("MSMQ.MSMQMessage")
    Set objMessage = objQueue.Receive(ReceiveTimeout:=100)
        
    Do Until objMessage Is Nothing
        Set objMessage = objQueue.Receive(ReceiveTimeout:=100)
    Loop
    
    objQueue.Close
    objQueueInfo.Close
        
    Set objMessage = Nothing
    Set objQueue = Nothing
    Set objQueueInfo = Nothing
    
    
    Exit Sub
    
ErrorHandler:

    Set objMessage = Nothing
    Set objQueue = Nothing
    Set objQueueInfo = Nothing
        
    Err.Clear
End Sub

'Exclusão de infrmações de alerta de erro do A7

Private Function flExcluirInformacaoAlerta(ByVal plngSouce As enumSourceAlerta) As Boolean

Dim objSharedPropMem                        As Object
Dim xmlAlerta                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strxmlAlerta                            As String
Dim blnExisteAlerta                         As Boolean

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    strxmlAlerta = objSharedPropMem.GetSPMProperty("A7BUSALERTA", "ALERTA")
    Set objSharedPropMem = Nothing
        
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Trim(strxmlAlerta) = "" Or Trim(strxmlAlerta) = "0" Then Exit Function
    
    xmlAlerta.loadXML strxmlAlerta
        
    For Each xmlNode In xmlAlerta.selectNodes("Alerta/Repet_Alerta/*")
         Select Case plngSouce
            Case enumSourceAlerta.MQSeries
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MQSERIES" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
            Case enumSourceAlerta.MSMQ
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "MSMQ" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
            Case enumSourceAlerta.Oracle
                If UCase(xmlNode.selectSingleNode("DE_FONT_ERRO").Text) = "ORACLE" Then
                    xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").removeChild xmlNode
                End If
         End Select
    Next
    
    Call flGerarPropriedadeAlerta(xmlAlerta.xml)
    
    flExcluirInformacaoAlerta = True
    
    Set xmlAlerta = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlAlerta = Nothing
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Verifica se o processamento de alguma fila esta ociosa

Private Function flVerificaThreadSemMensagem(ByVal pstrNomeObjeto As String) As Long

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem
Dim strControleThread                       As String
Dim xmlControleThread                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
        
    strControleThread = flGetControleThread
        
    If Trim(strControleThread) = vbNullString Then Exit Function
        
    Set xmlControleThread = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlControleThread.loadXML strControleThread
    
    With xmlControleThread.documentElement
        
        Set xmlNode = .selectSingleNode("Grupo_Thread[@NomeObjeto='" & UCase(Trim(pstrNomeObjeto)) & "']")
        
        If Not xmlNode Is Nothing Then
            If DateDiff("s", fgDtHrStr_To_DateTime(xmlNode.selectSingleNode("@UltimaAtualizacao").Text), Now) > 30 Then
                xmlNode.selectSingleNode("@QuantidadeThreadsAtivas").Text = 0
                xmlNode.selectSingleNode("@UltimaAtualizacao").Text = Format(Now, "yyyymmddHHmmss")
            End If
        End If
    End With
        
    Call flSetControleThread(xmlControleThread.xml)
        
    Set xmlControleThread = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlControleThread = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

'Converção de um data string com formato YYYYMMDDHHMMSS para o formato data hora

Public Function fgDtHrStr_To_DateTime(ByVal strYYYYMMDDHHMMSS As String) As Date

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long
Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer
Dim dthAux                                  As Date

On Error GoTo ErrorHandler

    If Len(strYYYYMMDDHHMMSS) <> 14 Then
        If Trim(strYYYYMMDDHHMMSS) = "" Then
            fgDtHrStr_To_DateTime = "00:00:00"
            Exit Function
        Else
            'Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss
             Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss"
        End If
    End If

    lngAno = Mid(strYYYYMMDDHHMMSS, 1, 4)
    lngMes = Mid(strYYYYMMDDHHMMSS, 5, 2)
    lngDia = Mid(strYYYYMMDDHHMMSS, 7, 2)

    intHora = Mid(strYYYYMMDDHHMMSS, 9, 2)
    intMinuto = Mid(strYYYYMMDDHHMMSS, 11, 2)
    intSegundo = Mid(strYYYYMMDDHHMMSS, 13, 2)

    dthAux = DateSerial(lngAno, lngMes, lngDia)
    
    dthAux = DateAdd("H", intHora, dthAux)
    dthAux = DateAdd("n", intMinuto, dthAux)
    dthAux = DateAdd("s", intSegundo, dthAux)
        
    fgDtHrStr_To_DateTime = dthAux

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

'Ler as informções de controle de thread no Shared Property memory

Public Function flGetControleThread() As String

Dim objSharedPropMem                        As Object 'A6A7A8.clsSharedPropMem

On Error GoTo ErrorHandler
    
    Set objSharedPropMem = CreateObject("A6A7A8.clsSharedPropMem")
    flGetControleThread = objSharedPropMem.GetSPMProperty("A6A7A8", "CONTROLETHREAD")
    Set objSharedPropMem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objSharedPropMem = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function


'Obter o nome da fila

Private Function flObterNomeFila(ByVal plngenumFila As enumFilasEntrada) As String
    
    Select Case plngenumFila
        Case enumA7QEENTRADA
            flObterNomeFila = "A7Q.E.ENTRADA"
        Case enumA7QEMENSAGEMRECEBIDA
            flObterNomeFila = "A7Q.E.MENSAGEMRECEBIDA"
        Case enumA7QEREPORT
            flObterNomeFila = "A7Q.E.REPORT"
        Case enumA8QEENTRADA
            flObterNomeFila = "A8Q.E.ENTRADA"
        Case enumA6QEREMESSASUBRESERVA
            flObterNomeFila = "A6Q.E.REMESSASUBRESERVA"
        Case enumA6QEREMESSAFUTURO
            flObterNomeFila = "A6Q.E.REMESSAFUTURO"
    End Select

End Function

'Gravar as mensagens com erro nas filas de erro

Private Sub flGravaMesgFilaErro(ByVal pstrNomeFila As String, _
                                ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strNomeFila                              As String

On Error GoTo ErrorHandler

    strNomeFila = Replace(pstrNomeFila, ".", "")

    strNomeArquivoLogErro = App.Path & "\MsgErro\" & pstrNomeFila & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, pstrErro
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

Public Function fgCompletaString(ByVal pstrTexto As String, pstrCaractere As String, pintNumCasas As Integer, Optional ByVal pblnEsquerda As Boolean) As String

    If IsNull(pblnEsquerda) Then
        fgCompletaString = Trim(pstrTexto) + String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere)
    Else
        If pblnEsquerda Then
            fgCompletaString = String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere) + Trim(pstrTexto)
        Else
            fgCompletaString = Trim(pstrTexto) + String(pintNumCasas - Len(Trim(pstrTexto)), pstrCaractere)
        End If
    End If

End Function

Private Function flVerificaJanelaVerificacao() As Boolean

Dim dtmDataHoraAtual                        As Date
Dim dtmDataHoraInicio                       As Date
Dim dtmDataHoraFim                          As Date

On Error GoTo ErrorHandler

    dtmDataHoraAtual = Now
    dtmDataHoraInicio = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraInicioVerificacao, "0", 4, True) & "00")
    dtmDataHoraFim = fgDtHrStr_To_DateTime(Format$(Date, "YYYYMMDD") & fgCompletaString(strHoraFimVerificacao, "0", 4, True) & "00")
    
    If dtmDataHoraAtual >= dtmDataHoraInicio And dtmDataHoraAtual <= dtmDataHoraFim Then
        flVerificaJanelaVerificacao = False
    Else
        flVerificaJanelaVerificacao = True
    End If
    
    Exit Function

ErrorHandler:
        
    flGravaMesgFilaErro "VerificaJanela", Err.Description
    
    flVerificaJanelaVerificacao = True
        
    Err.Clear

End Function

Private Sub flInicializaValidaRemessa(ByVal pblnInicializa As Boolean)

On Error GoTo ErrorHandler
        
    If pblnInicializa Then
        Call Shell("NET START A6A8AtivaValidaRemessa", vbHide)
    Else
        Call Shell("NET STOP A6A8AtivaValidaRemessa", vbHide)
   End If
    
   Exit Sub

ErrorHandler:
        
    Err.Clear

End Sub

Private Function flVerificaCacheUsuario() As String

Dim objFuncao                               As Object 'A6A7A8.clsA6A7A8Funcoes

On Error GoTo ErrorHandler
            
    Set objFuncao = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    flVerificaCacheUsuario = objFuncao.ObterValorParametrosGerais("CACHE_CONTROLE_ACESSO/PERIODICIDADE_EM_MINUTOS")
    Set objFuncao = Nothing

    Exit Function

ErrorHandler:
    flVerificaCacheUsuario = "N"
    Err.Clear

End Function

Private Function flRemoveUsuarioCache() As Long

Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa
Dim objParam                                As Object

On Error GoTo ErrorHandler
                
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    Call objValidaRemessa.RemoveUsuariosCache
    Set objValidaRemessa = Nothing
                            
    Set objParam = CreateObject("A7Server.clsParametrosGerais")
    Call objParam.AtualizaParametroCacheAcesso
    Set objParam = Nothing
    
    Exit Function

ErrorHandler:
    Set objValidaRemessa = Nothing
    Set objParam = Nothing
    
    Err.Clear

End Function

Private Sub flRenovaCache()

Dim objValidaRemessa                        As Object 'A6A8ValidaRemessa.clsValidaRemessa

On Error GoTo ErrorHandler
                    
    'Pikachu - Implementação de recarregar o cache sem parar o serviço - Ativa Valida Remessa
    '29/07/2005
                
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    Call objValidaRemessa.RecarregarCache
    Set objValidaRemessa = Nothing
    
    Exit Sub

ErrorHandler:
    Set objValidaRemessa = Nothing
    
    Err.Clear

End Sub

