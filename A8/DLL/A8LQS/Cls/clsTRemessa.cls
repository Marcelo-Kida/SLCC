VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável por controlar as transações no recebimento de mensagens

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Public objValidaRemessa                     As Object 'A6A8ValidaRemessa.clsValidaRemessa

Private ObjectContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

'Gerenciar recebimento das remessas do legado , PJ/PK das filas MQSeries (Get)
'Classe Transacional
Public Function ReceberMensagemMQ(ByVal pstrQName As String, _
                                  ByRef pstrLogErro As String, _
                                  ByRef pstrMessageIDHexOut As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorrelationID As String, _
                                  ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim strMessageDescriptor                    As String
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strErros                                As String
Dim dblLog                                  As Double

On Error GoTo ErrorHandler

    'fgLogText "RMMQ", dblLog

'    fgGravaArquivo "Cria clsMQ", vbNullString
    
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")

    objMQAX200.MQInit
    objMQAX200.MQOpen pstrQName, MQOO_INPUT_SHARED

    lngBufferLen = Len(strBuffer)
    
    If objMQAX200.MQGet(lngBufferLen, _
                        strBuffer, _
                        True, _
                        Binario, _
                        pstrMessageIDHexOut, _
                        pstrCorrelationID, _
                        strMessageDescriptor, _
                        plngBackOutCount) Then
        
        objMQAX200.MQClose

        Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlRemessa.preserveWhiteSpace = True
        
        If Not xmlRemessa.loadXML(strBuffer) Then
            ' 3032 - XML da Operação Inválido.
            lngCodigoErroNegocio = 3032
            GoTo ErrorHandler
        End If
    
        strErros = vbNullString

        'Chamada do ValidarRemessa
        'fgLogText "RMMQ Mensagem recebida", dblLog
        If flValidarRemessa(xmlRemessa, strErros) Then
            'fgLogText "RMMQ Mensagem validada", dblLog
            flVerificarRemessa xmlRemessa, strErros
            'fgLogText "RMMQ Mensagem processada", dblLog
        End If

        ReceberMensagemMQ = True
        
        If strErros <> vbNullString Then
            Err.Description = strErros
            GoTo ErrorHandler
        End If
    Else
        ReceberMensagemMQ = False
    End If

    objMQAX200.MQClose
    objMQAX200.MQEnd

    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    'fgLogText "RMMQ fim", dblLog

Exit Function
ErrorHandler:

    pstrLogErro = Err.Description
    
    If pstrLogErro = vbNullString And lngCodigoErroNegocio <> 0 Then
        pstrLogErro = lngCodigoErroNegocio & " - " & fgObterDescricaoErro(lngCodigoErroNegocio)
    End If
    
    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If

End Function

'Altera o XML de remessa para incluir os campos básicos de operação
Private Sub flConfiguraRemessaComoOperacao(ByRef pxmlremessa As MSXML2.DOMDocument40)

Dim strDocumentoNomeBase                    As String
Dim lngTipoOperacao                         As Long

On Error GoTo ErrorHandler

    With pxmlremessa.documentElement

        strDocumentoNomeBase = .baseName

        If .selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
            fgAppendNode pxmlremessa, strDocumentoNomeBase, "NU_SEQU_OPER_ATIV", ""
        End If

        flObterTipoOperacaoPorMsgBUS pxmlremessa, lngTipoOperacao

        If .selectSingleNode("TP_OPER") Is Nothing Then
            fgAppendNode pxmlremessa, strDocumentoNomeBase, "TP_OPER", lngTipoOperacao
        End If

        If .selectSingleNode("NO_CNPT") Is Nothing Then
            fgAppendNode pxmlremessa, strDocumentoNomeBase, "NO_CNTP", vbNullString
        End If

        If .selectSingleNode("NU_COMD_OPER") Is Nothing Then
            fgAppendNode pxmlremessa, strDocumentoNomeBase, "NU_COMD_OPER", vbNullString
        End If

        If .selectSingleNode("VA_OPER_ATIV") Is Nothing Then
            fgAppendNode pxmlremessa, strDocumentoNomeBase, "VA_OPER_ATIV", vbNullString
        End If

    End With

Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConfiguraRemessaComoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Gerenciar o recebimento da remessa por tipo de mensagem

Private Function flVerificarRemessa(ByRef pxmlremessa As MSXML2.DOMDocument40, _
                                    ByRef pstrErros As String) As Boolean

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

Dim objProcessaMensagemSTR                  As A8LQS.clsProcessoMensagemSTR

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objProcessoMensagem                     As A8LQS.clsProcessoMensagem
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objOperacao                             As A8LQS.clsOperacao
Dim objLivreMovimentacao                    As A8LQS.clsLivreMovimentacao
Dim objProcessoMensagemSTR                  As A8LQS.clsProcessoMensagemSTR

Dim strNumeroControleIF                     As String
Dim strMensagem                             As String
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strHeader                               As String

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim xmlErro                                 As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    If Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemSPB Then

        Set objProcessoMensagem = CreateObject("A8LQS.clsProcessoMensagem")
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

        strMensagem = pxmlremessa.documentElement.selectSingleNode("TX_MESG/*").xml
        
        strHeader = Mid(Trim(pxmlremessa.documentElement.selectSingleNode("TX_MESG").Text), 1, 200)

        xmlMensagem.loadXML strMensagem
        objProcessoMensagem.GerenciarRecebimentoMensagem xmlMensagem, strHeader

        Set objProcessoMensagem = Nothing
        Set xmlMensagem = Nothing

    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemSTR0010R2PZA8 Then

        Set objProcessoMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
        objProcessoMensagemSTR.GerenciarRecebimentoMensagem pxmlremessa, vbNullString
        Set objProcessoMensagemSTR = Nothing

    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemSPBErro Then

        Set objProcessoMensagem = CreateObject("A8LQS.clsProcessoMensagem")
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

        If pxmlremessa.documentElement.selectSingleNode("TX_MESG/*") Is Nothing Then
            
            'Pikachu -14 / 6 / 2004
            'Inclusão tratamento das mensagens Rejeitadas pelo Piloto - WorkFlow
            'Se o codigo do erro igual 5156 - Mensagem Cancelada - Workflow
            
            strHeader = pxmlremessa.documentElement.selectSingleNode("TX_MESG").Text
            udtProtocoloErroNZAux.String = strHeader
            LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
            
            '------------------------------------------------------------------------------------
            If IsNumeric(Trim$(udtProtocoloErroNZ.CodigoErro1)) Then
                If CLng("0" & Trim$(udtProtocoloErroNZ.CodigoErro1)) = 5156 Then
                    
                    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                    Set objOperacao = CreateObject("A8LQS.clsOperacao")
                    
                    strNumeroControleIF = Trim$(udtProtocoloErroNZ.ControleRemessaNZ)
                                                   
                    Call xmlMensagem.loadXML(objMesgRecbEnviSPB.LerMensagemIda(strNumeroControleIF))
                                                                                   
                    If xmlMensagem.xml = vbNullString Then
                        Err.Description = "Não foi encontrada a mensagem " & Trim$(udtProtocoloErroNZ.CodigoMensagem) & " de ida. Número de Controle IF da mensagem de erro inválido."
                        GoTo ErrorHandler
                    End If
                
                    Set xmlMensagem = objOperacao.LerXMLOperacao(xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                                                                                   
                    fgAppendNode xmlMensagem, xmlMensagem.documentElement.nodeName, "SISMSG", ""
                    
                    If Right$(Trim(udtProtocoloErroNZ.CodigoMensagem), 1) = "E" Then
                        fgAppendNode xmlMensagem, "SISMSG", "CodMsg", udtProtocoloErroNZ.CodigoMensagem
                    Else
                        fgAppendNode xmlMensagem, "SISMSG", "CodMsg", Trim$(udtProtocoloErroNZ.CodigoMensagem) & "E"
                    End If
                    
                    fgAppendNode xmlMensagem, "SISMSG", "NumCtrlIF", strNumeroControleIF
                    fgAppendNode xmlMensagem, "SISMSG", "DtMovto", udtProtocoloErroNZ.DataRemessa
                    
                    strMensagem = xmlMensagem.xml
                    Set objMesgRecbEnviSPB = Nothing
                    Set objOperacao = Nothing
                    
                Else
                '------------------------------------------------------------------------------------
                    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
                    Call fgAdicionaErro(xmlErro, 3039)
                    Err.Description = xmlErro.xml
                    Set xmlErro = Nothing
                    GoTo ErrorHandler
                End If
            Else
                If UCase$(Left$(strHeader, 3)) = "STR" Then
                    Set objProcessaMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
                    
                    Call xmlMensagem.loadXML(pxmlremessa.xml)
                    Call objProcessaMensagemSTR.GerenciarMensagemErroNZ(xmlMensagem, strHeader)
                    Set objProcessaMensagemSTR = Nothing
                    
                    GoTo ExitFunction
                End If
            End If
        Else
            strMensagem = pxmlremessa.documentElement.selectSingleNode("TX_MESG/*").xml
            strHeader = Mid(Trim(pxmlremessa.documentElement.selectSingleNode("TX_MESG").Text), 1, 638)
            
            'se a mensagem de erro for do NZ, chama a função flIncluirAtributoCodErroTagsErroNZ para deixar a sinalização
            'das tags com erro, igual à forma das mensagens de erro do bacen (mensagens "E")
            strMensagem = objProcessoMensagem.flIncluirAtributoCodErroTagsErroNZ(strMensagem, strHeader)
        End If

        xmlMensagem.loadXML strMensagem
        objProcessoMensagem.GerenciarRecebimentoMensagem xmlMensagem, strHeader

        Set objProcessoMensagem = Nothing
        Set xmlMensagem = Nothing

        
    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemPZErro Then
        
        strHeader = Mid(pxmlremessa.documentElement.selectSingleNode("TX_MESG").Text, 21)

        Set objProcessaMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
        Call objProcessaMensagemSTR.GerenciarMensagemErroPZ(strHeader)
        Set objProcessaMensagemSTR = Nothing
    
    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemPZR1 Then
        
        strHeader = pxmlremessa.documentElement.selectSingleNode("//TX_HEAD_NZ").Text

        Set objProcessaMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
        Call objProcessaMensagemSTR.GerenciarMensagemR1(pxmlremessa, strHeader)
        Set objProcessaMensagemSTR = Nothing
    
    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.MensagemPZR2 Then
        
        strHeader = Mid(pxmlremessa.documentElement.selectSingleNode("TX_MESG").Text, 21)

        Set objProcessaMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
        Call objProcessaMensagemSTR.ProcessaMensagemSTR0008R2Comti(strHeader)
        Set objProcessaMensagemSTR = Nothing
    
    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.CargaVeiculoLegal Then

        Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
        objVeiculoLegal.ProcessarRemessa pxmlremessa
        Set objVeiculoLegal = Nothing

    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.RetornoContaCorrenteBG Or _
           Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.RetornoContaCorrenteDV Then

        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        objContaCorrente.GerenciarRecebimento pxmlremessa
        Set objContaCorrente = Nothing

    ElseIf Val(pxmlremessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.LivreMovimentacao Then

        Set objLivreMovimentacao = CreateObject("A8LQS.clsLivreMovimentacao")
        objLivreMovimentacao.GerenciarRecebimento pxmlremessa, pstrErros
        Set objLivreMovimentacao = Nothing

    Else

        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.GerenciarRecebimentoOperacao pxmlremessa, pstrErros
        Set objProcessoOperacao = Nothing
        
        If pstrErros <> vbNullString Then
            flVerificarRemessa = False
            Err.Description = pstrErros
            GoTo ErrorHandler
        End If

    End If

ExitFunction:
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    pstrErros = fgLimpaCaracterEspecial(Err.Description)
    
    Set objMesgRecbEnviSPB = Nothing
    Set objProcessoOperacao = Nothing
    Set objProcessoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlMensagem = Nothing
    Set xmlErro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
End Function

'Inclusão da remessa na tabela de remessa rejeitada e geração de alerta

Public Function RemessaRejeitada(ByVal pstrQName As String, _
                                 ByRef pstrErros As String, _
                                 ByRef pstrMessageIDHexOut As String, _
                                 ByRef plngBackOutCount As Long, _
                                 ByVal xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim objRemessaRejeitada                     As A8LQS.clsRemessaRejeitada
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlErros                                As MSXML2.DOMDocument40
Dim xmlRemessaRejeitada                     As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strMessageDescriptor                    As String
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strErros                                As String
Dim blnVeiculoLegalValido                   As String

Dim objFilaErro                             As Object 'A7Server.clsFilaErro

    On Error GoTo ErrorHandler

    ' ************** Tirar a mensagem da fila
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")

    objMQAX200.MQInit
    objMQAX200.MQOpen pstrQName, MQOO_INPUT_SHARED

    lngBufferLen = Len(strBuffer)
    
    objMQAX200.MQGet lngBufferLen, _
                     strBuffer, _
                     True, _
                     Binario, _
                     pstrMessageIDHexOut, _
                     , _
                     strMessageDescriptor, _
                     plngBackOutCount

    objMQAX200.MQClose
    ' ************** Tirar a mensagem da fila

    Set xmlRemessaRejeitada = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlRemessaRejeitada, "", "REMESSAREJEITADA", ""
    fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TX_XML_REME", ""

    Set xmlNode = xmlRemessaRejeitada.documentElement.selectSingleNode("TX_XML_REME")
    xmlNode.appendChild fgCreateCDATASection(xmlRemessa.xml)

    fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TX_XML_ERRO", ""

    Set xmlNode = xmlRemessaRejeitada.documentElement.selectSingleNode("TX_XML_ERRO")
    xmlNode.appendChild fgCreateCDATASection(pstrErros)

    fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "TP_MESG", xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "CO_EMPR", xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text
    fgAppendNode xmlRemessaRejeitada, "REMESSAREJEITADA", "SG_SIST_INFO", xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text

    Set objRemessaRejeitada = CreateObject("A8LQS.clsRemessaRejeitada")
    objRemessaRejeitada.Incluir xmlRemessaRejeitada

    'Carga de veículo legal grava somente a remessa rejeitada
    If Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.CargaVeiculoLegal And _
       Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.MensagemSPB And _
       Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.MensagemSPBNaoTratada And _
       Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.MensagemSPBErro And _
       Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.MensagemPZR1 Then
        
        'Para tratar o Veículo Legal Inválido.
        Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
        xmlErros.loadXML pstrErros
        blnVeiculoLegalValido = True
        For Each xmlNode In xmlErros.selectNodes("Erro/Grupo_ErrorInfo")
            If xmlNode.selectSingleNode("Number").Text = "4007" Then
                blnVeiculoLegalValido = False
            End If
        Next
    
        'Enviar remessa rejeitada para o Legado
        fgRemessaRejeitadaLegado xmlRemessa, xmlErros, True

        If blnVeiculoLegalValido Then
            If Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.MensagemSPB Then
                If xmlRemessa.documentElement.selectSingleNode("//CO_FATO_GERA_ALER") Is Nothing Then
                    Set objAlerta = CreateObject("A8LQS.clsAlerta")
                    Call flConfiguraRemessaComoOperacao(xmlRemessa)
                    objAlerta.GerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.RemessaLegadoDadosInvalidos
                    Set objAlerta = Nothing
                Else
                    Set objAlerta = CreateObject("A8LQS.clsAlerta")
                    Call flConfiguraRemessaComoOperacao(xmlRemessa)
                    objAlerta.GerarAlerta xmlRemessa, , , xmlRemessa.documentElement.selectSingleNode("//CO_FATO_GERA_ALER").Text
                    Set objAlerta = Nothing
                End If

            End If
        Else
            Set objAlerta = CreateObject("A8LQS.clsAlerta")
            Call flConfiguraRemessaComoOperacao(xmlRemessa)
            objAlerta.GerarAlertaVeiculoLegalInvalido xmlRemessa
            Set objAlerta = Nothing
        
            Set objFilaErro = CreateObject("A7Server.clsFilaErro")
            
            Call objFilaErro.PutFilaErro(pstrQName, _
                                         "A7Q.E.ERRO", _
                                         xmlRemessa.xml, _
                                         xmlErros.xml, _
                                         pstrMessageIDHexOut, _
                                         True)
            
            Set objFilaErro = Nothing
        
        End If
    End If

    Set xmlErros = Nothing
    Set xmlRemessaRejeitada = Nothing
    Set objRemessaRejeitada = Nothing

    pstrErros = vbNullString

    objMQAX200.MQEnd
    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    pstrErros = Err.Description

    Set objRemessaRejeitada = Nothing
    Set objAlerta = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
End Function

'Obter tipo de operaçao da remessa

Private Function flObterTipoOperacaoPorMsgBUS(ByRef pxmlremessa As MSXML2.DOMDocument40, _
                                              ByRef plngTipoOperacao As Long) As Long

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao

On Error GoTo ErrorHandler

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Call objTipoOperacao.ObterTipoOperacaoPorMsgBUS(pxmlremessa, plngTipoOperacao, "", "")
    Set objTipoOperacao = Nothing

    
    Exit Function
ErrorHandler:

    Set objTipoOperacao = Nothing
    
    Err.Clear
    
    plngTipoOperacao = 0
End Function

'Camada para a rotina de liberação de mensagens agendadas

Public Function LiberarAgendamento() As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida

    On Error GoTo ErrorHandler

    'Liberar as operações
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    objMensagemEnviadaRecebida.LiberarAgendamento
    Set objMensagemEnviadaRecebida = Nothing

    LiberarAgendamento = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'reprocessamento de operações de cancelamento TO_SOLI = 3
Public Function ReprocessaOperacaoCancelamento(ByVal pxmlremessa As MSXML2.DOMDocument40, _
                                               ByVal pstrNomeFilaOrigem As String, _
                                               ByVal pstrNomeFilaDestino As String, _
                                               ByVal pstrMessageID As String, _
                                               ByRef pstrErros As String) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim lngPrioridade                           As Long
Dim blnTemMensagem                          As Boolean

    On Error GoTo ErrorHandler

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    
    Select Case pxmlremessa.selectSingleNode("//TP_SOLI").Text
        Case enumTipoSolicitacao.Complementacao
            lngPrioridade = 9
        Case enumTipoSolicitacao.Cancelamento, enumTipoSolicitacao.CancelamentoComMensagem
            lngPrioridade = 7
        Case Else
            lngPrioridade = 9
    End Select
    
    objMQAX200.MQInit

    'Retirar a mensagem original da fila
    objMQAX200.MQOpen pstrNomeFilaOrigem, MQOO_INPUT_SHARED
    
    If objMQAX200.MQGet(lngBufferLen, strBuffer, True, True, pstrMessageID) Then
        
        strBuffer = pxmlremessa.xml
        
        objMQAX200.MQClose
        objMQAX200.MQOpen pstrNomeFilaDestino, MQOO_OUTPUT
        lngBufferLen = Len(strBuffer)
        
        objMQAX200.MQPut strBuffer, _
                         lngBufferLen, _
                         enumPutOptions.Binario, _
                         "", _
                         "", _
                         "", _
                         "", _
                         0, _
                         lngPrioridade
        
        objMQAX200.MQClose
    
    End If
    
    objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    pstrErros = vbNullString

    Exit Function

ErrorHandler:
    pstrErros = Err.Description
    
    Set xmlRemessa = Nothing
    Set objMQAX200 = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
End Function

'Chamada para a rotina de liquidacao TQ

Public Function LiquidarOperacaoTQ(ByVal plngSequenciaOperacao As Long) As Boolean

    Dim objConciliacao              As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler

    'Liberar as operações
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    objConciliacao.LiquidarOperacaoTQ plngSequenciaOperacao
    Set objConciliacao = Nothing

    LiquidarOperacaoTQ = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarOperacaoTQ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Chamada para a rotina de verificação de fluxo de operação com liquidação futura
Public Function GerenciarOperacaoLiquidacaoFutura(ByVal plngSequenciaOperacao As Long) As Boolean

Dim objWorkflow                             As A8LQS.clsWorkflow
Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(plngSequenciaOperacao)
    Call fgAppendNode(xmlOperacao, "MESG", "CO_MESG", vbNullString)
    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.EmSer
    
    Call objOperacao.AlterarStatusAcao(plngSequenciaOperacao, enumStatusOperacao.EmSer, vbNullString, , , False)
    
    Call xmlControleProcessoOperacao.loadXML( _
                objProcessoOperacao.LerTodos("RecebimentoBoleto", _
                                             CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)))

    If xmlControleProcessoOperacao.xml <> vbNullString Then
        Call objProcessoOperacao.EnviarPrevisao(xmlOperacao, xmlControleProcessoOperacao)
    End If

    Call objWorkflow.GerenciarChamadas(xmlOperacao, "RecebimentoOperacao")
    
    GerenciarOperacaoLiquidacaoFutura = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Set objWorkflow = Nothing
    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objWorkflow = Nothing
    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarOperacaoLiquidacaoFutura Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function EnviarConsultaPZ(ByVal plngCodigoEmpresa As Long, _
                                 ByVal pstrConsultaPZ As String) As Boolean

Dim objProcessaMensagemSTR              As A8LQS.clsProcessoMensagemSTR

On Error GoTo ErrorHandler

    'Enviar Consulta PZ
    Set objProcessaMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
    Call objProcessaMensagemSTR.EnviarConsultaPZ(plngCodigoEmpresa, pstrConsultaPZ)
    Set objProcessaMensagemSTR = Nothing

    EnviarConsultaPZ = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objProcessaMensagemSTR = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarConsultaPZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function


'Função do ValidarRemessa
Private Function flValidarRemessa(ByRef pxmlremessa As MSXML2.DOMDocument40, _
                                  ByRef pstrErros As String) As Boolean



On Error GoTo ErrorHandler
    
    If Not pxmlremessa.selectSingleNode("//IN_VALIDA_REME") Is Nothing Then
        pstrErros = vbNullString
        flValidarRemessa = True
        Exit Function
    End If
    
    'If objValidaRemessa Is Nothing Then
'        fgGravaArquivo "Cria clsValidaRemessa", vbNullString
        Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")
    'End If
        
    flValidarRemessa = False
    
    pstrErros = objValidaRemessa.ValidarMensagemA8(pxmlremessa.xml)
    
    Set objValidaRemessa = Nothing
    
    If pstrErros = vbNullString Then
        flValidarRemessa = True
    End If


Exit Function
ErrorHandler:

    flValidarRemessa = False
    
    pstrErros = ""
    
    Set objValidaRemessa = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basA8LQS", "fgValidarRemessa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function



