VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável por controlar o recebimento de mensagens do MQSeries.

Option Explicit


Public objValidaRemessa                     As Object 'A6A8ValidaRemessa.clsValidaRemessa

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

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

'Gerenciar recebimento das remessas do legado , PJ/PK para o A8
'Classe não transacional

Public Function ReceberMensagemMQ(ByVal pstrQName As String, _
                                  ByRef pstrLogErro As String, _
                                  ByRef pstrMessageIDHexOut As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorralationID As String) As Boolean

Dim objTRemessa                             As A8LQS.clsTRemessa
Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim xmlLogErro                              As MSXML2.DOMDocument40
Dim blnEnviaParaRemessaRejeitada            As Boolean
Dim strNomeFilaOriginal                     As String
Dim strNomeFilaDestino                      As String

On Error GoTo ErrorHandler
    
'    fgGravaArquivo "Cria XML", vbNullString
    
    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
        
'    fgGravaArquivo "Cria clsTRemessa", vbNullString
    
    Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
    
    Set objTRemessa.objValidaRemessa = objValidaRemessa
    
    ReceberMensagemMQ = objTRemessa.ReceberMensagemMQ(pstrQName, _
                                                      pstrLogErro, _
                                                      pstrMessageIDHexOut, _
                                                      plngBackOutCount, _
                                                      pstrCorralationID, _
                                                      xmlRemessa)
    Set objTRemessa = Nothing

    If pstrLogErro <> vbNullString Then
        
        blnEnviaParaRemessaRejeitada = True
        
        If Not xmlRemessa Is Nothing Then
            'Se remessa OK
            If xmlRemessa.parseError.errorCode = 0 Then
                
                Set xmlLogErro = CreateObject("MSXML2.DOMDocument.4.0")
                
                'Se Log erro OK
                If xmlLogErro.loadXML(pstrLogErro) Then
                
                    If Not xmlLogErro.selectSingleNode("//Number") Is Nothing Then
                        '3009 - Operação não localizada no SLCC para cancelamento.
                        '3110 - Mensagem de complementação sem inclusão
                        '4151 - Operações de especificação vinculadas a uma operaçao interna não podem ser canceladas.
                        '4264 - Não foi possível encontrar a mensagem BMC0012R1.
                        If xmlLogErro.selectSingleNode("//Number").Text = 3009 Or _
                           xmlLogErro.selectSingleNode("//Number").Text = 3110 Or _
                           xmlLogErro.selectSingleNode("//Number").Text = 4151 Or _
                           xmlLogErro.selectSingleNode("//Number").Text = 4264 Then
                            
                            'Se for o primeiro reprocessamento
                            If xmlRemessa.selectSingleNode("//NU_PROC") Is Nothing Then
                                
                                If xmlRemessa.selectSingleNode("//TP_SOLI") Is Nothing Then
                                    Call fgAppendNode(xmlRemessa, "SISMSG", "NU_PROC", "1")
                                    Call fgAppendNode(xmlRemessa, "MESG", "TP_SOLI", enumTipoSolicitacao.Complementacao)
                                Else
                                    Call fgAppendNode(xmlRemessa, "MESG", "NU_PROC", "1")
                                    Call fgAppendNode(xmlRemessa, "MESG", "IN_VALIDA_REME", "1")
                                End If
                                
                                blnEnviaParaRemessaRejeitada = False
                                
                                'Enviar para Fila de reprocessamento
                                Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
                                                                         
                                If xmlLogErro.selectSingleNode("//Number").Text = 3009 Then
                                    strNomeFilaOriginal = "A8Q.E.ENTRADA"
                                ElseIf xmlLogErro.selectSingleNode("//Number").Text = 4264 Then
                                    strNomeFilaOriginal = "A8Q.E.ENTRADA_BACEN"
                                Else
                                    strNomeFilaOriginal = "A8Q.E.ENTRADA_OPERACAO"
                                End If
                                
                                Call objTRemessa.ReprocessaOperacaoCancelamento(xmlRemessa, _
                                                                                strNomeFilaOriginal, _
                                                                                "A8Q.E.ESPERA_REPROC", _
                                                                                pstrMessageIDHexOut, _
                                                                                pstrLogErro)
                                                                                
                                Set objTRemessa = Nothing
                                
                            Else
                                
                                strNomeFilaOriginal = pstrQName
                                
                                If pstrQName = "A8Q.E.ESPERA_REPROC" Then
                                    strNomeFilaDestino = "A8Q.E.REPROCESSAMENTO"
                                Else
                                    strNomeFilaDestino = "A8Q.E.ESPERA_REPROC"
                                End If
                                     
                                'Numero de Tentativas para reprocessar as operações de cancelmento
                                Select Case xmlRemessa.documentElement.selectSingleNode("//TP_SOLI").Text
                                
                                    Case "2"
                                        
                                        If xmlLogErro.selectSingleNode("//Number").Text <> 4151 Then
                                            xmlRemessa.selectSingleNode("//NU_PROC").Text = xmlRemessa.selectSingleNode("//NU_PROC").Text + 1
                                            blnEnviaParaRemessaRejeitada = False
                                            'Enviar para Fila de reprocessamento
                                            Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
                                            Call objTRemessa.ReprocessaOperacaoCancelamento(xmlRemessa, _
                                                                                            strNomeFilaOriginal, _
                                                                                            strNomeFilaDestino, _
                                                                                            pstrMessageIDHexOut, _
                                                                                            pstrLogErro)
                                            Set objTRemessa = Nothing
                                            
                                        Else
                                            If xmlRemessa.selectSingleNode("//NU_PROC").Text >= 3 Then
                                                'Se não achar a operacao original apos 3 tentativas de reprocessamento ,
                                                'enviar para remessa rejeitada
                                                blnEnviaParaRemessaRejeitada = True
                                            Else
                                                blnEnviaParaRemessaRejeitada = False
                                                xmlRemessa.selectSingleNode("//NU_PROC").Text = xmlRemessa.selectSingleNode("//NU_PROC").Text + 1
                                                
                                                'Enviar para Fila de reprocessamento
                                                Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
                                                Call objTRemessa.ReprocessaOperacaoCancelamento(xmlRemessa, _
                                                                                                strNomeFilaOriginal, _
                                                                                                strNomeFilaDestino, _
                                                                                                pstrMessageIDHexOut, _
                                                                                                pstrLogErro)
                                                Set objTRemessa = Nothing
                                            End If
                                        End If
                                        
                                    Case "3", "4"
                                    
                                        If xmlRemessa.selectSingleNode("//NU_PROC").Text >= 3 Then
                                            'Se não achar a operacao original apos 3 tentativas de reprocessamento ,
                                            'enviar para remessa rejeitada
                                            blnEnviaParaRemessaRejeitada = True
                                        Else
                                            blnEnviaParaRemessaRejeitada = False
                                            
                                            xmlRemessa.selectSingleNode("//NU_PROC").Text = xmlRemessa.selectSingleNode("//NU_PROC").Text + 1
                                            'Enviar para Fila de reprocessamento
                                            Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
                                            Call objTRemessa.ReprocessaOperacaoCancelamento(xmlRemessa, _
                                                                                            strNomeFilaOriginal, _
                                                                                            strNomeFilaDestino, _
                                                                                            pstrMessageIDHexOut, _
                                                                                            pstrLogErro)
                                            Set objTRemessa = Nothing
                                            
                                        End If
                                End Select
                                
                            End If
                        
                        Else
                            'Verificar se é Erro de Oracle ou MQSeries
                            'Se erro de Oracle ou MQ não gravar no remessa rejeitada
                             If fgVerificaErroOracleMQSeries(pstrLogErro) Then
                                blnEnviaParaRemessaRejeitada = False
                             Else
                                blnEnviaParaRemessaRejeitada = True
                             End If
                        End If
                    Else
                        blnEnviaParaRemessaRejeitada = True
                    End If
                Else
                    blnEnviaParaRemessaRejeitada = True
                End If
            End If
        Else
            blnEnviaParaRemessaRejeitada = False
        End If
    End If
            
    If blnEnviaParaRemessaRejeitada Then
        
        If Not xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
            xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vbNullString
        Else
            fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", ""
        End If
        
        Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
        
        objTRemessa.RemessaRejeitada pstrQName, _
                                     pstrLogErro, _
                                     pstrMessageIDHexOut, _
                                     plngBackOutCount, _
                                     xmlRemessa
         Set objTRemessa = Nothing
         
         pstrLogErro = ""
    End If

    Set xmlRemessa = Nothing

Exit Function
ErrorHandler:

    pstrLogErro = Err.Description
    
    Set objTRemessa = Nothing

End Function

