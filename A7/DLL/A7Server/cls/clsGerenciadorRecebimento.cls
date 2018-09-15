VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorRecebimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF7546400CD"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelo recebimento de mensagens no sistema A7.
Option Explicit

Implements ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private objMQAX200                          As A6A7A8CA.clsMQAX200
Private ObjectContext                       As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'Método responsável por retirar a mensagem da fila MQSeries;
'Executar o método clsMQAX200.MQGet;
Public Function ReceberMensagemMQ(ByVal pstrNomeFila As String, _
                                  ByRef pstrRetornoLogErro As String, _
                                  ByRef pstrMessageIdHex As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorrelationID As String, _
                         Optional ByVal pstrProtocolo As String = vbNullString) As Boolean
    

Dim lngLenMensagemRecebida                   As Long
Dim strMensagemRecebida                      As String
Dim strMessageDescriptor                     As String

On Error GoTo ErrorHandler
    
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    
    Call objMQAX200.MQInit
    
    Call objMQAX200.MQOpen(Trim(pstrNomeFila), MQOO_INPUT_SHARED)
     
    If objMQAX200.MQGet(lngLenMensagemRecebida, _
                        strMensagemRecebida, _
                        True, _
                        True, _
                        pstrMessageIdHex, _
                        pstrCorrelationID, _
                        strMessageDescriptor, _
                        plngBackOutCount) Then
        
        objMQAX200.MQClose
        
        Call flProcessaMensagem(pstrProtocolo, _
                                strMensagemRecebida, _
                                strMessageDescriptor, _
                                Trim(pstrNomeFila), _
                                pstrCorrelationID)
              
        ReceberMensagemMQ = True
    Else
        ReceberMensagemMQ = False
    End If
        
    Call objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing
       
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    
    'Caso ocorra algum erro não tratado, desfazer a transação e devolver o erro
    'para fazer o put na fila de erro para que o processo não pare
    
    pstrRetornoLogErro = Err.Description
    
    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If

End Function

'Validar a mensagem recebida consistindo dados como tipo de mensagem, regra de transporte, sistema e empresa.
Private Function flAutenticarMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objRegraTransporte                        As A7Server.clsRegraTransporte

On Error GoTo ErrorHandler
    
    Set objRegraTransporte = CreateObject("A7Server.clsRegraTransporte")
    flAutenticarMensagem = objRegraTransporte.AutenticarMensagem(pxmlMensagem)
    Set objRegraTransporte = Nothing
    
    Exit Function
ErrorHandler:

    Set objRegraTransporte = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAutenticarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar a mensagem recebida.
Private Function flSalvarMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                  ByVal plngCodigoOcorrencia As Long) As Long

Dim objMensagem                       As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    flSalvarMensagem = objMensagem.SalvarMensagem(pxmlMensagem, plngCodigoOcorrencia)
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Processar os dados da mensagem recebida consistindo seus dados e formatando-a para o formato do sistema A7.
Private Function flProcessaMensagem(ByVal pstrProtocolo As String, _
                                    ByVal pstrMensagemRecebida As String, _
                                    ByVal pstrMessageDescriptor As String, _
                                    ByVal pstrNomeFila As String, _
                                    ByRef pstrCorrelationID As String) As Boolean

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim lngCodigoMensagem                       As Long
Dim strDetlaheOcorrenciaErro                As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim udtProtocoloNZ                          As udtProtocoloNZ
Dim udtProtocoloNZAux                       As udtProtocoloNZAux

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim udtHeaderDV                             As udtHeaderDV
Dim udtHeaderDVAux                          As udtHeaderDVAux

Dim udtPZW0001                              As udtPZW0001
Dim udtPZW0001Aux                           As udtPZW0001Aux

Dim udtPZO00140                             As udtPZO00140
Dim udtPZO00140Aux                          As udtPZO00140Aux

Dim udtPZO00141                             As udtPZO00141
Dim udtPZO00141Aux                          As udtPZO00141Aux

Dim udtPZW0916                              As udtPZW0916
Dim udtPZW0916Aux                           As udtPZW0916Aux

Dim udtSTR0010                              As udtSTR0010
Dim udtSTR0010aux                           As udtSTR0010aux

Dim udtConsultaPZ                           As udtConsultaPZ
Dim udtConsultaPZAux                        As udtConsultaPZAux

Dim udtSTR0008R2                            As udtSTR0008R2
Dim udtSTR0008R2Aux                         As udtSTR0008R2Aux

Dim rsRegraSistDest                         As ADODB.Recordset

'Envio do carimbo para o sistema NZ
Dim objCarimbador                           As A7Server.clsCarimbador
Dim objRespostaOcorrencia                   As A7Server.clsRespostaOcorrencia
Dim lngTipoFluxoCarimbo                     As enumFluxoMonitor
Dim lngStatusCarimbo                        As enumStatusMonitor
Dim blnEnviarCarimbo                        As Boolean
Dim strNumCtrlIF                            As String
Dim strCodMsd                               As String
Dim lngEmpresa                              As Long
Dim blnMensagemR1                           As Boolean

Dim strMensagem                             As String
Dim xmlTestaMensagem                        As MSXML2.DOMDocument40


On Error GoTo ErrorHandler
        
    pstrMensagemRecebida = Replace(pstrMensagemRecebida, "&", " ")
    
    blnEnviarCarimbo = False
    blnMensagemR1 = False
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML flObterPropriedadesMensagem
        
    'Se pstrProtocolo <> "" o sistema origem é o NZ e destino A8
    If Trim(pstrProtocolo) = vbNullString Then
                
        'Protocolo Operação
        udtProtocoloAux.String = Mid(pstrMensagemRecebida, 1, Len(udtProtocoloAux.String))
        LSet udtProtocolo = udtProtocoloAux
        
        With xmlMensagem.documentElement
            .selectSingleNode("TP_MESG").Text = udtProtocolo.TipoMensagem
            .selectSingleNode("CO_EMPR_ORIG").Text = udtProtocolo.CodigoEmpresa
            .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
            .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
        End With
        
        Set objNode = xmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR")
        objNode.appendChild flCreateCDATASection(pstrMensagemRecebida)
    Else
        
        'Protocolo Operação
        udtProtocoloAux.String = pstrProtocolo
        LSet udtProtocolo = udtProtocoloAux
                
        If CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemA8NZR1_HD Then
                
            'Protocolo SPB - NZ - Tipo Mensagem 1000
            udtProtocoloNZAux.String = Mid(pstrMensagemRecebida, 1, Len(udtProtocoloNZAux.String))
            LSet udtProtocoloNZ = udtProtocoloNZAux
            
            strMensagem = Mid(pstrMensagemRecebida, Len(udtProtocoloNZAux.String) + 1)
            
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(enumTipoMensagemEntrada.MensagemNZA8)
                .selectSingleNode("CO_EMPR_ORIG").Text = udtProtocoloNZ.CodigoEmpresa
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
                
            pstrMensagemRecebida = flCompoeMensagemConsulta(udtProtocoloNZ.CodigoEmpresa, _
                                                            udtProtocoloNZAux.String, _
                                                            udtProtocoloNZ.NuOP, _
                                                            pstrNomeFila)
                
            'Informações para o carimbador NZ
            blnEnviarCarimbo = True
            strCodMsd = udtProtocoloNZ.CodigoMensagem
            strNumCtrlIF = udtProtocoloNZ.ControleRemessaNZ
            lngEmpresa = udtProtocoloNZ.CodigoEmpresa
            
            lngTipoFluxoCarimbo = FLUXO_MONITOR_NORMAL
            lngStatusCarimbo = MONITOR_RECEB_R1
            blnMensagemR1 = True
            
                
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemNZA8 Then
        
            'Protocolo SPB - NZ - Tipo Mensagem 1000
            udtProtocoloNZAux.String = Mid(pstrMensagemRecebida, 1, Len(udtProtocoloNZAux.String))
            LSet udtProtocoloNZ = udtProtocoloNZAux
            
            'Protocolo PZ - Tipo Mensagem 1010
            udtPZW0916Aux.String = udtProtocoloNZAux.String
            LSet udtPZW0916 = udtPZW0916Aux
            
            'Fazer put para A7NET nas mensagens novas da CAM
            If Left(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 3) = "CAM" _
            And (Left(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 7) <> "CAM0001" _
            And Left(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 7) <> "CAM0002" _
            And Left(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 7) <> "CAM0003" _
            And Left(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 7) <> "CAM0004") Then
            
                With objMQAX200
                    .MQOpen "A7Q.E.ENTRADA_NET", enumMQOO_Open.MQOO_OUTPUT
                    .MQPut pstrProtocolo & pstrMensagemRecebida, Len(pstrProtocolo & pstrMensagemRecebida), enumPutOptions.Binario
                    .MQClose
                End With
                
                Exit Function
            
            End If
            
            strMensagem = Mid(pstrMensagemRecebida, Len(udtProtocoloNZAux.String) + 1)
            
            Set xmlTestaMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            
            'If (Trim$(udtPZW0916.NumeroControleLegado) <> String$(23, "0") And _
            '    Trim$(udtPZW0916.NumeroControleLegado) <> vbNullString) Then
            
            If Not xmlTestaMensagem.loadXML(strMensagem) Then

                blnEnviarCarimbo = False

                With xmlMensagem.documentElement
                    .selectSingleNode("TP_MESG").Text = Trim(enumTipoMensagemEntrada.MensagemPZR1)
                    .selectSingleNode("CO_EMPR_ORIG").Text = udtPZW0916.CodigoEmpresa
                    .selectSingleNode("SG_SIST_ORIG").Text = "PZ"
                    .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
                End With

                udtProtocolo.CodigoEmpresa = udtPZW0916.CodigoEmpresa
                udtProtocolo.SiglaSistemaOrigem = "PZ "
                udtProtocolo.SiglaSistemaDestino = "A8 "
                udtProtocolo.TipoMensagem = enumTipoMensagemEntrada.MensagemPZR1

                LSet udtProtocoloAux = udtProtocolo

                pstrMensagemRecebida = udtProtocoloAux.String & pstrMensagemRecebida

            Else
                With xmlMensagem.documentElement
                    .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                    .selectSingleNode("CO_EMPR_ORIG").Text = udtProtocoloNZ.CodigoEmpresa
                    .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                    .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
                End With
                
                If InStr(1, pstrNomeFila, "CONS_HD") > 0 Then
                    pstrMensagemRecebida = flCompoeMensagemConsulta(udtProtocoloNZ.CodigoEmpresa, _
                                                                    udtProtocoloNZAux.String, _
                                                                    udtProtocoloNZ.NuOP, _
                                                                    pstrNomeFila)
                End If
                
                If CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemSTR0010R2PZA8 Then
                    pstrMensagemRecebida = udtProtocoloAux.String & pstrMensagemRecebida
                End If
                    
                'Informações para o carimbador NZ
                blnEnviarCarimbo = True
                strCodMsd = udtProtocoloNZ.CodigoMensagem
                strNumCtrlIF = udtProtocoloNZ.ControleRemessaNZ
                lngEmpresa = udtProtocoloNZ.CodigoEmpresa
            
            End If
            
            Set xmlTestaMensagem = Nothing
            
            If Right(Trim$(UCase$(udtProtocoloNZ.CodigoMensagem)), 2) = "R1" Then
                lngTipoFluxoCarimbo = FLUXO_MONITOR_NORMAL
                lngStatusCarimbo = MONITOR_RECEB_R1
                blnMensagemR1 = True
            Else
                'R2 , AVISOS E INFORMAÇÕES
                lngTipoFluxoCarimbo = FLUXO_MONITOR_MSG_EXTERNA
                lngStatusCarimbo = MONITOR_RECEB_MENS_EXTERNA
            End If
            
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemSTR0010R2PZA8 Then

            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim$(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = "558"
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With

            pstrMensagemRecebida = udtProtocoloAux.String & pstrMensagemRecebida
        
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemRetornoPZOk Then

            Set xmlMensagem = Nothing
            Exit Function

        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemErroNZA8 Then
            
            'Protocolo Erro - NZ - Tipo Mensagem 1003
            udtProtocoloErroNZAux.String = Mid(pstrMensagemRecebida, 1, Len(udtProtocoloErroNZAux.String))
            LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
            
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = udtProtocoloErroNZ.CodigoEmpresa
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
            
            'ENVIO DO CARIMBO PARA SISTEMA NZ
            blnEnviarCarimbo = True
            strCodMsd = udtProtocoloErroNZ.CodigoMensagem
            strNumCtrlIF = udtProtocoloErroNZ.ControleRemessaNZ
            lngEmpresa = udtProtocoloErroNZ.CodigoEmpresa
            lngTipoFluxoCarimbo = FLUXO_MONITOR_NORMAL
            If Trim$(udtProtocoloErroNZ.CodigoErro1) = "5156" Then
                'RECEBIMENTO DO CANCELAMENTO DO PILOTO
                lngStatusCarimbo = MONITOR_RECEB_CANCEL_PILOTO
            Else
                If Right$(Trim$(UCase$(udtProtocoloErroNZ.CodigoMensagem)), 1) = "E" Then
                    'Erro externo
                    lngStatusCarimbo = MONITOR_RECEB_ERRO_EXTERNO
                Else
                    'Erro interno
                    lngStatusCarimbo = MONITOR_RECEB_ERRO_NZ
                End If
            End If
        
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemRetornoDV Then
        
            'Protocolo Retorno - DV - Tipo Mensagem 1006
            udtHeaderDVAux.String = Mid(pstrMensagemRecebida, 1, Len(udtHeaderDVAux.String))
            LSet udtHeaderDV = udtHeaderDVAux
        
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = udtHeaderDV.CodigoEmpresa
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
            
            udtProtocolo.CodigoEmpresa = udtHeaderDV.CodigoEmpresa
            LSet udtProtocoloAux = udtProtocolo
            
            pstrMensagemRecebida = udtProtocoloAux.String & fgLimpaCaracterInvalido(pstrMensagemRecebida)
            
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemRetornoBG Then
            
            'Protocolo Retorno - BG - Tipo Mensagem 1008
            
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = "00558"
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
            
            udtProtocolo.CodigoEmpresa = "00558"
            LSet udtProtocoloAux = udtProtocolo
            
             pstrMensagemRecebida = Replace(pstrMensagemRecebida, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & "?>", vbNullString)
             pstrMensagemRecebida = Replace(pstrMensagemRecebida, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>", vbNullString)
             pstrMensagemRecebida = fgLimpaCaracterInvalido(pstrMensagemRecebida)
             
             pstrMensagemRecebida = flTrataMensagemRetornoBG(pstrMensagemRecebida, pstrCorrelationID)
             
             'Caso a mensagem de retorno do BG com XML inválido.
             If pstrMensagemRecebida = vbNullString Then
                             
                Call flSalvarMensagemRejeitada(pstrMensagemRecebida, _
                                               pstrMessageDescriptor, _
                                               "Mensagem retorno do sistema BG com XML Inválido.", _
                                               pstrNomeFila)
                
                Set xmlMensagem = Nothing
                Exit Function
             End If
             
             pstrMensagemRecebida = udtProtocoloAux.String & pstrMensagemRecebida
             
       
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemPZErro Then
        
            'Protocolo Erro PZ - Tipo Mensagem 1010
            udtPZW0001Aux.String = pstrMensagemRecebida
            LSet udtPZW0001 = udtPZW0001Aux
                        
            lngEmpresa = enumCodigoEmpresa.Meridional
            
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = lngEmpresa
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
            
            udtProtocolo.CodigoEmpresa = lngEmpresa
            LSet udtProtocoloAux = udtProtocolo
            
            pstrMensagemRecebida = udtProtocoloAux.String & fgLimpaCaracterInvalido(pstrMensagemRecebida)
        
        ElseIf CLng(udtProtocolo.TipoMensagem) = enumTipoMensagemEntrada.MensagemPZR2 Then
                
            If Mid(pstrMensagemRecebida, 38, 9) = "STR0008R2" Then
            
                Select Case CLng(Mid(pstrMensagemRecebida, 1, 3))
                    Case 33
                        lngEmpresa = enumCodigoEmpresa.Banespa
                    Case 351
                        lngEmpresa = enumCodigoEmpresa.Bozano
                    Case 8
                        lngEmpresa = enumCodigoEmpresa.Meridional
                    Case 353
                        lngEmpresa = enumCodigoEmpresa.Santander
                End Select
            
            Else
                udtSTR0008R2Aux.String = pstrMensagemRecebida
                LSet udtSTR0008R2 = udtSTR0008R2Aux
                
                lngEmpresa = udtSTR0008R2.CO_EMPR
            End If
                
            lngEmpresa = enumCodigoEmpresa.Meridional
            
            With xmlMensagem.documentElement
                .selectSingleNode("TP_MESG").Text = Trim(udtProtocolo.TipoMensagem)
                .selectSingleNode("CO_EMPR_ORIG").Text = lngEmpresa
                .selectSingleNode("SG_SIST_ORIG").Text = UCase(Trim(udtProtocolo.SiglaSistemaOrigem))
                .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(udtProtocolo.SiglaSistemaDestino))
            End With
            
            udtProtocolo.CodigoEmpresa = lngEmpresa
            LSet udtProtocoloAux = udtProtocolo
            
            pstrMensagemRecebida = udtProtocoloAux.String & fgLimpaCaracterInvalido(pstrMensagemRecebida)
                
        End If
    
        Set objNode = xmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR")
        objNode.appendChild flCreateCDATASection(pstrMensagemRecebida)
    
    End If
    
    If blnEnviarCarimbo Then
        'Envio do carimbo para sistema NZ
        Set objCarimbador = CreateObject("A7Server.clsCarimbador")
        objCarimbador.IncluirAvisoMonitor objMQAX200, _
                                          lngEmpresa, _
                                          strNumCtrlIF, _
                                          strCodMsd, _
                                          lngTipoFluxoCarimbo, _
                                          lngStatusCarimbo, _
                                          "A8", _
                                          "A8", _
                                          pstrNomeFila, _
                                          fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
        
        Set objCarimbador = Nothing
    End If
    
    If flAutenticarMensagem(xmlMensagem) Then
                
        'Enviar Somente a um sistema destino
        If Trim(xmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text) <> vbNullString Then
                
            'Salvar Evento em Transito
'            xmlMensagem.documentElement.selectSingleNode("CO_MESG_SAID").Text = 1
            'lngCodigoMensagem = flSalvarMensagem(xmlMensagem, enumOcorrencia.RecebimentoBemSucedido)
                   
            'xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = lngCodigoMensagem
            
            Call flProcessaEnvioMensagem(xmlMensagem)
        
        'Se o sistema destino for nulo enviar a todos os sistemas cadastrados
        Else
            
            Set rsRegraSistDest = flObterRegraSistemaDestino(xmlMensagem)
                        
            While Not rsRegraSistDest.EOF
                
                udtProtocoloAux.String = Mid(pstrMensagemRecebida, 1, Len(udtProtocoloAux.String))
                LSet udtProtocolo = udtProtocoloAux
                                
                udtProtocolo.SiglaSistemaDestino = UCase(Trim(rsRegraSistDest!SG_SIST_DEST))
                                
                With xmlMensagem.documentElement
                    .selectSingleNode("SG_SIST_DEST").Text = UCase(Trim(rsRegraSistDest!SG_SIST_DEST))
                End With
                
                LSet udtProtocoloAux = udtProtocolo
                
                pstrMensagemRecebida = Mid(pstrMensagemRecebida, Len(udtProtocolo) + 1)
                pstrMensagemRecebida = udtProtocoloAux.String & pstrMensagemRecebida
            
                xmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text = vbNullString
                
                Set objNode = xmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR")
                objNode.appendChild flCreateCDATASection(pstrMensagemRecebida)
                
                'Salvar Evento em Transito
                'lngCodigoMensagem = flSalvarMensagem(xmlMensagem, enumOcorrencia.RecebimentoBemSucedido)
                       
                'xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = lngCodigoMensagem
                
                Call flProcessaEnvioMensagem(xmlMensagem)
                
                rsRegraSistDest.MoveNext
            Wend
        End If
    Else
        
        Call flSalvarMensagemRejeitada(pstrMensagemRecebida, _
                                       pstrMessageDescriptor, _
                                       xmlMensagem.selectSingleNode("//TX_DTLH_OCOR_ERRO").Text, _
                                       pstrNomeFila)
    
        Set objRespostaOcorrencia = CreateObject("A7Server.clsRespostaOcorrencia")
        objRespostaOcorrencia.EnviarRespostaOcorrencia xmlMensagem, enumOcorrencia.RejeicaoNaoAutenticidade
        Set objRespostaOcorrencia = Nothing
        
    End If
    
    Set xmlMensagem = Nothing
    
    Exit Function

ErrorHandler:
    If InStr(1, Err.Description, "An invalid character was found in text content") > 0 Then
        strDetlaheOcorrenciaErro = "Foram encontrados Caracteres inválidos."
        xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO").appendChild fgCreateCDATASection(strDetlaheOcorrenciaErro)
        Call flSalvarMensagemRejeitada(pstrMensagemRecebida, _
                                       pstrMessageDescriptor, _
                                       xmlMensagem.selectSingleNode("//TX_DTLH_OCOR_ERRO").Text, _
                                       pstrNomeFila)
        Set xmlMensagem = Nothing
        Exit Function
    End If
    
    Set xmlTestaMensagem = Nothing
    Set objRespostaOcorrencia = Nothing
    Set objCarimbador = Nothing
    Set xmlMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar a mensagem recebida para o processamento do gerenciador de envio.
Private Function flProcessaEnvioMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim strNomeFila                              As String
Dim objGerEnvio                              As A7Server.clsGerenciadorEnvio


On Error GoTo ErrorHandler
    
    strNomeFila = "A7Q.E.MENSAGEMRECEBIDA"
        
    Call flProcessaEnvio(pxmlMensagem)
    
    'objMQAX200.MQOpen strNomeFila, enumMQOO_Open.MQOO_OUTPUT
       
    'objMQAX200.MQPut pxmlMensagem.xml, _
                     Len(pxmlMensagem.xml), _
                     enumPutOptions.Binario
    
    'objMQAX200.MQClose


    Exit Function
ErrorHandler:

    Set objGerEnvio = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaEnvioMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Nome Fila:" & strNomeFila)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_MESG.
Private Function flObterPropriedadesMensagem()

Dim objMensagem                             As A7Server.clsMensagem

On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A7Server.clsMensagem")
    flObterPropriedadesMensagem = objMensagem.ObterPropriedades
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterPropriedadesMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar mensagens rejeitadas no recebimento do sistema A7.
Private Sub flSalvarMensagemRejeitada(ByVal pstrMensagemRecebida As String, _
                                      ByVal pstrMessageDescriptor As String, _
                                      ByVal pstrDetalheOcorrenciaErro As String, _
                             Optional ByVal pstrNomeFila As String = vbNullString)

Dim xmlMessageDescriptor                     As MSXML2.DOMDocument40
Dim objMensagemRejeitada                     As A7Server.clsMensagemRejeitada
Dim objFilaErro                              As A7Server.clsFilaErro

Dim strNomeAplicativo                        As String
Dim strMessageIdHex                          As String
Dim dtmDataPut                               As Date

On Error GoTo ErrorHandler

    Set xmlMessageDescriptor = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMessageDescriptor.loadXML pstrMessageDescriptor
    
    With xmlMessageDescriptor.documentElement
        dtmDataPut = fgDtHrStr_To_DateTime(.selectSingleNode("PutDateTime").Text)
        strNomeAplicativo = Trim(.selectSingleNode("PutApplicationName").Text)
        strMessageIdHex = Trim(.selectSingleNode("MessageIdHex").Text)
    End With
    
    Set xmlMessageDescriptor = Nothing
    
    Set objMensagemRejeitada = CreateObject("A7Server.clsMensagemRejeitada")
    
    Call objMensagemRejeitada.SalvarMensagemRejeitada(pstrMensagemRecebida, _
                                                      enumOcorrencia.RejeicaoNaoAutenticidade, _
                                                      strMessageIdHex, _
                                                      strNomeAplicativo, _
                                                      dtmDataPut, _
                                                      pstrDetalheOcorrenciaErro)
    Set objMensagemRejeitada = Nothing

    Set objFilaErro = CreateObject("A7Server.clsFilaErro")
    
    Call objFilaErro.PutFilaErro(pstrNomeFila, _
                                 "A7Q.E.ERRO", _
                                 pstrMensagemRecebida, _
                                 pstrDetalheOcorrenciaErro, _
                                 strMessageIdHex, _
                                 True)
    
    Set objFilaErro = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set xmlMessageDescriptor = Nothing
    Set objMensagemRejeitada = Nothing
    Set objFilaErro = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarMensagemRejeitada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Função auxiliar para a criação de campos do tipo CDATA em XMLs.
Private Function flCreateCDATASection(ByRef psConteudo As String) As MSXML2.IXMLDOMCDATASection

Dim xmlDoc                               As MSXML2.DOMDocument40
Dim objnodeCDATA                         As IXMLDOMCDATASection
Dim xmlNode                              As IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objnodeCDATA = xmlDoc.createCDATASection(psConteudo)

    Set flCreateCDATASection = objnodeCDATA
    
    Set xmlDoc = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlDoc = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCreateCDATASection Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Obtém os sistemas destinos da mensagem de recebida no sistema A7.
Private Function flObterRegraSistemaDestino(ByRef xmlMensagem As MSXML2.DOMDocument40) As ADODB.Recordset

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    With xmlMensagem.documentElement
        strSQL = " SELECT B.SG_SIST_DEST" & _
                 "   FROM A7.TB_REGR_TRAP_MESG A," & _
                 "        A7.TB_REGR_SIST_DEST B " & _
                 "  WHERE B.DH_INIC_VIGE_REGR_TRAP = A.DH_INIC_VIGE_REGR_TRAP " & _
                 "    AND B.CO_EMPR_ORIG = A.CO_EMPR_ORIG " & _
                 "    AND  B.SG_SIST_ORIG = A.SG_SIST_ORIG " & _
                 "    AND  B.TP_MESG = A.TP_MESG " & _
                 "    AND  A.TP_MESG      =  '" & IIf(IsNumeric(.selectSingleNode("TP_MESG").Text), Val(.selectSingleNode("TP_MESG").Text), Trim(.selectSingleNode("TP_MESG").Text)) & "'" & _
                 "    AND  A.SG_SIST_ORIG = '" & Trim(.selectSingleNode("SG_SIST_ORIG").Text) & "'" & _
                 "    AND  A.CO_EMPR_ORIG =  " & CLng(.selectSingleNode("CO_EMPR_ORIG").Text) & _
                 "    AND (A.DT_FIM_VIGE_REGR_TRAP  IS NULL " & _
                 "     OR  A.DT_FIM_VIGE_REGR_TRAP  >=  " & fgDataHoraServidor_To_Oracle & ")"
    End With
         
    Set flObterRegraSistemaDestino = fgQuerySQL(strSQL)
     
    Exit Function
    
ErrorHandler:
      
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterRegraSistemaDestino Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Compor a mensagem recebida lendo as mensagens de detalhe da mesma.
Private Function flCompoeMensagemConsulta(ByVal plngCodigoEmpresa As Long, _
                                          ByVal pstrProtocoloNZ As String, _
                                          ByVal pstrNUOP As String, _
                                          ByVal pstrNomeFila As String) As String

Dim strMensagem                             As String
Dim strMensagemAux                          As String
Dim strNomeFileDetalhe                      As String
Dim blnTemMensagem                          As Boolean
Dim strMessageIdHex                         As String
Dim udtHeaderDetalheMensagem                As udtHeaderDetalheMensagem
Dim udtHeaderDetalheMensagemAux             As udtHeaderDetalheMensagemAux

Dim xmlFilas                                As MSXML2.DOMDocument40
On Error GoTo ErrorHandler
    

    Set xmlFilas = CreateObject("MSXML2.DOMDocument.4.0")
    
   'RATS 364 - Ocor 01
   'Paremetrizar Nome da fila com
       '<NomeFilaConsDT Empresa='523'>A8Q.E.CONS_DT.1</NomeFilaConsDT>
       '<NomeFilaConsDT Empresa='701'>A8Q.E.CONS_DT.2</NomeFilaConsDT>
    
    If xmlFilas.Load(App.Path & "\FilasNET.xml") Then
        
        If plngCodigoEmpresa = enumCodigoEmpresa.Banespa Then
            If Not xmlFilas.selectSingleNode("//NomeFilaConsDT[@Empresa='" & enumCodigoEmpresa.Banespa & "']") Is Nothing Then
                strNomeFileDetalhe = xmlFilas.selectSingleNode("//NomeFilaConsDT[@Empresa='" & enumCodigoEmpresa.Banespa & "']").Text
            Else
                strNomeFileDetalhe = "A8Q.E.CONS_DT"
            End If
        Else
            If Not xmlFilas.selectSingleNode("//NomeFilaConsDT[@Empresa='" & enumCodigoEmpresa.Santander & "']") Is Nothing Then
                strNomeFileDetalhe = xmlFilas.selectSingleNode("//NomeFilaConsDT[@Empresa='" & enumCodigoEmpresa.Santander & "']").Text
            Else
                strNomeFileDetalhe = "A8Q.E.CONS_DT"
            End If
        End If
    Else
        strNomeFileDetalhe = "A8Q.E.CONS_DT"
    End If
    
    Set xmlFilas = Nothing
        
    blnTemMensagem = True
         
    While blnTemMensagem
        strMensagem = vbNullString
        
        Call objMQAX200.MQOpen(strNomeFileDetalhe, MQOO_INPUT_SHARED)
        
        blnTemMensagem = objMQAX200.MQGet(Len(strMensagem), _
                                          strMensagem, _
                                          True, _
                                          True, _
                                          , _
                                          objMQAX200.flToEBCDIC(pstrNUOP & "0"))
        
        objMQAX200.MQClose
        
        If strMensagem <> vbNullString Then
            strMensagemAux = strMensagemAux & Mid(strMensagem, Len(udtHeaderDetalheMensagem) + 1)
        End If
    
    Wend
    
    strMensagemAux = pstrProtocoloNZ & strMensagemAux
    
    flCompoeMensagemConsulta = strMensagemAux
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCompoeMensagemConsulta Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Aplicar tratamento na mensagem recebida para o retorno de informações ao sistema BG. (CorrelationID da mensagem de integração)
Private Function flTrataMensagemRetornoBG(ByVal pstrMensagemRecebida As String, _
                                          ByVal pstrCorrelationID As String) As String

Dim xmlMensagemRetorno                      As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlMensagemRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMensagemRetorno.loadXML(pstrMensagemRecebida) Then
        flTrataMensagemRetornoBG = vbNullString
    Else
        Call fgAppendNode(xmlMensagemRetorno, "replyMsg", "NU_CTRL_ITGR", Trim(pstrCorrelationID))
        
        flTrataMensagemRetornoBG = xmlMensagemRetorno.xml
        
    End If

    Set xmlMensagemRetorno = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlMensagemRetorno = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTrataMensagemRetornoBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-----------------------------------------------------------------
' clsGerenciadorEnvio

Private Function flProcessaEnvio(ByRef pobjDomMensagem As MSXML2.DOMDocument40)

Dim objTradutor                             As A7Server.clsTradutor
Dim blnTraducaoBemSucedida                  As Boolean
Dim blnTraduzirMensagem                     As Boolean
Dim strSistemaOrigem                        As String
Dim strSistemaDestino                       As String
Dim strCodigoMensagem                       As String

    On Error GoTo ErrorHandler
    
    blnTraduzirMensagem = True
    
'    strSistemaOrigem = vbNullString
'    If Not pobjDomMensagem.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
'        strSistemaOrigem = Trim$(pobjDomMensagem.selectSingleNode("//SG_SIST_ORIG").Text)
'    End If
    
    'Inclui Tag de Repeticao na mensagem recebida pelo GPC ou R2
'    If strSistemaOrigem = "GPC" Or strSistemaOrigem = "R2" Then
'        blnTraduzirMensagem = flIncluiRepeticao("TX_CNTD_ENTR", pobjDomMensagem)
'    End If
    
    'Inclui Tag de Grupo na mensagem recebida pelo NZ e enviada para A8 - Somente CAM
'    If strSistemaOrigem = "NZ" Then
'        strCodigoMensagem = Trim(Mid(pobjDomMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 4, 9))
'        Select Case UCase(strCodigoMensagem)
'            Case "CAM0021R2", "CAM0022R2", "CAM0023R2", "CAM0024R2", "CAM0025R2", _
'                 "CAM0026R2", "CAM0030R2", "CAM0031R2", "CAM0032R2", "CAM0039R2"
'                blnTraduzirMensagem = flIncluiGrupo("TX_CNTD_ENTR", strCodigoMensagem, pobjDomMensagem)
'        End Select
'    End If
    
    If blnTraduzirMensagem Then
        Set objTradutor = CreateObject("A7Server.clsTradutor")
        blnTraducaoBemSucedida = objTradutor.TraduzirMensagem(pobjDomMensagem, "")
        Set objTradutor = Nothing
    Else
        blnTraducaoBemSucedida = False
    End If
    
    'Exclui Tag de Repeticao da mensagem recebida pelo GPC ou R2, para manter a mensagem original
'    If strSistemaOrigem = "GPC" Or strSistemaOrigem = "R2" Then
'        Call flRetiraRepeticao("TX_CNTD_ENTR", pobjDomMensagem)
'    End If
    
    'Exclui Tag de Grupo da mensagem recebida pelo NZ, para manter a mensagem original
'    If strSistemaOrigem = "NZ" Then
'        Select Case UCase(strCodigoMensagem)
'            Case "CAM0021R2", "CAM0022R2", "CAM0023R2", "CAM0024R2", "CAM0025R2", _
'                 "CAM0026R2", "CAM0030R2", "CAM0031R2", "CAM0032R2", "CAM0039R2"
'                 Call flRetiraGrupo("TX_CNTD_ENTR", strCodigoMensagem, pobjDomMensagem)
'        End Select
'    End If
            
    If blnTraducaoBemSucedida Then
        
        strSistemaDestino = vbNullString
        If Not pobjDomMensagem.selectSingleNode("//SG_SIST_DEST") Is Nothing Then
            strSistemaDestino = Trim$(pobjDomMensagem.selectSingleNode("//SG_SIST_DEST").Text)
        End If
        
        If strSistemaDestino <> "A7" Then
            
            'Exclui Tag de Repeticao da mensagem recebida pelo A8 e enviada ao GPC ou R2
'            If strSistemaDestino = "GPC" Or strSistemaDestino = "R2" Then
'                Call flRetiraRepeticao("TX_CNTD_SAID", pobjDomMensagem)
'            End If
            
            'Exclui Tag de Grupo da mensagem recebida pelo A8 e enviada para NZ - Somente CAM
'            If strSistemaDestino = "NZ" Then
'                If Not pobjDomMensagem.selectSingleNode("//TP_MESG") Is Nothing Then
'                    strCodigoMensagem = pobjDomMensagem.selectSingleNode("//TP_MESG").Text
'                End If
'                Select Case UCase(strCodigoMensagem)
'                    Case "CAM0021", "CAM0022", "CAM0023", "CAM0024", "CAM0025", "CAM0026", _
'                         "CAM0028", "CAM0030", "CAM0031", "CAM0032", "CAM0033", "CAM0039"
'                        Call flRetiraGrupo("TX_CNTD_SAID", strCodigoMensagem, pobjDomMensagem)
'                End Select
'            End If
        
            'Postar Mensagem Traduzida
            Call flPostarMensagemTraduzida(pobjDomMensagem)
        End If
        
        Call flSalvarMensagem(pobjDomMensagem, enumOcorrencia.PostagemBemSucedida)
    Else
        'Postar Mensagem Traduzida
        Call flSalvarMensagem(pobjDomMensagem, enumOcorrencia.CanceladaErroTradução)
        
        'Pikachu - Demanda 1026
        'Enviar Resposta para sistema Legado - Layout de Resposta
        Call EnviarMensagemRejeicaoLegado(pobjDomMensagem, enumOcorrencia.CanceladaErroTradução)
    End If

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaEnvioMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Preparar a memsagem recebida para o envio ao sistema destino.
Private Function flPostarMensagemTraduzida(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objSistema                              As MSXML2.DOMDocument40
Dim strSistema                              As String
Dim strAlerta                               As String
Dim strNomeFila                             As String

'Variaveis utilizadas pelo Carimbo NZ
Dim objCarimbador                           As A7Server.clsCarimbador
Dim xmlMensagemSPB                          As MSXML2.DOMDocument40
Dim strNumCtrlIF                            As String
Dim strValor                                As String
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
        
    strSistema = flObterSistema(pxmlMensagem)
    
    If Trim(strSistema) = vbNullString Then
        strAlerta = "Nome da Fila não cadastrada para o sistema: " & pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text
        Call fgGeraInformacaoAlerta(strAlerta, Format(Now, "dd/mm/yyyy HH:mm:ss"), "Gerenciador de envio", strAlerta)
        strNomeFila = ""
    Else
        
        If Trim(pxmlMensagem.documentElement.selectSingleNode("NO_FILA_MQSE_DEST").Text) = vbNullString Then
            Set objSistema = CreateObject("MSXML2.DOMDocument.4.0")
            objSistema.loadXML strSistema
            strNomeFila = Trim(objSistema.documentElement.selectSingleNode("Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text)
            Set objSistema = Nothing
        Else
            strNomeFila = Trim(pxmlMensagem.documentElement.selectSingleNode("NO_FILA_MQSE_DEST").Text)
        End If
    
        If pxmlMensagem.selectSingleNode("//SG_SIST_DEST").Text = "NZ" Then
            'Enviar carimbo
            On Error Resume Next
            
            'Obtem a mensagem de entrada pois esta contem as tags necessarias com nomes já padronizados pelo A8
            Set xmlMensagemSPB = CreateObject("MSXML2.DOMDocument.4.0")
            xmlMensagemSPB.loadXML Mid$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 21)
           
            'Obtem o Numero de Controle IF
            If Not xmlMensagemSPB.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
                strNumCtrlIF = xmlMensagemSPB.selectSingleNode("//NU_CTRL_IF").Text
            Else
                strNumCtrlIF = vbNullString
            End If
            
            'Obtem o Código da Mensagem
            If Not xmlMensagemSPB.selectSingleNode("//CO_MESG") Is Nothing Then
                strCodigoMensagem = xmlMensagemSPB.selectSingleNode("//CO_MESG").Text
            ElseIf Not IsNumeric(Left$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 9)) Then
                strCodigoMensagem = Trim$(Left$(pxmlMensagem.selectSingleNode("//TX_CNTD_ENTR").Text, 9))
            Else
                strCodigoMensagem = vbNullString
            End If
            
            'Obtem o Valor da Mensagem
            If Not xmlMensagemSPB.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                strValor = xmlMensagemSPB.selectSingleNode("//VA_OPER_ATIV").Text
            Else
                strValor = vbNullString
            End If
            
            Set objCarimbador = CreateObject("A7Server.clsCarimbador")
            objCarimbador.IncluirAvisoMonitor objMQAX200, _
                                              CLng(pxmlMensagem.selectSingleNode("//CO_EMPR_ORIG").Text), _
                                              strNumCtrlIF, _
                                              strCodigoMensagem, _
                                              FLUXO_MONITOR_NORMAL, _
                                              MONITOR_ENVIADA_NZ, _
                                              "A8", "NZ", _
                                              strNomeFila, _
                                              fgDataHoraServidor(enumFormatoDataHoraAux.DataAux), _
                                              strValor
            Set objCarimbador = Nothing
            Set xmlMensagemSPB = Nothing
        
            Err.Clear
            On Error GoTo 0
        
        End If
    
        If strNomeFila = vbNullString Then
            strAlerta = "Nome da Fila não cadastrada para o sistema: " & pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text
            Call fgGeraInformacaoAlerta(strAlerta, Format(Now, "dd/mm/yyyy HH:mm:ss"), "Gerenciador de envio", strAlerta)
            strNomeFila = ""
        End If
    
    End If
   
    Call flPostarMQSeries(pxmlMensagem, strNomeFila)
    
    Exit Function
ErrorHandler:
    
    Set objSistema = Nothing
    Set objCarimbador = Nothing
    Set xmlMensagemSPB = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPostarMensagemTraduzida Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



'Salvar a mensagem de saída (após a sua tradução).
Private Function flSalvarMensagemDadoSaida(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objMensagem                       As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    Call objMensagem.SalvarMensagemDadosSaida(pxmlMensagem)
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarMensagemDadoSaida Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir nova situação para a mensagem transitada.
Private Function flSalvarSituacaoMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                          ByVal plngCodigoOcorrencia As Long) As Boolean
                                                 
Dim objSituacaoMensagem                   As A7Server.clsSituacaoMensagem
    
On Error GoTo ErrorHandler
    
    Set objSituacaoMensagem = CreateObject("A7Server.clsSituacaoMensagem")
    Call objSituacaoMensagem.SalvarSituacaoMensagem(pxmlMensagem, plngCodigoOcorrencia)
    Set objSituacaoMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objSituacaoMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSalvarSituacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Enviar mensagem de rejeição para o Legado
'   - Montar protocolo de integração A7
'   - Montar Remessa com os erros identificados

Public Function EnviarMensagemRejeicaoLegado(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByVal plngCodigoOcorrencia As Long) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim intCount                                As Integer
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objErro                                 As A6A7A8CA.clsLogErro
Dim lngTipoMensagem                         As Long
Dim xmlMensagemSaida                        As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strTipoMensagemRetorno                  As String
Dim strSiglaSistemaOrigem                   As String

On Error GoTo ErrorHandler
    
    Set xmlMensagemSaida = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagemSaida.loadXML xmlMensagem.selectSingleNode("//TX_CNTD_SAID").Text
    xmlMensagemSaida.loadXML xmlMensagemSaida.selectSingleNode("//SaidaXML").xml
    
    
    strSiglaSistemaOrigem = UCase$(Trim$(xmlMensagem.documentElement.selectSingleNode("SG_SIST_ORIG").Text))
    
    If strSiglaSistemaOrigem = "WZ" Then
        
        If Not xmlMensagemSaida.selectSingleNode("//TP_SOLI") Is Nothing Then
            
            If Trim$(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) <> vbNullString And _
               IsNumeric(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) Then
                
                If Val(xmlMensagemSaida.selectSingleNode("//TP_SOLI").Text) = enumTipoSolicitacao.Inclusao Then
                    Exit Function
                End If
                
            Else
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If


    Select Case strSiglaSistemaOrigem
        Case "A8", "NZ", "BG", "DV"
            Exit Function
    End Select

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")

    udtProtocolo.CodigoEmpresa = Format(CLng(xmlMensagem.documentElement.selectSingleNode("//CO_EMPR_ORIG").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = UCase$(fgCompletaString(xmlMensagem.documentElement.selectSingleNode("//SG_SIST_ORIG").Text, " ", 3, False))
    udtProtocolo.SiglaSistemaOrigem = UCase$(fgCompletaString("A8", " ", 3, False))
    
    
    fgAppendNode xmlMensagemSaida, _
                 xmlMensagemSaida.documentElement.nodeName, _
                 "TP_MESG", _
                 Val(xmlMensagem.selectSingleNode("//TP_MESG").Text), _
                 ""
    
    Call objTipoOperacao.ObterTipoOperacaoPorMsgBUS(xmlMensagemSaida, lngTipoMensagem, "", strTipoMensagemRetorno)
    
    lngTipoMensagem = Val(strTipoMensagemRetorno)
    
    If Trim$(udtProtocolo.SiglaSistemaDestino) = "WZ" Then
        lngTipoMensagem = lngTipoMensagem + 2000
    Else
        lngTipoMensagem = lngTipoMensagem
    End If
    
    udtProtocolo.TipoMensagem = fgCompletaString(lngTipoMensagem, "0", 9, True)
    
    LSet udtProtocoloAux = udtProtocolo

    fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "TP_RETN", "2"
    fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ULTI_SITU_PROC", "99"

    Set objErro = CreateObject("A6A7A8CA.clsLogErro")

    Select Case plngCodigoOcorrencia
        
        Case enumOcorrencia.CanceladaErroTradução
            'Erro de Tradução no sistema A7.
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ERRO1", 1017
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "DE_ERRO1", xmlMensagem.selectSingleNode("//TX_DTLH_OCOR_ERRO").Text
        
        Case enumOcorrencia.RejeicaoNaoAutenticidade
        
            'Erro de não autenticidade no sistema A7.
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "CO_ERRO1", 1018
            fgAppendNode xmlMensagemSaida, xmlMensagemSaida.documentElement.nodeName, "DE_ERRO1", objErro.ObterDescErroNegocio(1018)
    
    End Select
    
    Set objErro = Nothing
    
    objMQAX200.MQOpen "A7Q.E.ENTRADA", enumMQOO_Open.MQOO_OUTPUT
    
    objMQAX200.MQPut udtProtocoloAux.String & xmlMensagemSaida.xml, _
                     Len(udtProtocoloAux.String & xmlMensagemSaida.xml), _
                     enumPutOptions.Binario
    
    objMQAX200.MQClose

    Exit Function
ErrorHandler:
    
    Set objTipoOperacao = Nothing
    Set objErro = Nothing
    
    Err.Clear

'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemRejeicao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function





'Obter os dados referentes ao sistema destino da mensagem recebida.
Private Function flObterSistema(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim objSistema                              As A7Server.clsParametrosSistema

On Error GoTo ErrorHandler

    Set objSistema = CreateObject("A7Server.clsParametrosSistema")
    
    flObterSistema = objSistema.Ler(pxmlMensagem.documentElement.selectSingleNode("SG_SIST_DEST").Text, _
                               CLng(pxmlMensagem.documentElement.selectSingleNode("CO_EMPR_ORIG").Text))

    Set objSistema = Nothing

    Exit Function
ErrorHandler:

    Set objSistema = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterSistema Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Enviar a mensagem traduzida para a fila do seu sistema destino.
Private Function flPostarMQSeries(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                  ByVal pstrNomeFila As String) As Boolean

Dim strMessageDescriptor                    As String
Dim strMensagemTraduzida                    As String
Dim strCorrelationID                        As String
Dim strMessageId                            As String
Dim strSistemaOrigem                        As String

Dim xmlMensagemPostada                      As DOMDocument40

Dim lngPrioridade                           As Long

On Error GoTo ErrorHandler

    Set xmlMensagemPostada = New DOMDocument40
    
    xmlMensagemPostada.loadXML pxmlMensagem.selectSingleNode("//TX_CNTD_SAID").Text
    
    'Pika
    'É aqui que vc vai decidir se é boleto ou Operação
    
    lngPrioridade = 9
    
    If Not xmlMensagemPostada.selectSingleNode("//TP_SOLI") Is Nothing Then
        
        If pxmlMensagem.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemEntrada.MensagemIDADV And _
           pxmlMensagem.selectSingleNode("//TP_MESG").Text <> enumTipoMensagemEntrada.MensagemIDABG Then
            
            Select Case CLng(xmlMensagemPostada.selectSingleNode("//TP_SOLI").Text)
                Case 1
                    pstrNomeFila = "A8Q.E.ENTRADA_BOLETO"
                Case 2
                    strSistemaOrigem = vbNullString
                    If Not xmlMensagemPostada.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                        strSistemaOrigem = Trim$(xmlMensagemPostada.selectSingleNode("//SG_SIST_ORIG").Text)
                    End If
                    'Se Sistema GPC ou R2 Enviar A8NET
'                    If strSistemaOrigem = "GPC" Or strSistemaOrigem = "R2" Then
'                        pstrNomeFila = "A8Q.E.ENTRADA_NET"
'                    Else
                        pstrNomeFila = "A8Q.E.ENTRADA_OPERACAO"
'                    End If
                Case Else
                    pstrNomeFila = "A8Q.E.ENTRADA"
            End Select
            
            
        End If
        
    Else
        If Not xmlMensagemPostada.selectSingleNode("//TP_MESG") Is Nothing Then
    
            Select Case CLng(xmlMensagemPostada.selectSingleNode("//TP_MESG").Text)
                Case enumTipoMensagemEntrada.MensagemNZA8, _
                     enumTipoMensagemEntrada.MensagemErroNZA8, _
                     enumTipoMensagemEntrada.MensagemPZErro, _
                     enumTipoMensagemEntrada.MensagemPZR1, _
                     enumTipoMensagemEntrada.MensagemPZR2, _
                     enumTipoMensagemEntrada.MensagemSTR0010R2PZA8
                     
'                    If UCase(Mid(xmlMensagemPostada.selectSingleNode("//CodMsg").Text, 1, 3)) = "CAM" Then
'                        If UCase(Mid(xmlMensagemPostada.selectSingleNode("//CodMsg").Text, 8, 1)) <> "E" Then
'                            pstrNomeFila = "A8Q.E.ENTRADA_NET"
'                        Else
'                            pstrNomeFila = "A8Q.E.ENTRADA_BACEN"
'                        End If
'                    Else
                        pstrNomeFila = "A8Q.E.ENTRADA_BACEN"
'                    End If
                
                Case enumTipoMensagemEntrada.MensagemRetornoDV, _
                     enumTipoMensagemEntrada.MensagemRetornoBG
                
                    pstrNomeFila = "A8Q.E.ENTRADA_CC"
                
            End Select
        
        End If
    
    End If
             
    'Incluir rotina ROBO quando as mensagens forem destinadas ao NZ.
     ' Nick 26/06/2016
     'If (objMQAX200.isProduction = False) Then
     '    If Left(UCase(Trim(pstrNomeFila)), 2) = "NZ" Then
     '
     '        If (flObterParametroRobo(pxmlMensagem)) Then
     '            pstrNomeFila = "A7Q.E.Robo_R1"
     '        End If
     '
     '    End If
     'End If
            
    Set xmlMensagemPostada = Nothing
    
    strMessageDescriptor = ""
    strCorrelationID = ""
   
    With pxmlMensagem.documentElement
        
        If Left(UCase(Trim(pstrNomeFila)), 3) <> "AEA" Then
            strMensagemTraduzida = .selectSingleNode("TX_CNTD_SAID").Text
        Else
            strMensagemTraduzida = flObterMensagemBG(.selectSingleNode("TX_CNTD_SAID").Text)
        End If
            
        strMensagemTraduzida = fgLimpaCaracterInvalido(strMensagemTraduzida)

        objMQAX200.MQOpen pstrNomeFila, MQOO_OUTPUT
        
        If Left(UCase(Trim(pstrNomeFila)), 2) = "NZ" Or _
           Left(UCase(Trim(pstrNomeFila)), 2) = "PZ" Or _
           Left(UCase(Trim(pstrNomeFila)), 2) = "DV" Or _
           Trim(UCase(pstrNomeFila)) = "A7Q.E.SAIDAMV" Or _
           Trim(UCase(pstrNomeFila)) = "A7Q.E.SAIDASIG" Then
        
            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.MainFrame, _
                             "A7Q.E.REPORT", _
                             "", _
                             strMessageId, , , _
                             9
        
        ElseIf Left(UCase(Trim(pstrNomeFila)), 3) = "AEA" Then

            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.MainFrame, _
                              , _
                             "", _
                             strMessageId
                             
        Else
        
            objMQAX200.MQPut strMensagemTraduzida, _
                             Len(strMensagemTraduzida), _
                             enumPutOptions.Binario, _
                             "A7Q.E.REPORT", _
                             strCorrelationID, _
                             strMessageId, , , _
                             lngPrioridade
        End If
        
        objMQAX200.MQClose
            
        .selectSingleNode("CO_MESG_MQSE").Text = strMessageId
        'Call flAtualizaMessageIdHex(CLng(.selectSingleNode("CO_MESG").Text), strMessageId)
        
        If Left(UCase(Trim(pstrNomeFila)), 2) = "DV" Or _
           Left(UCase(Trim(pstrNomeFila)), 3) = "AEA" Then
            
            If Not .selectSingleNode("TX_CNTD_ENTR") Is Nothing Then
                Call flAtualizaControleIntegracaoDV(Mid(.selectSingleNode("TX_CNTD_ENTR").Text, 21), strMessageId)
            End If
            
        End If
                
        'Call flSalvarSituacaoMensagem(pxmlMensagem, enumOcorrencia.PostagemBemSucedida)
    
    End With
    
    Exit Function

ErrorHandler:
    
    Set xmlMensagemPostada = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPostarMQSeries Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os dados referentes aos Parametros de comnicação com o Robo
Private Function flObterParametroRobo(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim objSistema                              As A7Server.clsParametrosSistema
Dim cdMesg                                  As String
Dim strXML                                  As String
Dim sistOrig                                As String
Dim xmlTratado                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set objSistema = CreateObject("A7Server.clsParametrosSistema")
    Set xmlTratado = New DOMDocument40
    
    cdMesg = pxmlMensagem.documentElement.selectSingleNode("TP_MESG").Text
    
    strXML = Mid(pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text, 21, Len(pxmlMensagem.documentElement.selectSingleNode("TX_CNTD_ENTR").Text))
    xmlTratado.loadXML (strXML)
    sistOrig = xmlTratado.documentElement.selectSingleNode("SG_SIST_ORIG").Text
        
    flObterParametroRobo = objSistema.ObterParametroRobo(sistOrig, cdMesg)

    Set objSistema = Nothing

    Exit Function
ErrorHandler:

    Set objSistema = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterSistema Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a mensagem de integração com o sistema BG através do correlationID.
Private Function flObterMensagemBG(ByVal pstrMensagemTraduzida As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pstrMensagemTraduzida
    
    flObterMensagemBG = xmlMensagem.selectSingleNode("//requestMsg").xml
    
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:

    Set xmlMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterMensagemBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Atualizar o MessageID da mensagem enviada.
Private Function flAtualizaMessageIdHex(ByVal plngCodigoMensagem As Long, _
                                        ByVal pstrMessageIdHex As String) As Boolean
                                                 
Dim objMensagem                       As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    Call objMensagem.AtualizaMessageID(plngCodigoMensagem, pstrMessageIdHex)
    Set objMensagem = Nothing

    Exit Function
ErrorHandler:

    Set objMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAtualizaMessageIdHex Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAtualizaControleIntegracaoDV(ByVal pxmlMensagemEntrada As String, _
                                                ByVal pstrMessageId As String) As Boolean
                                                 
Dim objLancamentoCCContabil                 As Object ' A8LQS.clsLancamentoCCContabil
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim lngNumSequLancamento                    As Long
Dim lngTipoLancamentoIntegracao             As Long
    
    On Error GoTo ErrorHandler
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pxmlMensagemEntrada
    
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    If Not xmlMensagem.documentElement.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
        lngNumSequLancamento = xmlMensagem.documentElement.selectSingleNode("//NR_SEQU_LANC").Text
    Else
        lngNumSequLancamento = 0
    End If
    
    lngTipoLancamentoIntegracao = 1
    If Not xmlMensagem.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
        lngTipoLancamentoIntegracao = Val(xmlMensagem.selectSingleNode("//TP_LANC_ITGR").Text)
    End If
    
    Call objLancamentoCCContabil.GravarControleIntegracao(xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                          lngTipoLancamentoIntegracao, _
                                                          Trim(pstrMessageId), lngNumSequLancamento)
    
    Set objLancamentoCCContabil = Nothing
    Set xmlMensagem = Nothing

    Exit Function
ErrorHandler:
    
    Set objLancamentoCCContabil = Nothing
    Set xmlMensagem = Nothing
    
    Err.Clear

End Function

'Private Function flRetiraRepeticao(ByVal pstrNomeTagMensagem As String, _
'                                   ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objDomMesgAux                           As MSXML2.DOMDocument40
'Dim objRepeticao                            As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim strCabecalhoSaida                       As String
'Dim strXmlSaida                             As String
'Dim strXmlAux                               As String
'Dim strRepeticao                            As String
'Dim strNomeTagPrincipal                     As String
'
'On Error GoTo ErrorHandler
'
'    flRetiraRepeticao = False
'
'    'Pega o xml da Mensagem
'    strXmlAux = pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text
'    strXmlSaida = strXmlAux
'
'    'Faz o Load do XML
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'        strCabecalhoSaida = Mid(strXmlAux, 1, 20)
'        strXmlSaida = Mid(strXmlAux, 21)
'        If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'            'XML da regra de tradução inválido.
'            lngCodigoErroNegocio = 1008
'            GoTo ErrorHandler
'        End If
'    End If
'
'    'Pega Nome da Tag Principal da Mensagem
'    strNomeTagPrincipal = objDomMesgSaida.documentElement.baseName
'
'    'Cria Mensagem Auxiliar para retirar a Repeticao
'    Set objDomMesgAux = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomMesgAux, "", strNomeTagPrincipal, "")
'
'    For Each objDomNode In objDomMesgSaida.documentElement.selectSingleNode("//" & strNomeTagPrincipal).childNodes
'
'        If UCase(Mid(objDomNode.nodeName, 1, 5)) <> "REPE_" Then
'            Call fgAppendNode(objDomMesgAux, strNomeTagPrincipal, objDomNode.nodeName, objDomNode.Text)
'        Else
'            Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'            strRepeticao = objDomNode.xml
'            objRepeticao.loadXML strRepeticao
'            Call flRetiraRepeticaoRecursivo(objRepeticao)
'            Call fgAppendXML(objDomMesgAux, strNomeTagPrincipal, Trim(objRepeticao.xml))
'        End If
'
'    Next objDomNode
'
'    strXmlSaida = objDomMesgAux.xml
'
'    strXmlSaida = Replace(strXmlSaida, "<REPET>", vbNullString)
'    strXmlSaida = Replace(strXmlSaida, "</REPET>", vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbFormFeed, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNullChar, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCr, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCrLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNewLine, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbVerticalTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbBack, vbNullString)
'
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text = ""
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).appendChild fgCreateCDATASection(strCabecalhoSaida & strXmlSaida)
'
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'
'    flRetiraRepeticao = True
'
'    Exit Function
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flRetiraRepeticao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Function
'
'Private Sub flRetiraRepeticaoRecursivo(ByRef pobjDomRepeticao As MSXML2.DOMDocument40)
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objRepeticao                            As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim objDomNodeGrupo                         As MSXML2.IXMLDOMNode
'Dim strXmlAux                               As String
'Dim strMesgSaida                            As String
'Dim strRepeticao                            As String
'
'On Error GoTo ErrorHandler
'
'    strMesgSaida = "<REPET>"
'
'    For Each objDomNode In pobjDomRepeticao.documentElement.childNodes
'
'        strXmlAux = objDomNode.xml
'
'        If InStr(1, UCase(strXmlAux), "REPE_") = 0 Then
'            strMesgSaida = strMesgSaida & objDomNode.xml
'        Else
'            strMesgSaida = strMesgSaida & "<" & objDomNode.nodeName & ">"
'            For Each objDomNodeGrupo In objDomNode.childNodes
'                If UCase(Mid(objDomNodeGrupo.nodeName, 1, 5)) <> "REPE_" Then
'                    strMesgSaida = strMesgSaida & objDomNodeGrupo.xml
'                Else
'                    Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'                    strRepeticao = objDomNodeGrupo.xml
'                    objRepeticao.loadXML strRepeticao
'                    Call flRetiraRepeticaoRecursivo(objRepeticao)
'                    strMesgSaida = strMesgSaida & objRepeticao.xml
'                End If
'             Next objDomNodeGrupo
'             strMesgSaida = strMesgSaida & "</" & objDomNode.nodeName & ">"
'        End If
'
'    Next objDomNode
'
'    strMesgSaida = strMesgSaida & "</REPET>"
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    objDomMesgSaida.loadXML strMesgSaida
'
'    Set pobjDomRepeticao = objDomMesgSaida
'
'    Set objDomMesgSaida = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    Set objDomNodeGrupo = Nothing
'
'    Exit Sub
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    Set objDomNodeGrupo = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flRetiraRepeticaoRecursivo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Sub
'
'Private Function flIncluiRepeticao(ByVal pstrNomeTagMensagem As String, _
'                                   ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objDomMesgAux                           As MSXML2.DOMDocument40
'Dim objRepeticao                            As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim strXmlAux                               As String
'Dim strCabecalhoSaida                       As String
'Dim strXmlSaida                             As String
'Dim strRepeticao                            As String
'Dim strNomeTagPrincipal                     As String
'
'On Error GoTo ErrorHandler
'
'    flIncluiRepeticao = False
'
'    'Pega o xml da Mensagem
'    strXmlAux = pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text
'    strXmlSaida = strXmlAux
'
'    'Faz o Load do XML
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'        strCabecalhoSaida = Mid(strXmlAux, 1, 20)
'        strXmlSaida = Mid(strXmlAux, 21)
'        If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'            'XML da regra de tradução inválido.
'            lngCodigoErroNegocio = 1008
'            GoTo ErrorHandler
'        End If
'    End If
'
'    'Pega Nome da Tag Principal da Mensagem
'    strNomeTagPrincipal = objDomMesgSaida.documentElement.baseName
'
'    'Cria Mensagem Auxiliar para incluir a Repeticao
'    Set objDomMesgAux = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomMesgAux, "", strNomeTagPrincipal, "")
'
'    For Each objDomNode In objDomMesgSaida.documentElement.selectSingleNode("//" & strNomeTagPrincipal).childNodes
'
'        If UCase(Mid(objDomNode.nodeName, 1, 3)) <> "GR_" Then
'            Call fgAppendNode(objDomMesgAux, strNomeTagPrincipal, objDomNode.nodeName, objDomNode.Text)
'        Else
'            If Not objDomMesgAux.selectSingleNode("//" & strNomeTagPrincipal & "/" & "REPE" & Mid(objDomNode.nodeName, 3)) Is Nothing Then
'                Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'                strRepeticao = objDomNode.xml
'                objRepeticao.loadXML strRepeticao
'                Call fgAppendXML(objDomMesgAux, strNomeTagPrincipal & "/" & "REPE" & Mid(objDomNode.nodeName, 3), Trim(objRepeticao.xml))
'            Else
'                Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'                strRepeticao = objDomNode.xml
'                objRepeticao.loadXML strRepeticao
'                Call flIncluiRepeticaoRecursivo(objRepeticao)
'                Call fgAppendXML(objDomMesgAux, strNomeTagPrincipal, Trim(objRepeticao.xml))
'            End If
'        End If
'
'    Next objDomNode
'
'    strXmlSaida = objDomMesgAux.xml
'
'    strXmlSaida = Replace(strXmlSaida, vbFormFeed, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNullChar, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCr, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCrLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNewLine, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbVerticalTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbBack, vbNullString)
'
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text = ""
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).appendChild fgCreateCDATASection(strCabecalhoSaida & strXmlSaida)
'
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'
'    flIncluiRepeticao = True
'
'    Exit Function
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluiRepeticao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Function
'
'Private Sub flIncluiRepeticaoRecursivo(ByRef pobjDomRepeticao As MSXML2.DOMDocument40)
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objRepeticao                            As MSXML2.DOMDocument40
'Dim objXmlAux                               As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim objDomNodeGrupo                         As MSXML2.IXMLDOMNode
'Dim strXmlAux                               As String
'Dim strMesgSaida                            As String
'Dim strRepeticao                            As String
'
'On Error GoTo ErrorHandler
'
'    strMesgSaida = "<" & "REPE" & Mid(pobjDomRepeticao.documentElement.baseName, 3) & ">"
'    strMesgSaida = strMesgSaida & "<" & pobjDomRepeticao.documentElement.baseName & ">"
'
'    For Each objDomNode In pobjDomRepeticao.documentElement.childNodes
'
'        If UCase(Mid(objDomNode.nodeName, 1, 3)) <> "GR_" Then
'            strMesgSaida = strMesgSaida & objDomNode.xml
'        Else
'            If InStr(1, strMesgSaida, "</" & "REPE" & Mid(objDomNode.nodeName, 3)) <> 0 Then
'                strRepeticao = objDomNode.xml
'                If InStr(2, strRepeticao, "<" & "GR_") <> 0 Then
'                    Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'                    objRepeticao.loadXML strRepeticao
'                    Call flIncluiRepeticaoRecursivo(objRepeticao)
'                    strRepeticao = objRepeticao.xml
'                    strMesgSaida = Mid(strMesgSaida, 1, InStr(1, strMesgSaida, "</" & "REPE" & Mid(objDomNode.nodeName, 3)) - 1) & Mid(strRepeticao, InStr(1, strRepeticao, objDomNode.nodeName) - 1)
'                Else
'                    strMesgSaida = Mid(strMesgSaida, 1, InStr(1, strMesgSaida, "</" & "REPE" & Mid(objDomNode.nodeName, 3)) - 1) & strRepeticao & Mid(strMesgSaida, InStr(1, strMesgSaida, "</" & "REPE" & Mid(objDomNode.nodeName, 3)))
'                End If
'            Else
'                Set objRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
'                strRepeticao = objDomNode.xml
'                objRepeticao.loadXML strRepeticao
'                Call flIncluiRepeticaoRecursivo(objRepeticao)
'                strMesgSaida = strMesgSaida & objRepeticao.xml
'            End If
'        End If
'
'    Next objDomNode
'
'    strMesgSaida = strMesgSaida & "</" & pobjDomRepeticao.documentElement.baseName & ">"
'    strMesgSaida = strMesgSaida & "</" & "REPE" & Mid(pobjDomRepeticao.documentElement.baseName, 3) & ">"
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    objDomMesgSaida.loadXML strMesgSaida
'
'    Set pobjDomRepeticao = objDomMesgSaida
'
'    Set objDomMesgSaida = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    Set objDomNodeGrupo = Nothing
'
'    Exit Sub
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objRepeticao = Nothing
'    Set objDomNode = Nothing
'    Set objDomNodeGrupo = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluiRepeticaoRecursivo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Sub
'
'Private Function flRetiraGrupo(ByVal pstrNomeTagMensagem As String, _
'                               ByVal pstrCodigoMensagem As String, _
'                               ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean
'
'Dim strXmlAux                               As String
'Dim strXmlSaida                             As String
'
'On Error GoTo ErrorHandler
'
'    flRetiraGrupo = False
'
'    'Pega o xml da Mensagem
'    strXmlAux = pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text
'    strXmlSaida = strXmlAux
'
'    Select Case pstrCodigoMensagem
'        Case "CAM0021", "CAM0022", "CAM0023", "CAM0024", "CAM0025", "CAM0026", _
'             "CAM0021R2", "CAM0022R2", "CAM0023R2", "CAM0024R2", "CAM0025R2", "CAM0026R2"
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_CodClausEspfcoIF>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_CodClausEspfcoIF>", vbNullString)
'
'        Case "CAM0028"
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_NumDespc>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_NumDespc>", vbNullString)
'
'        Case "CAM0030", "CAM0031", "CAM0032", "CAM0030R2", "CAM0031R2", "CAM0032R2"
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_NumDespc>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_NumDespc>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_CodClausEspfcoIF>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_CodClausEspfcoIF>", vbNullString)
'
'        Case "CAM0033"
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_RegOpCamlVincd>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_RegOpCamlVincd>", vbNullString)
'
'        Case "CAM0039", "CAM0039R2"
'            strXmlSaida = Replace(strXmlSaida, "<Grupo_" & pstrCodigoMensagem & "_RegOpCaml>", vbNullString)
'            strXmlSaida = Replace(strXmlSaida, "</Grupo_" & pstrCodigoMensagem & "_RegOpCaml>", vbNullString)
'
'    End Select
'
'    strXmlSaida = Replace(strXmlSaida, vbFormFeed, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNullChar, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCr, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCrLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNewLine, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbVerticalTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbBack, vbNullString)
'
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text = ""
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).appendChild fgCreateCDATASection(strXmlSaida)
'
'    flRetiraGrupo = True
'
'    Exit Function
'ErrorHandler:
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flRetiraGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Function
'
'Private Function flIncluiGrupo(ByVal pstrNomeTagMensagem As String, _
'                               ByVal pstrCodigoMensagem As String, _
'                               ByRef pobjDomMensagem As MSXML2.DOMDocument40) As Boolean
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objDomMesgAux                           As MSXML2.DOMDocument40
'Dim objGrupo                                As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim strXmlAux                               As String
'Dim strCabecalhoSaida                       As String
'Dim strXmlSaida                             As String
'Dim strGrupo                                As String
'Dim strNomeTagPrincipal                     As String
'
'On Error GoTo ErrorHandler
'
'    flIncluiGrupo = False
'
'    'Pega o xml da Mensagem
'    strXmlAux = pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text
'    strXmlSaida = strXmlAux
'
'    'Faz o Load do XML
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'        strCabecalhoSaida = Mid(strXmlAux, 1, 200)
'        strXmlSaida = Mid(strXmlAux, 201)
'        If Not objDomMesgSaida.loadXML(strXmlSaida) Then
'            'XML da regra de tradução inválido.
'            lngCodigoErroNegocio = 1008
'            GoTo ErrorHandler
'        End If
'    End If
'
'    'Pega Nome da Tag Principal da Mensagem
'    strNomeTagPrincipal = objDomMesgSaida.documentElement.baseName
'
'    'Cria Mensagem Auxiliar para incluir a Repeticao
'    Set objDomMesgAux = CreateObject("MSXML2.DOMDocument.4.0")
'    Call fgAppendNode(objDomMesgAux, "", strNomeTagPrincipal, "")
'    Call fgAppendNode(objDomMesgAux, strNomeTagPrincipal, pstrCodigoMensagem, "")
'
'    For Each objDomNode In objDomMesgSaida.documentElement.selectSingleNode("//" & strNomeTagPrincipal & "/" & pstrCodigoMensagem).childNodes
'
'        If UCase(Mid(objDomNode.nodeName, 1, 6)) <> "REPET_" Then
'            Call fgAppendNode(objDomMesgAux, strNomeTagPrincipal & "/" & pstrCodigoMensagem, objDomNode.nodeName, objDomNode.Text)
'        Else
'            If objDomMesgSaida.selectSingleNode("//" & strNomeTagPrincipal & "/" & pstrCodigoMensagem & "/" & objDomNode.nodeName & "/" & "Grupo" & Mid(objDomNode.nodeName, 6)) Is Nothing Then
'                Set objGrupo = CreateObject("MSXML2.DOMDocument.4.0")
'                strGrupo = objDomNode.xml
'                objGrupo.loadXML strGrupo
'                Call flIncluiGrupoRecursivo(objGrupo)
'                Call fgAppendXML(objDomMesgAux, strNomeTagPrincipal & "/" & pstrCodigoMensagem, Trim(objGrupo.xml))
'            Else
'                Call fgAppendXML(objDomMesgAux, strNomeTagPrincipal & "/" & pstrCodigoMensagem, objDomMesgSaida.selectSingleNode("//" & strNomeTagPrincipal & "/" & pstrCodigoMensagem & "/" & objDomNode.nodeName).xml)
'            End If
'        End If
'
'    Next objDomNode
'
'    strXmlSaida = objDomMesgAux.xml
'
'    strXmlSaida = Replace(strXmlSaida, vbFormFeed, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNullChar, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCr, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbCrLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbLf, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbNewLine, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbVerticalTab, vbNullString)
'    strXmlSaida = Replace(strXmlSaida, vbBack, vbNullString)
'
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).Text = ""
'    pobjDomMensagem.documentElement.selectSingleNode(pstrNomeTagMensagem).appendChild fgCreateCDATASection(strCabecalhoSaida & strXmlSaida)
'
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objGrupo = Nothing
'    Set objDomNode = Nothing
'
'    flIncluiGrupo = True
'
'    Exit Function
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objDomMesgAux = Nothing
'    Set objGrupo = Nothing
'    Set objDomNode = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluiGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Function
'
'Private Sub flIncluiGrupoRecursivo(ByRef pobjDomRepeticao As MSXML2.DOMDocument40)
'
'Dim objDomMesgSaida                         As MSXML2.DOMDocument40
'Dim objDomNode                              As MSXML2.IXMLDOMNode
'Dim strXmlAux                               As String
'Dim strMesgSaida                            As String
'Dim strRepeticao                            As String
'
'On Error GoTo ErrorHandler
'
'    strMesgSaida = "<" & pobjDomRepeticao.documentElement.baseName & ">"
'
'    For Each objDomNode In pobjDomRepeticao.documentElement.childNodes
'
'        strMesgSaida = strMesgSaida & "<Grupo" & Mid(pobjDomRepeticao.documentElement.baseName, 6) & ">"
'        strMesgSaida = strMesgSaida & objDomNode.xml
'        strMesgSaida = strMesgSaida & "</Grupo" & Mid(pobjDomRepeticao.documentElement.baseName, 6) & ">"
'
'    Next objDomNode
'
'    strMesgSaida = strMesgSaida & "</" & pobjDomRepeticao.documentElement.baseName & ">"
'
'    Set objDomMesgSaida = CreateObject("MSXML2.DOMDocument.4.0")
'    objDomMesgSaida.loadXML strMesgSaida
'
'    Set pobjDomRepeticao = objDomMesgSaida
'
'    Set objDomMesgSaida = Nothing
'    Set objDomNode = Nothing
'
'    Exit Sub
'ErrorHandler:
'    Set objDomMesgSaida = Nothing
'    Set objDomNode = Nothing
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluiGrupoRecursivo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'End Sub

