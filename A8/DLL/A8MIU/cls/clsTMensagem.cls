VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao tratamento de Mensagens.
' É responsável também, pela ligação da camada intermediária, à camada de negócios.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

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

' Este método aciona a gravação de mensagens, na camada de negócios.

Public Function EnviarMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                      Optional ByVal pstrChaveBMC0112 As String = "") As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
    
    If pstrChaveBMC0112 <> vbNullString Then
        Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "CO_CHAV_BMC", pstrChaveBMC0112)
    End If
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    EnviarMensagem = objMensagem.Salvar(xmlMensagem)
    Set objMensagem = Nothing
        
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a gravação de mensagens não tratadas pelo SLCC, na camada de negócios.

Public Function EnviarMensagemNaoTratada(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                         ByVal plngCodigoEmpresa As Long, _
                                         ByVal pintTipoBackOffice As enumTipoBackOffice, _
                                         ByVal plngCodigoLocalLiquidacao As Long, _
                                         ByVal pstrCodigoVeiculoLegal As String, _
                                         ByVal pstrSiglaSistema As String, _
                                         ByVal pstrAgendamento As String) As Boolean

Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim strComandoSelic                         As String
Dim lngISPBMesg                             As Long

On Error GoTo ErrorHandler
        
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    If Not xmlMensagem.selectSingleNode("//NumOpSEL") Is Nothing Then
        strComandoSelic = xmlMensagem.selectSingleNode("//NumOpSEL").Text
    End If
    
    Select Case Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3)
        Case "CTP"
            plngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
        Case "BMA"
            plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
        Case "SEL", "RDC"
            plngCodigoLocalLiquidacao = enumLocalLiquidacao.SELIC
        Case "LDL"
            If Not xmlMensagem.selectSingleNode("//ISPBLDL") Is Nothing Then
                lngISPBMesg = Val(xmlMensagem.selectSingleNode("//ISPBLDL").Text)
            ElseIf Not xmlMensagem.selectSingleNode("//ISPBIF") Is Nothing Then
                lngISPBMesg = Val(xmlMensagem.selectSingleNode("//ISPBIF").Text)
            End If
            
            Select Case lngISPBMesg
                Case enumISPB.IspbBMA
                    plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
                Case enumISPB.IspbCETIP
                    plngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
                Case enumISPB.IspbBMC
                    plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
            End Select
        Case "LTR"
            If Not xmlMensagem.selectSingleNode("//ISPBLTR") Is Nothing Then
                Select Case Val(xmlMensagem.selectSingleNode("//ISPBLTR").Text)
                    Case enumISPB.IspbBMA
                        plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
                    Case enumISPB.IspbCETIP
                        plngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
                    Case enumISPB.IspbBMC
                        plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
                End Select
            Else
                If Not xmlMensagem.selectSingleNode("//ISPBIF") Is Nothing Then
                    Select Case Val(xmlMensagem.selectSingleNode("//ISPBIF").Text)
                        Case enumISPB.IspbBMA
                            plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
                        Case enumISPB.IspbCETIP
                            plngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
                        Case enumISPB.IspbBMC
                            plngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
                    End Select
                End If
            End If
        Case "STR"
            plngCodigoLocalLiquidacao = enumLocalLiquidacao.SSTR
    End Select
    
    EnviarMensagemNaoTratada = objMensagem.Salvar(xmlMensagem, _
                                                  plngCodigoEmpresa, _
                                                  pintTipoBackOffice, _
                                                  plngCodigoLocalLiquidacao, _
                                                  pstrCodigoVeiculoLegal, _
                                                  pstrSiglaSistema, _
                                                  strComandoSelic, , , , _
                                                  pstrAgendamento)
    
    Set objMensagem = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMensagemNaoTratada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a alteração de agendamento de mensagens, na camada de negócios.

Public Function AlterarAgendamento(ByVal pstrNumeroControleIF As String, _
                                   ByVal pstrDTRegistroMensagemSPB As String, _
                                   ByVal pstrHorarioAgendamento As String, _
                                   ByVal pstrDHUltimaAtualizacao As String, _
                                   ByVal plngCodigoMensagemXML As Long, _
                                   ByVal pstrCodigoMensagem As String, _
                                   ByVal pintStatusMensagem As enumStatusMensagem) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
    
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Call objMensagemEnviadaRecebida.AlterarAgendamento(pstrNumeroControleIF, _
                                                       pstrDTRegistroMensagemSPB, _
                                                       pstrHorarioAgendamento, _
                                                       pstrDHUltimaAtualizacao, _
                                                       plngCodigoMensagemXML, _
                                                       pstrCodigoMensagem, _
                                                       pintStatusMensagem)
    Set objMensagemEnviadaRecebida = Nothing
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a confirmação de mensagens, na camada de negócios.

Public Function Confirmar(ByVal pstrNumeroControleIF As String, _
                          ByVal pstrDTRegistroMensagemSPB As String, _
                          ByVal pintStatusMensagem As enumStatusMensagem, _
                          ByVal pstrDHUltimaAtualizacao As String, _
                          ByVal pvntProtocolo As Variant) As String
    
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
    
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Call objMensagemEnviadaRecebida.Confirmar(pstrNumeroControleIF, _
                                              pstrDTRegistroMensagemSPB, _
                                              pintStatusMensagem, _
                                              pstrDHUltimaAtualizacao, _
                                              pvntProtocolo)
                                              
    Set objMensagemEnviadaRecebida = Nothing

    Confirmar = vbNullString        '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Confirmar = Err.Description
    
End Function

' Este método aciona a liberação de mensagens, na camada de negócios.

Public Function Liberar(ByVal pstrNumeroControleIF As String, _
                        ByVal pstrDTRegistroMensagemSPB As String, _
                        ByVal pvntNumeroSequencia As Variant, _
                        ByVal pintStatusMensagem As enumStatusMensagem, _
                        ByVal pstrDHUltimaAtualizacao As String, _
                        ByVal pblnIgnoraGradeHorario As Boolean) As String
    
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
    
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Call objMensagemEnviadaRecebida.Liberar(pstrNumeroControleIF, _
                                            pstrDTRegistroMensagemSPB, _
                                            pvntNumeroSequencia, _
                                            pintStatusMensagem, _
                                            pstrDHUltimaAtualizacao, _
                                            pblnIgnoraGradeHorario)
    Set objMensagemEnviadaRecebida = Nothing

    Liberar = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    Liberar = Err.Description
    
End Function

Public Function LiberarMensagemControleAlcada(ByVal pxmlMensagem As MSXML2.IXMLDOMNode, _
                                              ByVal pblnIgnoraGradeHorario As Boolean) As String
    
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objVeiculoLegal                         As A8MIU.clsMensagem
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlRetorno                              As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strCNPJ                                 As String
Dim strNomeCNPJ                             As String
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objVeiculoLegal = CreateObject("A8MIU.clsMensagem")
    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(pxmlMensagem.selectSingleNode("NumeroControleIF").Text)
        
        If Not xmlMensagem.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val("0" & xmlMensagem.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
                       
                If Not xmlMensagem.selectSingleNode("//CO_ORIG_RECU") Is Nothing Then
                    Call xmlVeiculoLegal.loadXML(objVeiculoLegal.LerTodosDominioTabela(xmlMensagem.selectSingleNode("//CO_EMPR").Text, _
                                         "A8.TB_VEIC_LEGA", xmlMensagem.selectSingleNode("//CO_VEIC_LEGA").Text, "", "", "", ""))
                    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
                    'If Not xmlOperacao.Load(pxmlMensagem) Then
                        '100 - Documento XML Inválido.
                    '    lngCodigoErroNegocio = 100
                    '    GoTo ErrorHandler
                    'End If

                    'Call xmlOperacao.loadXML(pxmlMensagem)
                    Call fgAppendNode(xmlOperacao, "", "Grupo_Lote", "")
                    If Val("0" & xmlMensagem.selectSingleNode("//CO_ORIG_RECU").Text) = enumOrigemRecurso.Isenta Then
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "TP_CONT_DEBT", xmlMensagem.selectSingleNode("//TP_CONT_DEBT_ISEN").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "NU_CONT_DEBT", xmlMensagem.selectSingleNode("//NU_CONT_DEBT_ISEN").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "TP_PESS_DEBT", xmlMensagem.selectSingleNode("//TP_PESS_DEBT_ISEN").Text)
                        'Trazer da tabela de Veiculo Legal LerTodosDominioTabela
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "CO_CNPJ_CPF_DEBT_1", xmlVeiculoLegal.selectSingleNode("//CO_CNPJ_VEIC_LEGA").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "NO_TITU_DEBT_1", xmlVeiculoLegal.selectSingleNode("//DESCRICAO").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "CO_AGEN_DEBT", xmlMensagem.selectSingleNode("//CO_AGEN_DEBT_ISEN").Text)
                    Else
                        
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "TP_CONT_DEBT", xmlMensagem.selectSingleNode("//TP_CONT_DEBT_TRIB").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "NU_CONT_DEBT", xmlMensagem.selectSingleNode("//NU_CONT_DEBT_TRIB").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "TP_PESS_DEBT", xmlMensagem.selectSingleNode("//TP_PESS_DEBT_TRIB").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "CO_CNPJ_CPF_DEBT_1", xmlVeiculoLegal.documentElement.selectSingleNode("//CO_CNPJ_VEIC_LEGA").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "NO_TITU_DEBT_1", xmlVeiculoLegal.selectSingleNode("//DESCRICAO").Text)
                        Call fgAppendNode(xmlOperacao, "Grupo_Lote", "CO_AGEN_DEBT", xmlMensagem.selectSingleNode("//CO_AGEN_DEBT_TRIB").Text)
                    End If
                    Set xmlMensagem = Nothing
                    Set xmlVeiculoLegal = Nothing
            End If
                 
        End If
    End If

      
        Call objMensagemEnviadaRecebida.LiberarMensagemControleAlcada(pxmlMensagem, pblnIgnoraGradeHorario, xmlOperacao)
        Set objMensagemEnviadaRecebida = Nothing
        Set objVeiculoLegal = Nothing
        Set xmlOperacao = Nothing
        LiberarMensagemControleAlcada = vbNullString          '<-- Processamento efetuado sem erros
    
        If Not ObjectContext Is Nothing Then
            ObjectContext.SetComplete
        End If

    
    Exit Function
    
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiberarMensagemControleAlcada = Err.Description
    
End Function

' Este metodo aciona a liberacao de mensagens de contingencia, na camada de negocios.

Public Function LiberarContingencia(ByVal pstrNumeroControleIF As String, _
                                    ByVal pstrDTRegistroMensagemSPB As String, _
                                    ByVal pvntNumeroSequencia As Variant, _
                                    ByVal pintStatusMensagem As enumStatusMensagem, _
                                    ByVal pstrDHUltimaAtualizacao As String, _
                                    ByVal pblnIgnoraGradeHorario As Boolean) As String
    
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
    
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Call objMensagemEnviadaRecebida.LiberarContingencia(pstrNumeroControleIF, _
                                                        pstrDTRegistroMensagemSPB, _
                                                        pvntNumeroSequencia, _
                                                        pintStatusMensagem, _
                                                        pstrDHUltimaAtualizacao, _
                                                        pblnIgnoraGradeHorario)
    Set objMensagemEnviadaRecebida = Nothing

    LiberarContingencia = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiberarContingencia = Err.Description
    
End Function

' Este método aciona o cancelamento de uma mensagem enviada manualmente, na camada de negócios.

Public Function CancelarEntradaManual(ByVal pstrNumeroControleIF As String, _
                                      ByVal pstrDTRegistroMensagemSPB As String, _
                                      ByVal pintStatus As Integer, _
                                      ByVal pstrDHUltimaAtualizacao As String) As String
                                      
    
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
    
    On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Call objMensagemEnviadaRecebida.CancelarEntradaManual(pstrNumeroControleIF, pstrDTRegistroMensagemSPB, pintStatus, pstrDHUltimaAtualizacao)
    Set objMensagemEnviadaRecebida = Nothing

    CancelarEntradaManual = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    CancelarEntradaManual = Err.Description
    
End Function

' Este método aciona o reenvio de mensagens, na camada de negócios.

Public Function Reenviar(ByVal pstrNumeroControleIF As String, _
                         ByVal pstrDHUltimaAtualizacao As String) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem

    On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Reenviar = objMensagem.ReenviarMensagem(pstrNumeroControleIF, pstrDHUltimaAtualizacao)
    Set objMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método aciona o reprocessamento de mensagens inconsistentes, na camada de negócios.

Public Function ReprocessarMensagemInconsistente(ByVal pstrNumeroControleIF As String, _
                                                 ByVal pstrDTRegistroMensagemSPB As String, _
                                                 ByVal pvntNumeroSequencia As Variant, _
                                                 ByVal pstrDHUltimaAtualizacao As String) As String
    
Dim objProcessoMensagemBMA              As A8LQS.clsProcessoMensagemBMA
    
On Error GoTo ErrorHandler

    Set objProcessoMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    
    Call objProcessoMensagemBMA.ReprocessarMensagemInconcistente(pstrNumeroControleIF, _
                                                                 pstrDTRegistroMensagemSPB, _
                                                                 pvntNumeroSequencia, _
                                                                 pstrDHUltimaAtualizacao)
                                            
    Set objProcessoMensagemBMA = Nothing

    ReprocessarMensagemInconsistente = vbNullString          '<-- Processamento efetuado sem erros
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    ReprocessarMensagemInconsistente = Err.Description
    
    Set objProcessoMensagemBMA = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
        
End Function

