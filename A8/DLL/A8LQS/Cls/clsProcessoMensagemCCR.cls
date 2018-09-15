VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessoMensagemCCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Public Function GerenciarRecebimentoMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    If strCodigoMensagem = "CCR0001R1" Or _
       strCodigoMensagem = "CCR0005R1" Then
        Call flProcessaR1Consulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "CCR0002R1" Or _
       strCodigoMensagem = "CCR0003R1" Or _
       strCodigoMensagem = "CCR0004R1" Then
        Call flProcessaMensagemR1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "CCR0006" Then
        Call flProcessaMensagemCCR0006(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "CCR0007" Then
        Call flProcessaMensagemCCR0007(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "CCR0002E" Or _
      (strCodigoMensagem = "CCR0002" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaErroConsulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "CCR0003E" Or _
      (strCodigoMensagem = "CCR0003" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaErroConsulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "CCR0004E" Or _
      (strCodigoMensagem = "CCR0004" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaErroConsulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "CCR0001E" Or _
      (strCodigoMensagem = "CCR0001" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaErroConsulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "CCR0005E" Or _
      (strCodigoMensagem = "CCR0005" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaErroConsulta(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
        
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flProcessaErroConsulta(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                        ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa


Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgE                                As MSXML2.DOMDocument40

Dim strCodMsg                               As String
Dim strMensagem                             As String

On Error GoTo ErrorHandler
    
    
    Set xmlMesgE = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    '-----------------------------  MENSAGEM  -----------------------------------------------
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    Set xmlMesgE = objMensagemEnvRec.LerXMLMensagem(udtProtocoloErroNZ.ControleRemessaNZ)
    
    If xmlMesgE.xml = vbNullString Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    Call objMensagemEnvRec.AlterarStatus(xmlMesgE.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgE.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.ErroNegocioSPB, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgE.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgE.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgE.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgE.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgE.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgE.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.CCR, _
                                      xmlMesgE.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgE.selectSingleNode("//SG_SIST").Text)
        
    Set objMensagemEnvRec = Nothing
    
    
    '-----------------------------  OPERAÇÃO -----------------------------------------------
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    If xmlMesgE.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgE.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
        
    Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       enumStatusOperacao.Rejeitada, _
                                       vbNullString, , , _
                                       False)
    
    
    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Rejeitada
    
    Set objLegado = CreateObject("A8LQS.clsLegado")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    strMensagem = objLegado.EnviarRetorno(xmlOperacao)
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
    
    
    If Mid$(strCodMsg, 1, 7) = "CCR0003" Then
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
        Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
        Set objGestaoCaixa = Nothing
    End If
    
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgE = Nothing
    
    
    Exit Function
ErrorHandler:
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgE = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaErroConsulta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Private Function flProcessaR1Consulta(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                      ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim strMensagem                             As String
Dim strDataOperacao                         As String
Dim strDataRegistroMensagem                 As String

On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlRemessa.xml

    If Not xmlRemessa.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlRemessa.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
    
    '---------- MENSAGEM ------------------------------------------------------
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(udtHeaderMensagem.ControleRemessaNZ))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                          Trim(udtHeaderMensagem.ControleRemessaNZ), _
                                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                          0, _
                                                          Trim(xmlRemessa.selectSingleNode("//CodMsg").Text), _
                                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//TP_BKOF").Text, _
                                                          udtHeaderMensagem.CodigoEmpresa, _
                                                          vbNullString, _
                                                          vbNullString, _
                                                          enumStatusMensagem.R1, _
                                                          enumIndicadorSimNao.Nao, _
                                                          "", _
                                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//SG_SIST").Text, _
                                                          1)

        
    objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                     enumStatusMensagem.Respondida, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                     , _
                                     False, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
    
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    Call objConciliacao.Incluir(Trim(udtHeaderMensagem.ControleRemessaNZ), _
                                strDataRegistroMensagem, _
                                Replace(xmlRemessa.selectSingleNode("//DtMovto").Text, "-", ""), _
                                0, 0, _
                                "", _
                                0, 0, _
                                "0", _
                                1, , _
                                enumTipoDebitoCredito.Debito, _
                                "", _
                                0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                xmlRemessa.selectSingleNode("//CNPJBaseEntRespons").Text)
    Set objConciliacao = Nothing
    '---------- operação ------------------------------------------------------
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    
    Call objOperacao.AlterarStatusAcao(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       enumStatusOperacao.Registrada, _
                                       vbNullString, , , False)
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)) Then
        fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
    End If
    
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    
    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
    
    Set objLegado = CreateObject("A8LQS.clsLegado")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlRemessa)
    Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
    
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMesgTratadaR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)



End Function

Private Function flProcessaMensagemR1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                      ByVal pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objLegado                               As A8LQS.clsLegado

Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40

Dim vntSequenciaOperacao                    As Variant
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim lngSeqControleRepeticao                 As Long
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strNumeroControleCamara                 As String
Dim strCodigoMensagemSPB                    As String
Dim strSituacaoMensagemSPB                  As String
Dim strAgendamento                          As String
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strComandoOperacao                      As String
Dim strMensagemLegado                       As String

Dim strSituacaoRecebida                     As String
Dim strHistSituacaoOperacao                 As String

Dim xmlParametrosGerais                     As MSXML2.DOMDocument40

Dim lngStatusAtualMensagemOriginal          As Long
Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem

    On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    
    lngStatusMensagem = enumStatusMensagem.R1
    lngStatusMensagemOriginal = enumStatusMensagem.Respondida
    
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strAgendamento = ""
    lngSeqControleRepeticao = 1
    
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
        
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    vntSequenciaOperacao = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)

    vntSequenciaOperacao = xmlOperacao.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    strCodigoVeiculoLegal = xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("//SG_SIST_ORIG").Text
    lngCodigoLocalLiquidacao = xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text
    lngTipoBackOffice = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//TP_BKOF").Text
    lngStatusAtualMensagemOriginal = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
        fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
    End If
    
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
        
    'Obter Status ---------------------
    Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
    xmlStatus.loadXML flObterStatusOperacaoMensagem(xmlMensagem)
    '----------------------------------
    
    lngStatusMensagem = enumStatusMensagem.R1
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlCCR") Is Nothing Then
        strNumeroControleCamara = Trim(xmlMensagem.selectSingleNode("//NumCtrlCCR").Text)
    End If
        

    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            strSituacaoMensagemSPB, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara)
    
    strNumeroControleCamara = ""
    If Not xmlMensagem.selectSingleNode("//CodReemb") Is Nothing Then
        strNumeroControleCamara = Trim(xmlMensagem.selectSingleNode("//CodReemb").Text)
    End If

    If strNumeroControleCamara <> vbNullString Then
        Call objOperacao.GravarControleSPBOriginal(vntSequenciaOperacao, strNumeroControleCamara)
    End If
        
    'Atualiza Status da Mensagem Original -------------------------------------------------------------------
    If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
        lngStatusMensagem = xmlStatus.selectSingleNode("//Status_Mensagem").Text
    Else
        lngStatusMensagem = enumStatusMensagem.Respondida
    End If
    
    strSituacaoRecebida = vbNullString
    If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
        strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
    End If
    
    If lngStatusMensagem = 0 Then lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    
    'RATS 1029 - O sistema não deve sobrepor o status da operação quando receber
    'um status de Pendente de Registro ou Pendente de Aceite.
    'Esta ação visa minimizar o impacto do recebimento das mensagens com status finais antes
    'das mensgaens de pendencia de registro e aceite.
    'RATS 1038 - Bruno Oliveira - 11/jan/2011 - esse RATS foi aberto para corrigir uma falha na implementação do RATS 1029
    If Not ((lngStatusMensagem = enumStatusMensagem.PendenteDeAceite Or lngStatusMensagem = enumStatusMensagem.PendenteDeRegistro) And (lngStatusAtualMensagemOriginal = enumStatusMensagem.MensagemRejeitada Or lngStatusAtualMensagemOriginal = enumStatusMensagem.Excluida Or lngStatusAtualMensagemOriginal = enumStatusMensagem.Recolhida Or lngStatusAtualMensagemOriginal = enumStatusMensagem.Reembolsada Or lngStatusAtualMensagemOriginal = enumStatusMensagem.Registrada)) Then
    
        objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         lngStatusMensagem, _
                                         xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                         , _
                                         False, _
                                         xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
    Else
        'apesar da Mensagem não ter seu status alterado, o Histórico da Situacao da Mensagem deve ser armazenado
        Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        objHistoricoMensagem.Incluir xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                     lngStatusMensagem, _
                                     0, _
                                     , _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
                                     
        Set objHistoricoMensagem = Nothing
    End If
    '-------------------------------------------------------------------
            
    Call flGerenciarChamada(xmlOperacao, xmlMensagem, _
                            CLng(IIf(xmlStatus.selectSingleNode("//Status_Operacao").Text = "", 0, xmlStatus.selectSingleNode("//Status_Operacao").Text)), _
                            strNumeroControleIF)
    
    Set objMensagem = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objLegado = Nothing
        Set objMensagem = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set objHistoricoMensagem = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



Private Function flObterStatusOperacaoMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim objProcessoMensagem                     As A8LQS.clsProcessoMensagem

Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlTagDominioStatus                     As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlTagSituacao                          As MSXML2.DOMDocument40
Dim xmlNodeTagSitu                          As MSXML2.IXMLDOMNode

Dim strSituacaoMensagemSPB                  As String
Dim strGrupoMensagem                        As String
Dim strTagDominioStatus                     As String
Dim strNomeTag                              As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler
     
    Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
     
    Call fgAppendNode(xmlStatus, "", "Grupo_Status", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Status_Operacao", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Status_Mensagem", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Situacao_Mensagem", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Situacao_Recebida", "")
     
        
    strGrupoMensagem = Mid(Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text), 1, 3)
    
    Set xmlTagSituacao = CreateObject("MSXML2.DOMDocument.4.0")
        
    Set objProcessoMensagem = CreateObject("A8LQS.clsProcessoMensagem")
    strTagDominioStatus = objProcessoMensagem.ObterNomeTagSituacao(strGrupoMensagem)
    Set objProcessoMensagem = Nothing

    If strTagDominioStatus = vbNullString Then
        '3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
        lngCodigoErroNegocio = 3017
        GoTo ErrorHandler
    End If
    
    xmlTagSituacao.loadXML strTagDominioStatus
        
    For Each xmlNodeTagSitu In xmlTagSituacao.documentElement.childNodes
        
        strNomeTag = xmlNodeTagSitu.selectSingleNode("NO_TAG").Text
        
        If Not pxmlMensagem.documentElement.selectSingleNode("//" & strNomeTag) Is Nothing Then
        
            Set objProcessoMensagem = CreateObject("A8LQS.clsProcessoMensagem")
            strTagDominioStatus = objProcessoMensagem.LerTodosTagDominioStatus(strGrupoMensagem, strNomeTag)
            Set objProcessoMensagem = Nothing
            
            If strTagDominioStatus = vbNullString Then
                '3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
                lngCodigoErroNegocio = 3017
                GoTo ErrorHandler
            End If
            
            Set xmlTagDominioStatus = CreateObject("MSXML2.DOMDocument.4.0")
            xmlTagDominioStatus.loadXML strTagDominioStatus
        
            If pxmlMensagem.documentElement.selectSingleNode("//" & strNomeTag) Is Nothing Then
                strSituacaoMensagemSPB = ""
            Else
                strSituacaoMensagemSPB = pxmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text
            End If
                
            For Each xmlNode In xmlTagDominioStatus.documentElement.selectNodes("//Grupo_TagDominioStatus")
                If pxmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text = xmlNode.selectSingleNode("DE_DOMI").Text Then
                    lngStatusOperacao = xmlNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                    lngStatusMensagem = xmlNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text
                    Exit For
                End If
            Next
                            
            xmlStatus.selectSingleNode("//Status_Operacao").Text = IIf(lngStatusOperacao = 0, "", lngStatusOperacao)
            xmlStatus.selectSingleNode("//Status_Mensagem").Text = IIf(lngStatusMensagem = 0, "", lngStatusMensagem)
            xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = strSituacaoMensagemSPB
            
            If Not xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG") Is Nothing Then
                strSituacaoRecebida = Format$(xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG").Text, "00000") & "|" & _
                                      strSituacaoMensagemSPB
                xmlStatus.selectSingleNode("//Situacao_Recebida").Text = strSituacaoRecebida
            End If
            
            Set xmlTagDominioStatus = Nothing
            
        End If
    Next
    
    flObterStatusOperacaoMensagem = xmlStatus.xml
    
    Set xmlStatus = Nothing
   
    Exit Function
ErrorHandler:
    
    Set xmlStatus = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusOperacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'-----------------------------------------------------------------------------
'Gerenciamento de Envio :
'   - Remessas A6 e PJ
'   - Mensagens de resposta ao legado
'   - Disponibilizar lançamento conta corrente
'   - Disponibilizar Alerta ao Usuário
'-----------------------------------------------------------------------------

Private Function flGerenciarChamada(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByVal plngStatusOperacao As Long, _
                                    ByVal pstrNumeroControleIF As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objLegado                               As A8LQS.clsLegado
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa

Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim lngTipoOperacao                         As Long
Dim strCodigoMensagem                       As String
Dim lngStatusAtualOperacao                  As Long
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao
    
On Error GoTo ErrorHandler

    
    If plngStatusOperacao = 0 Then plngStatusOperacao = enumStatusOperacao.Inconsistencia
    
    lngStatusAtualOperacao = xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text
    
    'RATS 1029 - O sistema não deve sobrepor o status da operação quando receber
    'um status de Pendente de Registro ou Pendente de Aceite.
    'Esta ação visa minimizar o impacto do recebimento das mensagens com status finais antes
    'das mensgaens de pendencia de registro e aceite.
    'RATS 1038 - Bruno Oliveira - 26/jan/2011 - esse RATS foi aberto para corrigir uma falha na implementação do RATS 1029
    If Not ((plngStatusOperacao = enumStatusOperacao.PendenteDeAceite Or plngStatusOperacao = enumStatusOperacao.PendenteDeRegistro) And (lngStatusAtualOperacao = enumStatusOperacao.Rejeitada Or lngStatusAtualOperacao = enumStatusOperacao.Excluida Or lngStatusAtualOperacao = enumStatusOperacao.Recolhida Or lngStatusAtualOperacao = enumStatusOperacao.Reembolsada Or lngStatusAtualOperacao = enumStatusOperacao.Registrada)) Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      plngStatusOperacao, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , _
                                      , _
                                      False
        Set objOperacao = Nothing
    Else
        'apesar da Operação não ter seu status alterado, o Histórico da Situacao da Operação deve ser armazenado
         Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
         objHistoricoOperacao.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      plngStatusOperacao, _
                                      , _
                                      0
         Set objHistoricoOperacao = Nothing
    End If
    
    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao

    If xmlOperacao Is Nothing Then
        Set objOperacao = Nothing
        Exit Function
    End If
    
    If plngStatusOperacao = 0 Then
        Exit Function
    End If

    'Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXML = objProcessoOperacao.LerTodos("RecebimentoR1", _
                                          CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text), _
                                          0, plngStatusOperacao)

    'If xmlControleProcessoOperacao.documentElement.selectSingleNode("//Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    Set objLegado = Nothing
    'End If

    If strXML = vbNullString Then
        Exit Function
    End If
    
    xmlControleProcessoOperacao.loadXML strXML
    

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("//Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
        Set objProcessoOperacao = Nothing
        Set objGestaoCaixa = Nothing
    End If


    
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    Set objHistoricoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flProcessaMensagemCCR0006(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim lngSeqControleRepeticao                 As Long

Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String

Dim vntCNPJNLiqdant                         As Variant
Dim intIndDebitoCredito                     As Integer
Dim strAux                                  As String

On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   "")
    
    lngSeqControleRepeticao = 1
      
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            enumStatusMensagem.AConciliar, _
                                                            enumIndicadorSimNao.Nao, _
                                                            vbNullString, _
                                                            enumLocalLiquidacao.CCR, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            xmlMensagem.selectSingleNode("//CodReemb").Text)
    Set objMesgRecbEnviSPB = Nothing

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    strAux = xmlMensagem.selectSingleNode("//TpComerc").Text & _
             xmlMensagem.selectSingleNode("//TpInstntoCCR").Text
             
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtOp").Text, _
                                0, 0, _
                                strAux, _
                                0, 0, _
                                xmlMensagem.selectSingleNode("//VlrOp").Text, _
                                lngSeqControleRepeticao, , _
                                enumTipoDebitoCredito.Debito, _
                                xmlMensagem.selectSingleNode("//DtVenc_Exprc").Text, _
                                0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                xmlMensagem.selectSingleNode("//CNPJBaseEntRespons").Text)
                                
    Set objConciliacao = Nothing
    
    Call flConciliarMensagem(Trim(xmlMensagem.selectSingleNode("//CodMsg").Text), _
                             strNumeroControleIF, _
                             strDataRegistroMensagem)
    
    
    
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCCR0006 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Sub flConciliarMensagem(ByVal pstrtCodigoMensagem As String, _
                                ByVal pstrNumeroControleIF As String, _
                                ByVal pstrDHRegistroMensagem As String)


Dim rsMensagem                              As ADODB.Recordset
Dim rsOperacao                              As ADODB.Recordset
Dim strSQL                                  As String
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao


On Error GoTo ErrorHandler
    
    Set rsMensagem = flLerDadosMensagem(pstrNumeroControleIF, pstrDHRegistroMensagem)
    
    If rsMensagem.EOF Then
        Set rsMensagem = Nothing
        Exit Sub
    End If
                
    strSQL = " SELECT                               " & vbCrLf & _
             "          OPER.NU_SEQU_OPER_ATIV,     " & vbCrLf & _
             "          OPER.CO_OPER_ATIV,          " & vbCrLf & _
             "          OPER.DT_OPER_ATIV,          " & vbCrLf & _
             "          OPER.NU_ATIV_MERC,          " & vbCrLf & _
             "          OPER.DT_VENC_ATIV,          " & vbCrLf & _
             "          OPER.VA_OPER_ATIV,          " & vbCrLf & _
             "          OPER.NU_CTRL_MESG_SPB_ORIG, " & vbCrLf & _
             "          OPER.CO_ISPB_IF_CNPT,       " & vbCrLf & _
             "          OPER.IN_ENTR_MANU           " & vbCrLf & _
             " FROM     A8.TB_OPER_ATIV OPER        " & vbCrLf & _
             " WHERE    OPER.TP_OPER                = " & enumTipoOperacaoLQS.NegociacaoOperacaoCCR & vbCrLf & _
             " AND      OPER.CO_ULTI_SITU_PROC IN (" & enumStatusOperacao.EmSer & "," & _
                                                       enumStatusOperacao.ManualEmSer & ")" & vbCrLf & _
             " AND      OPER.NU_CTRL_MESG_SPB_ORIG  ='" & Trim(rsMensagem!NU_CTRL_CAMR) & "'" & vbCrLf & _
             " AND      OPER.VA_OPER_ATIV          >= " & fgVlrToDBServer(rsMensagem!VA_FINC) & vbCrLf & _
             " AND      OPER.DT_OPER_ATIV           = " & fgDtXML_To_Oracle(rsMensagem!DT_OPER)
    
    Set rsOperacao = fgQuerySQL(strSQL)
    
    If rsOperacao.EOF Then
        Set rsMensagem = Nothing
        Set rsOperacao = Nothing
        Exit Sub
    End If
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    Call objMesgRecbEnviSPB.AlterarStatus(pstrNumeroControleIF, _
                                          pstrDHRegistroMensagem, _
                                          enumStatusMensagem.Conciliada, _
                                          "", _
                                          "", _
                                          False, _
                                          1)
       
    Call objMesgRecbEnviSPB.AlterarSequenciaOperacao(pstrNumeroControleIF, _
                                                     fgDtHrStr_To_DateTime(pstrDHRegistroMensagem), _
                                                     rsOperacao!NU_SEQU_OPER_ATIV, _
                                                     1)
    
    'If rsOperacao!IN_ENTR_MANU <> enumIndicadorSimNao.Sim Then
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        
        Call objProcessoOperacao.GerenciarChamadaViaTela(rsOperacao!NU_SEQU_OPER_ATIV, _
                                                         enumFuncionalidade.Confirmacao, _
                                                         True)
                                
        Set objProcessoOperacao = Nothing
    'End If
    Set objMesgRecbEnviSPB = Nothing
    Set rsMensagem = Nothing
    Set rsOperacao = Nothing
    
    
    Exit Sub
    
ErrorHandler:
    
    Set objProcessoOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set rsMensagem = Nothing
    Set rsOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Function flLerDadosMensagem(ByVal pstrNumeroControleIF As String, _
                                    ByVal pstrDataRegistroMensagemSPB As String) As ADODB.Recordset

Dim strSQL                                  As String
Dim strXML                                  As String
Dim rsMensagem                              As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT                                   " & vbCrLf & _
             "          A.NU_CTRL_IF,                   " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,             " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,            " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,            " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,            " & vbCrLf & _
             "          A.NU_CTRL_CAMR,                 " & vbCrLf & _
             "          A.CO_MESG_SPB,                  " & vbCrLf & _
             "          TO_CHAR(B.DT_OPER,'YYYYMMDD') AS DT_OPER,   " & vbCrLf & _
             "          B.NU_ATIV_MERC,                 " & vbCrLf & _
             "          B.VA_FINC,                      " & vbCrLf & _
             "          B.DT_VENC,                      " & vbCrLf & _
             "          B.CO_ISPB_IF_CNPT               " & vbCrLf & _
             " FROM    A8.TB_MESG_RECB_ENVI_SPB A,      " & vbCrLf & _
             "         A8.TB_MESG_RECB_SPB_CNCL B       " & vbCrLf & _
             " WHERE   A.NU_CTRL_IF         = B.NU_CTRL_IF        (+) " & vbCrLf & _
             " AND     A.NU_SEQU_CNTR_REPE  = B.NU_SEQU_CNTR_REPE (+) " & vbCrLf & _
             " AND     A.DH_REGT_MESG_SPB   = B.DH_REGT_MESG_SPB  (+) " & vbCrLf

    strSQL = strSQL & _
             "   AND   A.NU_CTRL_IF              = '" & pstrNumeroControleIF & "'" & vbCrLf & _
             "   AND   A.DH_REGT_MESG_SPB        = " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB) & vbCrLf & _
             "   AND   A.NU_SEQU_CNTR_REPE       = 1 " & vbCrLf & _
             "   AND   A.CO_MESG_SPB             = 'CCR0006'               " & vbCrLf & _
             "   ORDER BY A.DH_REGT_MESG_SPB DESC "
    
    Set flLerDadosMensagem = fgQuerySQL(strSQL)


    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerDadosMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function ConciliarOperacao(ByVal plngNumeroSequenciaOperacao As Long) As Boolean


Dim rsMensagem                              As ADODB.Recordset
Dim rsOperacao                              As ADODB.Recordset
Dim strSQL                                  As String
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim pstrNumeroControleIF As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT                               " & vbCrLf & _
             "          OPER.NU_SEQU_OPER_ATIV,     " & vbCrLf & _
             "          OPER.CO_OPER_ATIV,          " & vbCrLf & _
             "          OPER.DT_OPER_ATIV,          " & vbCrLf & _
             "          OPER.NU_ATIV_MERC,          " & vbCrLf & _
             "          OPER.DT_VENC_ATIV,          " & vbCrLf & _
             "          OPER.VA_OPER_ATIV,          " & vbCrLf & _
             "          OPER.NU_CTRL_MESG_SPB_ORIG, " & vbCrLf & _
             "          OPER.CO_ISPB_IF_CNPT,       " & vbCrLf & _
             "          OPER.IN_ENTR_MANU           " & vbCrLf & _
             " FROM     A8.TB_OPER_ATIV OPER        " & vbCrLf & _
             " WHERE    OPER.TP_OPER                = " & enumTipoOperacaoLQS.NegociacaoOperacaoCCR & vbCrLf & _
             " AND      OPER.NU_SEQU_OPER_ATIV      = " & plngNumeroSequenciaOperacao
    
    Set rsOperacao = fgQuerySQL(strSQL)
    
    If rsOperacao.EOF Then
        Set rsOperacao = Nothing
        Exit Function
    End If
    
    strSQL = " SELECT                                   " & vbCrLf & _
             "          A.NU_CTRL_IF,                   " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,             " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,            " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,            " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,            " & vbCrLf & _
             "          A.NU_CTRL_CAMR,                 " & vbCrLf & _
             "          A.CO_MESG_SPB,                  " & vbCrLf & _
             "          B.DT_OPER,                      " & vbCrLf & _
             "          B.NU_ATIV_MERC,                 " & vbCrLf & _
             "          B.VA_FINC,                      " & vbCrLf & _
             "          B.DT_VENC,                      " & vbCrLf & _
             "          B.CO_ISPB_IF_CNPT               " & vbCrLf & _
             " FROM    A8.TB_MESG_RECB_ENVI_SPB A,      " & vbCrLf & _
             "         A8.TB_MESG_RECB_SPB_CNCL B       " & vbCrLf & _
             " WHERE   A.NU_CTRL_IF         = B.NU_CTRL_IF        (+) " & vbCrLf & _
             " AND     A.NU_SEQU_CNTR_REPE  = B.NU_SEQU_CNTR_REPE (+) " & vbCrLf & _
             " AND     A.DH_REGT_MESG_SPB   = B.DH_REGT_MESG_SPB  (+) " & vbCrLf

    strSQL = strSQL & _
             "   AND   A.NU_SEQU_CNTR_REPE       = 1 " & vbCrLf & _
             "   AND   A.CO_MESG_SPB             = 'CCR0006'               " & vbCrLf & _
             "   AND   A.NU_CTRL_CAMR            = '" & Trim(rsOperacao!NU_CTRL_MESG_SPB_ORIG) & "'" & vbCrLf & _
             "   AND   B.DT_OPER                 =  " & fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))) '& vbCrLf & _
             --"   AND   B.VA_FINC                <= " & fgVlrToDBServer(rsOperacao!VA_OPER_ATIV)
    
    Set rsMensagem = fgQuerySQL(strSQL)
    
    If rsMensagem.EOF Then
        Set rsMensagem = Nothing
        ConciliarOperacao = False
        Exit Function
    End If
    
    ConciliarOperacao = True
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    pstrNumeroControleIF = rsMensagem!NU_CTRL_IF
    
    
    If rsMensagem!CO_ULTI_SITU_PROC <> enumStatusMensagem.Conciliada Then
        Call objMesgRecbEnviSPB.AlterarStatus(pstrNumeroControleIF, _
                                              fgDtHr_To_Xml(rsMensagem!DH_REGT_MESG_SPB), _
                                              enumStatusMensagem.Conciliada, _
                                              "", _
                                              "", _
                                              False, _
                                              1)
           
        Call objMesgRecbEnviSPB.AlterarSequenciaOperacao(pstrNumeroControleIF, _
                                                         rsMensagem!DH_REGT_MESG_SPB, _
                                                         rsOperacao!NU_SEQU_OPER_ATIV, _
                                                         1)
    End If
    Set objMesgRecbEnviSPB = Nothing
    Set rsMensagem = Nothing
    Set rsOperacao = Nothing
    
    
    Exit Function
    
ErrorHandler:
    
    Set objProcessoOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set rsMensagem = Nothing
    Set rsOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flProcessaMensagemCCR0007(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim lngSeqControleRepeticao                 As Long

Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String

Dim vntCNPJNLiqdant                         As Variant
Dim intIndDebitoCredito                     As Integer
Dim strAux                                  As String

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   "")
    
    lngSeqControleRepeticao = 1
      
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            enumStatusMensagem.Informação, _
                                                            enumIndicadorSimNao.Nao, _
                                                            vbNullString, _
                                                            enumLocalLiquidacao.CCR, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            vbNullString)
    

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                Replace(xmlMensagem.selectSingleNode("//DtMovto").Text, "-", ""), _
                                0, 0, _
                                strAux, _
                                0, 0, _
                                "0", _
                                lngSeqControleRepeticao, , _
                                enumTipoDebitoCredito.Debito, _
                                "", _
                                0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                xmlMensagem.selectSingleNode("//CNPJBaseEntRespons").Text)
    
    For Each xmlNode In xmlMensagem.selectNodes("//Repet_CCR0007_ResumDiario/*")
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                                Trim(strNumeroControleIF), _
                                                                0, _
                                                                0, _
                                                                xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                xmlNode.selectSingleNode("NumCtrlCCROr").Text, _
                                                                vbNullString, _
                                                                enumStatusMensagem.Informação, _
                                                                enumIndicadorSimNao.Nao, _
                                                                vbNullString, _
                                                                enumLocalLiquidacao.CCR, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                lngSeqControleRepeticao, _
                                                                xmlNode.selectSingleNode("CodReemb").Text)
    
        strAux = xmlMensagem.selectSingleNode("//TpComerc").Text & _
                 fgCompletaString(xmlMensagem.selectSingleNode("//TpInstntoCCR").Text, " ", 3, False) & _
                 fgCompletaString(xmlMensagem.selectSingleNode("//TpOpComercExtr").Text, " ", 3, False)
                 
                 
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    Replace(xmlNode.selectSingleNode("DtOp").Text, "-", ""), _
                                    0, 0, _
                                    strAux, _
                                    0, 0, _
                                    xmlNode.selectSingleNode("VlrOp").Text, _
                                    lngSeqControleRepeticao, , _
                                    enumTipoDebitoCredito.Debito, _
                                    "", _
                                    0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                    xmlMensagem.selectSingleNode("//CNPJBaseEntRespons").Text)
                                
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCCR0007 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



