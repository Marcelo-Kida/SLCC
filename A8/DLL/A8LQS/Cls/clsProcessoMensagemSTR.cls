VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessoMensagemSTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo gerenciamento dos fluxos operacionais das mensagens STR

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private ObjectContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Public Function VerificaMensagemSTR0008R2Comti() As String

Dim udtConsultaPZ                           As udtConsultaPZ
Dim udtConsultaPZAux                        As udtConsultaPZAux

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim lngCodigoEmpresa                        As Long
Dim objTRemessa                             As A8LQS.clsTRemessa
Dim xmlControleHora                         As MSXML2.DOMDocument40
Dim vntMensagemPZ                           As Variant
Dim intContador                             As Integer

On Error GoTo ErrorHandler

    Set xmlControleHora = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Dir(App.Path & "\xmlControlePZ.xml") = vbNullString Then
        Call fgAppendNode(xmlControleHora, "", "ControleHoraConsultaPZ", "")
        Call fgAppendNode(xmlControleHora, "ControleHoraConsultaPZ", "Data", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        Call fgAppendNode(xmlControleHora, "ControleHoraConsultaPZ", "UltimaConsulta", "00.00.00")
        
        xmlControleHora.save App.Path & "\xmlControlePZ.xml"
    Else
        xmlControleHora.Load App.Path & "\xmlControlePZ.xml"
        
        If xmlControleHora.selectSingleNode("//Data").Text <> fgDt_To_Xml(fgDataHoraServidor(DataAux)) Then
            xmlControleHora.selectSingleNode("//Data").Text = fgDt_To_Xml(fgDataHoraServidor(DataAux))
            xmlControleHora.selectSingleNode("//UltimaConsulta").Text = "00.00.00"
        End If
        
    End If
    
    strSQL = " SELECT   CO_LOCA_LIQU,   " & vbNewLine & _
             "          CO_EMPR,        " & vbNewLine & _
             "          CO_AGEN_COTR,   " & vbNewLine & _
             "          NU_CC_COTR      " & vbNewLine & _
             " FROM     A8.TB_LOCA_LIQU_CNTA_COTR " & vbNewLine & _
             " WHERE    CO_LOCA_LIQU  IN (" & enumLocalLiquidacao.CLBCAcoes & ", " & _
                                              enumLocalLiquidacao.BMD & ")" & vbNewLine & _
             " ORDER BY CO_LOCA_LIQU    "
             
    Set objRS = fgQuerySQL(strSQL)
    
    'Mensagens que serão consideradas para o envio
    'RATS 1224, os gestores da Custódia solicitam retirar do processo de busca de mensagens R2 no PZ da SRT0005R2 e da PAG0107R2.
    vntMensagemPZ = Split("STR0008R2|STR0006R2|STR0007R2|PAG0108R2", "|")
    
    If Not objRS.EOF Then
        While Not objRS.EOF
                            
            For intContador = 0 To UBound(vntMensagemPZ)
                udtConsultaPZ.AGENCIA = Format$(objRS!CO_AGEN_COTR, "00000")
                lngCodigoEmpresa = objRS!CO_EMPR
    '            udtConsultaPZ.BANCO = Format$(fgObterBanco(objRS!CO_EMPR), "000")
                udtConsultaPZ.BANCO = "033"
                udtConsultaPZ.CO_MENSAGEM = vntMensagemPZ(intContador)
                udtConsultaPZ.AGENCIA = IIf(IsNull(objRS!CO_AGEN_COTR), "00000", Format$(objRS!CO_AGEN_COTR, "00000"))
                udtConsultaPZ.CONTA = IIf(IsNull(objRS!NU_CC_COTR), "0000000000000", Format$(objRS!NU_CC_COTR, "0000000000000"))
                udtConsultaPZ.DT_MENSAGE = fgDt_To_Xml(fgDataHoraServidor(DataAux))
                udtConsultaPZ.Hora = xmlControleHora.selectSingleNode("//UltimaConsulta").Text
                udtConsultaPZ.SG_SIST = "A8 "
                udtConsultaPZ.FILLER = String$(20, " ")
                udtConsultaPZ.QT_ITEM = "00000"
                udtConsultaPZ.RC_ROTINA = "00000"
                udtConsultaPZ.MSG_RC_ROTINA = String$(Len(udtConsultaPZ.MSG_RC_ROTINA), " ")
                                                        
                LSet udtConsultaPZAux = udtConsultaPZ
                
                Set objTRemessa = CreateObject("A8LQS.clsTRemessa")
                Call objTRemessa.EnviarConsultaPZ(lngCodigoEmpresa, udtConsultaPZAux.String)
                Set objTRemessa = Nothing
            
            Next intContador
                                                        
            objRS.MoveNext
        
        Wend
    End If
    
    objRS.Close

    Exit Function
ErrorHandler:
    
    VerificaMensagemSTR0008R2Comti = Err.Description

    Set objTRemessa = Nothing
End Function


Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

'Gerenciar fluxo de recebimento de Retorno mensagens PZ - STR

Public Function GerenciarRecebimentoMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim strCodigoMensagem                       As String
Dim objProcessoMesgBMA                      As A8LQS.clsProcessoMensagemBMA

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim$(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    'Erro de Negocio BACEN e Erro PJ/PK
    If Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ) Then
        Call GerenciarMensagemErroExterno(xmlMensagem, pstrHeaderNZ)
    Else
        
        If strCodigoMensagem = "STR0010R2" Or strCodigoMensagem = "PAG0111R2" Then
            Call flProcessaMensagemDevolucaoTED(xmlMensagem)
        Else
            Set objProcessoMesgBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
            Call objProcessoMesgBMA.ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
            Set objProcessoMesgBMA = Nothing
        End If
    
    End If

    Exit Function
ErrorHandler:
    
    Set objProcessoMesgBMA = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o fluxo de mensagens STR R2 - Conti

Public Function ProcessaMensagemSTR0008R2Comti(ByVal pstrMensagem As String) As Boolean

Dim udtSTR0008R2                            As udtSTR0008R2
Dim udtSTR0008R2Aux                         As udtSTR0008R2Aux

Dim udtConsultaPZ                           As udtConsultaPZ
Dim udtConsultaPZAux                        As udtConsultaPZAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objConciliacaoLote                      As A8LQS.clsConciliacaoLote

Dim xmlControleHora                         As MSXML2.DOMDocument40

Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strNumeroControleCamara                 As String
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim lngSeqControleRepeticao                 As Long

Dim vntValorFinc                            As Variant
Dim lngAgencia                              As Long
Dim vntConta                                As Variant
Dim vntCNPJContraparte                      As Variant
Dim strNomeContraparte                      As String
Dim strISPBContraparte                      As String
Dim strDataLiquidacao                       As String
Dim strHora                                 As String

On Error GoTo ErrorHandler
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacaoLote = CreateObject("A8LQS.clsConciliacaoLote")

    strCodigoMensagemSPB = Mid(pstrMensagem, 38, 9)
    
    If strCodigoMensagemSPB = "STR0008R2" _
    Or strCodigoMensagemSPB = "STR0005R2" _
    Or strCodigoMensagemSPB = "PAG0107R2" _
    Or strCodigoMensagemSPB = "STR0006R2" _
    Or strCodigoMensagemSPB = "STR0007R2" _
    Or strCodigoMensagemSPB = "PAG0108R2" Then
    
        udtConsultaPZAux.String = pstrMensagem
        LSet udtConsultaPZ = udtConsultaPZAux
        
        'Verifica se rotina na alta esta OK
        If Val(udtConsultaPZ.RC_ROTINA) <> 0 Then
        
            Call fgGravaArquivo("ERRO_CONSULTA_" & Mid(strCodigoMensagemSPB, 1, 7), _
                                "ERRO : " & udtConsultaPZ.RC_ROTINA & " - " & udtConsultaPZ.MSG_RC_ROTINA)
            
            Exit Function
        End If
        
        If Trim$(udtConsultaPZ.MSG_RC_ROTINA) <> vbNullString Then
            
            Call fgGravaArquivo("ERRO_CONSULTA_" & Mid(strCodigoMensagemSPB, 1, 7), _
                                "ERRO : " & udtConsultaPZ.RC_ROTINA & " - " & udtConsultaPZ.MSG_RC_ROTINA)
            
            Exit Function
        End If
        
        Set objMesgRecbEnviSPB = Nothing
        
        Exit Function
    Else
        udtSTR0008R2Aux.String = pstrMensagem
        LSet udtSTR0008R2 = udtSTR0008R2Aux
        
        strNumeroControleIF = Trim$(udtSTR0008R2.NU_CTRL_EXTE_RECB)
        
        'Verifica se a mensagem ja existe
        If objMesgRecbEnviSPB.Ler(strNumeroControleIF) <> vbNullString Then
            Set objMesgRecbEnviSPB = Nothing
            Exit Function
        End If
    
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    strMensagem = pstrMensagem
    lngTipoBackOffice = 0
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strDataRegistroMensagem = ""
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.SSTR
    lngStatusMensagem = enumStatusMensagem.AConciliar
    strDataLiquidacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    
    
    lngCodigoEmpresa = udtSTR0008R2.CO_EMPR
    strNumeroControleCamara = Trim$(udtSTR0008R2.NU_RECB_PGTO)
    strNumeroControleIF = Trim$(udtSTR0008R2.NU_CTRL_EXTE_RECB)
    strCodigoMensagemSPB = Trim$(udtSTR0008R2.CO_MESG)
    
    vntValorFinc = Val(Left$(Trim$(udtSTR0008R2.VA_RECB_PGTO), 16)) & "," & Right$(Trim$(udtSTR0008R2.VA_RECB_PGTO), 2)
    
    lngAgencia = IIf(Trim(udtSTR0008R2.CO_AGEN_REMT) = "", 0, CLng("0" & udtSTR0008R2.CO_AGEN_REMT))
    vntConta = IIf(Trim(udtSTR0008R2.NU_CNTA_REMT) = "", 0, CDbl("0" & udtSTR0008R2.NU_CNTA_REMT))
    
    lngCodigoLocalLiquidacao = ObterLocalLiquidacao(lngCodigoEmpresa, IIf(Trim(udtSTR0008R2.CO_AGEN_DEST) = "", 0, CLng("0" & udtSTR0008R2.CO_AGEN_DEST)), udtSTR0008R2.NU_CNTA_DEST)
    
    vntCNPJContraparte = CDbl("0" & udtSTR0008R2.NU_CNPJ_CPF_REMT)
    strNomeContraparte = Trim(udtSTR0008R2.NO_REMT)
    strISPBContraparte = udtSTR0008R2.CO_ISPB_REMT
    strHora = Trim$(udtSTR0008R2.HO_RECB_MESG)
    lngSeqControleRepeticao = 1
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            "", _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara)

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntValorFinc, _
                                lngSeqControleRepeticao, _
                                strDataLiquidacao, _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                vntCNPJContraparte, _
                                strNomeContraparte, _
                                0, _
                                "", 0, 0, "", _
                                strISPBContraparte, _
                                0, "", "", _
                                lngAgencia, _
                                vntConta)

   'verifica se a R2 concilia com algum lote de Recebimento de corretoras
    If strCodigoMensagemSPB = "STR0006R2" Or strCodigoMensagemSPB = "STR0007R2" Or strCodigoMensagemSPB = "STR0008R2" Or strCodigoMensagemSPB = "PAG0108R2" Then
        
        If objConciliacaoLote.VerificarConciliacaoR2PagtoCorretora(pstrMensagem, lngCodigoLocalLiquidacao) = True Then
            Call objMesgRecbEnviSPB.AlterarStatus(Trim(strNumeroControleIF), strDataRegistroMensagem, _
                                                  enumStatusMensagem.MensagemLiquidadaAutomatica, vbNullString, , False, lngSeqControleRepeticao)
        End If
    
    End If
    
    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    Set objConciliacaoLote = Nothing
    
    Set xmlControleHora = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Dir(App.Path & "\xmlControlePZ.xml") <> vbNullString Then
        
        xmlControleHora.Load App.Path & "\xmlControlePZ.xml"
        
        If xmlControleHora.selectSingleNode("//Data").Text <> fgDt_To_Xml(fgDataHoraServidor(DataAux)) Then
            xmlControleHora.selectSingleNode("//Data").Text = fgDt_To_Xml(fgDataHoraServidor(DataAux))
        End If
        
        strHora = Mid(strHora, 1, 2) & "." & Mid(strHora, 3, 2) & "." & Mid(strHora, 5, 2)
        
        xmlControleHora.selectSingleNode("//UltimaConsulta").Text = strHora
        
        xmlControleHora.save App.Path & "\xmlControlePZ.xml"
        
    End If
    
    Set xmlControleHora = Nothing
    
    Exit Function
ErrorHandler:

    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    Set objConciliacaoLote = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaMensagemSTR0008R2Comti Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function





'Controlar o fluxo de mensagens STR R2 - Fila MQ

Private Function flProcessaMensagemSTR0008R2Fila(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                                 ByRef pstrHeaderNZ As String) As Boolean


Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strNumeroControleCamara                 As String
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim lngSeqControleRepeticao                 As Long

Dim vntValorFinc                            As Variant
Dim lngAgencia                              As Long
Dim vntConta                                As Variant
Dim vntCNPJContraparte                      As Variant
Dim strNomeContraparte                      As String
Dim strISPBContraparte                      As String
Dim strDataLiquidacao                       As String

On Error GoTo ErrorHandler
    
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strDataRegistroMensagem = ""
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.SSTR
        
    If Not xmlMensagem.selectSingleNode("//NumCtrlSTR") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlSTR").Text
    Else
        strNumeroControleCamara = vbNullString
    End If
    
    strDataLiquidacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntValorFinc = xmlMensagem.selectSingleNode("//VlrLanc").Text
    lngAgencia = xmlMensagem.selectSingleNode("//AgCredtd").Text
    vntConta = xmlMensagem.selectSingleNode("//CtCredtd").Text
    vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJ_CPFCliDebtd").Text
    strNomeContraparte = xmlMensagem.selectSingleNode("//NomCliDebtd").Text
    strISPBContraparte = xmlMensagem.selectSingleNode("//ISPBIFDebtd").Text
    
    If flVerificaAgenciaContaTratada(lngCodigoLocalLiquidacao, _
                                     lngCodigoEmpresa, _
                                     lngAgencia, _
                                     vntConta) Then
                                     
        lngStatusMensagem = enumStatusMensagem.AConciliar
        lngCodigoLocalLiquidacao = enumLocalLiquidacao.SSTR
    Else
        lngStatusMensagem = enumStatusMensagem.AConciliar
        lngCodigoLocalLiquidacao = 0
    End If
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    lngSeqControleRepeticao = 1
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            "", _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara)

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntValorFinc, _
                                lngSeqControleRepeticao, _
                                strDataLiquidacao, _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                vntCNPJContraparte, _
                                strNomeContraparte, _
                                0, _
                                "", 0, 0, "", _
                                strISPBContraparte, _
                                0, "", "", _
                                lngAgencia, _
                                vntConta)

                            
    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    
    Exit Function
ErrorHandler:

    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemSTR0008R2Fila Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Controlar o fluxo de envio de informações para outros sistemas

Private Function flGerenciarChamada(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByVal plngStatusOperacao As Long, _
                                    ByVal plngStatusMensagem As Long, _
                                    ByVal pstrSituacaoMensagem As String, _
                                    ByVal plngEstorno As enumIndicadorSimNao, _
                                    ByVal pstrNumeroControleIF As String, _
                           Optional ByVal pblnAtualizaStatusOperacao As Boolean = False) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objLegado                               As A8LQS.clsLegado
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String

On Error GoTo ErrorHandler

    'Alterar Status Operação
    If pblnAtualizaStatusOperacao Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      plngStatusOperacao, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , , False
        Set objOperacao = Nothing
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
    End If

    'Alterar o Status da Mensagem
    If plngStatusMensagem > 0 Then
        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Set xmlMensagemEnviadaRecebida = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagemEnviadaRecebida.loadXML objMensagemEnviadaRecebida.LerMensagemIda(pstrNumeroControleIF)
        objMensagemEnviadaRecebida.AlterarStatus xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                 plngStatusMensagem, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
        Set objMensagemEnviadaRecebida = Nothing
    End If
    
    'Carlos - Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    strXML = objProcessoOperacao.LerTodos("RecebimentoR1", _
                                          CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text), _
                                          plngEstorno, _
                                          plngStatusOperacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlOperacao.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        If Not xmlOperacao.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
        If xmlOperacao.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
            End If
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If

            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
            End If
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
        End If
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            Set objSubReserva = Nothing
        End If
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Disponibilizar xmlOperacao, False
            Set objContaCorrente = Nothing
        End If
    Else
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Cancelar xmlOperacao, False, False
            Set objContaCorrente = Nothing
        End If
    End If
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        Set objLegado = CreateObject("A8LQS.clsLegado")
        strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objLegado = Nothing
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_ALER").Text = enumIndicadorSimNao.Sim Then
        Set objAlerta = CreateObject("A8LQS.clsAlerta")
        objAlerta.GerarAlerta xmlOperacao, xmlMensagem, pstrSituacaoMensagem
        Set objAlerta = Nothing
    End If

    'Estornos de previsão
    objProcessoOperacao.EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        Set objSubReserva = Nothing
    End If

    If plngEstorno = enumIndicadorSimNao.Sim Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            strMensagem = objSubReserva.EstornarRealizadoConfirmado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            Set objSubReserva = Nothing
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If
    End If

    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
    '    ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'---------------------------------------------------------------------------
'Processamento das mensagens R2
'---------------------------------------------------------------------------

Public Function GerenciarMensagemR2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrHeaderNZ As String) As Boolean

Dim strCodigoMensagem                       As String
Dim objProcessoMesgBMA                      As A8LQS.clsProcessoMensagemBMA

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    Select Case strCodigoMensagem
        Case "STR0008R2"
            Call flProcessaMensagemSTR0008R2Fila(xmlMensagem, pstrHeaderNZ)
        Case Else
            Set objProcessoMesgBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
            Call objProcessoMesgBMA.ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemR2)
            Set objProcessoMesgBMA = Nothing
    End Select

    Exit Function
ErrorHandler:
    
    Set objProcessoMesgBMA = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemR2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Private Function flVerificaAgenciaContaTratada(ByVal plngLocalLiquidacao As Long, _
                                               ByVal plngEmpresa As Long, _
                                               ByVal plngAgencia As Long, _
                                               ByVal pvntConta As Variant) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   CO_AGEN_COTR " & vbCrLf & _
             "   FROM   A8.TB_LOCA_LIQU_CNTA_COTR " & vbCrLf & _
             "  WHERE   CO_LOCA_LIQU    = " & plngLocalLiquidacao & vbCrLf & _
             "    AND   CO_EMPR         = " & plngEmpresa & vbCrLf & _
             "    AND   CO_AGEN_COTR    = " & plngAgencia & vbCrLf & _
             "    AND   NU_CC_COTR      = " & pvntConta

    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        flVerificaAgenciaContaTratada = False
    Else
        flVerificaAgenciaContaTratada = True
    End If
    
    objRS.Close
    
    Exit Function
ErrorHandler:
    
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaAgenciaContaTratada Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'---------------------------------------------------------------------------
'Processamento das mensagens R1
'---------------------------------------------------------------------------

Public Function GerenciarMensagemR1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrHeaderNZ As String) As Boolean

Dim udtPZW0916                              As udtPZW0916
Dim udtPZW0916Aux                           As udtPZW0916Aux

Dim udtPZO00140                             As udtPZO00140
Dim udtPZO00140Aux                          As udtPZO00140Aux

Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objRegra                                As A8LQS.clsRegra
Dim objConciliacaoLote                      As A8LQS.clsTConciliacaoLote
                    
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlMensagemOrig                         As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40

Dim strDataRegistroMensagem                 As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim lngSeqControleRepeticao                 As Long
Dim strSitLancSTR                           As String
Dim strConciliacao                          As String
Dim strSituacaoMensagem                     As String
Dim strSequenciaOperacaoAnterior            As String
Dim strNumeroControleIFAnterior             As String
Dim intCodigoLayout                         As Integer
Dim strNumeroControleCamara                 As String

    On Error GoTo ErrorHandler
    
    Set objRegra = CreateObject("A8LQS.clsRegra")
    Set objConciliacaoLote = CreateObject("A8LQS.clsTConciliacaoLote")
    
    udtPZW0916Aux.String = pstrHeaderNZ
    LSet udtPZW0916 = udtPZW0916Aux
    
    strNumeroControleIF = Trim$(udtPZW0916.NumeroControleLegado)
    lngCodigoEmpresa = udtPZW0916.CodigoEmpresa
    strCodigoMensagemSPB = udtPZW0916.CodigoMensagem
    strSitLancSTR = Trim$(xmlMensagem.selectSingleNode("//SitLancSTR").Text)
       
    'Obter Status ---------------------
    Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
    xmlStatus.loadXML flObterStatusOperacaoMensagem(strSitLancSTR)
    '----------------------------------
        
    If xmlStatus.selectSingleNode("//Status_Operacao").Text <> vbNullString Then
        lngStatusOperacao = xmlStatus.selectSingleNode("//Status_Operacao").Text
    Else
        lngStatusOperacao = 0
    End If
            
    'Atualiza Status da Mensagem de Requisição  -------------------------------------------------------------------
    If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
        lngStatusMensagem = xmlStatus.selectSingleNode("//Status_Mensagem").Text
    Else
        lngStatusMensagem = enumStatusMensagem.Respondida
    End If
            
    Set xmlStatus = Nothing
    
    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagemOrig = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a mensagem de ida
    If Not xmlMensagemOrig.loadXML(objMensagemRecebidaEnviada.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlSTR") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlSTR").Text
        Call objMensagemRecebidaEnviada.AlterarNumeroControleCamara(strNumeroControleIF, strNumeroControleCamara)
    End If
    
    'Ler as operacões
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    If xmlMensagemOrig.selectSingleNode("//CO_MESG_SPB").Text = "STR0007" Or _
       xmlMensagemOrig.selectSingleNode("//CO_MESG_SPB").Text = "STR0004" Then
        strConciliacao = objConciliacao.LerTodos(xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    Else
        strConciliacao = objConciliacao.LerTodos(, , strNumeroControleIF)
    End If

    If strConciliacao = vbNullString Then
        Exit Function
    End If

    strSequenciaOperacaoAnterior = vbNullString
    strNumeroControleIFAnterior = vbNullString
    
    xmlConciliacao.loadXML strConciliacao

    For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")

        If strSequenciaOperacaoAnterior <> xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text And _
           Val(xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text) <> 0 Then
        
            strSequenciaOperacaoAnterior = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text

            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
    
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
                '3038 - Operação já liquidada.
                lngCodigoErroNegocio = 3038
                GoTo ErrorHandler
            End If
            
            intCodigoLayout = 0
            If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
                intCodigoLayout = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
            End If
            
            If lngStatusOperacao = enumStatusOperacao.Rejeitada Then
            
                If intCodigoLayout <> enumTipoMensagemBUS.EnvioTEDClientes And _
                   intCodigoLayout <> enumTipoMensagemBUS.EnvioPagDespesas Then
                    objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                  enumStatusOperacao.ConcordanciaBackoffice, _
                                                  xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                  , , _
                                                  False
                Else
                    objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                  lngStatusOperacao, _
                                                  xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                  , , _
                                                  False
                End If
                
                flGerenciarChamada xmlOperacao, _
                                   xmlMensagem, _
                                   lngStatusOperacao, _
                                   lngStatusMensagem, _
                                   strSituacaoMensagem, _
                                   enumIndicadorSimNao.Sim, _
                                   strNumeroControleIF
                                   
            Else
            
                flGerenciarChamada xmlOperacao, _
                                   xmlMensagem, _
                                   lngStatusOperacao, _
                                   lngStatusMensagem, _
                                   strSituacaoMensagem, _
                                   enumIndicadorSimNao.Nao, _
                                   strNumeroControleIF, _
                                   True
            
            End If
            
            If strNumeroControleIFAnterior = vbNullString Then
                strNumeroControleIFAnterior = strNumeroControleIF
                lngStatusMensagem = 0
            End If
        
        End If
    
    Next
    
    objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                       strNumeroControleIF, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                       strCodigoMensagemSPB, _
                                       xmlMensagemOrig.selectSingleNode("//TP_BKOF").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_EMPR").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_COMD_OPER").Text, _
                                       "", _
                                       enumStatusMensagem.R1, _
                                       enumIndicadorSimNao.Nao, _
                                       "", _
                                       xmlMensagemOrig.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       xmlMensagemOrig.selectSingleNode("//SG_SIST").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
    
    If strCodigoMensagemSPB = "STR0007R1" Then
        Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
        If xmlConciliacao.xml <> vbNullString Then
            For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_ConsultaConciliacao")
                If xmlNode.selectSingleNode("CO_MESG_SPB").Text = "BMC0112" Then
                    
                    Select Case lngStatusOperacao
                        Case enumStatusOperacao.Rejeitada
                            lngStatusMensagem = enumStatusMensagem.AConciliar
                        Case enumStatusOperacao.Liquidada
                            lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
                        Case enumStatusOperacao.Pendencia
                            lngStatusMensagem = enumStatusMensagem.MensagemPendente
                    End Select
                    
                    Call objMensagemRecebidaEnviada.AlterarStatus( _
                                                xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                lngStatusMensagem, _
                                                vbNullString, , _
                                                False)
                
                End If
            Next
        End If
        
        'verifica se as operações conciliadas a esta STR0007R1 são provenientes de um lote de Corretoras
        'e se ContaCorrente deve ser integrado automaticamente
        If Not xmlOperacao.selectSingleNode("//ID_LOTE") Is Nothing Then
            If objRegra.VerificarRegraAutomatica2(Val(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text), enumFuncaoSistema.IntegracaoCC) = True Then
                Call objConciliacaoLote.IntegrarContaCorrente(xmlOperacao.selectSingleNode("//ID_LOTE").Text, _
                                                              xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text, _
                                                              fgDtXML_To_Date(xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text))
        
            End If
        End If
    
    ElseIf strCodigoMensagemSPB = "STR0004R1" Then
        If xmlConciliacao.xml <> vbNullString Then
            For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")
                If xmlNode.selectSingleNode("CO_MESG_SPB").Text = "LTR0001" Then
                    
                    Call objMensagemRecebidaEnviada.AlterarStatus( _
                                                xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                enumStatusMensagem.MensagemLiquidada, _
                                                vbNullString, , _
                                                False)
                                                
                End If
            Next
        End If
        
    End If

    Set objMensagemRecebidaEnviada = Nothing
    
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objRegra = Nothing
    Set objConciliacaoLote = Nothing

    Exit Function
ErrorHandler:

    Set objConciliacao = Nothing
    Set objMensagemRecebidaEnviada = Nothing
    Set objOperacao = Nothing
    Set objRegra = Nothing
    Set objConciliacaoLote = Nothing
    Set xmlOperacao = Nothing

    Set xmlStatus = Nothing
    Set xmlOperacao = Nothing
    Set xmlConciliacao = Nothing
    Set xmlMensagemOrig = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens Erro Interno PZ
'---------------------------------------------------------------------------

Public Function GerenciarMensagemErroPZ(ByRef pstrMensagem As String) As Boolean

Dim udtPZW0001                              As udtPZW0001
Dim udtPZW0001Aux                           As udtPZW0001Aux

Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objAlerta                               As A8LQS.clsAlerta

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlMensagemOrig                         As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlMensagemPZ                           As MSXML2.DOMDocument40


Dim strDataRegistroMensagem                 As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoMensagemSPB                    As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim strConciliacao                          As String
Dim strNU_SEQU_OPER_ATIV_ANT                As String

    On Error GoTo ErrorHandler

    udtPZW0001Aux.String = pstrMensagem
    LSet udtPZW0001 = udtPZW0001Aux

    strNumeroControleIF = Trim$(udtPZW0001.ControleLegado)
    strCodigoMensagemSPB = Trim$(udtPZW0001.CodigoMensagem)

    lngStatusOperacao = enumStatusOperacao.Rejeitada
    lngStatusMensagem = enumStatusMensagem.MensagemRejeitada

    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagemOrig = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a mensagem de ida
    If Not xmlMensagemOrig.loadXML(objMensagemRecebidaEnviada.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    Set xmlMensagemPZ = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlMensagemPZ, "", "SISMSG", "")
    Call fgAppendNode(xmlMensagemPZ, "SISMSG", "CodMsg", strCodigoMensagemSPB)
    
    If Trim$(udtPZW0001.Erro1) <> "00000" Then
        Call fgAppendNode(xmlMensagemPZ, "SISMSG", "CodigoErro1", Trim$(udtPZW0001.Erro1) & " - " & fgObterDescricaoErro(Trim$(udtPZW0001.Erro1)))
    End If
    If Trim$(udtPZW0001.Erro2) <> "00000" Then
        Call fgAppendNode(xmlMensagemPZ, "SISMSG", "CodigoErro2", Trim$(udtPZW0001.Erro2) & " - " & fgObterDescricaoErro(Trim$(udtPZW0001.Erro2)))
    End If
    
    If Trim$(udtPZW0001.Erro3) <> "00000" Then
        Call fgAppendNode(xmlMensagemPZ, "SISMSG", "CodigoErro3", Trim$(udtPZW0001.Erro3) & " - " & fgObterDescricaoErro(Trim$(udtPZW0001.Erro3)))
    End If
    
    Call fgAppendNode(xmlMensagemPZ, "SISMSG", "MensagemRetorno", pstrMensagem)
    
    strMensagem = xmlMensagemPZ.xml
    
    Set xmlMensagemPZ = Nothing
    
    objMensagemRecebidaEnviada.Incluir strMensagem, _
                                       strNumeroControleIF, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                       strCodigoMensagemSPB, _
                                       xmlMensagemOrig.selectSingleNode("//TP_BKOF").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_EMPR").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_COMD_OPER").Text, _
                                       "", _
                                       enumStatusMensagem.ErroNegocioSPB, _
                                       enumIndicadorSimNao.Nao, _
                                       "", _
                                       xmlMensagemOrig.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       xmlMensagemOrig.selectSingleNode("//SG_SIST").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNTR_REPE").Text

    'Ler as operacões
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strConciliacao = objConciliacao.LerTodos(, , strNumeroControleIF)

    If strConciliacao = vbNullString Then
        If strCodigoMensagemSPB = "STR0007" Then
            strConciliacao = objConciliacao.LerTodos(, xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            If strConciliacao = vbNullString Then
                Exit Function
            End If
        ElseIf strCodigoMensagemSPB = "STR0004" Then
            strConciliacao = objConciliacao.LerTodos(xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
            If strConciliacao = vbNullString Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If

    strNU_SEQU_OPER_ATIV_ANT = vbNullString
    
    xmlConciliacao.loadXML strConciliacao

    For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")
        If strNU_SEQU_OPER_ATIV_ANT <> xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
            
            strNU_SEQU_OPER_ATIV_ANT = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
            
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                
    
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
                '3038 - Operação já liquidada.
                lngCodigoErroNegocio = 3038
                GoTo ErrorHandler
            End If
             
            If Not xmlOperacao.documentElement.selectSingleNode("//TP_MESG") Is Nothing Then
                If Val("0" & xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
                    Set objAlerta = CreateObject("A8LQS.clsAlerta")
                    Call objAlerta.GerarAlerta(xmlOperacao, , , enumFatorGeradorAlerta.MensagemSTRE)
                    Set objAlerta = Nothing
                    
                End If
            End If
            objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                          lngStatusOperacao, _
                                          xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                          , , False
            
            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
            
            flGerenciarChamada xmlOperacao, _
                               xmlMensagem, _
                               lngStatusOperacao, _
                               lngStatusMensagem, _
                               "", _
                               enumIndicadorSimNao.Sim, _
                               strNumeroControleIF
                               
        End If
    
        If strCodigoMensagemSPB = "STR0004" Then
            If xmlNode.selectSingleNode("CO_MESG_SPB").Text = "LTR0001" Then
                lngStatusMensagem = enumStatusMensagem.MensagemRejeitada
                
                Call objMensagemRecebidaEnviada.AlterarStatus( _
                                            xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                            xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                            lngStatusMensagem, _
                                            vbNullString, , _
                                            False)
            End If
        End If
        
    Next

    If strCodigoMensagemSPB = "STR0007" Then
        Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
        If xmlConciliacao.xml <> vbNullString Then
            For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_ConsultaConciliacao")
                If xmlNode.selectSingleNode("CO_MESG_SPB").Text = "BMC0112" Then
                    
                    Select Case lngStatusOperacao
                        Case enumStatusOperacao.Rejeitada
                            lngStatusMensagem = enumStatusMensagem.AConciliar
                        Case enumStatusOperacao.Liquidada
                            lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
                        Case enumStatusOperacao.Pendencia
                            lngStatusMensagem = enumStatusMensagem.MensagemPendente
                    End Select
                    
                    Call objMensagemRecebidaEnviada.AlterarStatus( _
                                                xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                lngStatusMensagem, _
                                                vbNullString, , _
                                                False)
                
                End If
            Next
        End If
    End If

    Set objOperacao = Nothing
    Set objMensagemRecebidaEnviada = Nothing
    Set objConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlConciliacao = Nothing

    Exit Function

ErrorHandler:
    Set xmlMensagemPZ = Nothing
    Set objConciliacao = Nothing
    Set objMensagemRecebidaEnviada = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Set xmlOperacao = Nothing
    Set xmlConciliacao = Nothing
    Set xmlMensagemOrig = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemErroPZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Obter o Status da Operação e mensagem
'------------------------------------------------------------------------------------------

Private Function flObterStatusOperacaoMensagem(ByVal pstrSitLancSTR As String) As String

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

On Error GoTo ErrorHandler
     
    Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
     
    Call fgAppendNode(xmlStatus, "", "Grupo_Status", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Status_Operacao", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Status_Mensagem", "")
    Call fgAppendNode(xmlStatus, "Grupo_Status", "Situacao_Mensagem", "")
     
    strGrupoMensagem = "STR"
    
    If IsNumeric(pstrSitLancSTR) Then
        pstrSitLancSTR = Format$(Val(pstrSitLancSTR), "0")
    End If
    
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
    
        For Each xmlNode In xmlTagDominioStatus.documentElement.selectNodes("//Grupo_TagDominioStatus")
            If Trim$(pstrSitLancSTR) = Trim$(xmlNode.selectSingleNode("DE_DOMI").Text) Then
                lngStatusOperacao = xmlNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                lngStatusMensagem = xmlNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text
                Exit For
            End If
        Next
                        
        xmlStatus.selectSingleNode("//Status_Operacao").Text = IIf(lngStatusOperacao = 0, "", lngStatusOperacao)
        xmlStatus.selectSingleNode("//Status_Mensagem").Text = IIf(lngStatusMensagem = 0, "", lngStatusMensagem)
        xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = strSituacaoMensagemSPB
        
        Set xmlTagDominioStatus = Nothing
    
    Next
    
    flObterStatusOperacaoMensagem = xmlStatus.xml
    
    Set xmlStatus = Nothing
   
    Exit Function
ErrorHandler:
    
    Set xmlStatus = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusOperacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens Erro NZ Externo STR
'---------------------------------------------------------------------------

Public Function GerenciarMensagemErroNZ(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                        ByVal pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZExt                   As udtProtocoloErroNZExt
Dim udtProtocoloErroNZExtAux                As udtProtocoloErroNZExtAux

Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlMensagemOrig                         As MSXML2.DOMDocument40

Dim strDataRegistroMensagem                 As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoMensagemSPB                    As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim strConciliacao                          As String

Dim strNU_SEQU_OPER_ATIV_ANT                As String

    On Error GoTo ErrorHandler

    udtProtocoloErroNZExtAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZExt = udtProtocoloErroNZExtAux
    
    strNumeroControleIF = Trim$(udtProtocoloErroNZExt.ControleRemessaNZ)
    strCodigoMensagemSPB = Trim$(udtProtocoloErroNZExt.CodigoMensagem)

    lngStatusOperacao = enumStatusOperacao.Rejeitada
    lngStatusMensagem = enumStatusMensagem.MensagemRejeitada

    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagemOrig = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a mensagem de ida
    If Not xmlMensagemOrig.loadXML(objMensagemRecebidaEnviada.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    Call objMensagemRecebidaEnviada.Incluir(pxmlMensagem.xml, _
                                            strNumeroControleIF, _
                                            xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                            xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                            strCodigoMensagemSPB, _
                                            xmlMensagemOrig.selectSingleNode("//TP_BKOF").Text, _
                                            xmlMensagemOrig.selectSingleNode("//CO_EMPR").Text, _
                                            xmlMensagemOrig.selectSingleNode("//NU_COMD_OPER").Text, _
                                            "", _
                                            enumStatusMensagem.R1, _
                                            enumIndicadorSimNao.Nao, _
                                            "", _
                                            xmlMensagemOrig.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                            xmlMensagemOrig.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                            xmlMensagemOrig.selectSingleNode("//SG_SIST").Text, _
                                            xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

    Set objMensagemRecebidaEnviada = Nothing

    'Ler as operacões
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strConciliacao = objConciliacao.LerTodos(, , strNumeroControleIF)

    If strConciliacao = vbNullString Then
        Exit Function
    End If

    Call xmlConciliacao.loadXML(strConciliacao)
    Call fgAppendNode(pxmlMensagem, "MESG", "CodMsg", strCodigoMensagemSPB)
    
    strNU_SEQU_OPER_ATIV_ANT = vbNullString

    For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")

        If strNU_SEQU_OPER_ATIV_ANT <> xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
            
            strNU_SEQU_OPER_ATIV_ANT = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
    
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
                '3038 - Operação já liquidada.
                lngCodigoErroNegocio = 3038
                GoTo ErrorHandler
            End If
    
            Call objOperacao.AlterarStatusAcao(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               lngStatusOperacao, _
                                               xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                               , , _
                                               False)
            
            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
            
            Call flGerenciarChamada(xmlOperacao, _
                                    pxmlMensagem, _
                                    lngStatusOperacao, _
                                    lngStatusMensagem, _
                                    "", _
                                    enumIndicadorSimNao.Sim, _
                                    strNumeroControleIF)
                                    
        End If

    Next

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMensagemRecebidaEnviada = Nothing
    Set objOperacao = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlConciliacao = Nothing
    Set xmlMensagemOrig = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemErroNZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens Erro Externo STR
'---------------------------------------------------------------------------

Public Function GerenciarMensagemErroExterno(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                             ByVal pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlMensagemOrig                         As MSXML2.DOMDocument40

Dim strDataRegistroMensagem                 As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoMensagemSPB                    As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim strConciliacao                          As String

On Error GoTo ErrorHandler

    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    strNumeroControleIF = Trim$(udtProtocoloErroNZ.ControleRemessaNZ)
    strCodigoMensagemSPB = Trim$(udtProtocoloErroNZ.CodigoMensagem)

    lngStatusOperacao = enumStatusOperacao.Rejeitada
    lngStatusMensagem = enumStatusMensagem.MensagemRejeitada

    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMensagemOrig = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a mensagem de ida
    If Not xmlMensagemOrig.loadXML(objMensagemRecebidaEnviada.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    objMensagemRecebidaEnviada.Incluir pxmlMensagem.xml, _
                                       strNumeroControleIF, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                       strCodigoMensagemSPB, _
                                       xmlMensagemOrig.selectSingleNode("//TP_BKOF").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_EMPR").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_COMD_OPER").Text, _
                                       "", _
                                       enumStatusMensagem.R1, _
                                       enumIndicadorSimNao.Nao, _
                                       "", _
                                       xmlMensagemOrig.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                       xmlMensagemOrig.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       xmlMensagemOrig.selectSingleNode("//SG_SIST").Text, _
                                       xmlMensagemOrig.selectSingleNode("//NU_SEQU_CNTR_REPE").Text

    Set objMensagemRecebidaEnviada = Nothing

    'Ler as operacões
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strConciliacao = objConciliacao.LerTodos(, , strNumeroControleIF)

    If strConciliacao = vbNullString Then
        Exit Function
    End If

    xmlConciliacao.loadXML strConciliacao

    For Each xmlNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")

        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

        If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If

        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      enumStatusOperacao.ConcordanciaBackoffice, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , , False
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
        
        flGerenciarChamada xmlOperacao, _
                           pxmlMensagem, _
                           lngStatusOperacao, _
                           lngStatusMensagem, _
                           "", _
                           enumIndicadorSimNao.Sim, _
                           strNumeroControleIF

    Next

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMensagemRecebidaEnviada = Nothing
    Set objOperacao = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlConciliacao = Nothing
    Set xmlMensagemOrig = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemErroExterno Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function EnviarConsultaPZ(ByVal plngCodigoEmpresa As Long, _
                                 ByVal pstrConsultaPZ As String) As String

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim lngBufferLen                            As Long
Dim strNomeFila                             As String
Dim xmlFilas                                As MSXML2.DOMDocument40
Dim strNomeFilaConsPZ                     As String

On Error GoTo ErrorHandler
    
    
'RATS 393 - Ocor 01
   'Paremetrizar com
        '<NomeFilaConsultaPZ Empresa='523'>NZQ.E.CONSULTA_PZ.1</NomeFilaConsultaPZ>
        '<NomeFilaConsultaPZ Empresa='701'>NZQ.E.CONSULTA_PZ.2</NomeFilaConsultaPZ>
    
    Set xmlFilas = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlFilas.Load(App.Path & "\FilasNET.xml") Then
        
        If plngCodigoEmpresa = enumCodigoEmpresa.Banespa Then
            If Not xmlFilas.selectSingleNode("//NomeFilaConsultaPZ[@Empresa='" & enumCodigoEmpresa.Banespa & "']") Is Nothing Then
                strNomeFilaConsPZ = xmlFilas.selectSingleNode("//NomeFilaConsultaPZ[@Empresa='" & plngCodigoEmpresa & "']").Text
            Else
                strNomeFilaConsPZ = "NZQ.E.CONSULTA_PZ.2"
            End If
        Else
            If Not xmlFilas.selectSingleNode("//NomeFilaConsultaPZ[@Empresa='" & enumCodigoEmpresa.Santander & "']") Is Nothing Then
                strNomeFilaConsPZ = xmlFilas.selectSingleNode("//NomeFilaConsultaPZ[@Empresa='" & plngCodigoEmpresa & "']").Text
            Else
                strNomeFilaConsPZ = "NZQ.E.CONSULTA_PZ.1"
            End If
        End If
    Else
        strNomeFilaConsPZ = "NZQ.E.CONSULTA_PZ"
    End If
    
    Set xmlFilas = Nothing
    
    strNomeFila = strNomeFilaConsPZ
    
    lngBufferLen = Len(pstrConsultaPZ)

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    objMQAX200.MQInit
    
    objMQAX200.MQOpen strNomeFila, enumMQOO_Open.MQOO_OUTPUT
    
    objMQAX200.MQPut pstrConsultaPZ, _
                     lngBufferLen, _
                     enumPutOptions.MainFrame, "", "", "", "", 30000
                     
    objMQAX200.MQClose

    objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing
    
    Exit Function
ErrorHandler:
    Set objMQAX200 = Nothing
    Set xmlFilas = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarConsultaPZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function ObterPontoVenda_UniOrg(ByVal plngEmpresa As Long, _
                                       ByVal plngAgencia As Long, _
                                       ByVal pvntConta As Variant) As Long

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   CO_UNI_ORG " & vbCrLf & _
             "   FROM   A8.TB_LOCA_LIQU_CNTA_COTR " & vbCrLf & _
             "  WHERE   CO_LOCA_LIQU    = " & enumLocalLiquidacao.SSTR & vbCrLf & _
             "    AND   CO_EMPR         = " & plngEmpresa & vbCrLf & _
             "    AND   CO_AGEN_COTR    = " & plngAgencia & vbCrLf & _
             "    AND   NU_CC_COTR      = " & pvntConta
    
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        ObterPontoVenda_UniOrg = objRS!CO_UNI_ORG
    Else
        Err.Raise vbError, "ObterPontoVenda_UniOrg", "Ponto de venda/Uniorg não configurada para Agência: " & plngAgencia & " e Conta: " & pvntConta & " - " & "(A8.TB_LOCA_LIQU_CNTA_COTR)"
    End If
    
    objRS.Close
    
    Set objRS = Nothing
    
    Exit Function
ErrorHandler:
    Set objRS = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPontoVenda_UniOrg Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterAgenciaContaDebitada(ByVal plngEmpresa As Long, _
                                          ByVal plngLocalLiquidacao As Long, _
                                          ByRef plngAgencia As Long, _
                                          ByRef pvntConta As Variant, _
                                          ByRef pvntCodUniorg As Variant)
                                          

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   " & _
             "          CO_LOCA_LIQU, " & vbCrLf & _
             "          CO_EMPR, " & vbCrLf & _
             "          CO_AGEN_COTR, " & vbCrLf & _
             "          NU_CC_COTR, " & vbCrLf & _
             "          CO_UNI_ORG  " & vbCrLf & _
             "   FROM   A8.TB_LOCA_LIQU_CNTA_COTR " & vbCrLf & _
             "  WHERE   CO_LOCA_LIQU    = " & plngLocalLiquidacao & vbCrLf & _
             "    AND   CO_EMPR         = " & plngEmpresa

    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        plngAgencia = objRS!CO_AGEN_COTR
        pvntConta = objRS!NU_CC_COTR
        pvntCodUniorg = objRS!CO_UNI_ORG
    Else
        Err.Raise vbError, "ObterAgenciaContaDebitada", "Agencia/Conta Debitada não cadastrada para empresa : " & plngEmpresa & " e Local Liquidacao: " & plngLocalLiquidacao & " - " & "(A8.TB_LOCA_LIQU_CNTA_COTR)"
    End If
    
    objRS.Close
    Set objRS = Nothing
    Exit Function
ErrorHandler:
    Set objRS = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterAgenciaContaDebitada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterLocalLiquidacao(ByVal plngEmpresa As Long, _
                                     ByVal plngAgencia As Long, _
                                     ByRef pvntConta As Variant) As Long

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   " & _
             "          CO_LOCA_LIQU, " & vbCrLf & _
             "          CO_EMPR, " & vbCrLf & _
             "          CO_AGEN_COTR, " & vbCrLf & _
             "          NU_CC_COTR, " & vbCrLf & _
             "          CO_UNI_ORG  " & vbCrLf & _
             "   FROM   A8.TB_LOCA_LIQU_CNTA_COTR " & vbCrLf & _
             "  WHERE   CO_EMPR         = " & plngEmpresa & vbCrLf & _
             "    AND   CO_AGEN_COTR    = " & plngAgencia & vbCrLf & _
             "    AND   NU_CC_COTR      = " & Val("0" & pvntConta)
             

    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        ObterLocalLiquidacao = objRS!CO_LOCA_LIQU
    Else
        Err.Raise vbError, "ObterLocalLiquidacao", "Local Liquidação não cadastrado para empresa : " & plngEmpresa & " e Agencia: " & plngAgencia & " e Conta:" & pvntConta & " - " & "(A8.TB_LOCA_LIQU_CNTA_COTR)"
    End If
    
    objRS.Close
    Set objRS = Nothing
    
    Exit Function
ErrorHandler:
    Set objRS = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterLocalLiquidacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens de devolução de TED a Clientes
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemDevolucaoTED(ByVal xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objAlerta                               As A8LQS.clsAlerta

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgEnviada                          As MSXML2.DOMDocument40

Dim lngStatusMensagem                       As Long

Dim vntNumSequOperacao                      As Variant
Dim strNumConcilOperacao                    As String
Dim strCodigoVeicLegal                      As String
Dim strSiglaSistema                         As String
Dim strCodigoEmpresa                        As String
Dim strCodigoOperacao                       As String
Dim intTipoBackOffice                       As Integer
Dim strDataRegistroMensagem                 As String
Dim lngLocalLiquidacao                      As Long

    On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objAlerta = CreateObject("A8LQS.clsAlerta")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgEnviada = CreateObject("MSXML2.DOMDocument.4.0")

    intTipoBackOffice = enumTipoBackOffice.Tesouraria
    strNumConcilOperacao = "0"
 
    Call xmlMesgEnviada.loadXML(objMensagemEnvRec.LerTodos(xmlRemessa.selectSingleNode("//NumCtrlIF").Text))
    
    If xmlMesgEnviada.xml = vbNullString Then
        'Não foi possível encontrar a mensagem TED original.
        lngCodigoErroNegocio = 4281
        GoTo ErrorHandler
    Else
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
    End If

    Set xmlMesgEnviada = objMensagemEnvRec.LerXMLMensagem(xmlRemessa.selectSingleNode("//NumCtrlIF").Text, , , _
                                                          Left$(xmlMesgEnviada.selectSingleNode("//CO_MESG_SPB").Text, 8))

    strNumConcilOperacao = xmlMesgEnviada.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text

    Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgEnviada.selectSingleNode("//NU_CTRL_IF").Text))

    If xmlOperacao.xml <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

        vntNumSequOperacao = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        strCodigoVeicLegal = xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
        strCodigoEmpresa = xmlOperacao.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
        lngLocalLiquidacao = xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text
        'Validação Automatização Pagamento Despesas
        If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val("0" & xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
                 If fgDtXML_To_Date(Left$(xmlMesgEnviada.selectSingleNode("//DH_REGT_MESG_SPB").Text, 8)) = fgDtXML_To_Date(xmlRemessa.selectSingleNode("//DtMovto").Text) Then
                    'fgDtXML_To_Date( xmlRemessa.selectSingleNode("//DtMovto").text) >
                                    
                                                
                 
                    Call objAlerta.GerarAlerta(xmlOperacao, , , enumFatorGeradorAlerta.MensagemDevolucaoTEDIndevida)
                 
                 End If
            End If
        End If
         
        
    
    End If

    strDataRegistroMensagem = objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                                           xmlRemessa.selectSingleNode("//NumCtrlIF").Text, _
                                                           vntNumSequOperacao, _
                                                           strNumConcilOperacao, _
                                                           xmlRemessa.selectSingleNode("//CodMsg").Text, _
                                                           intTipoBackOffice, _
                                                           xmlRemessa.selectSingleNode("//CO_EMPR").Text, _
                                                           vbNullString, _
                                                           vbNullString, _
                                                           lngStatusMensagem, _
                                                           enumIndicadorSimNao.Nao, _
                                                           vbNullString, _
                                                           lngLocalLiquidacao, _
                                                           strCodigoVeicLegal, _
                                                           strSiglaSistema)

    Call objConciliacao.Incluir(xmlRemessa.selectSingleNode("//NumCtrlIF").Text, _
                                strDataRegistroMensagem, _
                                xmlRemessa.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text, , _
                                xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text)

    Call fgAppendNode(xmlRemessa, "MESG", "TP_MESG", fgCompletaString("152", "0", 9, True))
    Call fgAppendNode(xmlRemessa, "MESG", "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
    Call fgAppendNode(xmlRemessa, "MESG", "SG_SIST_DEST", fgCompletaString(strSiglaSistema, " ", 3, False))
    Call fgAppendNode(xmlRemessa, "MESG", "CO_EMPR", Format$(strCodigoEmpresa, "00000"))
    Call fgAppendNode(xmlRemessa, "MESG", "DT_MESG", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
    Call fgAppendNode(xmlRemessa, "MESG", "HO_MESG", Format$(Now, "mmss"))
    Call fgAppendNode(xmlRemessa, "MESG", "CO_OPER_ATIV", strCodigoOperacao)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_VEIC_LEGA", strCodigoVeicLegal)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_MESG_STR_PAG", xmlOperacao.selectSingleNode("//CO_MESG_STR_PAG").Text)
    If Not xmlOperacao.selectSingleNode("//CO_UNID_REGI") Is Nothing Then
        Call fgAppendNode(xmlRemessa, "MESG", "CO_UNID_REGI", xmlOperacao.selectSingleNode("//CO_UNID_REGI").Text)
    Else
        Call fgAppendNode(xmlRemessa, "MESG", "CO_UNID_REGI", " ")
    End If
    If Not xmlOperacao.selectSingleNode("//ID_DEPO_JUDI") Is Nothing Then
        Call fgAppendNode(xmlRemessa, "MESG", "ID_DEPO_JUDI", xmlOperacao.selectSingleNode("//ID_DEPO_JUDI").Text)
    Else
        Call fgAppendNode(xmlRemessa, "MESG", "ID_DEPO_JUDI", " ")
    End If
    Call fgAppendNode(xmlRemessa, "MESG", "CO_MESG", xmlRemessa.selectSingleNode("//CodMsg").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "DH_BACEN", xmlRemessa.selectSingleNode("//DtHrBC").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_ISPB_IF_DEBT", xmlRemessa.selectSingleNode("//ISPBIFDebtd").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_ISPB_IF_CRED", xmlRemessa.selectSingleNode("//ISPBIFCredtd").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_DEVO_TRAN", xmlRemessa.selectSingleNode("//CodDevTransf").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "NU_CTRL_IF", xmlRemessa.selectSingleNode("//NumCtrlIF").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "DE_HIST", xmlRemessa.selectSingleNode("//Hist").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "DT_MOVI", xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "VA_OPER_ATIV", xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text)
    Call fgAppendNode(xmlRemessa, "MESG", "CO_FLND_TED", xmlOperacao.selectSingleNode("//CO_FLND_TED").Text)

    udtProtocolo.CodigoEmpresa = Format$(strCodigoEmpresa, "00000")
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString("152", "0", 9, True)

    LSet udtProtocoloAux = udtProtocolo
    
    Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
    
    Set objAlerta = Nothing
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    

    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgEnviada = Nothing

    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing

    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgEnviada = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemDevolucaoTED Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


