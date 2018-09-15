VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessoMensagemBMA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      21-set-2004 15:38:13
'-------------------------------------------------
'' Objeto responsável pelo processamento e liquidação das operações e mensagens da
'' BMA e CETIP
''
'' - Mensagens R1 , R2 , Aviso e Informação (CTP, LDL, LTR, BMA)
''
''
''
''

Option Explicit

Private Const NatuOp_Compra = "CP"   'CP Compra(Crédito)
Private Const NatuOp_Recompra = "RC" 'RC Recompra(Crédito)
Private Const NatuOp_Venda = "VD"    'VD Venda(Débito)
Private Const NatuOp_Revenda = "RV"  'RV Revenda(Débito)

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

'---------------------------------------------------------------------------
'Controlar o processamento das mensagens por tipo :
'   - Erro de Negocio (Erro BACEN ou PJ/PK)
'   - Mensagem R1
'   - Mensagem R2
'   - Mensagem de Aviso e Informação
'---------------------------------------------------------------------------

Public Function GerenciarRecebimentoMensagemBMA(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                                ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    If strCodigoMensagem = "BMC0001E" Or _
      (strCodigoMensagem = "BMC0001" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaMensagemBMC0001E(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "BMC0002E" Or _
      (strCodigoMensagem = "BMC0002" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaMensagemBMC0002E(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "BMC0003E" Or _
      (strCodigoMensagem = "BMC0003" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaMensagemBMC0003E(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
        
    If strCodigoMensagem = "BMC0012E" Or _
      (strCodigoMensagem = "BMC0012" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaMensagemBMC0012E(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "CAM0002E" Or _
      (strCodigoMensagem = "CAM0002" And Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ)) Then
        Call flProcessaMensagemCAM0002E(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    'Erro de Negocio BACEN e Erro PJ/PK
    If Len(pstrHeaderNZ) = Len(udtProtocoloErroNZ) Then
        Call GerenciarMensagemErroNegocio(xmlMensagem, pstrHeaderNZ)
    Else
        If Mid(strCodigoMensagem, 8, 2) = "R1" Then
            Call GerenciarMensagemR1(xmlMensagem, pstrHeaderNZ)
        ElseIf Mid(strCodigoMensagem, 8, 2) = "R2" Then
            Call GerenciarMensagemR2(xmlMensagem, pstrHeaderNZ)
        Else
            Call GerenciarMensagemAvisoInformacao(xmlMensagem, pstrHeaderNZ)
        End If
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoMensagemBMA Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens R1
'---------------------------------------------------------------------------

Public Function GerenciarMensagemR1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrHeaderNZ As String) As Boolean

Dim blnMensagemTratada                      As Boolean
Dim objMensagem                             As A8LQS.clsMensagem
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    
    If strCodigoMensagem = "BMC0001R1" Then
        Call flProcessaMensagemBMC0001R1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "BMC0002R1" Then
        Call flProcessaMensagemBMC0002R1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "BMC0003R1" Then
        Call flProcessaMensagemBMC0003R1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If
    
    If strCodigoMensagem = "BMC0012R1" Then
        Call flProcessaMensagemBMC0012R1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If

    If strCodigoMensagem = "CAM0002R1" Then
        Call flProcessaMensagemCAM0002R1(xmlMensagem, pstrHeaderNZ)
        Exit Function
    End If

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    blnMensagemTratada = objMensagem.VerificaMensagemTratadaSLCC(Mid(strCodigoMensagem, 1, 7))
    Set objMensagem = Nothing
        
    'BMF
    If Mid(strCodigoMensagem, 1, 7) = "LDL0004" Then
        blnMensagemTratada = False
    End If
        
    If blnMensagemTratada Then
        Call flProcessaMesgTratadaR1(xmlMensagem, pstrHeaderNZ)
    Else
        Call flProcessaMesgNaoTratadaR1(xmlMensagem, pstrHeaderNZ)
    End If
     
    Exit Function

ErrorHandler:
    Set objMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens de erro de negócio (BACEN e PJ/PK)
'---------------------------------------------------------------------------

Public Function GerenciarMensagemErroNegocio(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim blnMensagemTratada                      As Boolean
Dim objMensagem                             As A8LQS.clsMensagem
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    blnMensagemTratada = objMensagem.VerificaMensagemTratadaSLCC(Mid(strCodigoMensagem, 1, 7))
    Set objMensagem = Nothing
    
    If blnMensagemTratada Then
        Call flProcessaMesgTratadaErro(xmlMensagem, pstrHeaderNZ)
    Else
        Call flProcessaMesgNaoTratadaErro(xmlMensagem, pstrHeaderNZ)
    End If

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemErroNegocio Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens R2
'---------------------------------------------------------------------------

Public Function GerenciarMensagemR2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrHeaderNZ As String) As Boolean

Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    Select Case strCodigoMensagem
        Case "LTR0005R2"
            Call flProcessaMensagemLTR0005R2(xmlMensagem, pstrHeaderNZ)
        Case "LDL0005R2"
            Call flProcessaMensagemLDL0005R2(xmlMensagem, pstrHeaderNZ)
        Case "LDL0006R2"
            Call flProcessaMensagemLDL0006R2(xmlMensagem, pstrHeaderNZ)
        Case "LDL0009R2"
            Call flProcessaMensagemLDL0009R2(xmlMensagem, pstrHeaderNZ)
        Case "LTR0003R2"
            Call flProcessaMensagemLTR0003R2(xmlMensagem, pstrHeaderNZ)
        'KIDA - CBLC - 23/08/2008
        Case "LDL0015R2"
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
        Case Else
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemR2)
    End Select

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemR2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------
'Processamento das mensagens de Aviso e Informação
'---------------------------------------------------------------------------

Public Function GerenciarMensagemAvisoInformacao(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                                 ByRef pstrHeaderNZ As String) As Boolean

Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    strCodigoMensagem = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
        
    Select Case strCodigoMensagem
        Case "LDL0001"
            Call flProcessaMensagemLDL0001(xmlMensagem, pstrHeaderNZ)
        Case "LDL0007"
            Call flProcessaMensagemLDL0007(xmlMensagem, pstrHeaderNZ)
        Case "LTR0001"
            Call flProcessaMensagemLTR0001(xmlMensagem, pstrHeaderNZ)
        Case "LTR0007"
            Call flProcessaMensagemLTR0007(xmlMensagem, pstrHeaderNZ)
        Case "BMA0015"
            Call flProcessaMensagemBMA0015(xmlMensagem, pstrHeaderNZ)
        Case "BMA0002"
            Call flProcessaMensagemBMA0002(xmlMensagem, pstrHeaderNZ)
        Case "BMA0003"
            Call flProcessaMensagemBMA0003(xmlMensagem, pstrHeaderNZ)
        Case "BMA0013"
            Call flProcessaMensagemBMA0013(xmlMensagem, pstrHeaderNZ)
        Case "BMA0014"
            Call flProcessaMensagemBMA0014(xmlMensagem, pstrHeaderNZ)
        Case "CTP1002"
            Call flProcessaMensagemCTP1002(xmlMensagem, pstrHeaderNZ)
        Case "LDL1027"
            Call flProcessaMensagemLDL1027(xmlMensagem, pstrHeaderNZ)
        Case "LDL0028"
            Call flProcessaMensagemLDL0028(xmlMensagem, pstrHeaderNZ)
        Case "LDL1001"
            Call flProcessaMensagemLDL1001(xmlMensagem, pstrHeaderNZ)
        Case "BMC0005"
            Call flProcessaMensagemBMC0005(xmlMensagem, pstrHeaderNZ)
        Case "BMC0011"
            Call flProcessaMensagemBMC0011(xmlMensagem, pstrHeaderNZ)
        Case "BMC0010"
            Call flProcessaMensagemBMC0010(xmlMensagem, pstrHeaderNZ)
        Case "BMC0013"
            Call flProcessaMensagemBMC0013(xmlMensagem, pstrHeaderNZ)
        Case "BMC0101"
            Call flProcessaMensagemBMC0101(xmlMensagem, pstrHeaderNZ)
        Case "BMC0103"
            Call flProcessaMensagemBMC0103(xmlMensagem, pstrHeaderNZ)
        Case "BMC0112"
            Call flProcessaMensagemBMC0112(xmlMensagem, pstrHeaderNZ)
        Case "STR0011"
            Call flProcessaMensagemSTR0011(xmlMensagem)
        Case "LTR0003", "LTR0004"
            Call flProcessaMensagemLTR0003_LTR0004(xmlMensagem)
        Case "BMC0015"
            Call flProcessaMensagemBMC0015(xmlMensagem, pstrHeaderNZ)
        Case "CAM0003"
            Call flProcessaMensagemCAM0003(xmlMensagem, pstrHeaderNZ)
        Case Else
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
    End Select

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarMensagemAvisoInformacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMA0015 (BMA informa Lançamentos a serem efetuados no SELIC)
'   - Para cada Grupo de Titulo X Conta é gerado um registro no SLCC para conciliação.
'   - Veiculo Legal obtido pelo número da conta custodia .
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0015(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strNumeroControleCamara                 As String
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngStatusMensagemMae                    As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim lngSeqControleRepeticao                 As Long
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String

Dim strNomeTagRepetTitulo                   As String
Dim strNomeTagRepetConta                    As String

Dim xmlNodeTitulo                           As MSXML2.IXMLDOMNode
Dim xmlNodeConta                            As MSXML2.IXMLDOMNode

Dim strCtCstdiaAux                          As String
Dim strDtLiquid                             As String
Dim strIdentdTitSEL                         As String
Dim strDtVenc                               As String
Dim strCtCstdia                             As String
Dim strNumOpSEL                             As String
Dim strTpDeb_Cred                           As String
Dim vntQtd                                  As Variant
Dim vntPU                                   As Variant
Dim vntVlrFinanc                            As Variant
Dim vntCNPJContraparte                      As Variant
Dim strNomeVeicLega                         As String
Dim strSiglaSistemaVeicLega                 As String

On Error GoTo ErrorHandler
    
    '---------------------------------------------------------------------------------------
    'Somente as mensagens BMA0015 com SitResultLiqdBMA =  3 - definitivo ficam a Conciliar
    '---------------------------------------------------------------------------------------
    If Not xmlMensagem.selectSingleNode("//SitResultLiqdBMA") Is Nothing Then
        If Val("0" & xmlMensagem.selectSingleNode("//SitResultLiqdBMA").Text) <> enumSitResultLiqdBMA.definitivo Then
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
            Exit Function
        End If
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
    strCodigoVeiculoLegal = ""
    strDataRegistroMensagem = ""
    strSiglaSistema = udtHeaderMensagem.SiglaSistemaEnviouNZ
    
    strNomeTagRepetTitulo = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_BMA0015_Tit"
    strNomeTagRepetConta = ObterNomeTagRepeticao(strCodigoMensagemSPB, 2) '"Repet_BMA0015_Ct"
    
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlBMA").Text
    strDtLiquid = xmlMensagem.selectSingleNode("//DtLiquid").Text
    strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux)) 'xmlMensagem.selectSingleNode("//DtMovto").Text
    vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
   
    If fgObterCodigoVeiculoLegal("BMA00151", _
                                 vntCNPJContraparte, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
    
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vbNullString, _
                                xmlMensagem)

    End If

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        
    lngStatusMensagem = flObterStatusBMA0015(xmlMensagem, lngCodigoEmpresa)
        
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
       lngStatusMensagemMae = enumStatusMensagem.Aviso
    Else
       lngStatusMensagemMae = lngStatusMensagem
    End If
        
    'Salvar Mensagem Enviada Recebida Pai
    lngSeqControleRepeticao = 1
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumOpSEL, _
                                                            "", _
                                                            lngStatusMensagemMae, _
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
                                0, _
                                lngSeqControleRepeticao, _
                                strDtLiquid, _
                                0, _
                                "", _
                                0, _
                                "", _
                                0, _
                                "")
    
    
    

    'Repeticao Titulos
    For Each xmlNodeTitulo In xmlMensagem.selectNodes("//" & strNomeTagRepetTitulo & "/*")
        
        strCodigoVeiculoLegal = ""
        strSiglaSistemaVeicLega = ""
        lngTipoBackOffice = 0
        strNomeVeicLega = ""
       
        strIdentdTitSEL = xmlNodeTitulo.selectSingleNode("IdentdTitSEL").Text
        strDtVenc = xmlNodeTitulo.selectSingleNode("DtVenc").Text
        
        For Each xmlNodeConta In xmlNodeTitulo.selectNodes(strNomeTagRepetConta & "/*")
                        
            strMensagem = flMontaMensagemNodeBMA0015(strCodigoMensagemSPB, xmlMensagem, xmlNodeTitulo, xmlNodeConta, strNomeTagRepetTitulo, strNomeTagRepetConta)
                        
            lngSeqControleRepeticao = lngSeqControleRepeticao + 1
            
            strCtCstdia = xmlNodeConta.selectSingleNode("CtCstdia").Text
            strCtCstdiaAux = fgCompletaString(xmlNodeConta.selectSingleNode("CtCstdia").Text, "0", 9, True)
            strCtCstdiaAux = Left$(strCtCstdiaAux, 4)
            
            strNumOpSEL = xmlNodeConta.selectSingleNode("NumOpSEL").Text
            strTpDeb_Cred = xmlNodeConta.selectSingleNode("TpDeb_Cred").Text
            
            'KIDA - 18/02/2008
            If Not xmlNodeConta.selectSingleNode("QtdTit") Is Nothing Then
                vntQtd = xmlNodeConta.selectSingleNode("QtdTit").Text
            Else
                vntQtd = vbNullString
            End If
            
            If Not xmlNodeConta.selectSingleNode("PU") Is Nothing Then
                vntPU = xmlNodeConta.selectSingleNode("PU").Text
            Else
                vntPU = vbNullString
            End If
            
            If Not xmlNodeConta.selectSingleNode("VlrFinanc") Is Nothing Then
                vntVlrFinanc = xmlNodeConta.selectSingleNode("VlrFinanc").Text
            Else
                vntVlrFinanc = vbNullString
            End If
            
            If lngStatusMensagem = enumStatusMensagem.AConciliar Then
                strCodigoVeiculoLegal = ""
                strSiglaSistemaVeicLega = ""
                strNomeVeicLega = ""
                
                lngTipoBackOffice = flObterTipoBackOfficeContaBMA0015(strCtCstdia, _
                                                                      lngCodigoEmpresa, _
                                                                      strCodigoVeiculoLegal, _
                                                                      strSiglaSistemaVeicLega)
                
                'If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                             strCtCstdiaAux, _
                                             lngCodigoEmpresa, _
                                             strCodigoVeiculoLegal, _
                                             strSiglaSistemaVeicLega, _
                                             lngTipoBackOffice, _
                                             strNomeVeicLega) Then
                    
                'End If
            
            End If
            
            Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                             Trim(strNumeroControleIF), _
                                             0, _
                                             0, _
                                             strCodigoMensagemSPB, _
                                             lngTipoBackOffice, _
                                             lngCodigoEmpresa, _
                                             strNumOpSEL, _
                                             "", _
                                             lngStatusMensagem, _
                                             lngIndicadorEntradaManual, _
                                             "", _
                                             lngCodigoLocalLiquidacao, _
                                             strCodigoVeiculoLegal, _
                                             strSiglaSistemaVeicLega, _
                                             lngSeqControleRepeticao, _
                                             strNumeroControleCamara, _
                                             strDataRegistroMensagem)

                    
            'Salvar Dados Conciliacao
            Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                        strDataRegistroMensagem, _
                                        strDataOperacao, _
                                        0, _
                                        0, _
                                        strIdentdTitSEL, _
                                        vntPU, _
                                        vntQtd, _
                                        vntVlrFinanc, _
                                        lngSeqControleRepeticao, _
                                        strDtLiquid, _
                                        IIf(strTpDeb_Cred = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito), _
                                        strDtVenc, _
                                        vntCNPJContraparte, _
                                        "", _
                                        strCtCstdia, _
                                        "")
                                        
        Next
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------------
'Processamento das mensagesn LDL0001 (Câmara informa Resultado líquido de negociações)
'   - Para cada CNPJ não Liquidante é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0001(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagResultLiqd                    As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode

Dim strTpInf                                As String
Dim strDtLiquid                             As String
Dim vntVlrLanc                              As Variant
Dim intTpDeb_Cred                           As String
Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String
Dim lngTipoPagtoLDL                         As Long
Dim vntNumeroPartCamr                       As Variant

Dim xmlMensagemExistenteBMC                 As MSXML2.DOMDocument40
Dim lngStatusMensagemExistente              As Long

    On Error GoTo ErrorHandler
    
    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                   Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strSituacaoMensagemSPB = ""
    strDataRegistroMensagem = ""
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL0001_ResultLiqd"
    
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLDL").Text
    strTpInf = xmlMensagem.selectSingleNode("//TpInf").Text
    strDtLiquid = xmlMensagem.selectSingleNode("//DtLiquid").Text
    intTpDeb_Cred = IIf(xmlMensagem.selectSingleNode("//TpDeb_Cred").Text = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrLanc").Text
    
    If fgVlrXml_To_Decimal(vntVlrLanc) = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If
        
    If strTpInf = "D" Then
        lngStatusMensagem = flProcessaMensagemSegundaLDL0001(xmlMensagem, pstrHeaderNZ)
    Else
        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Or _
           lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC Or _
           lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP Then
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
            Exit Function
        Else
            lngStatusMensagem = flObterStatusLDL0001(xmlMensagem, lngCodigoEmpresa)
        End If
    End If
        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC And _
       strTpInf = "D" And _
       lngStatusMensagem = enumStatusMensagem.AConciliar Then
       
        Set xmlMensagemExistenteBMC = objMesgRecbEnviSPB.LerXMLMensagem(, , , _
                                                                        strCodigoMensagemSPB, _
                                                                        , , , _
                                                                        strDataOperacao, _
                                                                        strTpInf, _
                                                                        lngCodigoLocalLiquidacao)
        
        If xmlMensagemExistenteBMC.xml <> vbNullString Then
            lngStatusMensagemExistente = Val(xmlMensagemExistenteBMC.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
            
            If lngStatusMensagemExistente <> enumStatusMensagem.MensagemLiquidada And _
               lngStatusMensagemExistente <> enumStatusMensagem.AConciliar Then
                
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                
            Else
            
                If lngStatusMensagemExistente = enumStatusMensagem.AConciliar Then
                    With xmlMensagemExistenteBMC
                        Call objMesgRecbEnviSPB.AlterarStatus(.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              .selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              enumStatusMensagem.MensagemCancelada, _
                                                              vbNullString, , _
                                                              False, _
                                                              0, _
                                                              strCodigoMensagemSPB)
                    End With
                End If
                
            End If
        End If
        
    End If
    
    'Incluir Mensagem Mae
    strMensagem = xmlMensagem.xml
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
    lngSeqControleRepeticao = 1
    
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)

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
                                vntVlrLanc, _
                                lngSeqControleRepeticao, _
                                strDtLiquid, _
                                intTpDeb_Cred, _
                                "", _
                                0, _
                                "", _
                                0, _
                                strTpInf)
   
    If lngStatusMensagem <> enumStatusMensagem.AConciliar And _
       lngStatusMensagem <> enumStatusMensagem.MensagemInconsistente Then
        Set objConciliacao = Nothing
        Set objMesgRecbEnviSPB = Nothing
        Exit Function
    End If

    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodeResultLiqd.selectSingleNode("VlrResultLiqdNLiqdant").Text
        
        If Not xmlNodeResultLiqd.selectSingleNode("TpPgtoLDL") Is Nothing Then
            lngTipoPagtoLDL = xmlNodeResultLiqd.selectSingleNode("TpPgtoLDL").Text
        Else
            lngTipoPagtoLDL = 0
        End If
        
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntNumeroPartCamr = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
        Else
            vntNumeroPartCamr = 0
        End If
               
        If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice, _
                                     strNomeVeicLega) Then
                                     
            'lngStatusMensagem = enumStatusMensagem.AConciliar
            
            Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrResultLiqdNLiqdant, _
                                    xmlMensagem)
            
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        End If
           
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem, _
                                           0, _
                                           vntNumeroPartCamr)

        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    intTpDeb_Cred, _
                                    "", _
                                    0, _
                                    "", _
                                    0, _
                                    strTpInf)
                                    
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------------------------
'Processamento das mensagens LDL1001 (Margem liquidação prévia)
'   - Para cada CNPJ não Liquidante é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL1001(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagResultLiqd                    As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode

Dim strTpInf                                As String
Dim strDtLiquid                             As String
Dim vntVlrLanc                              As Variant
Dim intTpDeb_Cred                           As String
Dim vntCNPJEntRespons                       As Variant
Dim vntVlrCobrir                            As Variant
Dim strNomeVeicLega                         As String
Dim lngTipoPagtoLDL                         As Long
Dim vntNumeroPartCamr                       As Variant

On Error GoTo ErrorHandler
    
    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao <> enumLocalLiquidacao.BMD Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    If xmlMensagem.selectSingleNode("//TpTitlar").Text <> "PLD" And _
       xmlMensagem.selectSingleNode("//TpTitlar").Text <> "CO" Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strSituacaoMensagemSPB = ""
    strDataRegistroMensagem = ""
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL1001_ColatCobrir"
    
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLDL").Text
    strTpInf = "P"
    lngStatusMensagem = enumStatusMensagem.AConciliar
    intTpDeb_Cred = enumTipoDebitoCredito.Debito
    strDtLiquid = xmlMensagem.selectSingleNode("//DtMovto").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrTotCobrir").Text
    
    If fgVlrXml_To_Decimal(vntVlrLanc) = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        
    'Incluir Mensagem Mae
    strMensagem = xmlMensagem.xml
    vntCNPJEntRespons = fgObterCNPJEmpresa(lngCodigoEmpresa)
    lngSeqControleRepeticao = 1
    
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJEntRespons, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)

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
                                vntVlrLanc, _
                                lngSeqControleRepeticao, _
                                strDtLiquid, _
                                intTpDeb_Cred, _
                                "", _
                                0, _
                                "", _
                                0, _
                                strTpInf)
   
    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        vntVlrCobrir = xmlNodeResultLiqd.selectSingleNode("VlrCobrir").Text
        
        If Not xmlNodeResultLiqd.selectSingleNode("CodTitlar") Is Nothing Then
            If IsNumeric(xmlNodeResultLiqd.selectSingleNode("CodTitlar").Text) Then
                'KIDA - 24/08/2009
                vntNumeroPartCamr = CLng(Left$(Format$(xmlNodeResultLiqd.selectSingleNode("CodTitlar").Text, "000000000"), 3))
            Else
                vntNumeroPartCamr = 0
            End If
        Else
            vntNumeroPartCamr = 0
        End If
               
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem, _
                                           0, _
                                           vntNumeroPartCamr)

        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrCobrir, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    intTpDeb_Cred, _
                                    "", _
                                    0, _
                                    "", _
                                    0, _
                                    strTpInf)
                                    
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL1001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------
'Processamento das mensagesn BMA0002 (BMA informa Operação realizada)
'   - Para cada Conta x Titulo é gerado um registro para concilição de registro.
'---------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0002(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagRepetConta                    As String
Dim xmlNodeConta                            As MSXML2.IXMLDOMNode

Dim strDtLiquid                             As String
Dim strIdentdTitSEL                         As String
Dim strDtVenc                               As String
Dim lngCtCstdia                             As Long
Dim strNumOpSEL                             As String
Dim intTpDeb_Cred                           As Integer
Dim vntQtd                                  As Variant
Dim vntPU                                   As Variant
Dim vntVlrFinanc                            As Variant
Dim vntCNPJContraparte                      As Variant
Dim vntTaxaNegociada                        As Variant
Dim vntCodTitularCutd                       As Variant
Dim lngTipoNegoBMA                          As Long
Dim strNomeVeicLega                         As String
Dim strNumOpSELMigr                         As String
Dim strDataRetornoOperacao                  As String
Dim vntPrazoDiasRetornoOperacao             As Variant
Dim vntValorRetornoOperacao                 As Variant

On Error GoTo ErrorHandler
    
    '---------------------------------------------------------------------------------------
    'Somente as mensagens BMA0002 com SitResultLiqdBMA =  3 - definitivo ficam a Conciliar
    '---------------------------------------------------------------------------------------
    
    If Not xmlMensagem.selectSingleNode("//SitOpBMA") Is Nothing Then
        If Val("0" & xmlMensagem.selectSingleNode("//SitOpBMA").Text) <> enumSitOpBMA.EspecificacaoTotal Then
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
            Exit Function
        End If
    End If
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
    lngTipoBackOffice = 0
    strSituacaoMensagemSPB = ""
    strDataRegistroMensagem = ""
    vntCodTitularCutd = 0
   
    If Not xmlMensagem.selectSingleNode("//DtLiquid") Is Nothing Then
        strDtLiquid = xmlMensagem.selectSingleNode("//DtLiquid").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//VlrFinanc") Is Nothing Then
        vntVlrFinanc = xmlMensagem.selectSingleNode("//VlrFinanc").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumOpNegcBMA") Is Nothing Then
        strNumOpSEL = xmlMensagem.selectSingleNode("//NumOpNegcBMA").Text
    End If
    
    strDataOperacao = fgSelectSingleNode(xmlMensagem, "//DtMovto").Text
    
    If Not xmlMensagem.selectSingleNode("//CNPJEntRespons") Is Nothing Then
        vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlBMA") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlBMA").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//TpNegcBMA") Is Nothing Then
        lngTipoNegoBMA = xmlMensagem.selectSingleNode("//TpNegcBMA").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//TaxNegcd") Is Nothing Then
        vntTaxaNegociada = xmlMensagem.selectSingleNode("//TaxNegcd").Text
    End If
        
    If Not xmlMensagem.selectSingleNode("//NumOpSELMigr") Is Nothing Then
        strNumOpSELMigr = xmlMensagem.selectSingleNode("//NumOpSELMigr").Text
    Else
        strNumOpSELMigr = vbNullString
    End If
        
    If Not xmlMensagem.selectSingleNode("//VlrFinancRet") Is Nothing Then
        vntValorRetornoOperacao = xmlMensagem.selectSingleNode("//VlrFinancRet").Text
    Else
        vntValorRetornoOperacao = 0
    End If
    
    If Not xmlMensagem.selectSingleNode("//PrzRetOp") Is Nothing Then
        vntPrazoDiasRetornoOperacao = xmlMensagem.selectSingleNode("//PrzRetOp").Text
    Else
        vntPrazoDiasRetornoOperacao = 0
    End If

    If Not xmlMensagem.selectSingleNode("//DtOpRet") Is Nothing Then
        strDataRetornoOperacao = xmlMensagem.selectSingleNode("//DtOpRet").Text
    Else
        strDataRetornoOperacao = ""
    End If
        
    Select Case UCase(xmlMensagem.selectSingleNode("//NatuOp").Text)
        Case NatuOp_Compra, NatuOp_Recompra
            intTpDeb_Cred = enumTipoDebitoCredito.Credito
        Case NatuOp_Venda, NatuOp_Revenda
            intTpDeb_Cred = enumTipoDebitoCredito.Debito
    End Select
                
                
    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(xmlMensagem, "//CodMsg").Text, _
                                 vntCNPJContraparte, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
        
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntVlrFinanc, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If

    If strNumOpSELMigr <> "" Then
        If lngTipoNegoBMA = enumTipoNegociacaoBMA.MigracaoIdaCamara Or _
            lngTipoNegoBMA = enumTipoNegociacaoBMA.MigracaoVoltaCamara Then
            lngStatusMensagem = enumStatusMensagem.Aviso
        End If
    End If

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    'Incluir Mensagem Pai
    lngSeqControleRepeticao = 0

    strNomeTagRepetConta = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) ' "Repet_BMA0002_CtTit"
            
    strMensagem = xmlMensagem.xml
    
    'Repeticao Titulos
    For Each xmlNodeConta In xmlMensagem.selectNodes("//" & strNomeTagRepetConta & "/*")
                
        'strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeConta)
                
        If Not xmlNodeConta.selectSingleNode("IdentdTitSEL") Is Nothing Then
            strIdentdTitSEL = xmlNodeConta.selectSingleNode("IdentdTitSEL").Text
        Else
            strIdentdTitSEL = ""
        End If
        
        If Not xmlNodeConta.selectSingleNode("DtVenc") Is Nothing Then
            strDtVenc = xmlNodeConta.selectSingleNode("DtVenc").Text
        Else
            strDtVenc = ""
        End If
        
        If Not xmlNodeConta.selectSingleNode("CtCstdia") Is Nothing Then
            lngCtCstdia = Val(xmlNodeConta.selectSingleNode("CtCstdia").Text)
        Else
            lngCtCstdia = 0
        End If
        
        If Not xmlNodeConta.selectSingleNode("QtdTit") Is Nothing Then
            vntQtd = xmlNodeConta.selectSingleNode("QtdTit").Text
        Else
            vntQtd = ""
        End If
        
        If Not xmlNodeConta.selectSingleNode("PU") Is Nothing Then
            vntPU = xmlNodeConta.selectSingleNode("PU").Text
        Else
            vntPU = ""
        End If
        
        If Not xmlNodeConta.selectSingleNode("CodTitlarCstdte") Is Nothing Then
            vntCodTitularCutd = Val(xmlNodeConta.selectSingleNode("CodTitlarCstdte").Text)
        Else
            vntCodTitularCutd = "0"
        End If
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                                Trim(strNumeroControleIF), _
                                                                0, _
                                                                0, _
                                                                strCodigoMensagemSPB, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                strNumOpSEL, _
                                                                "", _
                                                                lngStatusMensagem, _
                                                                lngIndicadorEntradaManual, _
                                                                "", _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                lngSeqControleRepeticao, _
                                                                strNumeroControleCamara, _
                                                                "")
                
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    CLng("0" & strNumOpSELMigr), _
                                    strIdentdTitSEL, _
                                    vntPU, _
                                    vntQtd, _
                                    vntVlrFinanc, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    intTpDeb_Cred, _
                                    strDtVenc, _
                                    vntCNPJContraparte, _
                                    "", _
                                    lngCtCstdia, _
                                    "", _
                                    vntCodTitularCutd, _
                                    lngTipoNegoBMA, _
                                    vntTaxaNegociada, _
                                    , , , , , , _
                                    strDataRetornoOperacao, vntPrazoDiasRetornoOperacao, vntValorRetornoOperacao _
                                    )
        
        
        If lngStatusMensagem = enumStatusMensagem.AConciliar Then
            If strNumOpSELMigr = "" Then
                'Chamada para a Rotina de Cociliacao automatica
                Call flConciliarMensagem(strCodigoMensagemSPB, _
                                         Trim(strNumeroControleIF), _
                                         lngSeqControleRepeticao, _
                                         strDataRegistroMensagem)
            End If
        End If
    Next
   
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0002 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'---------------------------------------------------------------------------------------
'Processamento das mensagens CTP1002 (CETIP informa operação pendente de lançamento)
'   - Conciliação de Operações com Ordem de Pagamento
'---------------------------------------------------------------------------------------

Private Function flProcessaMensagemCTP1002(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

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

Dim vntIdentPartCamr                        As Variant
Dim strNumComando                           As String
Dim strTpCompra_Venda                       As String
Dim strIdentdTitCTP                         As String
Dim vntQtdCTP                               As Variant
Dim intTpDeb_Cred                           As Integer
Dim strSubTpAtv                             As String
Dim lngModLiquid                            As Long
Dim strCodOpCTP                             As String
Dim vntPU                                   As Variant
Dim vntValorFinanc                          As Variant
Dim vntIdentdCtrapartCamr                   As Variant

On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
    strCodigoMensagemSPB = Trim(fgSelectSingleNode(xmlMensagem, "//CodMsg").Text)
    lngTipoBackOffice = 0
    strCodigoVeiculoLegal = ""
    strSituacaoMensagemSPB = ""
   
    strNumeroControleCamara = fgSelectSingleNode(xmlMensagem, "//NumCtrlCTP").Text
    strDataOperacao = fgSelectSingleNode(xmlMensagem, "//DtMovto").Text
    strTpCompra_Venda = fgSelectSingleNode(xmlMensagem, "//TpCompra_Venda").Text
    strIdentdTitCTP = fgSelectSingleNode(xmlMensagem, "//IdentdTitCTP").Text
    vntQtdCTP = fgSelectSingleNode(xmlMensagem, "//QtdCTP").Text
    strNumComando = fgSelectSingleNode(xmlMensagem, "//NumOpPart").Text
    strCodOpCTP = fgSelectSingleNode(xmlMensagem, "//CodOpCTP").Text
    lngModLiquid = fgSelectSingleNode(xmlMensagem, "//ModLiquid").Text
    
    Select Case lngModLiquid
    
        Case enumTipoLiquidacao.SemModalidade
            
            lngModLiquid = enumTipoLiquidacao.SemModalidade
        
        Case enumTipoLiquidacao.Bilateral, _
             enumTipoLiquidacao.BilateralBT, _
             enumTipoLiquidacao.BilateralSTR
            
            lngModLiquid = enumTipoLiquidacao.Bilateral
            
        Case enumTipoLiquidacao.Bruta, _
             enumTipoLiquidacao.BrutaBT, _
             enumTipoLiquidacao.BrutaBTAUT, _
             enumTipoLiquidacao.BrutaSTR
        
            lngModLiquid = enumTipoLiquidacao.Bruta
        
        Case enumTipoLiquidacao.Multilateral
            
            lngModLiquid = enumTipoLiquidacao.Multilateral
                
        Case Else
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
            Exit Function
    End Select
    
    If Not xmlMensagem.selectSingleNode("//SubTpAtv") Is Nothing Then
        strSubTpAtv = xmlMensagem.selectSingleNode("//SubTpAtv").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//PUNegc") Is Nothing Then
        vntPU = xmlMensagem.selectSingleNode("//PUNegc").Text
    Else
        vntPU = 0
    End If
    
    If Not xmlMensagem.selectSingleNode("//VlrFinanc") Is Nothing Then
        vntValorFinanc = xmlMensagem.selectSingleNode("//VlrFinanc").Text
    Else
        vntValorFinanc = 0
    End If

    If Not xmlMensagem.selectSingleNode("//IdentdCtrapartCamr") Is Nothing Then
        vntIdentdCtrapartCamr = xmlMensagem.selectSingleNode("//IdentdCtrapartCamr").Text
    Else
        vntIdentdCtrapartCamr = ""
    End If


    vntIdentPartCamr = fgSelectSingleNode(xmlMensagem, "//IdentdPartCamr").Text
    
    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(xmlMensagem, "//CodMsg").Text, _
                                 vntIdentPartCamr, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice) Then
        
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
    
    If strTpCompra_Venda = enumTipoCompraVendaCetip.Compra Then
        intTpDeb_Cred = enumTipoDebitoCredito.Credito
    ElseIf strTpCompra_Venda = enumTipoCompraVendaCetip.VendaCarteiraPropria Then
        intTpDeb_Cred = enumTipoDebitoCredito.Debito
    End If
            
    lngSeqControleRepeticao = 1
            
    'Salvar Mensagem Enviada Recebida
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumComando, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara, _
                                                            "")
    
    Set objMesgRecbEnviSPB = Nothing
            
    'Salvar Dados Conciliacao
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                strIdentdTitCTP, _
                                vntPU, _
                                vntQtdCTP, _
                                vntValorFinanc, _
                                lngSeqControleRepeticao, _
                                "", _
                                intTpDeb_Cred, _
                                "", _
                                0, _
                                "", _
                                CLng(vntIdentPartCamr), _
                                "", _
                                0, _
                                0, _
                                "", _
                                "", _
                                lngModLiquid, _
                                strSubTpAtv, _
                                strCodOpCTP, 0, 0, "", 0, 0, "")
                                
    Set objConciliacao = Nothing
    
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        'Vertifica se concilia automatico
        Call flConciliarMensagem(strCodigoMensagemSPB, _
                                 strNumeroControleIF, _
                                 lngSeqControleRepeticao, _
                                 strDataRegistroMensagem)
    End If
    
    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCTP1002 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LDL0005R2 (Transferência Resultado Líquido Negociação LDL)
'   - Para cada CNPJ não Liquidante é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0005R2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagResultLiqd                    As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode

Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String
Dim vntVlrLanc                              As Variant
Dim vntNumeroPartCamr                       As Variant

On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

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
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL0005R2_ResultLiqd"
        
    If Not xmlMensagem.selectSingleNode("//NumCtrlLDLOr") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text
    Else
        strNumeroControleCamara = vbNullString
    End If
    
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrLanc").Text
        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    'Incluir Mensagem Pai
    
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
        lngStatusMensagem = flObterStatusLDL0005R2CBLC(xmlMensagem, lngCodigoEmpresa)
    'BMF
    ElseIf lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD Then
        lngStatusMensagem = flObterStatusLDL0005R2BMF(xmlMensagem, lngCodigoEmpresa)
    Else
        lngStatusMensagem = flObterStatusLDL0005R2_LDL0009R2(xmlMensagem, lngCodigoEmpresa)
    End If
    
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
    
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        

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
                                                            "")

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntVlrLanc, _
                                lngSeqControleRepeticao, _
                                "", _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                0, _
                                "", _
                                0, _
                                "")

    
    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodeResultLiqd.selectSingleNode("VlrResultLiqdNLiqdant").Text
           
        'KIDA - CBLC - 24/09/2008
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntNumeroPartCamr = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
        Else
            vntNumeroPartCamr = 0
        End If
           
        If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice) Then
            
            'lngStatusMensagem = enumStatusMensagem.AConciliar
            
            Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrResultLiqdNLiqdant, _
                                    xmlMensagem)
            
        'Else
            'lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        End If
           
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem, _
                                           0, _
                                           vntNumeroPartCamr) 'KIDA - CBLC
        
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    "", _
                                    enumTipoDebitoCredito.Credito, _
                                    "", _
                                    vntCNPJNLiqdant, _
                                    "", _
                                    0, _
                                    "")
                                    
    Next

    Exit Function
ErrorHandler:
    
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0005R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LTR0005R2 (Transferência Resultado Negociação LTR)
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0005R2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim strNomeVeicLega                         As String
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
Dim strAgendamento                          As String
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim lngSeqControleRepeticao                 As Long

Dim vntCNPJNLiqdant                         As Variant
Dim strNumOpLTR                             As String
Dim vntVlrLanc                              As Variant

Dim lngTipoLiquidacao                       As Long
Dim vntIdentificadorPartCamara              As String

On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strAgendamento = ""
    
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLTR").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntCNPJNLiqdant = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrLanc").Text
    
    If Not xmlMensagem.selectSingleNode("//NumOpLTR") Is Nothing Then
        strNumOpLTR = xmlMensagem.selectSingleNode("//NumOpLTR").Text
        strNumOpLTR = UCase$(strNumOpLTR)
        lngTipoLiquidacao = enumTipoLiquidacao.Bruta
    Else
        strNumOpLTR = vbNullString
        lngTipoLiquidacao = enumTipoLiquidacao.Bilateral
    End If
       
    If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
        vntIdentificadorPartCamara = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
    End If
           
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega, _
                                 vntIdentificadorPartCamara) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntVlrLanc, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    
    End If

    lngSeqControleRepeticao = 1

    'Salvar Mensagem Enviada Recebida
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumOpLTR, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara, _
                                                            , , _
                                                            vntIdentificadorPartCamara)

    Set objMesgRecbEnviSPB = Nothing

    'Salvar Dados Conciliacao
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntVlrLanc, _
                                lngSeqControleRepeticao, _
                                "", _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                vntCNPJNLiqdant, _
                                "", _
                                0, _
                                "", , , , , _
                                lngTipoLiquidacao)

    Set objConciliacao = Nothing

    'Verifica se concilia automatico
    Call flConciliarMensagem(strCodigoMensagemSPB, _
                             strNumeroControleIF, _
                             lngSeqControleRepeticao, _
                             strDataRegistroMensagem)

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0005R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LTR0003R2
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0003R2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao

Dim xmlLeitura                              As MSXML2.DOMDocument40

Dim lngCodigoLocalLiquidacao                As Long
Dim lngTipoBackOffice                       As Long
Dim lngStatusMensagem                       As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String
Dim strDataRegistroMensagem                 As String

    On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    flProcessaMensagemLTR0003R2 = True
    
    xmlMensagem.selectSingleNode("//CodMsg").Text = Trim$(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, xmlMensagem.selectSingleNode("//CodMsg").Text)
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Call fgAppendNode(xmlLeitura, vbNullString, "Repeat_Filtros", vbNullString)
    
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_NumeroControleLTR", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_NumeroControleLTR", "NumeroControleLTR", xmlMensagem.selectSingleNode("//NumCtrlLTROr").Text)
    
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_Data", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(xmlMensagem.selectSingleNode("//DtMovto").Text))
    
    Call fgAppendNode(xmlLeitura, "Repeat_Filtros", "Grupo_Status", vbNullString)
    Call fgAppendNode(xmlLeitura, "Grupo_Status", "Status", enumStatusOperacao.Registrada)
    Call fgAppendNode(xmlLeitura, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica)
    
    Call xmlLeitura.loadXML(objOperacao.ObterDetalheOperacao(xmlLeitura, True))
    
    If xmlLeitura.xml = vbNullString Then
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        lngTipoBackOffice = 0
        strCodigoVeiculoLegal = vbNullString
        strSiglaSistemaVeicLega = vbNullString
        strNomeVeicLega = vbNullString
    
        If pstrHeaderNZ = vbNullString Then
            flProcessaMensagemLTR0003R2 = False
            Exit Function
        End If
        
    Else
        lngStatusMensagem = enumStatusMensagem.AConciliar
        lngTipoBackOffice = xmlLeitura.selectSingleNode("//TP_BKOF").Text
        strCodigoVeiculoLegal = xmlLeitura.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistemaVeicLega = xmlLeitura.selectSingleNode("//SG_SIST").Text
        strNomeVeicLega = xmlLeitura.selectSingleNode("//NO_VEIC_LEGA").Text
    
        Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                xmlMensagem.selectSingleNode("//VlrLanc").Text, _
                                xmlMensagem)
    End If

    If pstrHeaderNZ <> vbNullString Then
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                                Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                                                0, _
                                                                0, _
                                                                xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                                lngTipoBackOffice, _
                                                                udtHeaderMensagem.CodigoEmpresa, _
                                                                Right$(xmlMensagem.selectSingleNode("//NumCtrlSTR").Text, 10), _
                                                                vbNullString, _
                                                                lngStatusMensagem, _
                                                                enumIndicadorSimNao.Nao, _
                                                                vbNullString, _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                1, _
                                                                xmlMensagem.selectSingleNode("//NumCtrlLTROr").Text)
    
        Call objConciliacao.Incluir(Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                    strDataRegistroMensagem, _
                                    xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                    0, _
                                    0, _
                                    vbNullString, _
                                    0, _
                                    0, _
                                    xmlMensagem.selectSingleNode("//VlrLanc").Text, _
                                    1, _
                                    vbNullString, _
                                    enumTipoDebitoCredito.Credito)
    
    Else
        
        Call xmlMensagem.loadXML(objMesgRecbEnviSPB.LerXMLMensagem(, _
                                                                   , _
                                                                   , _
                                                                   "LTR0003R2", _
                                                                   xmlMensagem.selectSingleNode("//NumCtrlLTROr").Text).xml)
        
        If lngStatusMensagem <> Val(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text) Then
            
            Call objMesgRecbEnviSPB.Alterar(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                            xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                            lngStatusMensagem, _
                                            , , , , _
                                            False, _
                                            , _
                                            strCodigoVeiculoLegal, _
                                            strSiglaSistemaVeicLega, _
                                            lngTipoBackOffice)
                                        
        End If
        
    End If

    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        
        Call flConciliarMensagem(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 udtHeaderMensagem.ControleRemessaNZ, _
                                 1, _
                                 strDataRegistroMensagem)
    
    End If

    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set xmlLeitura = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set xmlLeitura = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0003R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0005 (BMC informa Situação da proposta de operação interbancária de câmbio)
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0005(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlLeitura                              As MSXML2.DOMDocument40

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim lngSeqControleRepeticao                 As Long

Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String

Dim vntCNPJNLiqdant                         As Variant
Dim intIndDebitoCredito                     As Integer
Dim strCodigoPraca                          As String
Dim strNumeroComando                        As String

    On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlBMC") Is Nothing Then
        Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlLeitura = objMesgRecbEnviSPB.LerXMLMensagem(, , , "BMC0010", _
                                                                 xmlMensagem.selectSingleNode("//NumCtrlBMC").Text)
        
        If xmlLeitura.xml <> vbNullString Then
            Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao, xmlLeitura.selectSingleNode("//NU_CTRL_IF").Text)
            Set objMesgRecbEnviSPB = Nothing
            Set xmlLeitura = Nothing
            Exit Function
        End If
    End If

    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar

        Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                0, _
                                xmlMensagem)

    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
           
    If Not xmlMensagem.selectSingleNode("//SitLancBMC") Is Nothing Then
        If Val(xmlMensagem.selectSingleNode("//SitLancBMC").Text) = 5 Then
            lngStatusMensagem = enumStatusMensagem.MensagemRejeitada
        End If
    End If
    
    If Not xmlMensagem.selectSingleNode("//RegOpCaml") Is Nothing Then
        If IsNumeric(xmlMensagem.selectSingleNode("//RegOpCaml").Text) Then
            xmlMensagem.selectSingleNode("//RegOpCaml").Text = Val(xmlMensagem.selectSingleNode("//RegOpCaml").Text)
            strNumeroComando = Val(xmlMensagem.selectSingleNode("//RegOpCaml").Text)
        End If
    End If
    
    lngSeqControleRepeticao = 1
      
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumeroComando, _
                                                            vbNullString, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao)
    Set objMesgRecbEnviSPB = Nothing

    If xmlMensagem.selectSingleNode("//TpOpCAM").Text = "C" Then
        intIndDebitoCredito = enumTipoDebitoCredito.Credito
    Else
        intIndDebitoCredito = enumTipoDebitoCredito.Debito
    End If

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                0, _
                                lngSeqControleRepeticao, , _
                                intIndDebitoCredito, _
                                vbNullString, _
                                vntCNPJNLiqdant, _
                                , , , , , , , , , , , , , , , , , , , , _
                                strNumeroComando, , , , , , _
                                xmlMensagem.selectSingleNode("//ChACAM").Text)
                                
    Set objConciliacao = Nothing
    
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        'Verifica se concilia automatico
        Call flConciliarMensagem(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 strNumeroControleIF, _
                                 lngSeqControleRepeticao, _
                                 strDataRegistroMensagem, _
                                 strNomeVeicLega)
    End If

    Exit Function
ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0005 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0011
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0011(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objRegra                                As A8LQS.clsRegra
Dim xmlChaveConcilia                        As MSXML2.DOMDocument40
Dim blnGrupoUsuario                         As Boolean

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim lngSeqControleRepeticao                 As Long

Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String

Dim vntCNPJNLiqdant                         As Variant
Dim intIndDebitoCredito                     As Integer
Dim strCodigoPraca                          As String

Dim blnGerarAlerta                          As Boolean

    On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objRegra = CreateObject("A8LQS.clsRegra")
    Set xmlChaveConcilia = CreateObject("MSXML2.DOMDocument.4.0")
    
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
'        Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
'                                lngTipoBackOffice, _
'                                strNomeVeicLega, _
'                                xmlMensagem.selectSingleNode("//VlrMN").Text, _
'                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
           
    lngSeqControleRepeticao = 1
    strNumeroControleIF = Trim$(strNumeroControleIF)
      
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            strNumeroControleIF, _
                                                            0, _
                                                            0, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            xmlMensagem.selectSingleNode("//NumCtrlBMC").Text, _
                                                            , , , _
                                                            xmlMensagem.selectSingleNode("//DtHrBMC").Text)

    If xmlMensagem.selectSingleNode("//TpOpCAM").Text = "C" Then
        intIndDebitoCredito = enumTipoDebitoCredito.Credito
    Else
        intIndDebitoCredito = enumTipoDebitoCredito.Debito
    End If

    Call objConciliacao.Incluir(strNumeroControleIF, _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                xmlMensagem.selectSingleNode("//VlrMN").Text, _
                                lngSeqControleRepeticao, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                intIndDebitoCredito, _
                                , _
                                vntCNPJNLiqdant, _
                                , , , , , _
                                xmlMensagem.selectSingleNode("//TaxCam").Text, _
                                , , , , , , , , , , , _
                                xmlMensagem.selectSingleNode("//CodMoeda").Text, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                xmlMensagem.selectSingleNode("//CodPcaIF").Text, _
                                , _
                                xmlMensagem.selectSingleNode("//VlrME").Text, _
                                , _
                                xmlMensagem.selectSingleNode("//CodSISBACENCortr").Text, _
                                xmlMensagem.selectSingleNode("//CanOpInterbanc").Text)
                                
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        Call fgAppendNode(xmlChaveConcilia, vbNullString, "Grupo_Mensagem", vbNullString)
        Call fgAppendNode(xmlChaveConcilia, "Grupo_Mensagem", "CO_EMPR", lngCodigoEmpresa)
        Call fgAppendNode(xmlChaveConcilia, "Grupo_Mensagem", "NU_CTRL_IF", strNumeroControleIF)
        Call fgAppendNode(xmlChaveConcilia, "Grupo_Mensagem", "DH_REGT_MESG_SPB", strDataRegistroMensagem)
        Call fgAppendNode(xmlChaveConcilia, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", lngSeqControleRepeticao)

        blnGerarAlerta = False
        
        Call objConciliacao.ConciliarRodaDolarPronto(enumAcaoConciliacao.BOConcordar, _
                                                     xmlChaveConcilia, _
                                                     True, _
                                                     blnGerarAlerta)
    
        If blnGerarAlerta Then
            Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    xmlMensagem.selectSingleNode("//VlrMN").Text, _
                                    xmlMensagem)
        End If
        
    End If
    
    Set objRegra = Nothing
    Set xmlChaveConcilia = Nothing
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function

ErrorHandler:
    Set objRegra = Nothing
    Set xmlChaveConcilia = Nothing
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0011 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0015 (BMC requisita Registro de Operação Interbancária de Câmbio)
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0015(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlDetOperacao                          As MSXML2.DOMDocument40
Dim xmlDetMensagem                          As MSXML2.DOMDocument40

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String
Dim strChACAM                               As String
Dim strCodigoPraca                          As String

Dim lngSeqControleRepeticao                 As Long
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long

Dim vntNrIdefNegoBMC                        As Variant
Dim vntCNPJNLiqdant                         As Variant

Dim intIndDebitoCredito                     As Integer

    On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)

    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then

        lngStatusMensagem = enumStatusMensagem.AConciliar
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If

    lngSeqControleRepeticao = 1
    
    vntNrIdefNegoBMC = Val(xmlMensagem.selectSingleNode("//IdentdNegBMC").Text)
    xmlMensagem.selectSingleNode("//IdentdNegBMC").Text = vntNrIdefNegoBMC
    
    strChACAM = xmlMensagem.selectSingleNode("//ChACAM").Text
    
    'Buscar Dados da Operacao
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If xmlMensagem.xml <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(, , , , , , , , , vntNrIdefNegoBMC)
        If xmlOperacao.xml = vbNullString Then
            Set xmlOperacao = objOperacao.LerXMLOperacao(, , , , , , , , , , strChACAM)
        End If
    End If

    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            xmlMensagem.selectSingleNode("//NumCtrlBMC").Text)

    'Compra
    If xmlMensagem.selectSingleNode("//TpOpCAM").Text = "C" Then
        intIndDebitoCredito = enumTipoDebitoCredito.Credito
    'Venda
    Else
        intIndDebitoCredito = enumTipoDebitoCredito.Debito
    End If

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                xmlMensagem.selectSingleNode("//VlrMN").Text, _
                                lngSeqControleRepeticao, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                intIndDebitoCredito, _
                                vbNullString, _
                                vntCNPJNLiqdant, _
                                , , , , , _
                                xmlMensagem.selectSingleNode("//TaxCam").Text, _
                                , , , , , , , , , , , _
                                xmlMensagem.selectSingleNode("//CodMoeda").Text, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                strCodigoPraca, _
                                , _
                                xmlMensagem.selectSingleNode("//VlrME").Text, , , , _
                                vntNrIdefNegoBMC, _
                                strChACAM)

    'Monta Operacao para Conciliacao
    If xmlOperacao.xml <> vbNullString Then
        If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro Then
    
            Set xmlDetOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlDetOperacao, "", "Repeat_Conciliacao", "")
            Call fgAppendNode(xmlDetOperacao, "Repeat_Conciliacao", "Grupo_Operacao", "")
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "CO_EMPR", xmlOperacao.selectSingleNode("//CO_EMPR").Text)
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "NU_SEQU_OPER_ATIV", xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "TP_OPER", xmlOperacao.selectSingleNode("//TP_OPER").Text)
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "DH_ULTI_ATLZ", xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text)
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "LocalLiquidacao", enumLocalLiquidacao.BMC)
            Call fgAppendNode(xmlDetOperacao, "Grupo_Operacao", "TipoConfirmacao", enumTipoConfirmacao.Operacao)
        
        
            'Busca Detalhes da Mensagem Para Conciliacao
            Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
            Set xmlMesgRecbEnviSPB = objMesgRecbEnviSPB.LerXMLMensagem(Trim$(strNumeroControleIF), _
                                                                       lngSeqControleRepeticao, _
                                                                       , _
                                                                       xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                                       , _
                                                                       strDataRegistroMensagem)
            'Monta Detalhes da Mensagem Para Conciliacao
            Set xmlDetMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlDetMensagem, "", "Repeat_Conciliacao", "")
            Call fgAppendNode(xmlDetMensagem, "Repeat_Conciliacao", "Grupo_Mensagem", "")
            Call fgAppendNode(xmlDetMensagem, "Grupo_Mensagem", "NU_CTRL_IF", strNumeroControleIF)
            Call fgAppendNode(xmlDetMensagem, "Grupo_Mensagem", "DH_REGT_MESG_SPB", xmlMesgRecbEnviSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text)
            Call fgAppendNode(xmlDetMensagem, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            Call fgAppendNode(xmlDetMensagem, "Grupo_Mensagem", "DH_ULTI_ATLZ", xmlMesgRecbEnviSPB.selectSingleNode("//DH_ULTI_ATLZ").Text)
            Call fgAppendNode(xmlDetMensagem, "Grupo_Mensagem", "TP_LIQU_OPER_ATIV", xmlMesgRecbEnviSPB.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
        
        
            If lngStatusMensagem = enumStatusMensagem.AConciliar Then
                'Verifica se concilia automatico
                If xmlDetOperacao.xml <> vbNullString _
                And xmlDetMensagem.xml <> vbNullString Then
                    'Processo de Conciliacao automatica
                    If objConciliacao.ConciliarRegistroOperacao(enumAcaoConciliacao.BOConcordar, _
                                                                  xmlDetOperacao, _
                                                                  xmlDetMensagem, _
                                                                  True) <> "" Then
                        
                        Call flGeraAlertaCamara("BMC0015DOPER", _
                                                lngTipoBackOffice, _
                                                strNomeVeicLega, _
                                                xmlMensagem.selectSingleNode("//VlrMN").Text, _
                                                xmlMensagem)
                    
                    End If
                End If
            End If
        End If
    Else
        Call flGeraAlertaCamara("BMC0015SOPER", _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                xmlMensagem.selectSingleNode("//VlrMN").Text, _
                                xmlMensagem)
    
    End If

    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing

    Exit Function
ErrorHandler:
    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens CAM0003 (CAM avisa Operação Anulada)
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemCAM0003(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlBMC0015                              As MSXML2.DOMDocument40
Dim xmlCAM0002                              As MSXML2.DOMDocument40

Dim lngSeqControleRepeticao                 As Long
Dim lngSequencialConciliacaoOperacao        As Long
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long

Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String
Dim strChACAM                               As String
Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim strBMC0015                              As String

Dim vntCNPJNLiqdant                         As Variant
Dim vntNrSeqOperacao                        As Variant
Dim vntNrConcilOperacao                    As Variant

    On Error GoTo ErrorHandler

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
    strChACAM = xmlMensagem.selectSingleNode("//ChACAM").Text
    lngSeqControleRepeticao = 1
    vntNrConcilOperacao = 0
    
    'Busca Dados da Operacao
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If xmlMensagem.xml <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(, , , , , , , , , , strChACAM)
    End If
    
    If xmlOperacao.xml <> vbNullString Then
        
        vntNrSeqOperacao = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        
        'Cancela Mensagens somente se a mensagem nao for cancelada
        If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarAceite Then
            
            lngStatusMensagem = enumStatusMensagem.MensagemCancelada
            
            'Atualiza status da mensagem BMC0015 para Cancelada
            strBMC0015 = objMesgRecbEnviSPB.LerMensagemConciliada(Val(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text), _
                                                                  "BMC0015", _
                                                                  xmlOperacao.selectSingleNode("//CD_ASSO_CAMB").Text)
            
            Set xmlBMC0015 = CreateObject("MSXML2.DOMDocument.4.0")
            If xmlBMC0015.loadXML(strBMC0015) Then
                Call objMesgRecbEnviSPB.AlterarStatus(xmlBMC0015.selectSingleNode("//NU_CTRL_IF").Text, _
                                                      xmlBMC0015.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                      lngStatusMensagem, _
                                                      xmlBMC0015.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                      , _
                                                      False, _
                                                      xmlBMC0015.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
            
            'Atualiza status da mensagem CAM0002 para Cancelada
            Set xmlCAM0002 = objMesgRecbEnviSPB.LerXMLMensagem(, , , "CAM0002", _
                                                               , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                               , , , , xmlOperacao.selectSingleNode("//CD_ASSO_CAMB").Text)
            
            If xmlCAM0002.xml <> vbNullString Then
                lngStatusMensagem = enumStatusMensagem.MensagemCancelada
                Call objMesgRecbEnviSPB.AlterarStatus(xmlCAM0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                                      xmlCAM0002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                      lngStatusMensagem, _
                                                      xmlCAM0002.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                      , _
                                                      False, _
                                                      xmlCAM0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
        
            'Atualiza status da operacao para A Conciliar Registro
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.AConciliarRegistro, _
                                               xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text)
            
            Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            Set objFluxoProcessamento = Nothing
            
        Else
            'Atualiza status da operacao para Cancelada
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.CanceladaCamara, _
                                               xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text)
            
            Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
            
            If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
            Else
                xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
            End If
            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            
            Set objFluxoProcessamento = Nothing
        End If
    
    Else
    
        lngStatusMensagem = enumStatusMensagem.MensagemCancelada
        
        'Atualiza status da mensagem BMC0015 para Cancelada
        Set xmlBMC0015 = objMesgRecbEnviSPB.LerXMLMensagem(, , , _
                                                           "BMC0015", , , , , , , , _
                                                           strChACAM)
        
        If xmlBMC0015.xml <> vbNullString Then
            Call objMesgRecbEnviSPB.AlterarStatus(xmlBMC0015.selectSingleNode("//NU_CTRL_IF").Text, _
                                                  xmlBMC0015.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                  lngStatusMensagem, _
                                                  xmlBMC0015.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                  , _
                                                  False, _
                                                  xmlBMC0015.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        End If
        
    End If
    
    'Grava Mensagem CAM0003
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            vntNrSeqOperacao, _
                                                            vntNrConcilOperacao, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            enumStatusMensagem.Aviso, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            xmlMensagem.selectSingleNode("//NumCtrlCAM").Text)

    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlBMC0015 = Nothing
    Set xmlCAM0002 = Nothing

    Exit Function
ErrorHandler:
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlBMC0015 = Nothing
    Set xmlCAM0002 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCAM0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0101
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0101(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String

Dim vntCNPJNLiqdant                         As Variant

Dim xmlMensagemExistenteBMC                 As MSXML2.DOMDocument40
Dim strTpInf                                As String
Dim strCodigoMensagemSPB                    As String
Dim strDataOperacao                         As String
Dim lngStatusMensagemExistente              As Long
Dim vntVlrLanc                              As Variant

    On Error GoTo ErrorHandler

    strTpInf = xmlMensagem.selectSingleNode("//TpInf").Text
    If strTpInf <> "D" Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If
    
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrME").Text
    If fgVlrXml_To_Decimal(vntVlrLanc) = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If
        
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
    
    strCodigoMensagemSPB = xmlMensagem.selectSingleNode("//CodMsg").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
        
    If fgObterCodigoVeiculoLegal(strCodigoMensagemSPB, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                xmlMensagem.selectSingleNode("//VlrME").Text, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
           
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    If strTpInf = "D" And lngStatusMensagem = enumStatusMensagem.AConciliar Then
       
        Set xmlMensagemExistenteBMC = objMesgRecbEnviSPB.LerXMLMensagem(, , , _
                                                                        strCodigoMensagemSPB, _
                                                                        , , , _
                                                                        strDataOperacao, _
                                                                        strTpInf, _
                                                                        lngCodigoLocalLiquidacao)
        
        If xmlMensagemExistenteBMC.xml <> vbNullString Then
            lngStatusMensagemExistente = Val(xmlMensagemExistenteBMC.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
            
            If lngStatusMensagemExistente <> enumStatusMensagem.MensagemLiquidada And _
               lngStatusMensagemExistente <> enumStatusMensagem.AConciliar Then
                
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                
            Else
            
                If lngStatusMensagemExistente = enumStatusMensagem.AConciliar Then
                    With xmlMensagemExistenteBMC
                        Call objMesgRecbEnviSPB.AlterarStatus(.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              .selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              enumStatusMensagem.MensagemCancelada, _
                                                              vbNullString, , _
                                                              False, _
                                                              0, _
                                                              strCodigoMensagemSPB)
                    End With
                End If
                
            End If
        End If
        
    End If
    
    lngSeqControleRepeticao = 1
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao)
    Set objMesgRecbEnviSPB = Nothing

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                0, _
                                lngSeqControleRepeticao, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                enumTipoDebitoCredito.Credito, _
                                vbNullString, _
                                vntCNPJNLiqdant, _
                                , , strTpInf, , , , , _
                                enumTipoLiquidacao.Multilateral, _
                                , , , , , , , , , _
                                xmlMensagem.selectSingleNode("//CodMoeda").Text, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                , , _
                                xmlMensagem.selectSingleNode("//VlrME").Text)
                                
    Set objConciliacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0101 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0103
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0103(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String

Dim vntCNPJNLiqdant                         As Variant

    On Error GoTo ErrorHandler

    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                xmlMensagem.selectSingleNode("//VlrME").Text, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
           
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
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao)
    Set objMesgRecbEnviSPB = Nothing

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                0, _
                                lngSeqControleRepeticao, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                enumTipoDebitoCredito.Debito, _
                                vbNullString, _
                                vntCNPJNLiqdant, _
                                , , , , , , , _
                                enumTipoLiquidacao.Multilateral, _
                                , , , , , , , , , _
                                xmlMensagem.selectSingleNode("//CodMoeda").Text, _
                                xmlMensagem.selectSingleNode("//DtLiquid").Text, _
                                , , _
                                xmlMensagem.selectSingleNode("//VlrME").Text)
                                
    Set objConciliacao = Nothing
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0103 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LTR0001 (Câmara informa Resultado bruto ou bilateral de negociações)
'   - Para  CNPJ não Liquidante  é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0001(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlMesgLTR0001                          As MSXML2.DOMDocument40
Dim lngAcaoMensagem                         As enumTipoAcao

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

Dim vntCNPJNLiqdant                         As Variant
Dim strNumOpLTR                             As String
Dim vntVlrLanc                              As Variant
Dim strISPBancoLiquidante                   As String
Dim strNomeVeicLega                         As String
Dim lngTipoLiquidacao                       As Long
Dim vntIdentificadorPartCamara              As String
Dim vntCodParticipanteCamara                As Variant
Dim strISPBDebitado                         As String

    On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                   Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
        
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLTR").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrLanc").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntCNPJNLiqdant = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
    
    If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
        vntCodParticipanteCamara = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
    Else
        vntCodParticipanteCamara = xmlMensagem.selectSingleNode("//ISPBIFDebtd").Text
    End If

    If Not xmlMensagem.selectSingleNode("//ISPBIFDebtd") Is Nothing Then
        strISPBDebitado = xmlMensagem.selectSingleNode("//ISPBIFDebtd").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//ISPBIFCredtd") Is Nothing Then
        strISPBancoLiquidante = xmlMensagem.selectSingleNode("//ISPBIFCredtd").Text
    Else
        strISPBancoLiquidante = ""
    End If
        
    If Not xmlMensagem.selectSingleNode("//NumOpLTR") Is Nothing Then
        strNumOpLTR = xmlMensagem.selectSingleNode("//NumOpLTR").Text
        strNumOpLTR = UCase$(strNumOpLTR)
        lngTipoLiquidacao = enumTipoLiquidacao.Bruta
    Else
        strNumOpLTR = vbNullString
        lngTipoLiquidacao = enumTipoLiquidacao.Bilateral
    End If
       
    If strISPBancoLiquidante <> vbNullString And strNumOpLTR = vbNullString Then
        Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlMesgLTR0001 = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
        Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
        Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP)
            
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBBancoLiquidanteContraparte", "")
        Call fgAppendNode(xmlFiltros, "Grupo_ISPBBancoLiquidanteContraparte", "ISPBBancoLiquidanteContraparte", strISPBancoLiquidante)
            
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0001")
        
        Call xmlMesgLTR0001.loadXML(objMensagemSLCC.ObterDetalheMensagem(xmlFiltros, True))
        
        If xmlMesgLTR0001.xml <> vbNullString Then
            lngAcaoMensagem = enumTipoAcao.LTR0001ComISPBJaExistente
            
            If Val(xmlMesgLTR0001.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) <> lngAcaoMensagem Then
                Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgLTR0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                                      xmlMesgLTR0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                      Val(xmlMesgLTR0001.selectSingleNode("//CO_ULTI_SITU_PROC").Text), _
                                                      vbNullString, , _
                                                      False, , , _
                                                      lngAcaoMensagem)
            End If
        End If
                
        Set xmlFiltros = Nothing
        Set xmlMesgLTR0001 = Nothing
        Set objMensagemSLCC = Nothing
        Set objMesgRecbEnviSPB = Nothing
    End If
        
    If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
        vntIdentificadorPartCamara = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
    End If
           
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega, _
                                 vntIdentificadorPartCamara) Then
                                 
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntVlrLanc, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    
    End If
           
    lngSeqControleRepeticao = 1
      
    'Salvar Mensagem Enviada Recebida
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumOpLTR, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara, _
                                                            , , _
                                                            vntCodParticipanteCamara, , _
                                                            lngAcaoMensagem, _
                                                            strISPBDebitado)

    Set objMesgRecbEnviSPB = Nothing

    'Salvar Dados Conciliacao
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")

    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntVlrLanc, _
                                lngSeqControleRepeticao, _
                                "", _
                                enumTipoDebitoCredito.Debito, _
                                "", _
                                vntCNPJNLiqdant, _
                                "", _
                                0, _
                                "", _
                                0, _
                                0, _
                                "", _
                                strISPBancoLiquidante, _
                                lngTipoLiquidacao)
                                
    Set objConciliacao = Nothing
    
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        'Verifica se concilia automatico
        Call flConciliarMensagem(strCodigoMensagemSPB, _
                                 strNumeroControleIF, _
                                 lngSeqControleRepeticao, _
                                 strDataRegistroMensagem)
    End If
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processamento das LTR0007 (Câmara informa Resultado bruto ou bilateral de negociação entre Não Liquidantes da mesma IF)
'   -  Para cada CNPJ não Liquidante Creditado e CNPJ não Liquidante debitado é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0007(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlMesgLTR0007                          As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lngAcaoMensagem                         As enumTipoAcao

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
Dim lngIndicadorConfirmMesg                 As Long

Dim vntCNPJNLiqdant                         As Variant
Dim strNumOpLTR                             As String
Dim vntVlrLanc                              As Variant
Dim strNomeVeicLega                         As String
Dim vntCNPJContraparte                      As Variant
Dim vntCNPJParticipante                     As Variant
Dim vntIdentContraparteCamara               As Variant
Dim vntCodParticipanteCamara                As Variant
Dim strISPBancoLiquidante                   As String

Dim lngTipoLiquidacao                       As Long

    On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
       
    strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLTR").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrLanc = xmlMensagem.selectSingleNode("//VlrLanc").Text
    lngIndicadorConfirmMesg = IIf(UCase(xmlMensagem.selectSingleNode("//IndrConfc").Text) = "S", enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao)
    
    lngSeqControleRepeticao = 1
    
    '---------------------------------- Grava Mensagem CNPJ Debitada -------------------------------------------------------
    strNumOpLTR = vbNullString
    
    If Not xmlMensagem.selectSingleNode("//NumOpLTR") Is Nothing Then
        strNumOpLTR = xmlMensagem.selectSingleNode("//NumOpLTR").Text
    End If
      
    If strNumOpLTR = vbNullString Then
        If Not xmlMensagem.selectSingleNode("//NumOpLTRDebtd") Is Nothing Then
            strNumOpLTR = xmlMensagem.selectSingleNode("//NumOpLTRDebtd").Text
        End If
    End If
    
    lngTipoLiquidacao = IIf(Trim$(strNumOpLTR) = vbNullString, _
                            enumTipoLiquidacao.Bilateral, _
                            enumTipoLiquidacao.Bruta)
    
    If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd") Is Nothing Then
        vntCNPJNLiqdant = Val(xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text)
        vntCNPJParticipante = Val(xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text)
    End If

    If Not xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd") Is Nothing Then
        vntCodParticipanteCamara = xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd").Text
    Else
        vntCodParticipanteCamara = vntCNPJParticipante
    End If
    
    If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd") Is Nothing Then
        vntCNPJContraparte = Val(xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd").Text)
    Else
        vntCNPJContraparte = "0"
    End If
    
    If Not xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd") Is Nothing Then
        vntIdentContraparteCamara = xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd").Text
    Else
        vntIdentContraparteCamara = 0
    End If

    If fgObterCodigoVeiculoLegal(strCodigoMensagemSPB, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega, _
                                 vntCodParticipanteCamara) Then
                                 
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntVlrLanc, _
                                xmlMensagem)
        
        lngStatusMensagem = enumStatusMensagem.AConciliar
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If
    
    If Not xmlMensagem.selectSingleNode("//ISPBIF") Is Nothing Then
        strISPBancoLiquidante = xmlMensagem.selectSingleNode("//ISPBIF").Text
        
        Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlMesgLTR0007 = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
        Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
        Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP)
            
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))))
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_ISPBBancoLiquidanteContraparte", "")
        Call fgAppendNode(xmlFiltros, "Grupo_ISPBBancoLiquidanteContraparte", "ISPBBancoLiquidanteContraparte", strISPBancoLiquidante)
            
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "LTR0007")
        
        Call xmlMesgLTR0007.loadXML(objMensagemSLCC.ObterDetalheMensagem(xmlFiltros, True))
        
        If xmlMesgLTR0007.xml <> vbNullString Then
            lngAcaoMensagem = enumTipoAcao.LTR0007ComISPBJaExistente
            
            For Each objDomNode In xmlMesgLTR0007.selectNodes("//Repeat_DetalheMensagem/*")
                If Val(objDomNode.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text) <> lngAcaoMensagem Then
                    Call objMesgRecbEnviSPB.AlterarStatus(objDomNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                          objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                          Val(objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                                                          vbNullString, , _
                                                          False, , , _
                                                          lngAcaoMensagem, , , , , , _
                                                          False)
                End If
            Next
        End If
                
        Set xmlFiltros = Nothing
        Set xmlMesgLTR0007 = Nothing
        Set objMensagemSLCC = Nothing
        Set objMesgRecbEnviSPB = Nothing
        
    Else
        strISPBancoLiquidante = ""
    End If
        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    If vntCodParticipanteCamara = vntCNPJParticipante Then
        vntCodParticipanteCamara = vbNullString
    End If
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumOpLTR, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara, _
                                                            strDataRegistroMensagem, _
                                                            lngIndicadorConfirmMesg, _
                                                            vntCodParticipanteCamara, , _
                                                            lngAcaoMensagem)

    'Salvar Dados Conciliacao
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                strDataOperacao, _
                                0, _
                                0, _
                                "", _
                                0, _
                                0, _
                                vntVlrLanc, _
                                lngSeqControleRepeticao, , _
                                enumTipoDebitoCredito.Debito, , _
                                vntCNPJContraparte, , , , , , , _
                                strISPBancoLiquidante, _
                                lngTipoLiquidacao, , , , , , , , _
                                vntIdentContraparteCamara)

    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        'Verifica se concilia automatico
        Call flConciliarMensagem(strCodigoMensagemSPB, _
                                 strNumeroControleIF, _
                                 lngSeqControleRepeticao, _
                                 strDataRegistroMensagem)
    End If
    '---------------------------------- Grava Mensagem CNPJ Creditada -------------------------------------------------------
    
    Call fgWait(3)
    
    lngSeqControleRepeticao = 1
    
     'Participante Credito ----------------------------------------------------------
     
    If Not xmlMensagem.selectSingleNode("//NumOpLTRCredtd") Is Nothing Then
        If xmlMensagem.selectSingleNode("//NumOpLTRCredtd").Text <> vbNullString Then
            strNumOpLTR = xmlMensagem.selectSingleNode("//NumOpLTRCredtd").Text
        End If
    End If
    
    If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd") Is Nothing Then
        vntCNPJNLiqdant = Val(xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd").Text)
    Else
        vntCNPJNLiqdant = ""
    End If
    
    If vntCNPJNLiqdant <> "" Then
        
        If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd") Is Nothing Then
            vntCNPJContraparte = Val(xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text)
        Else
            vntCNPJContraparte = "0"
        End If
        
        If Not xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd") Is Nothing Then
            vntCodParticipanteCamara = xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd").Text
        Else
            vntCodParticipanteCamara = vntCNPJNLiqdant
        End If
        
        If Not xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd") Is Nothing Then
            vntIdentContraparteCamara = xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd").Text
        Else
            vntIdentContraparteCamara = vbNullString
        End If

        If fgObterCodigoVeiculoLegal(strCodigoMensagemSPB, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice, _
                                     strNomeVeicLega, _
                                     vntCodParticipanteCamara) Then
                                     
            Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrLanc, _
                                    xmlMensagem)
        
            lngStatusMensagem = enumStatusMensagem.AConciliar
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        End If
        
        If vntCodParticipanteCamara = vntCNPJNLiqdant Then
            vntCodParticipanteCamara = vbNullString
        End If
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                                Trim(strNumeroControleIF), _
                                                                0, _
                                                                0, _
                                                                strCodigoMensagemSPB, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                strNumOpLTR, _
                                                                "", _
                                                                lngStatusMensagem, _
                                                                lngIndicadorEntradaManual, _
                                                                "", _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                lngSeqControleRepeticao, _
                                                                strNumeroControleCamara, _
                                                                "", _
                                                                lngIndicadorConfirmMesg, _
                                                                vntCodParticipanteCamara, , _
                                                                lngAcaoMensagem)

        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrLanc, _
                                    lngSeqControleRepeticao, , _
                                    enumTipoDebitoCredito.Credito, , _
                                    vntCNPJContraparte, , , , , , , _
                                    strISPBancoLiquidante, _
                                    lngTipoLiquidacao, , , , , , , , _
                                    vntIdentContraparteCamara)
        
    End If
    
    If lngStatusMensagem = enumStatusMensagem.AConciliar Then
        'Verifica se concilia automatico
        Call flConciliarMensagem(strCodigoMensagemSPB, _
                                 strNumeroControleIF, _
                                 lngSeqControleRepeticao, _
                                 strDataRegistroMensagem)
    End If
    
    Set objMesgRecbEnviSPB = Nothing
    Set objConciliacao = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0007 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------
'Gerenciamento de envio de alerta ao usuário
'------------------------------------------------------------------

Private Function flEnviarAlertaR1(ByRef pxmlMensagem As MSXML2.DOMDocument40) As String

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim strNumeroIF                             As String
Dim strMensagemIda                          As String
Dim xmlMensagemIda                          As MSXML2.DOMDocument40
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim xmlStatusMensagem                       As MSXML2.DOMDocument40

Dim vntValor                                As Variant
Dim lngTipoBackOffice                       As Long
Dim strNomeVeiculoLegal                     As String
Dim strSiglaSistema                         As String
Dim strCodigoMensagemSPB                    As String
Dim vntCodigoVeiculoLegal                   As Variant
Dim strXmlVeiculoLegal                      As String
Dim strXmlStatusMensagem                    As String

Dim lngStatusMensagem                       As Long
Dim blnEnviaAlerta                          As Boolean

    On Error GoTo ErrorHandler

    If Not pxmlMensagem.selectSingleNode("//NumCtrlIF") Is Nothing Then
        strNumeroIF = pxmlMensagem.selectSingleNode("//NumCtrlIF").Text
    ElseIf Not pxmlMensagem.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
        strNumeroIF = pxmlMensagem.selectSingleNode("//NU_CTRL_IF").Text
    End If

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strMensagemIda = objMensagemEnviadaRecebida.LerMensagemIda(strNumeroIF)
    Set objMensagemEnviadaRecebida = Nothing
    
    Set xmlMensagemIda = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlMensagemIda.loadXML(strMensagemIda) Then
        'Mensagem de Ida não encontrada
        lngCodigoErroNegocio = 9999
        GoTo ErrorHandler
    End If
    
    strCodigoMensagemSPB = pxmlMensagem.selectSingleNode("//CodMsg").Text
    
    If xmlMensagemIda.selectSingleNode("//TP_BKOF").Text = "" Then xmlMensagemIda.selectSingleNode("//TP_BKOF").Text = enumTipoBackOffice.Todos
    lngTipoBackOffice = xmlMensagemIda.selectSingleNode("//TP_BKOF").Text
    vntCodigoVeiculoLegal = xmlMensagemIda.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlMensagemIda.selectSingleNode("//SG_SIST").Text
    
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    strXmlVeiculoLegal = objVeiculoLegal.Ler(vntCodigoVeiculoLegal, strSiglaSistema)
    Set objVeiculoLegal = Nothing
    
    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlVeiculoLegal.loadXML(strXmlVeiculoLegal)
    strNomeVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//NO_VEIC_LEGA").Text
    Set xmlVeiculoLegal = Nothing
    
    Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                            lngTipoBackOffice, _
                            strNomeVeiculoLegal, _
                            vntValor, _
                            pxmlMensagem, _
                            lngStatusMensagem)

    Set xmlMensagemIda = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagemEnviadaRecebida = Nothing
    Set xmlMensagemIda = Nothing
    Set objVeiculoLegal = Nothing
    Set xmlStatusMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarAlertaR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'----------------------------------------------------------------------------------------
'Processamento de Mensagens não são conciliadas com operações.
'----------------------------------------------------------------------------------------

Public Function ProcessaMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                 ByVal pstrHeaderNZ As String, _
                                 ByVal plngTipoMensagemSPB As enumTipoMensagemSPB, _
                        Optional ByVal pstrNumeroControleIFInformado As String = vbNullString) As String

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim xmlRetorno                              As MSXML2.DOMDocument40

Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim strAgendamento                          As String
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim lngSeqControleRepeticao                 As Long
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strNumCtrlCamr                          As String

Dim lngISPB                                 As Long

On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    If pstrNumeroControleIFInformado = vbNullString Then
        strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    Else
        strNumeroControleIF = pstrNumeroControleIFInformado
    End If
    
    strMensagem = xmlMensagem.xml
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strAgendamento = ""
    strCodigoVeiculoLegal = ""
    strSiglaSistema = udtHeaderMensagem.SiglaSistemaEnviouNZ
   
    lngSeqControleRepeticao = 1
    lngISPB = 0
        
    Select Case Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3)
        
        Case "STR"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.SSTR
            
            If Not xmlMensagem.selectSingleNode("//NumCtrlSTR") Is Nothing Then
                strNumCtrlCamr = xmlMensagem.selectSingleNode("//NumCtrlSTR").Text
            End If
        
        Case "PAG"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.PAG
            
        Case "CTP"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
            
        Case "SEL"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.SELIC
            
        Case "BMA"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
            
            If Not xmlMensagem.selectSingleNode("//NumCtrlLDL") Is Nothing Then
                strNumCtrlCamr = xmlMensagem.selectSingleNode("//NumCtrlBMA").Text
            End If
            
        Case "LDL"
            If Not xmlMensagem.selectSingleNode("//ISPBLDL") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBLDL").Text)
            ElseIf Not xmlMensagem.selectSingleNode("//ISPBIF_LDLDebtd") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBIF_LDLDebtd").Text)
            End If
            
            Select Case lngISPB
                Case enumISPB.IspbBMA
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
                Case enumISPB.IspbCETIP
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
                Case enumISPB.IspbBMC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
                Case enumISPB.IspbBMD
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD
                Case enumISPB.IspbCBLC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes
                Case enumISPB.IspbSELIC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.SELIC
            End Select
            
            If Not xmlMensagem.selectSingleNode("//NumCtrlLDL") Is Nothing Then
                strNumCtrlCamr = xmlMensagem.selectSingleNode("//NumCtrlLDL").Text
            End If
            
        Case "LTR"
            If Not xmlMensagem.selectSingleNode("//ISPBLTR") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBLTR").Text)
            ElseIf Not xmlMensagem.selectSingleNode("//ISPBIF_LTRDebtd") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBIF_LTRDebtd").Text)
            End If
            
            Select Case lngISPB
                Case enumISPB.IspbBMA
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
                Case enumISPB.IspbCETIP
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
                Case enumISPB.IspbBMC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
                Case enumISPB.IspbBMD
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD
                Case enumISPB.IspbCBLC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCTPub
                Case enumISPB.IspbSELIC
                    lngCodigoLocalLiquidacao = enumLocalLiquidacao.SELIC
            End Select
            
            If Not xmlMensagem.selectSingleNode("//NumCtrlLTR") Is Nothing Then
                strNumCtrlCamr = xmlMensagem.selectSingleNode("//NumCtrlLTR").Text
            End If
    
        Case "BMC"
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
            
            If Not xmlMensagem.selectSingleNode("//NumCtrlBMC") Is Nothing Then
                strNumCtrlCamr = xmlMensagem.selectSingleNode("//NumCtrlBMC").Text
            End If
        
    End Select
        
    Select Case plngTipoMensagemSPB
        Case enumTipoMensagemSPB.MensagemR1
            lngStatusMensagem = enumStatusMensagem.R1
        Case enumTipoMensagemSPB.MensagemR2
            lngStatusMensagem = enumStatusMensagem.R2
        Case enumTipoMensagemSPB.MensagemAviso, enumTipoMensagemSPB.MensagemInformacao
            Select Case flObterTipoMensagemSPB(xmlMensagem.selectSingleNode("//CodMsg").Text)
                Case TipoFluxo5
                    lngStatusMensagem = enumStatusMensagem.Informação
                Case TipoFluxo7
                    lngStatusMensagem = enumStatusMensagem.Aviso
                Case Else
                    lngStatusMensagem = enumStatusMensagem.Informação
                End Select
        Case enumTipoMensagemSPB.MensagemInativa
            lngStatusMensagem = enumStatusMensagem.Inativa
    End Select
                        
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
                    
    'Salvar Mensagem Enviada Recebida
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
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
                                                            "", _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao, _
                                                            strNumCtrlCamr)
        
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlRetorno, "", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlRetorno, "Grupo_Mensagem", "NU_CTRL_IF", strNumeroControleIF)
    Call fgAppendNode(xmlRetorno, "Grupo_Mensagem", "DH_REGT_MESG_SPB", strDataRegistroMensagem)
    Call fgAppendNode(xmlRetorno, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", lngSeqControleRepeticao)
        
    ProcessaMensagem = xmlRetorno.xml
    
    Set xmlRetorno = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function

ErrorHandler:
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-----------------------------------------------------------------------------
'Obter o Tipo de fluxo da mensagem
'-----------------------------------------------------------------------------

Private Function flObterTipoMensagemSPB(ByVal pstrCodigoMensagem As String) As Integer

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT SQ_TIPO_FLUX " & _
             "   FROM A8.TB_MENSAGEM A, " & _
             "        A8.TB_EVENTO B " & _
             "  WHERE A.SQ_EVEN = B.SQ_EVEN " & _
             "    AND A.CO_MESG = '" & Trim(pstrCodigoMensagem) & "'"
             
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        flObterTipoMensagemSPB = objRS.fields("SQ_TIPO_FLUX")
    End If
    
    objRS.Close

    Exit Function
ErrorHandler:


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterTipoMensagemSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-----------------------------------------------------------------------------
'Processamento de mensagens de erro de negócio geradas por um remessa do legado
'-----------------------------------------------------------------------------

Private Function flProcessaMesgTratadaErro(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByVal pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim strCodigoMensagemSPB                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim lngSeqControleRepeticao                 As Long
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strComandoOperacao                      As String
Dim blnAtualizaStatusOperGerenciarChamada   As Boolean
Dim strDataHoraRegistroMensagem             As String


On Error GoTo ErrorHandler

'-------------------------------------------

    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtProtocoloErroNZ.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    blnAtualizaStatusOperGerenciarChamada = True
    
    lngStatusMensagem = enumStatusMensagem.ErroNegocioSPB
    lngStatusMensagemOriginal = enumStatusMensagem.Respondida
    
    lngCodigoEmpresa = udtProtocoloErroNZ.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    lngSeqControleRepeticao = 1
    
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
    
    'Obter a Operação
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(strNumeroControleIF)) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
        
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlOperacao.loadXML(objOperacao.Ler(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
    
    'Salvar Mensagem Enviada Recebida
    'Obter o mensagem de requisicao
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                    
    vntSequenciaOperacao = xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    strCodigoVeiculoLegal = xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text
    
    'Bruno Oliveira - 19/abr/2012 - Mudança para ler SG_SIST da operação ao invés da mensagem
    If Not xmlOperacao.documentElement.selectSingleNode("//SG_SIST") Is Nothing Then
        strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("//SG_SIST").Text
    Else
        strSiglaSistema = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//SG_SIST").Text
    End If
    
    lngCodigoLocalLiquidacao = xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
    lngTipoBackOffice = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//TP_BKOF").Text
    strComandoOperacao = xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
    strDataHoraRegistroMensagem = xmlMesgRecbEnviSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text
    lngSeqControleRepeticao = Val(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strComandoOperacao, _
                                                            strSituacaoMensagemSPB, _
                                                            enumStatusMensagem.ErroNegocioSPB, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao, _
                                                            "", _
                                                            "")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
        fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
    End If
    
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
   
   
    If Trim(udtProtocoloErroNZ.CodigoErro1) = "5156" Then
        lngStatusOperacao = enumStatusOperacao.RejeitadaPiloto
    Else
        lngStatusOperacao = enumStatusOperacao.Rejeitada
    End If
   
    Select Case Mid(strCodigoMensagemSPB, 1, 7)
    
        Case "BMA0006", "BMA0005"
            
            Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
               
               Case enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                    enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
                    enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA
                   
                   Call objMesgRecbEnviSPB.AlterarStatusBMA0002LDL0001(vntSequenciaOperacao, enumStatusMensagem.AConciliar)
                   
            End Select
                    
                    
        Case Else
                    
    End Select
    
    GerenciarChamada xmlOperacao, _
                     xmlMensagem, _
                     lngStatusOperacao, _
                     enumStatusMensagem.MensagemRejeitada, _
                     "", _
                     enumIndicadorSimNao.Sim, _
                     strNumeroControleIF, _
                     , _
                     blnAtualizaStatusOperGerenciarChamada, _
                     "", _
                     pstrHeaderNZ
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    'Se erro mensagem CTP Com ordem de lançamento , alterar CTP1002 para A conciliar
    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
    
        Case enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao, _
             enumTipoOperacaoLQS.RetiradaCustodiaConciliacao, _
             enumTipoOperacaoLQS.TransferCustodiaConciliacao, _
             enumTipoOperacaoLQS.OperacaoDefinitivaVendaOrdem, _
             enumTipoOperacaoLQS.MovimExercDireitosDebenturesConciliacao, _
             enumTipoOperacaoLQS.ConversaoPermutaValorMobiliarioConciliacao, _
             enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP

            'Alterar status a CTP1002 para A conciliar
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        
            If xmlCociliacao.loadXML(objConciliacao.LerTodos(, vntSequenciaOperacao)) Then
            
                For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
                    'Numero Controle IF CTP1002
                    strNumeroControleIF = xmlNode.selectSingleNode("NU_CTRL_IF").Text
                    
                    xmlMesgRecbEnviSPB.loadXML objMesgRecbEnviSPB.LerMensagemIda(strNumeroControleIF)
                    
                    Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.selectSingleNode("//NU_CTRL_IF").Text, _
                                                          xmlMesgRecbEnviSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                          enumStatusMensagem.AConciliar, _
                                                          xmlMesgRecbEnviSPB.selectSingleNode("//DH_ULTI_ATLZ ").Text)

                Next
            
            End If
        
            Set objConciliacao = Nothing
            Set xmlCociliacao = Nothing
        
    End Select
    
    
    Set xmlCociliacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMesgTratadaErro Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'-----------------------------------------------------------------------------
'Processamento de mensagens de erro de negócio não geradas por uma remessa do legado
'-----------------------------------------------------------------------------

Private Function flProcessaMesgNaoTratadaErro(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                              ByVal pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objAlerta                               As A8LQS.clsAlerta
Dim objLegado                               As A8LQS.clsLegado
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlCociliacao                           As MSXML2.DOMDocument40
        
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim strCodigoMensagemSPB                    As String
Dim strCodigoMensagemMae                    As String
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim strSituacaoMensagemSPB                  As String
Dim lngStatusMensagem                       As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngIndicadorEntradaManual               As Long
Dim strAgendamento                          As String
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim lngSeqControleRepeticao                 As Long
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strNumeroComando                        As String
Dim blnGravaMensagem                        As Boolean
Dim blnAtualizaStatusOperGerenciarChamada   As Boolean
Dim strDataHoraRegistroMensagem             As String
Dim lngStatusOperacao                       As Long

'KIDA - CBLC - 03/10/2008
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim blnAcaoPagtoContingencia                As Boolean

On Error GoTo ErrorHandler

    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtProtocoloErroNZ.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngStatusMensagem = enumStatusMensagem.ErroNegocioSPB
    lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
    
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtProtocoloErroNZ.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strAgendamento = ""
    lngSeqControleRepeticao = 1
    strSituacaoMensagemSPB = ""
    
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
                    
    'Salvar Mensagem Enviada Recebida
    'Obter o mensagem de requisicao
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
                        
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
                        
    vntSequenciaOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    lngCodigoEmpresa = xmlMesgRecbEnviSPB.selectSingleNode("//CO_EMPR").Text
    strCodigoVeiculoLegal = xmlMesgRecbEnviSPB.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlMesgRecbEnviSPB.selectSingleNode("//SG_SIST").Text
    lngCodigoLocalLiquidacao = xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text
    lngTipoBackOffice = Val(xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text)
    strDataHoraRegistroMensagem = ""
    lngSeqControleRepeticao = Val(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            vntSequenciaConciliacaoOperacao, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumeroComando, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao, "", strDataHoraRegistroMensagem)
    
    Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                          enumStatusMensagem.MensagemRejeitada, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                          , _
                                          , _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    blnAtualizaStatusOperGerenciarChamada = True
    
    If Left$(strCodigoMensagemSPB, 3) <> "GEN" Then
        strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(strCodigoMensagemSPB, 1, 7)) ' LDL0001
    Else
        strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_MESG_SPB").Text, 1, 7)) ' LDL0001
    End If
    
    If strCodigoMensagemMae <> vbNullString Then
        
        strMensagem = flObterMensagemMae(vntSequenciaConciliacaoOperacao, strCodigoMensagemMae)
        
        If Trim$(strMensagem) <> vbNullString Then
            xmlMesgRecbEnviSPB.loadXML strMensagem
            
            strNumeroControleIF = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text
                    
            If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
                
                'KIDA - CBLC - 03/10/2008
                '-------------------------------------------------------------------------
                Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
                Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
                
                strMensagem = objHistMensagem.LerTodos(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                       xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                       xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                 
                xmlHistorico.loadXML strMensagem
                                
                blnAcaoPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
                
                Set objHistMensagem = Nothing
                Set xmlHistorico = Nothing
                '----------------------------------------------------------------------------
                
                If strCodigoMensagemSPB = "LTR0002E" Then
                    
                    lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                
                ElseIf strCodigoMensagemSPB = "LDL0008E" Then
                    
                    lngStatusMensagemOriginal = enumStatusMensagem.AConciliar
                
                ElseIf strCodigoMensagemSPB = "LDL0004E" Or _
                       strCodigoMensagemSPB = "LDL0004" Then
                    
                    lngStatusMensagemOriginal = enumStatusMensagem.AConciliar
                
                'KIDA - CBLC - 03/10/2008
                ElseIf strCodigoMensagemSPB = "LTR0004E" Or _
                       strCodigoMensagemSPB = "LTR0004" Then
                
                    If blnAcaoPagtoContingencia Then
                        lngStatusMensagemOriginal = enumStatusMensagem.AConciliar
                    Else
                        lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                    End If
                Else
                    lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                End If
            
            ElseIf lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC Then
                
                Select Case Left$(Trim(strCodigoMensagemSPB), 7)
                    Case "BMC0012"
                        lngStatusOperacao = enumStatusOperacao.AConciliar
                    
                    Case "LDL0004", "GEN0004"
                        
                        Select Case Val(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                            Case enumStatusMensagem.AConciliar
                                lngStatusMensagemOriginal = enumStatusMensagem.AConciliar
                                lngStatusOperacao = enumStatusOperacao.AConciliar
                            Case enumStatusMensagem.PagamentoLib
                                lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                                lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
                        End Select
                    
                    Case Else
                        lngStatusOperacao = enumStatusOperacao.Rejeitada
                
                End Select
                
            ElseIf lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP Then
                
                If strCodigoMensagemSPB = "LTR0002E" Then
                    lngStatusMensagemOriginal = Val(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                End If
            
            Else
            
                lngStatusOperacao = enumStatusOperacao.Rejeitada
            
            End If
            
            If lngStatusMensagemOriginal <> Val(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) Then
                Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                      lngStatusMensagemOriginal, _
                                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_RECB_ENVI_MESG_SPB").Text, _
                                                      False, _
                                                      0, _
                                                      strCodigoMensagemMae)
            End If
            
        End If
    End If
    '-------------------------------------------------------------------------------------------------------------
    
    If vntSequenciaConciliacaoOperacao <> "0" Then
    
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao, , strNumeroControleIF)) Then
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
            For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
                
                vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
                If Val(vntSequenciaOperacao) <> 0 Then
                    Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                    
                    'KIDA - CBLC
                    If xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
                        
                        If strCodigoMensagemSPB = "LDL0008E" Then
                            lngStatusOperacao = enumStatusOperacao.Rejeitada

                            
                            lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                            
                            Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                            
                            blnAtualizaStatusOperGerenciarChamada = False
                            
                        ElseIf strCodigoMensagemSPB = "LTR0002E" Then
                            
                            lngStatusOperacao = enumStatusOperacao.Rejeitada
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
                            blnAtualizaStatusOperGerenciarChamada = False
                        
                        ElseIf strCodigoMensagemSPB = "LDL0004E" Then
                            
                            lngStatusOperacao = enumStatusOperacao.Rejeitada
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
                            Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                            blnAtualizaStatusOperGerenciarChamada = False
                        
                        ElseIf strCodigoMensagemSPB = "LTR0004E" Then
                            
                            lngStatusOperacao = enumStatusOperacao.Rejeitada
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
                            Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                            blnAtualizaStatusOperGerenciarChamada = False
                        
                        'KIDA - CBLC - 09/10/2008
                        ElseIf strCodigoMensagemSPB = "LTR0008E" Then
                            
                            lngStatusOperacao = enumStatusOperacao.Rejeitada
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
                            blnAtualizaStatusOperGerenciarChamada = False
                            
                            If xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                                Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                            Else
                                lngStatusOperacao = Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                            End If
                        Else
                            lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemRejeitada
                            blnAtualizaStatusOperGerenciarChamada = False
                        End If
                        
                    ElseIf xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMC Then
                        
                        Select Case Left$(Trim(strCodigoMensagemSPB), 7)
                            Case "BMC0012", "BMC0102"
                                lngStatusOperacao = enumStatusOperacao.AConciliar
                            
                            Case Else
                                lngStatusOperacao = enumStatusOperacao.Rejeitada
                            
                        End Select
                        
                    ElseIf xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
                        
                        If strCodigoMensagemSPB = "LTR0002E" Then
                            lngStatusOperacao = Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                        
                            Set objAlerta = CreateObject("A8LQS.clsAlerta")
                            Call objAlerta.GerarAlerta(xmlOperacao, xmlMensagem, "Rejeitada", , enumStatusOperacao.Rejeitada)
                            Set objAlerta = Nothing
                        End If
                        
                    Else
                        
                        lngStatusOperacao = enumStatusOperacao.Rejeitada
                    
                    End If
                    
                    If lngStatusOperacao <> Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text) Then
                        GerenciarChamada xmlOperacao, xmlMensagem, _
                                         lngStatusOperacao, 0, _
                                         "", enumIndicadorSimNao.Sim, _
                                         strNumeroControleIF, _
                                         False, _
                                         blnAtualizaStatusOperGerenciarChamada
                    End If
                    
                End If
                
            Next
            
            Set objOperacao = Nothing
            Set xmlOperacao = Nothing
        
        End If
    
    Else
    
        If vntSequenciaOperacao <> 0 Then
            
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
            
            Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
               
                Case enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                     enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
                     enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA
                                        
                     Call objMesgRecbEnviSPB.AlterarStatusBMA0002LDL0001(vntSequenciaOperacao, enumStatusMensagem.AConciliar)
            
            End Select
            
            If strCodigoMensagemSPB = "CTP0100E" Then
            
                Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, _
                                                         xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         0, _
                                                         enumTipoAcao.CancelamentoRejeitado)
                
                xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Rejeitada
                
                Set objAlerta = CreateObject("A8LQS.clsAlerta")
                Call objAlerta.GerarAlerta(xmlOperacao, xmlMensagem, "Rejeitada", , enumStatusOperacao.Rejeitada)
                Set objAlerta = Nothing
                
                Set xmlNode = xmlOperacao.selectSingleNode("//SISMSG[*/CodMsg!='CTP0100E']")
                
                If Not xmlNode Is Nothing Then
                    xmlOperacao.documentElement.removeChild xmlNode
                End If
                       
                Set objLegado = CreateObject("A8LQS.clsLegado")
                Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
                
                strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem)
                Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS)
                
                Set objProcessoOperacao = Nothing
                Set objLegado = Nothing
                
            Else
                
                GerenciarChamada xmlOperacao, _
                                 xmlMensagem, _
                                 enumStatusOperacao.Rejeitada, _
                                 lngStatusMensagemOriginal, _
                                 "", _
                                 enumIndicadorSimNao.Sim, _
                                 strNumeroControleIF, _
                                 False, _
                                 blnAtualizaStatusOperGerenciarChamada
                    
            End If
                
            Set objOperacao = Nothing
            Set xmlOperacao = Nothing
        
        End If
    
    End If
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function

ErrorHandler:
    Set objProcessoOperacao = Nothing
    Set objLegado = Nothing
    
    Set objAlerta = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMesgNaoTratadaErro Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'-----------------------------------------------------------------------------
'Processamento de mensagens R1 geradas por uma remessa do legado
'-----------------------------------------------------------------------------

Private Function flProcessaMesgTratadaR1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                         ByVal pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao

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
Dim blnGerenciar                            As Boolean
Dim blnAtualizaStatusOper                   As Boolean
Dim strMensagemLegado                       As String

Dim strSituacaoRecebida                     As String
Dim strHistSituacaoOperacao                 As String

Dim xmlParametrosGerais                     As MSXML2.DOMDocument40

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
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
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

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
    
        Case enumStatusOperacao.Liquidada
            
            'Exceção para demanda Porteiro
            If strCodigoMensagemSPB = "CTP9002R1" Then
                                
                'Verifica parametrização porteiro
                Set xmlParametrosGerais = CreateObject("MSXML2.DOMDocument.4.0")
                Call xmlParametrosGerais.loadXML(fgSelectVarchar4000(0, False))
        
                If Not xmlParametrosGerais.selectSingleNode("//ALTERAR") Is Nothing Then
                    If UCase(xmlParametrosGerais.selectSingleNode("//ALTERAR").Text) = "S" Then
                    
                        'Não dá erro
                        Set xmlParametrosGerais = Nothing
                        GoTo ByPass1
                    
                    End If
                End If
                
                Set xmlParametrosGerais = Nothing
                
            End If
            
             'INICIO Nick - Projeto INR Liquidação Bruta
            ' Incluindo a situação 43 para não rejeitar a mensagem
            If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
                If Trim(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = "2" Then
                    If Trim(xmlMensagem.selectSingleNode("//SitOpCTP").Text) = "43" Then
                       GoTo ByPass1
                    End If
                End If
            End If
            'FIM Nick - Projeto INR Liquidação Bruta

            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        
ByPass1:

        Case enumStatusOperacao.Registrada, enumStatusOperacao.RegistradaAutomatica, _
             enumStatusOperacao.ConcordanciaBackoffice, enumStatusOperacao.ConcordanciaBackofficeAutomatico
        
            If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
            
            'INICIO Nick - Projeto INR Liquidação Bruta
                ' Incluindo a situação 43 para não rejeitar a mensagem
                If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
                    If Trim(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) <> "2" Then
                        If Trim(xmlMensagem.selectSingleNode("//SitOpCTP").Text) <> "43" Then
    
                            '3122 - Operação já Registrada.
                            lngCodigoErroNegocio = 3122
                            GoTo ErrorHandler
                        
                        End If
                    End If
                End If
            'FIM Nick - Projeto INR Liquidação Bruta

            End If
    
    End Select
    
    'Salvar Mensagem Enviada Recebida
    'Obter o mensagem de requisicao
    vntSequenciaOperacao = xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    strCodigoVeiculoLegal = xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
    lngCodigoLocalLiquidacao = xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
    lngTipoBackOffice = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//TP_BKOF").Text
   
    
    If Not xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        strComandoOperacao = xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
    End If
        
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
        
    If Left$(strCodigoMensagemSPB, 3) = "CTP" Then
        If Not xmlMensagem.selectSingleNode("//NumOpCTP") Is Nothing Then
            strComandoOperacao = xmlMensagem.selectSingleNode("//NumOpCTP").Text
        End If
    End If
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strComandoOperacao, _
                                                            strSituacaoMensagemSPB, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao)
    
        
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
    
    Select Case strCodigoMensagemSPB
        Case "CTP0052R1"
            Call fgWait(1)
        Case "CTP0001R1"
            If Val(Mid$(strSituacaoRecebida, 7)) = enumOperacaoCETIP_CTP0100.EstornoFinalizado Then
                Call fgWait(1)
            End If
    End Select
    
    objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                     lngStatusMensagem, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                     , _
                                     False, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
    
    If CLng(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        
        'RATS - 942
        If Not xmlMensagem.selectSingleNode("//NumCtrlCTP") Is Nothing Then
            strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlCTP").Text
        End If
        
        If Not xmlMensagem.selectSingleNode("//NumOpCTP") Is Nothing Then
            strComandoOperacao = xmlMensagem.selectSingleNode("//NumOpCTP").Text
        End If
        
        Dim strNumAtivMerc As String
        
        If Not xmlMensagem.selectSingleNode("//IdentdTitCTP") Is Nothing Then
            strNumAtivMerc = xmlMensagem.selectSingleNode("//IdentdTitCTP").Text
        End If
        
        Call objOperacao.GravarControleSPBOriginalCTP(vntSequenciaOperacao, _
                                                      strNumeroControleCamara, _
                                                      strComandoOperacao, _
                                                      strNumAtivMerc)
        
    End If

    '-------------------------------------------------------------------
            
    blnGerenciar = True
    blnAtualizaStatusOper = True
            
    Select Case xmlStatus.selectSingleNode("//Status_Operacao").Text
        
        Case enumStatusOperacao.Liquidada, _
             enumStatusOperacao.EmLancamento, _
             enumStatusOperacao.Pendencia, _
             enumStatusOperacao.Registrada, _
             enumStatusOperacao.PendenteEstrategia, _
             enumStatusOperacao.EnviadaSTR
                                                       
             Select Case strCodigoMensagemSPB
                                                       
                Case "BMA0006R1"
                    
                    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
                       
                       Case enumTipoOperacaoLQS.CancelamentoEspecificacaoBMA, _
                            enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA
                       
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                                'Status da Especificação original = Cancelada
                                Call flProcessaMensagemBMA0006R1(xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text)
                            End If
                       
                       Case enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
                            enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
                            enumTipoOperacaoLQS.EspecTermoIntermediacao, _
                            enumTipoOperacaoLQS.EspecTermoCobertura, _
                            enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                            enumTipoOperacaoLQS.OperacaoTermoCobertaBMA
                            
                            blnGerenciar = False
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                               Call flGerenciarEnvioMensagem(xmlMensagem, strNumeroControleIF, enumFuncaoSistema.Liberar)
                            Else
                               blnAtualizaStatusOper = False
                               blnGerenciar = True
                            End If
                        
                        Case enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, _
                             enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
                             enumTipoOperacaoLQS.EspecCompromissadaCobertura
                           
                            blnGerenciar = False
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                               Call flGerenciarEnvioMensagem(xmlMensagem, strNumeroControleIF, enumFuncaoSistema.LberarCancelamentoEspecificacaoCompromissada)
                            End If
                           
                    End Select
                
                Case "LDL1002R1", "LDL1006R1"
                    
                    blnGerenciar = False
                    
                    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
                        
                        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                               Call flGerenciarEnvioMensagem(xmlMensagem, strNumeroControleIF, enumFuncaoSistema.LiberarLiquidacaoLeilaoBMA)
                            End If
                            
                        Case Else
                        
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                               Call flGerenciarEnvioMensagem(xmlMensagem, strNumeroControleIF, enumFuncaoSistema.Liberar)
                            End If
                            
                    End Select
                    
                    If strCodigoMensagemSPB = "LDL1006R1" Then
                    
                        Set objLegado = CreateObject("A8LQS.clsLegado")
                
                        xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
                        
                        strMensagemLegado = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem)
                        objProcessoOperacao.EnviarMensagemMQ strMensagemLegado, enumIdentificadorFila.BUS
                
                        Set objLegado = Nothing
                    
                    End If
                    
                Case "LDL1016R1"
                    
                    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA Then
                        
                        blnAtualizaStatusOper = False
                        blnGerenciar = False
                       
                        Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, _
                                                                 enumStatusOperacao.Pendencia, _
                                                                 0, _
                                                                 0)
                        
                    End If
                
                 Case "CTP0001R1", "CTP0011R1"
                    
                    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.MovimentacaoInstrumentoFinanceiro Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.AplicacaoFundoInvestimentoCETIP Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DepositoFundoInvestimentoCETIP Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RetiradaCustodia Or _
                       xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RetiradaCustodiaConciliacao Then
                        
                        Call objOperacao.AlterarComandoOperacao(vntSequenciaOperacao, _
                                                                strComandoOperacao)
                    
                    End If
               
                 Case "CTP0012R1"
                 
                    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
                    Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(vntSequenciaOperacao))
                    
                    If xmlConciliacao.xml <> vbNullString Then
                        If xmlConciliacao.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                            
                            If Val(xmlConciliacao.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Nao Then
                                blnGerenciar = False
                                xmlStatus.selectSingleNode("//Status_Operacao").Text = "0"
                            End If
                        
                            If Val(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> 0 Then
                                Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                            
                                strHistSituacaoOperacao = objHistSituacaoOperacao.Ler( _
                                                                "A.NU_SEQU_OPER_ATIV = " & vntSequenciaOperacao & " AND " & _
                                                                "A.CO_SITU_PROC IN (" & xmlStatus.selectSingleNode("//Status_Operacao").Text & ")")
                                                        
                                If strHistSituacaoOperacao <> vbNullString Then
                                    blnGerenciar = False
                                    xmlStatus.selectSingleNode("//Status_Operacao").Text = 0
                                End If
                            
                                Set objHistSituacaoOperacao = Nothing
                            End If
                        
                        End If
                    End If
                    
                    Set objConciliacao = Nothing
                    Set xmlConciliacao = Nothing
                    
                 Case "CTP9001R1"
                 
                    'Demanda Porteiro
                    'Verifica se é retorno pendência de complemento de layout 94 com prêmio
                    If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                    
                        If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
                            If Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira Then
                            
                                'Verifica parametrização porteiro
                                Set xmlParametrosGerais = CreateObject("MSXML2.DOMDocument.4.0")
                                Call xmlParametrosGerais.loadXML(fgSelectVarchar4000(0, False))
                        
                                If Not xmlParametrosGerais.selectSingleNode("//ALTERAR") Is Nothing Then
                                    If UCase(xmlParametrosGerais.selectSingleNode("//ALTERAR").Text) = "S" Then
                                    
                                        'Gera CTP9002 de complemento
                                        Call fgAppendNode(xmlOperacao, "MESG", "CO_MESG", "CTP9002")
                                        strMensagem = objMensagem.MontarMensagem(xmlOperacao, True)
                                        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, True
                                    
                                    End If
                                End If
                                
                                Set xmlParametrosGerais = Nothing
                                
                            End If
                        End If
                        
                    End If
                 
             End Select
                                                                       
             If blnGerenciar Then
                GerenciarChamada xmlOperacao, _
                                 xmlMensagem, _
                                 CLng(IIf(xmlStatus.selectSingleNode("//Status_Operacao").Text = "", 0, xmlStatus.selectSingleNode("//Status_Operacao").Text)), _
                                 0, _
                                 IIf(xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = "", "", xmlStatus.selectSingleNode("//Situacao_Mensagem").Text), _
                                 enumIndicadorSimNao.Nao, _
                                 strNumeroControleIF, _
                                 False, _
                                 blnAtualizaStatusOper, _
                                 , _
                                 , _
                                 lngTipoBackOffice
             End If
       
            'Tratamento do fluxo automático para as mensagens CETIP
            If Trim$(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> vbNullString Then
                If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Registrada Then
                    Call flConciliarOperacao(xmlOperacao)
                End If
            End If
            
        Case enumStatusOperacao.Expirada, _
             enumStatusOperacao.Rejeitada, _
             enumStatusOperacao.Cancelada, _
             enumStatusOperacao.Estornada
            
             Select Case strCodigoMensagemSPB
                                                       
                Case "BMA0006R1"
                                        
                    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
                       
                       Case enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
                            enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
                            enumTipoOperacaoLQS.EspecTermoIntermediacao, _
                            enumTipoOperacaoLQS.EspecTermoCobertura, _
                            enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
                            enumTipoOperacaoLQS.EspecCompromissadaIntermediacao
                       
                            blnGerenciar = True
                            xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = xmlStatus.selectSingleNode("//Status_Operacao").Text
                            
                       Case enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                            enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
                            enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA
                            
                             'Damanda A822
                            'Alteração no SLCC para permitir o tratamento correto dos status no SLCC ligados à BMA0004 e à BMA006 (Eventos)
                            'A mensagem BMA0002 deverá ficar com status A Conciliar
                            Call objMesgRecbEnviSPB.AlterarStatusBMA0002LDL0001(vntSequenciaOperacao, enumStatusMensagem.AConciliar)
                            
                    End Select
             End Select
             
             If blnGerenciar Then
                If Val(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> enumStatusOperacao.Rejeitada Then
                    strSituacaoRecebida = vbNullString
                End If
                
                GerenciarChamada xmlOperacao, _
                                 xmlMensagem, _
                                 CLng(IIf(xmlStatus.selectSingleNode("//Status_Operacao").Text = "", 0, xmlStatus.selectSingleNode("//Status_Operacao").Text)), _
                                 0, _
                                 IIf(xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = "", "", xmlStatus.selectSingleNode("//Situacao_Mensagem").Text), _
                                 enumIndicadorSimNao.Sim, _
                                 strNumeroControleIF, _
                                 , _
                                 blnAtualizaStatusOper, _
                                 strSituacaoRecebida
            End If
    End Select
    
    Set objMensagem = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing
    Set objProcessoOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set objMensagem = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMesgTratadaR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-----------------------------------------------------------------------------
'Gerenciamento de Envio :
'   - Remessas A6 e PJ
'   - Mensagens de resposta ao legado
'   - Disponibilizar lançamento conta corrente
'   - Disponibilizar Alerta ao Usuário
'-----------------------------------------------------------------------------

Public Function GerenciarChamada(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                 ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                 ByVal plngStatusOperacao As Long, _
                                 ByVal plngStatusMensagem As Long, _
                                 ByVal pstrSituacaoMensagem As String, _
                                 ByVal plngEstorno As enumIndicadorSimNao, _
                                 ByVal pstrNumeroControleIF As String, _
                        Optional ByVal pblnPagamentoContingencia As Boolean = False, _
                        Optional ByVal pblnAtualizaStatusOperacao As Boolean = True, _
                        Optional ByVal pstrSituacaoRecebida As String = vbNullString, _
                        Optional ByVal pstrHeaderNZ As String, _
                        Optional ByVal pintTipoBackOffice As Integer) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objLegado                               As A8LQS.clsLegado
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objAlerta                               As A8LQS.clsAlerta
Dim objContabilidade                        As A8LQS.clsContabilidade

Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim lngStatusOperacao                       As Long
Dim blnOperSemModalidade                    As Boolean
Dim blnEnviaA6RelzConf                      As Boolean
Dim blnEnviaMensagemRetorno                 As Boolean
Dim lngTipoOperacao                         As Long
Dim strCodigoMensagem                       As String
Dim intCanalVenda                           As Integer
Dim str_CO_ULTI_SITU_PROC_ORIGINAL          As String
Dim blnOperCompromissadaCETIPLiquidComitente As Boolean
Dim intTipoContraparte                      As Integer
Dim strSitOpCTP                             As String
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim strLancamentoCCContabil                 As String
Dim str_CO_ULTI_SITU_PROC_GRAVADO           As String

    On Error GoTo ErrorHandler

    'obtem conteudo tags
    If Not xmlOperacao Is Nothing Then
        If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
            lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
        End If
        If Not xmlOperacao.selectSingleNode("//TP_CNPT") Is Nothing Then
            intTipoContraparte = Val(xmlOperacao.selectSingleNode("//TP_CNPT").Text)
        End If
    End If
    If Not xmlMensagem Is Nothing Then
        If Not xmlMensagem.selectSingleNode("//SitOpCTP") Is Nothing Then
            strSitOpCTP = xmlMensagem.selectSingleNode("//SitOpCTP").Text
        End If
    End If

    'Alterar Status Operação
    
    lngStatusOperacao = plngStatusOperacao
    
    If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        
        If Not xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            
            If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text = "0" Then
               
                blnOperSemModalidade = True
               
                If lngStatusOperacao = enumStatusOperacao.Registrada Then
                    lngStatusOperacao = enumStatusOperacao.Liquidada
                    plngStatusOperacao = enumStatusOperacao.Liquidada
                End If
                               
                If Not xmlMensagem Is Nothing Then
                    If Not xmlMensagem.selectSingleNode("//SitOpCTP") Is Nothing Then
                        If xmlMensagem.selectSingleNode("//SitOpCTP").Text = "23" Then
                            Exit Function
                        End If
                    End If
                End If
            Else
                blnOperSemModalidade = False
                
                If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroContratoTermo Then
                    
                    If lngStatusOperacao = enumStatusOperacao.Registrada Then
                        lngStatusOperacao = enumStatusOperacao.Liquidada
                        plngStatusOperacao = enumStatusOperacao.Liquidada
                    End If
                    
                Else
                
                ''PROJETO INR - NICK - LQC (29/12/2014)
                ' Incluindo a situação CTP0054R1 para não encerrar o fluxo de tratamento da mensagem

                    If Not xmlMensagem Is Nothing Then
                        If Not xmlMensagem.selectSingleNode("//SitOpCTP") Is Nothing Then
                            If CLng("0" & xmlMensagem.selectSingleNode("//SitOpCTP").Text) = enumOperacaoCETIP_CTP0100.EstornoFinalizado And _
                               xmlMensagem.selectSingleNode("//CodMsg").Text <> "CTP0100R1" And _
                               xmlMensagem.selectSingleNode("//CodMsg").Text <> "CTP0012R1" And _
                               xmlMensagem.selectSingleNode("//CodMsg").Text <> "CTP0054R1" Then
                                Exit Function
                            End If
                        End If
                    End If
                    
                End If
            End If
        Else
            
            blnOperSemModalidade = True
            
            If lngStatusOperacao = enumStatusOperacao.Registrada Then
                lngStatusOperacao = enumStatusOperacao.Liquidada
                plngStatusOperacao = enumStatusOperacao.Liquidada
            End If
            
            If Not xmlMensagem Is Nothing Then
                If Not xmlMensagem.selectSingleNode("//SitOpCTP") Is Nothing Then
                    If xmlMensagem.selectSingleNode("//SitOpCTP").Text = "23" Then
                        Exit Function
                    End If
                End If
            End If
            
            If lngStatusOperacao <> enumStatusOperacao.Cancelada And _
               lngStatusOperacao <> enumStatusOperacao.Rejeitada And _
               lngStatusOperacao <> enumStatusOperacao.Expirada Then
                
                If Not xmlOperacao Is Nothing Then
                    If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroContratoSWAP Then
                        If Not xmlOperacao.selectSingleNode("//IN_TIPO_CNTR") Is Nothing Then
                             If UCase$(xmlOperacao.selectSingleNode("//IN_TIPO_CNTR").Text) = "S" Then
                                If lngStatusOperacao <> enumStatusOperacao.Liquidada Then
                                    lngStatusOperacao = enumStatusOperacao.Pendencia
                                    plngStatusOperacao = enumStatusOperacao.Pendencia
                                End If
                             End If
                        End If
                    End If
                End If
            End If
        End If
    
    ElseIf xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
        
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            Call fgAppendNode(xmlOperacao, xmlOperacao.documentElement.nodeName, "TP_LIQU_OPER_ATIV", "")
        End If
        
    Else
        blnOperSemModalidade = False
    End If
        
    If lngStatusOperacao > 0 And pblnAtualizaStatusOperacao Then
    
     ''PROJETO INR - NICK - LQC (29/12/2014)
     ' Se a CTP0054R1 chegar, strSitOpCTP = 43 e for TP_LIQU_OPER_ATIV = 2 (Bruta) - altera o status para liquidada.
        
        If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            If Trim(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = "2" Then
                If strSitOpCTP = "43" Then
                    lngStatusOperacao = enumStatusOperacao.Liquidada
                    plngStatusOperacao = enumStatusOperacao.Liquidada
                End If
            End If
        End If
        
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      plngStatusOperacao, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , _
                                      , _
                                      False, _
                                      , _
                                      pstrSituacaoRecebida
        Set objOperacao = Nothing
        str_CO_ULTI_SITU_PROC_ORIGINAL = xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        str_CO_ULTI_SITU_PROC_GRAVADO = plngStatusOperacao
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
    End If
    
    '=================================================================
    'verifica regras específicas para antecipar a integração de C/C
    'se sim, após StatusOperacao ter sido atualizado (código acima),
    'então passa a tratar daqui em diante como StatusOperacao =Liquidada
    '=================================================================
    
    'alteração devido SGC (Bruno Oliveira - 18/out/2011)
    'Msg=CTP9001R1 E StatusOperacao=PendenteEstrategia
    If Not xmlMensagem Is Nothing Then
        If xmlMensagem.selectSingleNode("//CodMsg").Text = "CTP9001R1" Then
            If lngStatusOperacao = enumStatusOperacao.PendenteEstrategia Then
                lngStatusOperacao = enumStatusOperacao.Liquidada
                plngStatusOperacao = enumStatusOperacao.Liquidada
                xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
            End If
        End If
    End If
    
    'Compromissadas CETIP - Liquidação de Comitente (Bruno Oliveira - 08/fev/2012)
    'Oper=CetipsCompromissada E Backoffice=Tesouraria E TipoContraparte=Cliente1 E SitOpCTP = 33
    If (lngTipoOperacao = enumTipoOperacaoLQS.OperacaoCompromissadaCETIP _
    Or lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP _
    Or lngTipoOperacao = enumTipoOperacaoLQS.OperacaoRetornoCompromissadaCETIP) _
    And pintTipoBackOffice = enumTipoBackOffice.Tesouraria _
    And intTipoContraparte = enumTipoContraparte.Cliente1 Then
                
        blnOperCompromissadaCETIPLiquidComitente = True
        
        If strSitOpCTP = "33" Then
            lngStatusOperacao = enumStatusOperacao.Liquidada
            plngStatusOperacao = enumStatusOperacao.Liquidada
            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
        End If
            
    End If
    '=================================================================

   'Alterar o Status da Mensagem
    If plngStatusMensagem > 0 Then
        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Set xmlMensagemEnviadaRecebida = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagemEnviadaRecebida.loadXML objMensagemEnviadaRecebida.LerMensagemIda(pstrNumeroControleIF)
        
        objMensagemEnviadaRecebida.AlterarStatus xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                 plngStatusMensagem, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                 , _
                                                 , _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text
                                                 
        Set objMensagemEnviadaRecebida = Nothing
    End If
    
    If xmlOperacao Is Nothing Then
        Set objOperacao = Nothing
        Set objProcessoOperacao = Nothing
        Set xmlControleProcessoOperacao = Nothing
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
                                          plngEstorno, _
                                          plngStatusOperacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    '=================================================================
    'verifica se o C/C já foi integrado antes,
    'se sim vai direto para o envio de retorno ao legado
    '=================================================================
    
    'SGC (Bruno Oliveira - 18/out/2011)
    'se CO_ULTI_SITU_PROC =PendenteEstrategia e status da CTP9001R1 for =Liquidada, então vai direto para a função de EnviarRetorno para o legado
    If str_CO_ULTI_SITU_PROC_ORIGINAL = enumStatusOperacao.PendenteEstrategia And lngStatusOperacao = enumStatusOperacao.Liquidada Then
        GoTo ByPass_Para_EnviarRetorno
    End If
    
    'Compromissadas CETIP - Liquidação de Comitente (Bruno Oliveira - 08/fev/2012)
    If blnOperCompromissadaCETIPLiquidComitente = True And lngStatusOperacao = enumStatusOperacao.Liquidada Then
        
        Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
        strLancamentoCCContabil = objLancamentoCCContabil.Ler(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                              enumTipoLancamentoIntegracao.Normal, _
                                                              0)
 
        'strLancamentoCCContabil <> vbNullString significa que já foi feito lançamento de C/C para esta operação
        If strLancamentoCCContabil <> vbNullString Then
            GoTo ByPass_Para_EnviarRetorno
        End If
    End If
    '=================================================================

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
        
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
        
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'a integracao nao deve ser enviada ao A6 e PJ. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
    
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            Set objSubReserva = Nothing
        End If
        
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
        
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            blnEnviaA6RelzConf = False
            If xmlMensagem Is Nothing Then
                blnEnviaA6RelzConf = True
            ElseIf xmlMensagem.selectSingleNode("//CodMsg").Text <> "LTR0002R1" Then
                blnEnviaA6RelzConf = True
            End If
            If blnEnviaA6RelzConf Then
                Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                Set objSubReserva = Nothing
            End If
        End If
        
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If
        End If
    
        'Projeto Sisbacen Interbancario/Arbitragem (Bruno Oliveira - 22/06/2012)
        'Operações de Informa Liquidação Interbancária (CAM0016) sempre devem enviar remessa para o PJ
        If lngTipoOperacao = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
        End If
    
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If
    
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ_ME").Text = enumIndicadorSimNao.Sim Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EnviarRealizadoME(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If
    
    End If
    
    If pblnPagamentoContingencia Then
        xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Nao
    Else
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Nao
            End If
        End If
    End If
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
        'KIDA - CBLC
        If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
            If plngEstorno = enumIndicadorSimNao.Nao Then
                Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
                objContaCorrente.Disponibilizar xmlOperacao, False
                Set objContaCorrente = Nothing
            Else
                Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
                objContaCorrente.ExcluirDisponivel xmlOperacao
                Set objContaCorrente = Nothing
            End If
        Else
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Disponibilizar xmlOperacao, False
            Set objContaCorrente = Nothing
        End If
    End If
    
ByPass_Para_EnviarRetorno:
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        
        blnEnviaMensagemRetorno = True
        
        If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
            If Val(xmlOperacao.selectSingleNode("//IN_ENTR_MANU").Text) = enumIndicadorSimNao.Sim Then
                If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
                    If Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.DespesasCETIP Then
                        blnEnviaMensagemRetorno = False
                    End If
                End If
                If Not xmlOperacao.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                    'Se Operacao foi gerada via Entrada Manual e e uma Operacao de COMEX (SistOrig = GPC, R2, E2A), ou Sistema = E2 e Mensagem CAM, entao nao deve ser enviado mensagem de retorno para o legado // INcluiso BOL (NICK)
                    If Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "GPC" Or Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "R2" Or Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "BOL" Or Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "HQ" Or Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "CMX" Or Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "E2A" Then
                        blnEnviaMensagemRetorno = False
                    End If
                    If Trim(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "E2" And Left(xmlMensagem.selectSingleNode("//CodMsg").Text, 3) = "CAM" Then
                        blnEnviaMensagemRetorno = False
                    End If
                End If
            End If
        Else
            If Not xmlMensagem Is Nothing Then
                If xmlMensagem.selectSingleNode("//CodMsg").Text = "LTR0002R1" Then
                    If plngStatusMensagem = enumStatusMensagem.Discordada Then
                        blnEnviaMensagemRetorno = False
                    End If
                End If
            End If
        End If
        
        '-------------------------------------------------------------------------
        'tratamento específico para operações que fazem integração antecipada de C/C em situações de pendência,
        'para garantir que a situação enviada para o legado (CO_ULTI_SITU_PROC) nessa situação seja 20=Pendência
        'ao invés de 15=Liquidada
        If str_CO_ULTI_SITU_PROC_GRAVADO = enumStatusOperacao.PendenteEstrategia _
        Or blnOperCompromissadaCETIPLiquidComitente = True Then
        
            If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> str_CO_ULTI_SITU_PROC_GRAVADO _
            And str_CO_ULTI_SITU_PROC_GRAVADO <> vbNullString Then
                'o status PendenteEstrategia (=320) é uma exceção, não pode ser enviado para o legado, é um status reconhecido apenas dentro do SLCC,
                'ao invés dele deve ser enviado o status 20=Pendente
                If str_CO_ULTI_SITU_PROC_GRAVADO = enumStatusOperacao.PendenteEstrategia Then
                    str_CO_ULTI_SITU_PROC_GRAVADO = enumStatusOperacao.Pendencia
                End If
                xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = str_CO_ULTI_SITU_PROC_GRAVADO
            End If
        
        End If
        '-------------------------------------------------------------------------
        
        If blnEnviaMensagemRetorno Then
            Set objLegado = CreateObject("A8LQS.clsLegado")
            strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem, pstrHeaderNZ)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, _
                                                 enumIdentificadorFila.BUS, , , pstrHeaderNZ
            Set objLegado = Nothing
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_ALER").Text = enumIndicadorSimNao.Sim Then
        Set objAlerta = CreateObject("A8LQS.clsAlerta")
        objAlerta.GerarAlerta xmlOperacao, _
                              xmlMensagem, _
                              pstrSituacaoMensagem, _
                              0, _
                              plngStatusOperacao
        
        Set objAlerta = Nothing
    End If

    'Estornos de previsão
    objProcessoOperacao.EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao

    If plngEstorno = enumIndicadorSimNao.Sim Then
               
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
        
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
            
            blnEnviaA6RelzConf = True
            
            If Not xmlMensagem Is Nothing Then
                If Not xmlMensagem.selectSingleNode("//CodMsg") Is Nothing Then
                    If xmlMensagem.selectSingleNode("//CodMsg").Text = "LTR0002E" Then
                        blnEnviaA6RelzConf = False
                    End If
                End If
            End If
                       
            If blnEnviaA6RelzConf Then
                Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                Set objSubReserva = Nothing
            End If
        End If
        
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
        
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            strMensagem = objSubReserva.EstornarRealizadoConfirmado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            Set objSubReserva = Nothing
        End If
                
        'Conforme solicitado pela Aline/Ulysses em 30/08/2011
        'Se a tag TP_LIQU_OPER_ATIV nao existir ou estiver em branco,
        'as integracoes nao devem ser enviadas. - Ivan/Bruno
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
        Else
            If Trim$(xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If
        End If
                
        'Projeto Sisbacen Interbancario/Arbitragem (Bruno Oliveira - 22/06/2012)
        'Operações de Informa Liquidação Interbancária (CAM0016) sempre devem enviar remessa para o PJ
        If lngTipoOperacao = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
        End If
                
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If
        
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ_ME").Text = enumIndicadorSimNao.Sim Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            strMensagem = objGestaoCaixa.EstornarRealizadoME(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            Set objGestaoCaixa = Nothing
        End If

        If Not xmlOperacao.selectSingleNode("//CodMsg") Is Nothing Then
            strCodigoMensagem = xmlOperacao.selectSingleNode("//CodMsg").Text
        End If
        
        If strCodigoMensagem = "CTP0100R1" And lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
            Call objProcessoOperacao.EnviarPrevisao(xmlOperacao, xmlControleProcessoOperacao)
        End If
    
        intCanalVenda = 0
        If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
            intCanalVenda = Val(xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text)
        End If
        
        Select Case lngTipoOperacao
            Case enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
                 enumTipoOperacaoLQS.EventosJurosSWAP, _
                 enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, _
                 enumTipoOperacaoLQS.RegDadosComplemContratoSWAP, _
                 enumTipoOperacaoLQS.RegistroContratoSWAP, _
                 enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira
            
                If intCanalVenda = enumCanalDeVenda.SGC And lngStatusOperacao = enumStatusOperacao.Cancelada Then
                    If flVerificarContaCorrente(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) Then
                        
                        If xmlOperacao.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
                            Call fgAppendNode(xmlOperacao, "MESG", "TP_LANC_ITGR", vbNullString)
                        End If
                        
                        xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text = enumTipoLancamentoIntegracao.Estorno
                        
                        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
                        Call objContaCorrente.Disponibilizar(xmlOperacao)
                        Set objContaCorrente = Nothing
                    
                    End If
                End If
            
        End Select

    End If
    
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Verificar a situação do lançamento em conta corrente da operação
'-------------------------------------------------------------------------------------------------

Private Function flVerificarContaCorrente(ByVal pvntNumeroSequencial As Variant) As Boolean

Dim objHistSituacaoIntegracao               As A8LQS.clsHistSituacaoIntegracao
Dim xmlHistSituacaoIntegracao               As MSXML2.DOMDocument40
Dim strHistSituacaoIntegracao               As String

    On Error GoTo ErrorHandler

    flVerificarContaCorrente = False
    
    Set objHistSituacaoIntegracao = CreateObject("A8LQS.clsHistSituacaoIntegracao")

    strHistSituacaoIntegracao = objHistSituacaoIntegracao.Ler( _
                                    "A.NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & " AND " & _
                                    "A.CO_SITU_PROC IN (" & enumStatusIntegracao.Integrado & ", " & _
                                                            enumStatusIntegracao.IntegradoCC & ")")
                            
    If strHistSituacaoIntegracao <> vbNullString Then
        flVerificarContaCorrente = True
    End If

    Set objHistSituacaoIntegracao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objHistSituacaoIntegracao = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flVerificarContaCorrente", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'-----------------------------------------------------------------------------
'Obter o nome da tag de repetição das mensagens com repetições para loop de grupo.
'-----------------------------------------------------------------------------

Public Function ObterNomeTagRepeticao(ByVal pstrCodigoMensagem As String, _
                                      ByVal plngNivelRepeticao As Long) As String
                                      
Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
                                      
On Error GoTo ErrorHandler
                                          
    strSQL = " SELECT    NO_TAG_REPT " & _
             "   FROM    A8.TB_CTRL_RETN_MESG " & _
             "  WHERE    CO_MESG_RECB ='" & pstrCodigoMensagem & "' " & _
             "    AND    NU_NIVE_REPE = " & plngNivelRepeticao
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        ObterNomeTagRepeticao = vbNullString
        lngCodigoErroNegocio = 3049
        GoTo ErrorHandler
    Else
        ObterNomeTagRepeticao = Trim(objRS!NO_TAG_REPT)
    End If
                                      
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNomeTagRepeticao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                                      
End Function

'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LDL0009R2
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0009R2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice

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

Dim strNomeTagResultLiqd                    As String
Dim strNETCnpj                              As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeNet                              As MSXML2.IXMLDOMNode

Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String
Dim lngTipoPagamentoLDL                     As Long
Dim vntIdentContraparteCamara               As Variant

On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

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
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL0005R2_ResultLiqd"
        
    If Not xmlMensagem.selectSingleNode("//NumCtrlLDLOr") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text
    Else
        strNumeroControleCamara = vbNullString
    End If
    
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrResultLiqdNLiqdant = xmlMensagem.selectSingleNode("//VlrLanc").Text
        
    'Incluir Mensagem Pai
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
        lngStatusMensagem = flObterStatusLDL0007(xmlMensagem, lngCodigoEmpresa)
    Else
        lngStatusMensagem = flObterStatusLDL0005R2_LDL0009R2(xmlMensagem, lngCodigoEmpresa)
    End If
    
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        

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
                                vntVlrResultLiqdNLiqdant, _
                                lngSeqControleRepeticao, _
                                "", _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                0, _
                                , _
                                0, _
                                "", , 0, , , enumTipoLiquidacao.Multilateral)

    
    'Substitui os grupos de repeticao ,com os grupos por net (CNPJ)
    
    Set xmlNodeNet = flPreparaMensagemLDL0009R2(xmlMensagem, lngCodigoLocalLiquidacao)
    
    Set xmlNode = xmlMensagem.selectSingleNode("//" & strNomeTagResultLiqd)
    
    If Not xmlNode Is Nothing Then
        xmlMensagem.selectSingleNode("//" & xmlMensagem.documentElement.childNodes(0).nodeName).removeChild xmlNode
        xmlMensagem.selectSingleNode("//" & xmlMensagem.documentElement.childNodes(0).nodeName).appendChild xmlNodeNet
    End If
    
    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
            If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
                vntIdentContraparteCamara = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
            Else
                vntIdentContraparteCamara = vbNullString
            End If
        Else
            vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
        End If
        
        vntVlrResultLiqdNLiqdant = xmlNodeResultLiqd.selectSingleNode("VlrNLiqdant").Text
        
        If Not xmlNodeResultLiqd.selectSingleNode("TpPgtoLDL") Is Nothing Then
            lngTipoPagamentoLDL = xmlNodeResultLiqd.selectSingleNode("TpPgtoLDL").Text
        End If
           
        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
            
            Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
            lngTipoBackOffice = objTipoBackOffice.ObterTipoBackOfficeParticipanteCamara(vntIdentContraparteCamara)
            vntCNPJNLiqdant = vbNullString
            strCodigoVeiculoLegal = vbNullString
            strSiglaSistemaVeicLega = vbNullString
            strNomeVeicLega = vbNullString
            
            
            Set objTipoBackOffice = Nothing
        Else
           
            If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                         vntCNPJNLiqdant, _
                                         lngCodigoEmpresa, _
                                         strCodigoVeiculoLegal, _
                                         strSiglaSistemaVeicLega, _
                                         lngTipoBackOffice, _
                                         strNomeVeicLega) Then
            
            
                Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                        lngTipoBackOffice, _
                                        strNomeVeicLega, _
                                        vntVlrResultLiqdNLiqdant, _
                                        xmlMensagem)
            
            End If
        End If
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem, _
                                           0, vntIdentContraparteCamara)

        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    "", _
                                    enumTipoDebitoCredito.Credito, _
                                    "", _
                                    0, _
                                    "", _
                                    0, _
                                    "", 0, 0, "", "", _
                                    enumTipoLiquidacao.Multilateral, _
                                    "", 0, 0, 0, "", 0, 0, "", _
                                    lngTipoPagamentoLDL)
                                    
    Next

    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0009R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processamento das mensagens LTR0007 (Transferência Liquidação Eventos e Emissões LDL)
'   - Para cada CNPJ não Liquidante é gerado um registro no SLCC para conciliação e liquidação.
'   - Veiculo Legal obtido pelo número do CNPJ.
'--------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0007(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice

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

Dim strNomeTagResultLiqd                    As String
Dim strNETCnpj                              As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeNet                              As MSXML2.IXMLDOMNode

Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String
Dim strDtLiquid                             As String
Dim lngTipoPagtoLDL                         As Long
Dim vntCodPartCamara                        As Variant

On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

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
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL0007_EvtEms"
        
    If Not xmlMensagem.selectSingleNode("//NumCtrlLDL") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlLDL").Text
    Else
        strNumeroControleCamara = vbNullString
    End If
    
    strDtLiquid = xmlMensagem.selectSingleNode("//DtLiquid").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    vntVlrResultLiqdNLiqdant = xmlMensagem.selectSingleNode("//VlrLanc").Text
        
    'Incluir Mensagem Pai
    lngStatusMensagem = flObterStatusLDL0007(xmlMensagem, lngCodigoEmpresa)
    
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        

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
                                vntVlrResultLiqdNLiqdant, _
                                lngSeqControleRepeticao, _
                                strDtLiquid, _
                                enumTipoDebitoCredito.Credito, _
                                "", _
                                0, _
                                , _
                                0, _
                                "", _
                                0, _
                                0, _
                                "", _
                                "", _
                                enumTipoLiquidacao.Multilateral)

    
    'Substitui os grupos de repeticao ,com os grupos por net (CNPJ)
    
    Set xmlNodeNet = flPreparaMensagemLDL0009R2(xmlMensagem, lngCodigoLocalLiquidacao)
    
    Set xmlNode = xmlMensagem.selectSingleNode("//" & strNomeTagResultLiqd)
    
    If Not xmlNode Is Nothing Then
        xmlMensagem.selectSingleNode("//" & xmlMensagem.documentElement.childNodes(0).nodeName).removeChild xmlNode
    
        xmlMensagem.selectSingleNode("//" & xmlMensagem.documentElement.childNodes(0).nodeName).appendChild xmlNodeNet
    
    End If
    
    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        'vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodeResultLiqd.selectSingleNode("VlrNLiqdant").Text
        lngTipoPagtoLDL = xmlNodeResultLiqd.selectSingleNode("TpPgtoLDL").Text
        
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntCodPartCamara = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
        Else
            vntCodPartCamara = vbNullString
        End If
           
        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
            Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
            lngTipoBackOffice = objTipoBackOffice.ObterTipoBackOfficeParticipanteCamara(vntCodPartCamara)
            vntCNPJNLiqdant = vbNullString
            strCodigoVeiculoLegal = vbNullString
            strSiglaSistemaVeicLega = vbNullString
            strNomeVeicLega = vbNullString
            
            
            Set objTipoBackOffice = Nothing
        Else
           
            If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                         vntCNPJNLiqdant, _
                                         lngCodigoEmpresa, _
                                         strCodigoVeiculoLegal, _
                                         strSiglaSistemaVeicLega, _
                                         lngTipoBackOffice, _
                                         strNomeVeicLega) Then
                
                Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                        lngTipoBackOffice, _
                                        strNomeVeicLega, _
                                        vntVlrResultLiqdNLiqdant, _
                                        xmlMensagem)
                
            End If
        End If
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem, _
                                           0, _
                                           vntCodPartCamara)
        
                
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    enumTipoDebitoCredito.Credito, _
                                    "", _
                                    0, _
                                    "", _
                                    0, _
                                    "", , , , , _
                                    enumTipoLiquidacao.Multilateral, _
                                    "", "", 0, 0, "", 0, 0, "", _
                                    lngTipoPagtoLDL)
                                    
    Next

    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objTipoBackOffice = Nothing
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0007 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'------------------------------------------------------------------------------------------------------------
'Processamento das mensagem LDL0006R2 (Resposta a Terceiros sobre Transferência por Devolução de Créditos IF)
'------------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0006R2(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim strNomeTagDevCred                       As String
Dim xmlNodeDevCred                          As MSXML2.IXMLDOMNode

Dim strSiglaSistemaVeicLega                 As String
Dim strCodigoVeiculoLegal                   As String
Dim lngCodigoEmpresa                        As Long
Dim lngSeqControleRepeticao                 As Long
Dim lngTipoBackOffice                       As Long
Dim strCodigoMensagemSPB                    As String
Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String
Dim lngCodigoLocalLiquidacao                As Long

On Error GoTo ErrorHandler
        
    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD Or _
       lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemInformacao)
    Else
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemR2)
    End If
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    
    strNomeTagDevCred = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1)
        
    lngSeqControleRepeticao = 0
    
    'Repeticao Titulos
    For Each xmlNodeDevCred In xmlMensagem.selectNodes("//" & strNomeTagDevCred & "/*")
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        vntCNPJNLiqdant = xmlNodeDevCred.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodeDevCred.selectSingleNode("VlrNLiqdant").Text
           
        If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice, _
                                     strNomeVeicLega) Then
                                    
            Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrResultLiqdNLiqdant, _
                                    xmlMensagem)
        End If
           
    Next

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0006R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------------------------
'Geração de Alerta para o grupo de Usuario
'------------------------------------------------------------------------------------------------------------

Private Sub flGeraAlertaCamara(ByVal pstrCodigoMensagem As String, _
                               ByVal plngTipoBackOffice As Long, _
                               ByVal pstrNomeVeiculoLegal As String, _
                               ByVal pstrValor As String, _
                               ByRef xmlMensagemSPB As MSXML2.DOMDocument40, _
                      Optional ByVal plngStatusMensagem As Long = 0, _
                      Optional ByVal vntNumeroSeqOperacao As Variant = "")
                               
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlPropriedadesAlerta                   As MSXML2.DOMDocument40
Dim lngCodFatorGeraAlerta                   As Long
                              
On Error GoTo ErrorHandler
                               
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    Set xmlPropriedadesAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlPropriedadesAlerta.loadXML objAlerta.ObterPropriedades
    
    Select Case pstrCodigoMensagem
        Case "BMC0011"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0011
        Case "LDL0001"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0001
        Case "LDL0005R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0005R2
        Case "LDL0006R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0006R2
        Case "LDL0009R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0009R2
        Case "LTR0001"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0001
        Case "LTR0007"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0007
        Case "LTR0005R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0005R2
        Case "LTR0006R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0006R2
        Case "LTR0009R1"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0009R1
        Case "BMA0002"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMA0002
        Case "BMA0003"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMA0003
        Case "BMA0013"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMA0013
        Case "BMA0014"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMA0014
        Case "BMA0015"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMA0015
        Case "CTP1002"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemCTP1002
        Case "LDL1027"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL1027
        Case "CTP0100"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.CTPPendente
        Case "LDL0026R1"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0026R1
        Case "BMC0101"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0101
        Case "BMC0012"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0012
        Case "BMC0005"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0005
        Case "LDL0003R1"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLDL0003R1
        Case "LTR0003R2"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemLTR0003R2
        Case "BMC0013"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0013
        Case "BMC0002R1"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0002R1
        Case "BMC0002E"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0002E
        Case "BMC0003R1"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0003R1
        Case "BMC0003E"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0003E
        Case "BMC0015SOPER"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0015SemOperacao
        Case "BMC0015DOPER"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0015DivergOperacao
        Case "BMC0005SOPER"
            lngCodFatorGeraAlerta = enumFatorGeradorAlerta.MensagemBMC0005SemOperacao
        Case Else
            Exit Sub
    End Select
    
    xmlPropriedadesAlerta.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = vntNumeroSeqOperacao
    xmlPropriedadesAlerta.selectSingleNode("//CO_FATO_GERA_ALER").Text = lngCodFatorGeraAlerta
    xmlPropriedadesAlerta.selectSingleNode("//TP_BKOF").Text = plngTipoBackOffice
    xmlPropriedadesAlerta.selectSingleNode("//NO_VEIC_LEGA").Text = pstrNomeVeiculoLegal
    xmlPropriedadesAlerta.selectSingleNode("//VA_OPER_ATIV").Text = pstrValor
    xmlPropriedadesAlerta.selectSingleNode("//TX_ANEX").Text = xmlMensagemSPB.selectSingleNode("//SISMSG").xml

    Call objAlerta.GerarAlertaCamara(xmlPropriedadesAlerta)
                               
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    Exit Sub
ErrorHandler:
    
    Set objAlerta = Nothing
    Set xmlPropriedadesAlerta = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGeraAlertaCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'------------------------------------------------------------------------------------------------------------
'Processamento das mensagem LTR0009R1 (Resposta à Consulta Posição financeira LTR)
'------------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0009R1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagPosicao                       As String
Dim xmlNodePosicao                          As MSXML2.IXMLDOMNode

Dim strISPBIFLiqdant                        As String
Dim strNumOpLTR                             As String
Dim intTpDeb_Cred                           As String
Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String

On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = 0
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strSituacaoMensagemSPB = ""
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    
    strNomeTagPosicao = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) 'Repet_LTR0009R1_ConsPos
    
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
        
    lngSeqControleRepeticao = 1
    
    'Salvar Mensagem Enviada Recebida
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    'lngStatusMensagem = flObterStatusLTR0009R1_LDL0026R1(xmlMensagem, lngCodigoEmpresa)
    
    'Repeticao Titulos
    For Each xmlNodePosicao In xmlMensagem.selectNodes("//" & strNomeTagPosicao & "/*")
        
        strNumeroControleCamara = xmlNodePosicao.selectSingleNode("NumCtrlLTROr").Text
        vntCNPJNLiqdant = xmlNodePosicao.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodePosicao.selectSingleNode("VlrLanc").Text
        
        If xmlNodePosicao.selectSingleNode("TpDeb_Cred").Text = "D" Then
            intTpDeb_Cred = enumTipoDebitoCredito.Debito
        Else
            intTpDeb_Cred = enumTipoDebitoCredito.Credito
        End If
           
        If Not xmlNodePosicao.selectSingleNode("NumOpLTR") Is Nothing Then
            strNumOpLTR = xmlNodePosicao.selectSingleNode("NumOpLTR").Text
        Else
            strNumOpLTR = ""
        End If
           
        If Not xmlNodePosicao.selectSingleNode("ISPBIFLiqdant") Is Nothing Then
            strISPBIFLiqdant = xmlNodePosicao.selectSingleNode("ISPBIFLiqdant").Text
        Else
            strISPBIFLiqdant = ""
        End If
           
        If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice, _
                                     strNomeVeicLega) Then
                                        
            lngStatusMensagem = enumStatusMensagem.AConciliar
            
            Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrResultLiqdNLiqdant, _
                                    xmlMensagem, _
                                    0, _
                                    "")
            
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        End If
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodePosicao)
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                                Trim(strNumeroControleIF), _
                                                                0, _
                                                                0, _
                                                                strCodigoMensagemSPB, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                strNumOpLTR, _
                                                                "", _
                                                                lngStatusMensagem, _
                                                                lngIndicadorEntradaManual, _
                                                                "", _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                lngSeqControleRepeticao, _
                                                                strNumeroControleCamara, _
                                                                "", _
                                                                0)
        
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    "", _
                                    intTpDeb_Cred, _
                                    "", _
                                    vntCNPJNLiqdant, _
                                    "", _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    "", _
                                    strISPBIFLiqdant, _
                                    0, _
                                    "", _
                                    "")
                                    
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0009R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------------------------
'Montagem da mensagem somente com o node do grupo para as mensagens com repetição
'------------------------------------------------------------------------------------------------------------

Private Function flMontaMensagemNode(ByVal pstrCodigoMensagem As String, _
                                     ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                     ByRef pxmlNode As MSXML2.IXMLDOMNode) As String


Dim strNomeTagRepet                         As String
Dim xmlNodeRepeticao                        As MSXML2.DOMDocument40
Dim strNomeTagPrincipal                     As String
Dim blnAppendXML                            As Boolean

On Error GoTo ErrorHandler
    
    Set xmlNodeRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
    
    strNomeTagPrincipal = pxmlMensagem.documentElement.childNodes(0).nodeName
    
    Call fgAppendNode(xmlNodeRepeticao, "", "SISMSG", "")
    Call fgAppendNode(xmlNodeRepeticao, "SISMSG", strNomeTagPrincipal, "")
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodMsg", pstrCodigoMensagem)
    
    blnAppendXML = True
    
    Select Case pstrCodigoMensagem
        Case "BMA0015"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlBMA", pxmlMensagem.selectSingleNode("//NumCtrlBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtLiquid").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJEntRespons", pxmlMensagem.selectSingleNode("//CNPJEntRespons").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpTitlar", pxmlMensagem.selectSingleNode("//TpTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlar", pxmlMensagem.selectSingleNode("//CodTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CtCstdiaBMA", pxmlMensagem.selectSingleNode("//CtCstdiaBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "FinlddBMA", pxmlMensagem.selectSingleNode("//FinlddBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrSit", pxmlMensagem.selectSingleNode("//DtHrSit").Text)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "BMA0013"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlBMA", pxmlMensagem.selectSingleNode("//NumCtrlBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJEntRespons", pxmlMensagem.selectSingleNode("//CNPJEntRespons").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpTitlar", pxmlMensagem.selectSingleNode("//TpTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlar", pxmlMensagem.selectSingleNode("//CodTitlar").Text)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrSit", pxmlMensagem.selectSingleNode("//DtHrSit").Text)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "LDL0001"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLDL", pxmlMensagem.selectSingleNode("//NumCtrlLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpInf", pxmlMensagem.selectSingleNode("//TpInf").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtLiquid").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpDeb_Cred", pxmlMensagem.selectSingleNode("//TpDeb_Cred").Text)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "LDL1001"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLDL", pxmlMensagem.selectSingleNode("//NumCtrlLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlar", pxmlMensagem.selectSingleNode("//CodTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpInf", "P")
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpDeb_Cred", enumTipoDebitoCredito.Credito)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "BMA0002"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlBMA", pxmlMensagem.selectSingleNode("//NumCtrlBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlBMAEspecc", pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJEntRespons", pxmlMensagem.selectSingleNode("//CNPJEntRespons").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpTitlar", pxmlMensagem.selectSingleNode("//TpTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlar", pxmlMensagem.selectSingleNode("//CodTitlar").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJPart", pxmlMensagem.selectSingleNode("//CNPJPart").Text)
            
            'KIDA - 18/02/2008
            If Not pxmlMensagem.selectSingleNode("//CNPJMC") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJMC", pxmlMensagem.selectSingleNode("//CNPJMC").Text)
            End If
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "IdentdSistNegc", pxmlMensagem.selectSingleNode("//IdentdSistNegc").Text)
            
            If Not pxmlMensagem.selectSingleNode("//CNPJCtrapartNegc") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJCtrapartNegc", pxmlMensagem.selectSingleNode("//CNPJCtrapartNegc").Text)
            End If
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpNegcBMA", pxmlMensagem.selectSingleNode("//NumOpNegcBMA").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodOperadSistNegc", pxmlMensagem.selectSingleNode("//CodOperadSistNegc").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpTitlarOferte", pxmlMensagem.selectSingleNode("//TpTitlarOferte").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlarOferte", pxmlMensagem.selectSingleNode("//CodTitlarOferte").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "HrFchtSistNegc", pxmlMensagem.selectSingleNode("//HrFchtSistNegc").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpNegcBMA", pxmlMensagem.selectSingleNode("//TpNegcBMA").Text)
            
            If Not pxmlMensagem.selectSingleNode("//RodaNegcBMA") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "RodaNegcBMA", pxmlMensagem.selectSingleNode("//RodaNegcBMA").Text)
            End If
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NatuOp", pxmlMensagem.selectSingleNode("//NatuOp").Text)
            
            If Not pxmlMensagem.selectSingleNode("//NumOpSELMigr") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpSELMigr", pxmlMensagem.selectSingleNode("//NumOpSELMigr").Text)
            End If
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "LDL0005R2"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlSTR", pxmlMensagem.selectSingleNode("//NumCtrlSTR").Text)
            
            If Not pxmlMensagem.selectSingleNode("//NumCtrlLDLOr") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLDLOr", pxmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text)
            End If
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "LDL0009R2"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlSTR", pxmlMensagem.selectSingleNode("//NumCtrlSTR").Text)
            
            If Not pxmlMensagem.selectSingleNode("//DtHrBC") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrBC", pxmlMensagem.selectSingleNode("//DtHrBC").Text)
            End If
            
            If Not pxmlMensagem.selectSingleNode("//ISPBLDL") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrBC", pxmlMensagem.selectSingleNode("//ISPBLDL").Text)
            End If
            
            If Not pxmlMensagem.selectSingleNode("//NumCtrlLDLOr") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLDLOr", pxmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text)
            End If
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
            
        Case "LTR0009R1"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlIF", pxmlMensagem.selectSingleNode("//NumCtrlIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBLTR", pxmlMensagem.selectSingleNode("//ISPBLTR").Text)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
                        
            blnAppendXML = False
                        
        Case "LDL0026R1"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlIF", pxmlMensagem.selectSingleNode("//NumCtrlIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBLDL", pxmlMensagem.selectSingleNode("//ISPBLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpInf", pxmlMensagem.selectSingleNode("//TpInf").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrLDL", pxmlMensagem.selectSingleNode("//DtHrLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtLiquid").Text)
    
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
        
        Case "LTR0007D"
            
            xmlNodeRepeticao.selectSingleNode("//CodMsg").Text = "LTR0007"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLTR", pxmlMensagem.selectSingleNode("//NumCtrlLTR").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBLTR", pxmlMensagem.selectSingleNode("//ISPBLTR").Text)
                        
            If Not pxmlMensagem.selectSingleNode("//Hist") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "Hist", pxmlMensagem.selectSingleNode("//Hist").Text)
            End If
        
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJNLiqdantDebtd", pxmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text)
            
            If Not pxmlMensagem.selectSingleNode("//AgDebtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "AgDebtd", pxmlMensagem.selectSingleNode("//AgDebtd").Text)
            End If
            
            If Not pxmlMensagem.selectSingleNode("//CtDebtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CtDebtd", pxmlMensagem.selectSingleNode("//CtDebtd").Text)
            End If
 
            If Not pxmlMensagem.selectSingleNode("//IdentdPartCamrDebtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "IdentdPartCamrDebtd", pxmlMensagem.selectSingleNode("//IdentdPartCamrDebtd").Text)
            End If
 
            If Not pxmlMensagem.selectSingleNode("//NumOpLTRDebtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpLTRDebtd", pxmlMensagem.selectSingleNode("//NumOpLTRDebtd").Text)
            End If
 
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "VlrLanc", pxmlMensagem.selectSingleNode("//VlrLanc").Text)

            If Not pxmlMensagem.selectSingleNode("//NumOpLTR") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpLTR", pxmlMensagem.selectSingleNode("//NumOpLTR").Text)
            End If

            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "IndrConfc", pxmlMensagem.selectSingleNode("//IndrConfc").Text)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
        
        Case "LTR0007C"
            
            xmlNodeRepeticao.selectSingleNode("//CodMsg").Text = "LTR0007"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLTR", pxmlMensagem.selectSingleNode("//NumCtrlLTR").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBLTR", pxmlMensagem.selectSingleNode("//ISPBLTR").Text)
                        
            If Not pxmlMensagem.selectSingleNode("//Hist") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "Hist", pxmlMensagem.selectSingleNode("//Hist").Text)
            End If
        
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJNLiqdantCredtd", pxmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd").Text)
            
            If Not pxmlMensagem.selectSingleNode("//AgCredtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "AgCredtd", pxmlMensagem.selectSingleNode("//AgCredtd").Text)
            End If
            
            If Not pxmlMensagem.selectSingleNode("//CtCredtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CtCredtd", pxmlMensagem.selectSingleNode("//CtCredtd").Text)
            End If
 
            If Not pxmlMensagem.selectSingleNode("//IdentdPartCamrCredtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "IdentdPartCamrCredtd", pxmlMensagem.selectSingleNode("//IdentdPartCamrCredtd").Text)
            End If
 
            If Not pxmlMensagem.selectSingleNode("//NumOpLTRCredtd") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpLTRCredtd", pxmlMensagem.selectSingleNode("//NumOpLTRCredtd").Text)
            End If
 
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "VlrLanc", pxmlMensagem.selectSingleNode("//VlrLanc").Text)

            If Not pxmlMensagem.selectSingleNode("//NumOpLTR") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumOpLTR", pxmlMensagem.selectSingleNode("//NumOpLTR").Text)
            End If

            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "IndrConfc", pxmlMensagem.selectSingleNode("//IndrConfc").Text)
             
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
             
            blnAppendXML = False
        
        Case "LDL0007"
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlLDL", pxmlMensagem.selectSingleNode("//NumCtrlLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBLDL", pxmlMensagem.selectSingleNode("//ISPBLDL").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtLiquid").Text)
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "VlrLanc", pxmlMensagem.selectSingleNode("//VlrLanc").Text)
            
            strNomeTagRepet = ObterNomeTagRepeticao(pstrCodigoMensagem, 1)
            
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, strNomeTagRepet, "")
            
            Call fgAppendXML(xmlNodeRepeticao, strNomeTagRepet, pxmlNode.xml)
            
            If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
                Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
            End If
            
            blnAppendXML = False
    
    End Select
    
    If blnAppendXML Then
        Call fgAppendXML(xmlNodeRepeticao, strNomeTagPrincipal, pxmlNode.xml)
        
        If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
            Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
        End If
    End If
    
    flMontaMensagemNode = xmlNodeRepeticao.xml
    
    Set xmlNodeRepeticao = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlNodeRepeticao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaMensagemNode Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------------------------
'Obter o codigo da mensagem que esta sendo liquidada (  Ex. LTR0004 liquida a mensagem LTR0001)
'------------------------------------------------------------------------------------------------------------

Private Function flObterCodigoMensagemMae(ByVal pstrCodigoMensagemEnvio As String) As String

Dim objFluxoProcesso                        As A8LQS.clsFluxoProcessamento
Dim xmlFluxoProcesso                        As MSXML2.DOMDocument40
Dim strFluxoProcesso                        As String

On Error GoTo ErrorHandler:
    
    Set xmlFluxoProcesso = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objFluxoProcesso = CreateObject("A8LQS.clsFluxoProcessamento")
    
    strFluxoProcesso = objFluxoProcesso.LerTodos("", pstrCodigoMensagemEnvio, 0)
    
    If strFluxoProcesso <> "" Then
        xmlFluxoProcesso.loadXML strFluxoProcesso
    
        flObterCodigoMensagemMae = xmlFluxoProcesso.selectSingleNode("//CO_MESG_SPB_RECB").Text
    End If
    
    Set objFluxoProcesso = Nothing
    Set xmlFluxoProcesso = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlFluxoProcesso = Nothing
    Set objFluxoProcesso = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterCodigoMensagemMae Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'------------------------------------------------------------------
'Obter os dados da remessa (Operação Original) que gerou a mensagem
'------------------------------------------------------------------

Public Function ObterOperacaoOriginal(ByVal pstrNumeroControleIF As String, _
                                      ByVal pstrCodigoMensagem As String) As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsOperacao = CreateObject("ADODB.Recordset")
    
    Set ObterOperacaoOriginal = CreateObject("MSXML2.DOMDocument.4.0")

    strSQL = " Select   A.TP_OPER,                     " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          B.CO_TEXT_XML,                 " & vbCrLf & _
             "          B.DH_REGT_MESG_SPB             " & vbCrLf & _
             " From     A8.TB_OPER_ATIV             A, " & vbCrLf & _
             "          A8.TB_MESG_RECB_ENVI_SPB    B  " & vbCrLf & _
             " Where    A.NU_SEQU_OPER_ATIV     =   B.NU_SEQU_OPER_ATIV " & vbCrLf & _
             " And      B.NU_CTRL_IF            =   '" & pstrNumeroControleIF & "'" & vbCrLf & _
             " And      B.CO_MESG_SPB           =   '" & pstrCodigoMensagem & "'" & vbCrLf

    ObterOperacaoOriginal.loadXML fgQueryXMLLer("MESG", strSQL, "MESG")

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacaoOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens LDL0026R1 (Resposta à Consulta Resultado Líquido Negociação)
'-----------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0026R1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagResultLiqd                    As String
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode

Dim strTpInf                                As String
Dim strDtLiquid                             As String
Dim intTpDeb_Cred                           As String
Dim vntCNPJNLiqdant                         As Variant
Dim vntVlrResultLiqdNLiqdant                As Variant
Dim strNomeVeicLega                         As String


On Error GoTo ErrorHandler

    'Verifica se é uma camara já tratada
    lngCodigoLocalLiquidacao = VerificaCamaraAtiva(xmlMensagem, _
                                                     Trim(xmlMensagem.selectSingleNode("//CodMsg").Text))
    
    If lngCodigoLocalLiquidacao = 0 Then
        Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)
        Exit Function
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strSituacaoMensagemSPB = ""
    strDataRegistroMensagem = ""
    
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) '"Repet_LDL0026R1_ResultLiqd"
    strNumeroControleCamara = ""
    strTpInf = xmlMensagem.selectSingleNode("//TpInf").Text
    strDtLiquid = xmlMensagem.selectSingleNode("//DtLiquid").Text
    strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
        
    'Incluir Mensagem Pai
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
    lngStatusMensagem = enumStatusMensagem.R1
    
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNLiqdant, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistemaVeicLega, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        

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
                                0, _
                                lngSeqControleRepeticao, _
                                strDtLiquid, _
                                0, _
                                "", _
                                vntCNPJNLiqdant, _
                                strNomeVeicLega, _
                                0, _
                                strTpInf)
   
   
    'Repeticao Titulos
    For Each xmlNodeResultLiqd In xmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeResultLiqd)
        
        lngSeqControleRepeticao = lngSeqControleRepeticao + 1
        
        vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
        vntVlrResultLiqdNLiqdant = xmlNodeResultLiqd.selectSingleNode("VlrResultLiqdNLiqdant").Text
        intTpDeb_Cred = IIf(xmlNodeResultLiqd.selectSingleNode("TpDeb_CredNLiqdant").Text = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
            
        If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     lngCodigoEmpresa, _
                                     strCodigoVeiculoLegal, _
                                     strSiglaSistemaVeicLega, _
                                     lngTipoBackOffice, _
                                     strNomeVeicLega) Then
                                     
            lngStatusMensagem = enumStatusMensagem.AConciliar
            
            Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                    lngTipoBackOffice, _
                                    strNomeVeicLega, _
                                    vntVlrResultLiqdNLiqdant, _
                                    xmlMensagem)
            
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        End If
           
        'Salvar Mensagem Enviada Recebida
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
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
                                           strNumeroControleCamara, _
                                           strDataRegistroMensagem)

                
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrResultLiqdNLiqdant, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    intTpDeb_Cred, _
                                    "", _
                                    vntCNPJNLiqdant, _
                                    strNomeVeicLega, _
                                    0, _
                                    strTpInf)
                                    
    Next
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0026R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Obter o Status da Operação e mensagem
'------------------------------------------------------------------------------------------

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
     
    If Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text) = "LDL0003R1" Or _
       Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text) = "LTR0002R1" Then
        
        xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.Confirmada
    
    ElseIf Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text) = "LTR0008R1" Then
    
        xmlStatus.selectSingleNode("//Status_Operacao").Text = ""
        xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemLiquidada
        xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = ""
    
    ElseIf Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text) = "BMC0102R1" Then
    
        xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado
        xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.Confirmada
        xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = ""
    
    Else
        
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
    End If
    
    flObterStatusOperacaoMensagem = xmlStatus.xml
    
    Set xmlStatus = Nothing
   
    Exit Function
ErrorHandler:
    
    Set xmlStatus = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusOperacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens não geradas a partir de uma remessa do legado
'------------------------------------------------------------------------------------------

Private Function flProcessaMesgNaoTratadaR1(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                            ByVal pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim objLegado                               As A8LQS.clsLegado
Dim strMensagemLegado                       As String

Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strDataOperacao                         As String
Dim strNumeroComando                        As String
Dim strCodigoMensagemMae                    As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoVeiculoLegal                   As String
Dim strDataRegistroMensagem                 As String
Dim strCodigoMensagemSPB                    As String
Dim strSiglaSistema                         As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngStatusOperacao                       As Long
Dim lngSeqControleRepeticao                 As Long
Dim blnSalvaMensagem                        As Boolean
Dim blnAtualizaStatusOperacao               As Boolean
Dim blnAtualizaStatusOperGerenciarChamada   As Boolean
Dim blnGerarAlerta                          As Boolean

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler
                    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
                    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text)
    blnSalvaMensagem = True
    blnAtualizaStatusOperacao = True
    blnAtualizaStatusOperGerenciarChamada = True
    
    If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = pxmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
                    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    If CLng(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.MensagemLiquidada Then
        '3038 - Operação já liquidada.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
                    
    Select Case strCodigoMensagemSPB
        Case "LDL0026R1"
            
            blnSalvaMensagem = False
            blnAtualizaStatusOperacao = False
            
            'Atualiza Status da Mensagem Original
            lngStatusMensagemOriginal = enumStatusMensagem.Respondida
            objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                             lngStatusMensagemOriginal, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, "", False, 1
            
            
            Call flProcessaMensagemLDL0026R1(pxmlMensagem, pstrHeaderNZ)
            
        Case "LTR0009R1"
            
            blnSalvaMensagem = False
            blnAtualizaStatusOperacao = False
            
            'Atualiza Status da Mensagem Original
            lngStatusMensagemOriginal = enumStatusMensagem.Respondida
            objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                             lngStatusMensagemOriginal, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
            

            
            Call flProcessaMensagemLTR0009R1(pxmlMensagem, pstrHeaderNZ)
            
        
        Case "LDL0004R1", "LTR0004R1", "LTR0003R1", "LDL0008R1"
            
            blnSalvaMensagem = False
            blnAtualizaStatusOperacao = False
            
            Call flProcessaMensagemR1Pagamento(pxmlMensagem, pstrHeaderNZ)
            
        Case "LTR0002R1", "LTR0008R1", "LDL0003R1"
        
            blnSalvaMensagem = True
            blnAtualizaStatusOperacao = False
            
            Call flProcessaMensagemR1Confirmacao(pxmlMensagem, pstrHeaderNZ, blnGerarAlerta)
        
            If blnGerarAlerta Then
                Call flEnviarAlertaR1(pxmlMensagem)
            End If
        
        Case "CTP0100R1"
        
            blnSalvaMensagem = True
            blnAtualizaStatusOperacao = False
            
            Call flProcessaMensagemCTP0100R1(pxmlMensagem, pstrHeaderNZ)
        
        Case "BMC0102R1"
            
            blnSalvaMensagem = False
            blnAtualizaStatusOperacao = False
            
            Call flProcessaMensagemBMC0102R1(pxmlMensagem, pstrHeaderNZ)
        
        Case "CTP2002R1"

            blnSalvaMensagem = False
            blnAtualizaStatusOperacao = False

            Call flProcessaMensagemCTP2002R1(pxmlMensagem, pstrHeaderNZ)

    End Select
                    
    If blnSalvaMensagem Then
                    
        lngSeqControleRepeticao = 1
        lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
        strNumeroComando = xmlMesgRecbEnviSPB.selectSingleNode("//NU_COMD_OPER").Text
        vntSequenciaOperacao = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
        lngCodigoEmpresa = xmlMesgRecbEnviSPB.selectSingleNode("//CO_EMPR").Text
        strCodigoVeiculoLegal = xmlMesgRecbEnviSPB.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlMesgRecbEnviSPB.selectSingleNode("//SG_SIST").Text
        lngCodigoLocalLiquidacao = xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text
        lngTipoBackOffice = Val(xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text)
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(pxmlMensagem.xml, _
                                                                Trim(strNumeroControleIF), _
                                                                vntSequenciaOperacao, _
                                                                vntSequenciaConciliacaoOperacao, _
                                                                strCodigoMensagemSPB, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                strNumeroComando, _
                                                                "", _
                                                                enumStatusMensagem.R1, _
                                                                lngIndicadorEntradaManual, _
                                                                "", _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistema, _
                                                                lngSeqControleRepeticao)
        
        If xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = "" And _
           strCodigoMensagemSPB <> "LTR0008R1" And _
           strCodigoMensagemSPB <> "LTR0002R1" And _
           strCodigoMensagemSPB <> "LDL0003R1" Then
           
            blnAtualizaStatusOperacao = False
            
            lngStatusMensagemOriginal = enumStatusMensagem.Respondida
            objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                             lngStatusMensagemOriginal, _
                                             xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, , False
        
        End If
    
    End If
    
    If blnAtualizaStatusOperacao Then
        
        'Obter Status ---------------------
        Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
        xmlStatus.loadXML flObterStatusOperacaoMensagem(pxmlMensagem)
        '----------------------------------
            
        If xmlStatus.selectSingleNode("//Status_Operacao").Text <> vbNullString Then
            lngStatusOperacao = xmlStatus.selectSingleNode("//Status_Operacao").Text
        Else
            lngStatusOperacao = 0
            blnAtualizaStatusOperGerenciarChamada = False
        End If
                
        'Atualiza Status da Mensagem de Requisição  -------------------------------------------------------------------
        If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
            lngStatusMensagemOriginal = xmlStatus.selectSingleNode("//Status_Mensagem").Text
        Else
            lngStatusMensagemOriginal = enumStatusMensagem.Respondida
        End If
        
        objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                         xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                         lngStatusMensagemOriginal, _
                                         xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                         , _
                                         False
       
        strSituacaoRecebida = vbNullString
        If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
            strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
        End If
        
       '----------------------------------------------------------------------------------------------------------
            
        'Atualiza Status da Operacao  -------------------------------------------------------------------
        If xmlStatus.selectSingleNode("//Status_Operacao").Text <> "" Then
                
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
            
            Select Case strCodigoMensagemSPB
                
                Case "BMA0005R1"
                
                    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
        
                       Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
                            enumTipoOperacaoLQS.EspecCompromissadaIntermediacao

                            'Atualizar Numero Controle Especificação
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                                If Not pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc") Is Nothing Then
                                    Call objOperacao.GravarControleSPBOriginal(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                               pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text)
                                End If
                            End If
                            
                            'Se a especificação for de uma compromissada generica , atualizar para Registradq
                            Call flProcessaMensagemBMA0005R1(xmlOperacao.documentElement.selectSingleNode("//NU_COMD_OPER").Text, xmlStatus.selectSingleNode("//Status_Operacao").Text)
                    
                       Case enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA
                                                   
                            'Para compromissada Especifica Coberta não atualizar a operacao
                            blnAtualizaStatusOperGerenciarChamada = False
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada Then
                                
                                blnAtualizaStatusOperGerenciarChamada = True
                                xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = xmlStatus.selectSingleNode("//Status_Operacao").Text
                                
                                'Damanda A822
                                'Alteração no SLCC para permitir o tratamento correto dos status no SLCC ligados à BMA0004 e à BMA006 (Eventos)
                                'A mensagem BMA0002 deverá ficar com status A Conciliar
                                Call objMesgRecbEnviSPB.AlterarStatusBMA0002LDL0001(vntSequenciaOperacao, enumStatusMensagem.AConciliar)
                            End If
                    
                    End Select
                
                Case "BMA0004R1"
                    
                    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
        
                       Case enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
                            enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
                            enumTipoOperacaoLQS.EspecTermoIntermediacao, _
                            enumTipoOperacaoLQS.EspecTermoCobertura
                                        
                            'Atualizar Numero Controle Especificação
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                                If Not pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc") Is Nothing Then
                                    Call objOperacao.GravarControleSPBOriginal(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                               pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text)
                                End If
                            End If
                        
                        Case enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                             enumTipoOperacaoLQS.OperacaoTermoCobertaBMA
                             
                            'Atualizar Numero Controle Especificação
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                                'If Not pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc") Is Nothing Then
                                '    Call objOperacao.GravarControleSPBOriginal(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                               pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text)
                                'End If
                            End If
                                                    
                            'Para Definitiva e Termo  Coberta não atualizar a operacao
                            blnAtualizaStatusOperGerenciarChamada = False
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada Then
                                
                                xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = xmlStatus.selectSingleNode("//Status_Operacao").Text
                                
                                blnAtualizaStatusOperGerenciarChamada = True
                            
                                'Damanda A822
                                'Alteração no SLCC para permitir o tratamento correto dos status no SLCC ligados à BMA0004 e à BMA006 (Eventos)
                                'A mensagem BMA0002 deverá ficar com status A Conciliar
                                Call objMesgRecbEnviSPB.AlterarStatusBMA0002LDL0001(vntSequenciaOperacao, enumStatusMensagem.AConciliar)
                            End If
                                                        
                    End Select
                
            End Select
            
            Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
            Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
                fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
            End If
            
            Set objTipoOperacao = Nothing
            Set xmlTipoOperacao = Nothing
            
            If Not pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc") Is Nothing Then
                If xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                    Call fgAppendNode(xmlOperacao, "MESG", "NU_CTRL_MESG_SPB_ORIG", pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text)
                Else
                    xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text = pxmlMensagem.selectSingleNode("//NumCtrlBMAEspecc").Text
                End If
            End If
                        
            If lngStatusOperacao <> enumStatusOperacao.Rejeitada Then
                strSituacaoRecebida = vbNullString
            End If
            
            Call GerenciarChamada(xmlOperacao, _
                                  pxmlMensagem, _
                                  lngStatusOperacao, _
                                  0, _
                                  "", _
                                  IIf(lngStatusOperacao = enumStatusOperacao.Rejeitada, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao), _
                                  "", _
                                  , _
                                  blnAtualizaStatusOperGerenciarChamada, _
                                  strSituacaoRecebida)
            
            Set objOperacao = Nothing
            Set xmlOperacao = Nothing
        End If
            
        'Atualiza Status Mensagem Mae --------------------------------------------------------------------------------
        'LDL0001, LTR0001 , LTR0007
        
        If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
            
            If xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemLiquidada Then
                
                strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(strCodigoMensagemSPB, 1, 7))
                
                If strCodigoMensagemMae <> vbNullString Then
                    
                    strMensagem = flObterMensagemMae(vntSequenciaConciliacaoOperacao, strCodigoMensagemMae)
                    
                    xmlMesgRecbEnviSPB.loadXML strMensagem
                    
                    For Each xmlNode In xmlMesgRecbEnviSPB.selectNodes("//Repeat_MensagemRecebidaEnviada/*")
                    
                        Call objMesgRecbEnviSPB.AlterarStatus(xmlNode.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                              xmlStatus.selectSingleNode("//Status_Mensagem").Text, _
                                                              xmlNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                              xmlNode.selectSingleNode(" DH_RECB_ENVI_MESG_SPB").Text, _
                                                              False, _
                                                              1, _
                                                              strCodigoMensagemMae)
                    
                    Next
                End If
            End If
        End If
        '-------------------------------------------------------------------------------------------------------------
    
    End If
    
    Call flGerenciarEnvioMensagem(pxmlMensagem, strNumeroControleIF, enumFuncaoSistema.Liberar)
    
    If strCodigoMensagemSPB = "LDL1003R1" Then

        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        Set objLegado = CreateObject("A8LQS.clsLegado")
        Set objOperacao = CreateObject("A8LQS.clsOperacao")

        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        
        xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
                    
        'Se existir SISMSG da mensagem SEL1023R1 remover
        If Not xmlOperacao.selectSingleNode("//SISMSG[*/CodMsg='SEL1023R1']") Is Nothing Then
            Set xmlNode = xmlOperacao.selectSingleNode("//SISMSG[*/CodMsg='SEL1023R1']")
            xmlOperacao.documentElement.removeChild xmlNode
        End If
        
        strMensagemLegado = objLegado.EnviarRetorno(xmlOperacao, pxmlMensagem)
        
        objProcessoOperacao.EnviarMensagemMQ strMensagemLegado, enumIdentificadorFila.BUS

        Set objLegado = Nothing
        Set objProcessoOperacao = Nothing
        Set objOperacao = Nothing

    End If
    
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function

ErrorHandler:
    
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMesgNaoTratadaR1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Obter os dados da mensagem que esta sendo liquidada (Ex. LTR0004 liquida a mensagem LTR0001)
'------------------------------------------------------------------------------------------

Private Function flObterMensagemMae(ByVal pvntNumeroSequenciaConciliacao As Variant, _
                           Optional ByVal pstrCodigoMensagem As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   NU_CTRL_IF,                  " & vbCrLf & _
             "          DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          TP_BKOF,                     " & vbCrLf & _
             "          CO_EMPR,                     " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,       " & vbCrLf & _
             "          CO_MESG_SPB,                 " & vbCrLf & _
             "          NU_COMD_OPER,                " & vbCrLf & _
             "          CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          CO_TEXT_XML,                 " & vbCrLf & _
             "          HO_ENVI_MESG_SPB,            " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
             "          DH_ULTI_ATLZ,                " & vbCrLf & _
             "          IN_ENTR_MANU                 " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB     " & vbCrLf & _
             " WHERE    NU_SEQU_CNCL_OPER_ATIV_MESG  =  " & pvntNumeroSequenciaConciliacao & _
             "   AND    CO_MESG_SPB   =  '" & pstrCodigoMensagem & "'" & _
             "   AND    NU_SEQU_CNTR_REPE = 1"

    flObterMensagemMae = fgQueryXMLLerTodos("MensagemRecebidaEnviada", strSQL, App.EXEName & "." & TypeName(Me))

Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterMensagemMae Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'------------------------------------------------------------------------------------------
'Gerencimanto de envio automatico de mensagens .
'Ex. Ao chegar BMA0006R1, o SLCC envia a BMA0004
'------------------------------------------------------------------------------------------

Private Function flGerenciarEnvioMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                          ByVal pstrNumCtrlIF As String, _
                                          ByVal plngFuncaoSistema As enumFuncaoSistema) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objOperacao                             As A8LQS.clsOperacao

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlFluxoProcessamento                   As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim vntNumSeqOperAtiv                       As Variant
Dim strNumCtrlIF                            As String
Dim strFluxoProcessamento                   As String
Dim strCodMesg                              As String
Dim strCodMesgEnvio                         As String
Dim strMensagemEnvio                        As String
Dim blnEviaMensagem                         As Boolean
Dim lngStatusOperacao                       As Long
Dim lngAcao                                 As Long

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    strCodMesg = Trim(pxmlMensagem.selectSingleNode("//CodMsg").Text)
    blnEviaMensagem = True
    
    'Verifica se a mensagem gera uma nova mensagemn
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    strFluxoProcessamento = objFluxoProcessamento.LerTodos(strCodMesg, "", plngFuncaoSistema)
    Set objFluxoProcessamento = Nothing
    
    If strFluxoProcessamento = "" Then Exit Function
    
    Set xmlFluxoProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlFluxoProcessamento.loadXML strFluxoProcessamento
        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
        
    For Each xmlNode In xmlFluxoProcessamento.selectNodes("//Repeat_HistSituacaoIntegracao/*")
        
        strCodMesgEnvio = xmlNode.selectSingleNode("CO_MESG_SPB_ENVI").Text
        
        xmlMesgRecbEnviSPB.loadXML objMesgRecbEnviSPB.LerMensagemIda(Trim(pstrNumCtrlIF))
        
        vntNumSeqOperAtiv = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
       
        Set xmlOperacao = objOperacao.LerXMLOperacao(vntNumSeqOperAtiv)
                            
        strNumCtrlIF = fgObterNumeroControleIF()
            
        If blnEviaMensagem Then
        
            Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
            Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
                fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
            End If
            
            Set objTipoOperacao = Nothing
            Set xmlTipoOperacao = Nothing
        
            strMensagemEnvio = flMontarMensagemNZ(xmlOperacao, _
                                                  pxmlMensagem, _
                                                  strNumCtrlIF, _
                                                  strCodMesgEnvio, _
                                                  xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text)
            
            Call objMesgRecbEnviSPB.IncluirBMA(strMensagemEnvio, _
                                               Trim(strNumCtrlIF), _
                                               vntNumSeqOperAtiv, _
                                               0, _
                                               strCodMesgEnvio, _
                                               xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text, _
                                               xmlMesgRecbEnviSPB.selectSingleNode("//CO_EMPR").Text, _
                                               xmlMesgRecbEnviSPB.selectSingleNode("//NU_COMD_OPER").Text, _
                                               "", _
                                               enumStatusMensagem.Liberada, _
                                               0, _
                                               "", _
                                               xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text, _
                                               xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                               xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text, _
                                               1, _
                                               "", _
                                               "")
                        
            'Transformar o Tipo de Débito/Crédito para Mensageria
            If Not xmlOperacao.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                If IsNumeric(xmlOperacao.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED").Text) Then
                    xmlOperacao.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED").Text = IIf(CLng(xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, "D", "C")
                End If
            End If
                        
            Call fgAppendXML(xmlOperacao, "MESG", pxmlMensagem.xml)
                        
            'RATS 359
            If Not xmlOperacao.documentElement.selectSingleNode("//CO_MESG") Is Nothing Then
                xmlOperacao.documentElement.selectSingleNode("//CO_MESG").Text = Trim(xmlOperacao.documentElement.selectSingleNode("//CO_MESG").Text)
            End If
                        
            strMensagemEnvio = xmlOperacao.xml
            
            'Complementar as informações necessárias para enviar a operação para o PJ
            udtProtocolo.CodigoEmpresa = Format(CLng(xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
            udtProtocolo.SiglaSistemaDestino = fgCompletaString("NZ", " ", 3, False)
            udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
            udtProtocolo.TipoMensagem = fgCompletaString(Trim(strCodMesgEnvio), " ", 9, False)
            LSet udtProtocoloAux = udtProtocolo
            
            Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
            objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & strMensagemEnvio, enumIdentificadorFila.BUS
            Set objProcessoOperacao = Nothing
                            
            lngStatusOperacao = xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
                        
            Select Case strCodMesgEnvio
                Case "SEL1023"
                    lngAcao = enumTipoAcao.EnviadaSEL1023
                Case "LDL1003"
                    lngAcao = enumTipoAcao.EnviadaLDL1003
                Case "SEL1007"
                    lngAcao = enumTipoAcao.EnviadaSEL1007
            End Select
        End If
        
        objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            lngStatusOperacao, _
                                            , _
                                            lngAcao
        
    Next
        
    Set objOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarEnvioMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Montagem do Header e mensagem a ser enviada ao NZ / Mensageria
'------------------------------------------------------------------------------------------

Private Function flMontarMensagemNZ(ByRef pxmlOperacao As MSXML2.DOMDocument40, _
                                    ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrNumeroControleIF As String, _
                                    ByVal pstrTipoMensagem As String, _
                                    ByVal plngTipoBackOffice As Long) As String

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim strVeiculoLegal                         As String
Dim strHeaderNZ                             As String
Dim strNumeroControleReqIF                  As String
Dim strNumeroControleEspc                   As String
Dim vntValorFinanceiro                      As Variant
Dim vntVlrXML                               As Variant

On Error GoTo ErrorHandler

    strHeaderNZ = fgMontaHeaderMensageNZ(pstrTipoMensagem, _
                                         pxmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                         pstrNumeroControleIF)


    'Incluir Header NZ na mensagem para enviar BUS
    If pxmlOperacao.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "CO_MESG", fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
    Else
        pxmlOperacao.documentElement.selectSingleNode("CO_MESG").Text = fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
    End If
    
    If pxmlOperacao.documentElement.selectSingleNode("TX_HEAD_NZ") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "TX_HEAD_NZ", strHeaderNZ
    Else
        pxmlOperacao.documentElement.selectSingleNode("TX_HEAD_NZ").Text = strHeaderNZ
    End If
    
    If pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "NU_CTRL_IF", pstrNumeroControleIF
    Else
        pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_IF").Text = pstrNumeroControleIF
    End If
    
    If pxmlOperacao.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "TP_BKOF", plngTipoBackOffice
    Else
        pxmlOperacao.documentElement.selectSingleNode("TP_BKOF").Text = plngTipoBackOffice
    End If
    
    If pxmlOperacao.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "NIVE_PREF", ""
    Else
        pxmlOperacao.documentElement.selectSingleNode("NIVE_PREF").Text = ""
    End If
    
    If pxmlOperacao.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "DT_SIST", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    Else
        pxmlOperacao.documentElement.selectSingleNode("DT_SIST").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    End If

    If pxmlOperacao.documentElement.selectSingleNode("DH_SIST") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "DH_SIST", fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    Else
        pxmlOperacao.documentElement.selectSingleNode("DH_SIST").Text = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    End If
    
    strNumeroControleReqIF = fgObterNumeroControleIF
       
    If pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_REQU_IF") Is Nothing Then
        fgAppendNode pxmlOperacao, "MESG", "NU_CTRL_REQU_IF", strNumeroControleReqIF
    Else
        pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_REQU_IF").Text = strNumeroControleReqIF
    End If
    
    Select Case CLng(pxmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
         
         Case enumTipoOperacaoLQS.DepositoBMA
             
             pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text = "01"
            
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
         
             pxmlOperacao.selectSingleNode("//QT_ATIV_MERC").Text = pxmlMensagem.selectSingleNode("//QtdActo").Text
             pxmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = pxmlMensagem.selectSingleNode("//VlrActo").Text
         
         Case enumTipoOperacaoLQS.RetiradaBMA
              
             pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text = "01"
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
                              
             pxmlOperacao.selectSingleNode("//NU_COMD_OPER").Text = pxmlMensagem.selectSingleNode("//NumOpSEL").Text
             pxmlOperacao.selectSingleNode("//QT_ATIV_MERC").Text = pxmlMensagem.selectSingleNode("//QtdTit").Text
             
             vntValorFinanceiro = fgVlrXml_To_Decimal(pxmlOperacao.selectSingleNode("//PU_ATIV_MERC").Text * pxmlOperacao.selectSingleNode("//QT_ATIV_MERC").Text)
             
             If InStr(1, vntValorFinanceiro, ",") Then
                vntVlrXML = Split(vntValorFinanceiro, ",", , vbBinaryCompare)
                vntValorFinanceiro = vntVlrXML(0) & "," & Mid(vntVlrXML(1), 1, 2)
             End If
             
             pxmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = vntValorFinanceiro
        
        Case enumTipoOperacaoLQS.TransferenciasBMCDeposito
                  
             pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text = Format(pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text, "00")
                  
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
                
        Case enumTipoOperacaoLQS.TransferenciasBMCRetirada
         
             pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text = Format(pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text, "00")
                  
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
        
        Case enumTipoOperacaoLQS.TransferenciaBMA, _
             enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA
                
             pxmlOperacao.documentElement.selectSingleNode("TP_TRAF_LDL").Text = "01"
            
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
             
             If pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text
             Else
                 pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CESS").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text
             End If
            
        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            
            fgAppendNode pxmlOperacao, "MESG", "TP_TRAF_LDL", "01"
            
            If pxmlOperacao.documentElement.selectSingleNode("TP_ATIV") Is Nothing Then
                fgAppendNode pxmlOperacao, "MESG", "TP_ATIV", "4"
            Else
                pxmlOperacao.documentElement.selectSingleNode("TP_ATIV").Text = ""
            End If

            If pxmlOperacao.documentElement.selectSingleNode("TP_REQU") Is Nothing Then
                fgAppendNode pxmlOperacao, "MESG", "TP_REQU", "2"
            Else
                pxmlOperacao.documentElement.selectSingleNode("TP_REQU").Text = "2"
            End If
            
            If Not pxmlMensagem.selectSingleNode("//QtdActo") Is Nothing Then
                pxmlOperacao.selectSingleNode("//QT_ATIV_MERC").Text = pxmlMensagem.selectSingleNode("//QtdActo").Text
            End If

            If pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                
                If Not pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                    fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                End If
                
                If Not pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                    fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                End If
                
            ElseIf pxmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                
                If Not pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                    fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CEDE", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                End If
                
                If Not pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                    fgAppendNode pxmlOperacao, "MESG", "CO_CNTA_CESS", pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                End If
            
            End If
            
            If Not pxmlOperacao.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
                fgAppendNode pxmlOperacao, "MESG", "NIVE_PREF", ""
            End If
            
            
        Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
             enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
             enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
             enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
             enumTipoOperacaoLQS.EspecTermoCobertura, _
             enumTipoOperacaoLQS.EspecTermoIntermediacao
                
             strNumeroControleEspc = ObterNumeroControleEspecificacaoOriginal(pxmlOperacao.documentElement.selectSingleNode("//NU_COMD_OPER").Text, _
                                                                              True, _
                                                                              pxmlOperacao.documentElement.selectSingleNode("//CO_EMPR").Text, _
                                                                              pxmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
        
             If pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                 fgAppendNode pxmlOperacao, "MESG", "NU_CTRL_MESG_SPB_ORIG", strNumeroControleEspc
             Else
                 pxmlOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text = strNumeroControleEspc
             End If
        
        Case enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
             enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, _
             enumTipoOperacaoLQS.OperacaoTermoCobertaBMA
             
            If Not pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_DEPO") Is Nothing Then
                pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text = pxmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_DEPO").Text
            End If
        
    End Select
    
    If pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL") Is Nothing Then
       fgAppendNode pxmlOperacao, "MESG", "CO_ISPB_LDL", ""
    End If
    
    'ISPBLDL
    Select Case pxmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
        Case enumLocalLiquidacao.BMA
            pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMA, "00000000")
        Case enumLocalLiquidacao.CETIP
            pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbCETIP, "00000000")
        Case enumLocalLiquidacao.BMC
            pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMC, "00000000")
        Case enumLocalLiquidacao.BMD
            pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMD, "00000000")
        Case enumLocalLiquidacao.SELIC
            pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbSELIC, "00000000")
        Case enumLocalLiquidacao.CLBCAcoes, _
             enumLocalLiquidacao.CLBCTpPriv, _
             enumLocalLiquidacao.CLBCTPub
           pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbCBLC, "00000000")
        Case enumLocalLiquidacao.SSTR
           pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBACEN, "00000000")
        Case enumLocalLiquidacao.CIP
           pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(4391007, "00000000")
    End Select
    
    If pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_IF") Is Nothing Then
       fgAppendNode pxmlOperacao, "MESG", "CO_ISPB_IF", ""
    End If
        
    'ISPBIF
    pxmlOperacao.documentElement.selectSingleNode("CO_ISPB_IF").Text = fgObterISPBIF(CLng(pxmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text))

    If Not pxmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then
        
        Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
        Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        xmlVeiculoLegal.loadXML objVeiculoLegal.Ler(pxmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, pxmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

        If xmlVeiculoLegal.documentElement.selectSingleNode("//TP_TITL_BMA").Text <> "PNA" And _
            xmlVeiculoLegal.documentElement.selectSingleNode("//TP_TITL_BMA").Text <> "PLC" Then
            strVeiculoLegal = fgObterCNPJEmpresa(pxmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text)
        Else
            strVeiculoLegal = xmlVeiculoLegal.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text
        End If
        

        If pxmlOperacao.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA") Is Nothing Then
            fgAppendNode pxmlOperacao, "MESG", "CO_CNPJ_VEIC_LEGA", strVeiculoLegal
        Else
            pxmlOperacao.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text = strVeiculoLegal
        End If
        Set objVeiculoLegal = Nothing
        Set xmlVeiculoLegal = Nothing
    End If

    flMontarMensagemNZ = pxmlOperacao.xml

    
    Exit Function
ErrorHandler:
    

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMensagemNZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'------------------------------------------------------------------------------------------
'Verificação e processamento do fluxo automático de registro e liquidação de operações e mensagens.
'   - LTR0001
'   - LTR0007
'   - LTR0005R2
'   - BMA0002
'   - CTP1002
'   - BMC0005
'------------------------------------------------------------------------------------------

Private Sub flConciliarMensagem(ByVal pstrtCodigoMensagem As String, _
                                ByVal pstrNumeroControleIF As String, _
                                ByVal plngNumeroSequenciaRepeticao As Long, _
                                ByVal pstrDHRegistroMensagem As String, _
                       Optional ByVal pstrNomeVeicLega As String = vbNullString)

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

    On Error GoTo ErrorHandler

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")

    xmlMesgRecbEnviSPB.loadXML objMesgRecbEnviSPB.LerXMLMensagem(Trim$(pstrNumeroControleIF), _
                                                                 plngNumeroSequenciaRepeticao, _
                                                                 , _
                                                                 pstrtCodigoMensagem, _
                                                                 , _
                                                                 pstrDHRegistroMensagem).xml

    Call fgAppendNode(xmlConciliacao, "", "Repeat_Conciliacao", "")
    Call fgAppendNode(xmlConciliacao, "Repeat_Conciliacao", "Grupo_Mensagem", "")

    Call fgAppendNode(xmlConciliacao, "Grupo_Mensagem", "NU_CTRL_IF", pstrNumeroControleIF)
    Call fgAppendNode(xmlConciliacao, "Grupo_Mensagem", "DH_REGT_MESG_SPB", xmlMesgRecbEnviSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text)
    Call fgAppendNode(xmlConciliacao, "Grupo_Mensagem", "NU_SEQU_CNTR_REPE", xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    Call fgAppendNode(xmlConciliacao, "Grupo_Mensagem", "DH_ULTI_ATLZ", xmlMesgRecbEnviSPB.selectSingleNode("//DH_ULTI_ATLZ").Text)
    Call fgAppendNode(xmlConciliacao, "Grupo_Mensagem", "TP_LIQU_OPER_ATIV", xmlMesgRecbEnviSPB.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    
    Select Case pstrtCodigoMensagem
        
        Case "LTR0001", "LTR0007", "LTR0003R2", "LTR0005R2"
            Call objConciliacao.GerenciarConciliacaoAutomaticaCETIP(, xmlConciliacao)
        Case "BMA0002", "CTP1002"
            Call objConciliacao.ConciliarRegistroOperacao(enumAcaoConciliacao.BOConcordar, _
                                                          , _
                                                          xmlConciliacao, _
                                                          True)
        Case "BMC0005"
            If objConciliacao.ConciliarRegistroOperacao(enumAcaoConciliacao.BOConcordar, _
                                                        , _
                                                        xmlConciliacao, _
                                                        True) <> "" Then
                
                Call flGeraAlertaCamara("BMC0005SOPER", _
                                        enumTipoBackOffice.Tesouraria, _
                                        pstrNomeVeicLega, _
                                        0, _
                                        xmlMesgRecbEnviSPB)
            End If
            
    End Select
    
    Set xmlConciliacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlConciliacao = Nothing
    Set objConciliacao = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlConciliacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlConciliacao = Nothing
    Set objConciliacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'------------------------------------------------------------------------------------------
'Processamento da Mensahem BMA0004R1 (Resposta ao Requisitante de Especificação de Conta de Custódia)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0004R1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                             ByVal plngStatusOperacao As Long) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlNodeEspecConta                       As MSXML2.IXMLDOMNode
Dim strNomeTagEspecConta                    As String

On Error GoTo ErrorHandler
        
    strNomeTagEspecConta = ObterNomeTagRepeticao(Trim(xmlMensagem.selectSingleNode("//CodMsg").Text), 1)
            
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    For Each xmlNodeEspecConta In xmlMensagem.documentElement("//" & strNomeTagEspecConta & "/*")
                
        If Not xmlOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
            Call fgAppendNode(xmlOperacao, "MESG", "NU_CTRL_MESG_SPB_ORIG", xmlNodeEspecConta.selectSingleNode("NumCtrlBMAEspecc").Text)
        Else
            xmlOperacao.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text = xmlNodeEspecConta.selectSingleNode("NumCtrlBMAEspecc").Text
        End If
            
        If Not xmlOperacao.documentElement.selectSingleNode("CO_TITL_CUTD") Is Nothing Then
            Call fgAppendNode(xmlOperacao, "MESG", "CO_TITL_CUTD", xmlNodeEspecConta.selectSingleNode("CodTitlarCstdte").Text)
        Else
            xmlOperacao.documentElement.selectSingleNode("CO_TITL_CUTD").Text = xmlNodeEspecConta.selectSingleNode("CodTitlarCstdte").Text
        End If
                
        If Not xmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
            Call fgAppendNode(xmlOperacao, "MESG", "CO_CNTA_CUTD_SELIC_VEIC_LEGA", xmlNodeEspecConta.selectSingleNode("CtCstdia").Text)
        Else
            xmlOperacao.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text = xmlNodeEspecConta.selectSingleNode("CtCstdia").Text
        End If
                
        If Not xmlNodeEspecConta.selectSingleNode("VlrFinanc") Is Nothing Then
            If Not xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV") Is Nothing Then
                Call fgAppendNode(xmlOperacao, "MESG", "VA_OPER_ATIV", xmlNodeEspecConta.selectSingleNode("VlrFinanc").Text)
            Else
                xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = xmlNodeEspecConta.selectSingleNode("VlrFinanc").Text
            End If
        End If
        
        If Not xmlOperacao.documentElement.selectSingleNode("QT_ATIV_MERC") Is Nothing Then
            Call fgAppendNode(xmlOperacao, "MESG", "QT_ATIV_MERC", xmlNodeEspecConta.selectSingleNode("QtdTit").Text)
        Else
            xmlOperacao.documentElement.selectSingleNode("QT_ATIV_MERC").Text = xmlNodeEspecConta.selectSingleNode("QtdTit").Text
        End If
                
        If Not xmlNodeEspecConta.selectSingleNode("PUFinldd") Is Nothing Then
            If Not xmlOperacao.documentElement.selectSingleNode("PU_ATIV_MERC") Is Nothing Then
                Call fgAppendNode(xmlOperacao, "MESG", "PU_ATIV_MERC", xmlNodeEspecConta.selectSingleNode("PUFinldd").Text)
            Else
                xmlOperacao.documentElement.selectSingleNode("PU_ATIV_MERC").Text = xmlNodeEspecConta.selectSingleNode("PUFinldd").Text
            End If
        End If
                
        Call objOperacao.IncluirOperacao(xmlOperacao, plngStatusOperacao)
    
    Next
    
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set objOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0004R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Tratamento das mensagens de pagamento LDL0004R1, LTR0004R1, LDL0003R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemR1Pagamento(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                               ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strDataOperacao                         As String
Dim strNumeroComando                        As String
Dim strCodigoMensagemMae                    As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strNumeroControleCamara                 As String
Dim strCodigoVeiculoLegal                   As String
Dim strDataRegistroMensagem                 As String
Dim strCodigoMensagemSPB                    As String
Dim strSiglaSistema                         As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngStatusOperacao                       As Long
Dim lngSeqControleRepeticao                 As Long
Dim blnSalvaMensagem                        As Boolean
Dim blnAtualizaStatusOperacao               As Boolean
Dim blnEstorno                              As Boolean

'---
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim strRet                                  As String
Dim blnAcaoPagtoContingencia                As Boolean
Dim blnGerenciarChamada                     As Boolean

Dim vntSequenciaOperacaoAnterior            As Variant

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler
                    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
                    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    blnSalvaMensagem = True
    blnAtualizaStatusOperacao = True
    blnEstorno = False
    
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
                    
    strNumeroControleCamara = vbNullString
    If Not xmlMensagem.selectSingleNode("//NumCtrlSTR") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlSTR").Text
    End If
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
                    
    lngSeqControleRepeticao = 1
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strNumeroComando = xmlMesgRecbEnviSPB.selectSingleNode("//NU_COMD_OPER").Text
    vntSequenciaOperacao = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    lngCodigoEmpresa = xmlMesgRecbEnviSPB.selectSingleNode("//CO_EMPR").Text
    strCodigoVeiculoLegal = xmlMesgRecbEnviSPB.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlMesgRecbEnviSPB.selectSingleNode("//SG_SIST").Text
    lngCodigoLocalLiquidacao = xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text
    lngTipoBackOffice = Val(xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text)
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            vntSequenciaConciliacaoOperacao, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumeroComando, _
                                                            "", _
                                                            enumStatusMensagem.R1, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao, _
                                                            strNumeroControleCamara)
        
     'Obter Status ---------------------
     Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
     xmlStatus.loadXML flObterStatusOperacaoMensagem(xmlMensagem)
     '----------------------------------
         
     'Atualiza Status da Mensagem Original -------------------------------------------------------------------
     If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
         lngStatusMensagemOriginal = xmlStatus.selectSingleNode("//Status_Mensagem").Text
     Else
         lngStatusMensagemOriginal = enumStatusMensagem.Respondida
     End If
     
     objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                      lngStatusMensagemOriginal, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
    
    '----------------------------------------------------------------------------------------------------------
    
    If vntSequenciaConciliacaoOperacao = 0 Then
        Set xmlStatus = Nothing
        Set xmlMesgRecbEnviSPB = Nothing
        Set objMesgRecbEnviSPB = Nothing
        Exit Function
    End If
    
    strSituacaoRecebida = vbNullString
    If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
        strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
    End If
    
    'Atualiza Status Mensagem LDL0001 --------------------------------------------------------------------------------
    If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
        
        If xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemLiquidada Or _
           xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemRejeitada Or _
           xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemPendente Then
            
            strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(strCodigoMensagemSPB, 1, 7)) ' LDL0001
            
            If strCodigoMensagemMae <> vbNullString Then
                
                strMensagem = flObterMensagemMae(vntSequenciaConciliacaoOperacao, strCodigoMensagemMae)
                
                xmlMesgRecbEnviSPB.loadXML strMensagem
                
                '-------------------------------------------------------------------------
                Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
                Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
                                
                strNumeroControleIF = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text
                
                strRet = objHistMensagem.LerTodos(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                  xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                  xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                 
                xmlHistorico.loadXML strRet
                                
                blnAcaoPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
                
                Set objHistMensagem = Nothing
                Set xmlHistorico = Nothing
                '----------------------------------------------------------------------------
                
                If blnAcaoPagtoContingencia Then
                    xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.AConciliar
                                        
                    If strCodigoMensagemMae <> "LDL0001" Then
                        Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              xmlStatus.selectSingleNode("//Status_Mensagem").Text, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_RECB_ENVI_MESG_SPB").Text, _
                                                              False, _
                                                              1, _
                                                              strCodigoMensagemMae)
                    End If
                
                Else
                    
                    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
                       
                        'KIDA - CBLC - 05/10/2008
                        If xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemRejeitada Then
                            
                            If strCodigoMensagemSPB = "LDL0004R1" Or strCodigoMensagemSPB = "LDL0008R1" Then
                                lngStatusMensagemOriginal = enumStatusMensagem.AConciliar
                            
                            ElseIf strCodigoMensagemSPB = "LTR0004R1" Then
                                
                                lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                                
                            End If
                            
                        ElseIf xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemPendente Then
                            
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemPendente
                        
                        ElseIf xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemLiquidada Then
                            lngStatusMensagemOriginal = enumStatusMensagem.MensagemLiquidada
                        Else
                            lngStatusMensagemOriginal = 0
                        End If
                        
                    ElseIf lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC Then
                       
                        If xmlStatus.selectSingleNode("//Status_Mensagem").Text = enumStatusMensagem.MensagemRejeitada Then
                            If strCodigoMensagemSPB = "LDL0004R1" Then
                                lngStatusMensagemOriginal = enumStatusMensagem.ConcordanciaBackoffice
                            End If
                        End If
                            
                    End If
                    
                    If lngStatusMensagemOriginal <> 0 Then
                        Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              lngStatusMensagemOriginal, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                              xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_RECB_ENVI_MESG_SPB").Text, _
                                                              False, _
                                                              0, _
                                                              strCodigoMensagemMae)
                    End If
                    
                End If
                
            End If
        End If
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao, 0, strNumeroControleIF)) Then
    
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
        vntSequenciaOperacaoAnterior = vbNullString
        
        For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
            
            vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
            
            'Atualizar as operaçoes---------------------------
            If xmlStatus.selectSingleNode("//Status_Operacao").Text <> "" And vntSequenciaOperacao <> 0 Then
                
                If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Or _
                   xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada Or _
                   xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.RejeitadaLiquidacao Then
                    
                    Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                        
                    blnGerenciarChamada = True
                    
                    If blnAcaoPagtoContingencia Then
                        lngStatusOperacao = enumStatusOperacao.Registrada
                        
                        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC Or strCodigoMensagemSPB = "LTR0004R1" Then
                            blnGerenciarChamada = False
                        End If
                        
                        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Or strCodigoMensagemSPB = "LDL0004R1" Then
                            blnGerenciarChamada = False
                        End If
                        
                        'KIDA - CBLC - 17/09/2008
                        'KIDA - BMF  - 11/08/2009
                        If lngCodigoLocalLiquidacao <> enumLocalLiquidacao.CLBCAcoes And _
                           lngCodigoLocalLiquidacao <> enumLocalLiquidacao.BMD Then

                        
                            If Not xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
                                If Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> lngStatusOperacao Then
                                    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = lngStatusOperacao
                                    
                                    If lngStatusOperacao <> enumStatusOperacao.Rejeitada Then
                                        strSituacaoRecebida = vbNullString
                                    End If
                                    
                                    Call objOperacao.AlterarStatusAcao(vntSequenciaOperacao, _
                                                                       lngStatusOperacao, _
                                                                       vbNullString, , , _
                                                                       False, , _
                                                                       strSituacaoRecebida)
                                
                                End If
                            End If
                        End If
                    Else
                        
                        If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
                           
                           'KIDA - CBLC - 24/09/2008
                           If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                           
                                lngStatusOperacao = enumStatusOperacao.Liquidada
                                blnAtualizaStatusOperacao = True
                                blnEstorno = False
                                
                            End If
                           
                           'KIDA - CBLC - 03/10/2008
                           If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada Or _
                              xmlStatus.selectSingleNode("//Status_Operacao").Text = RejeitadaLiquidacao Then
                           
                                lngStatusOperacao = enumStatusOperacao.Rejeitada
                                blnAtualizaStatusOperacao = False
                                blnEstorno = True
                                
                                If strCodigoMensagemSPB = "LDL0004R1" Or strCodigoMensagemSPB = "LDL0008R1" Then
                                    Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Rejeitada
                                End If
                                
                                If strCodigoMensagemSPB = "LTR0004R1" Then
                                    Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, enumStatusOperacao.ConcordanciaBackoffice)
                                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Rejeitada
                                End If
                                
                                
                            End If
                        
                        ElseIf lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC Then
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Liquidada Then
                                lngStatusOperacao = enumStatusOperacao.DebitoMoedaNacionalLiquidado
                            End If
                            
                            If xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada Then
                                lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
                                blnEstorno = True
                            End If
                        
                        Else
                            If Val(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> 0 Then
                                lngStatusOperacao = Val(xmlStatus.selectSingleNode("//Status_Operacao").Text)
                            End If
                        
                            If lngStatusOperacao = enumStatusOperacao.Rejeitada Or _
                               lngStatusOperacao = enumStatusOperacao.RejeitadaLiquidacao Then
                                blnEstorno = True
                            Else
                                blnEstorno = False
                            End If
                        
                        End If
                        
                    End If
                    
                    If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
                        If Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Multilateral Then
                           If blnAcaoPagtoContingencia Then
                                blnGerenciarChamada = False
                           End If
                        End If
                    End If
                                            
                    If vntSequenciaOperacaoAnterior = vntSequenciaOperacao Or _
                       Val(vntSequenciaOperacao) = 0 Then
                        blnGerenciarChamada = False
                    End If
                    
                    'KIDA - CBLC - EVENTO
                    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
                       If xmlStatus.selectSingleNode("//Status_Operacao").Text <> enumStatusOperacao.Liquidada Then
                            blnAtualizaStatusOperacao = False
                            blnEstorno = False
                            blnGerenciarChamada = False
                        End If
                    End If
                    
                    If blnGerenciarChamada Then
                        vntSequenciaOperacaoAnterior = vntSequenciaOperacao
                        
                        If lngStatusOperacao <> enumStatusOperacao.Rejeitada Then
                            strSituacaoRecebida = vbNullString
                        End If
                
                        GerenciarChamada xmlOperacao, _
                                         xmlMensagem, _
                                         lngStatusOperacao, _
                                         0, _
                                         vbNullString, _
                                         IIf(blnEstorno, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao), _
                                         vbNullString, _
                                         blnAcaoPagtoContingencia, _
                                         blnAtualizaStatusOperacao, _
                                         strSituacaoRecebida
                    
                    End If
                    
                ElseIf xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Pendencia Or _
                       xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.PendenteLiquidacao Then
                    
                    Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                    
                    blnGerenciarChamada = True
                    If vntSequenciaOperacaoAnterior = vntSequenciaOperacao Or _
                       Val(vntSequenciaOperacao) = 0 Then
                        blnGerenciarChamada = False
                    End If
                    
                    If blnGerenciarChamada Then
                        vntSequenciaOperacaoAnterior = vntSequenciaOperacao
                        
                        GerenciarChamada xmlOperacao, _
                                         xmlMensagem, _
                                         xmlStatus.selectSingleNode("//Status_Operacao").Text, _
                                         0, _
                                         "", _
                                         IIf(blnEstorno, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao), _
                                         "", _
                                         blnAcaoPagtoContingencia, blnAtualizaStatusOperacao
                    
                    End If
                    
                End If
            
            End If
            
        Next
        
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
    
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set objHistMensagem = Nothing
    Set xmlHistorico = Nothing
    Set xmlOperacao = Nothing
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemR1Pagamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




'
'------------------------------------------------------------------------------------------
'Tratamento para o caso do SLCC receber uma segunda mensagem LDL0001 definitiva
'------------------------------------------------------------------------------------------
Private Function flProcessaMensagemSegundaLDL0001(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                                  ByRef pstrHeaderNZ As String) As Long

Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

Dim xmlLDL0001Conflitante                   As MSXML2.DOMDocument40
Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPBTodas                 As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim lngCodigoLocalLiquidacao                As Long
Dim lngCodigoEmpresa                        As Long
Dim strLDL0001Conflitante                   As String

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

On Error GoTo ErrorHandler
                        
'    Exit Function
                        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlLDL0001Conflitante = CreateObject("MSXML2.DOMDocument.4.0")
    
    Select Case Val(xmlMensagem.selectSingleNode("//ISPBLDL").Text)
        Case enumISPB.IspbBMA
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
        Case enumISPB.IspbCETIP
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
        Case enumISPB.IspbBMD
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD
    End Select

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    
    'Verificar se ja existe uma LDL0001 Para Local Liquidacao e Empresa e estiver paga
    strLDL0001Conflitante = objMesgRecbEnviSPB.ObterMensagemLDL0001Conflitante(lngCodigoEmpresa, lngCodigoLocalLiquidacao)
                                
    'BMF
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD Then
        If strLDL0001Conflitante <> vbNullString Then
            xmlLDL0001Conflitante.loadXML strLDL0001Conflitante
            
            If xmlMensagem.selectSingleNode("//TpDeb_Cred").Text = "D" Then
                If xmlLDL0001Conflitante.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                    strLDL0001Conflitante = vbNullString
                End If
            Else
                If xmlLDL0001Conflitante.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                    strLDL0001Conflitante = vbNullString
                End If
            End If
            
        End If
    End If
                            
    'Se naõ existir , fluxo normal
    If strLDL0001Conflitante = vbNullString Then
        flProcessaMensagemSegundaLDL0001 = enumStatusMensagem.AConciliar
        Set objMesgRecbEnviSPB = Nothing
        Set xmlLDL0001Conflitante = Nothing
        Exit Function
    End If
                        
    xmlLDL0001Conflitante.loadXML strLDL0001Conflitante
        
    'Se existir e não estiver paga (LDL0004)
    '1) Atualizar a LDL0001 para CANCELADA
    '2) Atualizar as Operações para REGISTRADA
    '3) Limpar ações na LDL0001 anterior
            
    If Not xmlLDL0001Conflitante.selectSingleNode("//JaEnviouPagamento") Is Nothing Then
        flProcessaMensagemSegundaLDL0001 = enumStatusMensagem.Aviso
        Set objMesgRecbEnviSPB = Nothing
        Set xmlLDL0001Conflitante = Nothing
        Exit Function
    End If
    
    flProcessaMensagemSegundaLDL0001 = enumStatusMensagem.MensagemCancelada
    
    strNumeroControleIF = xmlLDL0001Conflitante.selectSingleNode("//NU_CTRL_IF").Text
    strDataRegistroMensagem = xmlLDL0001Conflitante.selectSingleNode("//DH_REGT_MESG_SPB").Text
    
    Set xmlMesgRecbEnviSPBTodas = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlMesgRecbEnviSPBTodas.loadXML(objMesgRecbEnviSPB.LerTodos(strNumeroControleIF)) Then
                                                          
        For Each xmlNode In xmlMesgRecbEnviSPBTodas.selectNodes("//Repeat_MensagemEnviadaRecebida/*")
                                                      
            'Atualiza a LDL0001 anterior com as filhas
            Call objMesgRecbEnviSPB.AlterarStatus(xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                  xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                  enumStatusMensagem.MensagemCancelada, _
                                                  xmlNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                  xmlNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text, _
                                                  False, _
                                                  xmlNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        Next
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Val(vntSequenciaConciliacaoOperacao) <> 0 Then
    
        If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao)) Then
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
            For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
                
                vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
                xmlOperacao.loadXML objOperacao.Ler(vntSequenciaOperacao)
                objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                              enumStatusOperacao.Registrada, _
                                              xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                              , _
                                              , _
                                              False
            Next
            
            Set objOperacao = Nothing
            Set xmlOperacao = Nothing
        
        End If
    End If
    
    flProcessaMensagemSegundaLDL0001 = enumStatusMensagem.AConciliar
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemSegundaLDL0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMA0006R1 (Resposta ao Requisitante de Cancelamento de Especificação de Operação BMA)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0006R1(ByVal pstrNumeroControleCamaraOriginal As String) As Boolean

Dim objRS                                   As ADODB.Recordset
Dim objRSCompromissada                      As ADODB.Recordset

Dim strSQL                                  As String

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objMesgRecebEnviSPB                     As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMesgRecebEnviSPB                     As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim vntNumeroSequOperAtiv                   As Variant
Dim vntNumeroSequOperAtivCompro             As Variant

Dim strNumeroComando                        As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_SEQU_OPER_ATIV,TP_OPER,NU_COMD_OPER " & _
             "   FROM   A8.TB_OPER_ATIV " & _
             "  WHERE   NU_CTRL_MESG_SPB_ORIG = '" & pstrNumeroControleCamaraOriginal & "'" & _
             " order by TP_OPER "
             
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        objRS.Close
        Exit Function
    End If
        
    While Not objRS.EOF
        
        vntNumeroSequOperAtiv = objRS.fields("NU_SEQU_OPER_ATIV")
        strNumeroComando = objRS.fields("NU_COMD_OPER")
        
        If objRS!TP_OPER = enumTipoOperacaoLQS.CancelamentoEspecificacaoCompromissadaBMA Then
                
            strSQL = " SELECT NU_SEQU_OPER_ATIV " & _
                     "   FROM A8.TB_OPER_ATIV  " & _
                     "  WHERE NU_COMD_OPER = '" & strNumeroComando & "'" & _
                     "    AND TP_OPER = " & enumTipoOperacaoLQS.CompromissadaGenericaAVista & _
                     "    AND CO_ULTI_SITU_PROC = " & enumStatusOperacao.Registrada
                                 
            Set objRSCompromissada = fgQuerySQL(strSQL)
            
            If Not objRSCompromissada.EOF Then
                
                vntNumeroSequOperAtivCompro = objRSCompromissada!NU_SEQU_OPER_ATIV
                
                Set objOperacao = CreateObject("A8LQS.clsOperacao")
                Call objOperacao.AlterarStatusOperIndisp(vntNumeroSequOperAtivCompro, enumStatusOperacao.AConciliarBMA0013)
                Set objOperacao = Nothing
            
            End If
        
        Else
            
            Set xmlMesgRecebEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
            
            strSQL = " SELECT   NU_CTRL_IF,                  " & vbCrLf & _
                     "          DH_REGT_MESG_SPB,            " & vbCrLf & _
                     "          NU_SEQU_OPER_ATIV,           " & vbCrLf & _
                     "          NU_SEQU_CNTR_REPE,           " & vbCrLf & _
                     "          NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
                     "          TP_BKOF,                     " & vbCrLf & _
                     "          CO_EMPR,                     " & vbCrLf & _
                     "          CO_VEIC_LEGA,                " & vbCrLf & _
                     "          CO_LOCA_LIQU,                " & vbCrLf & _
                     "          SG_SIST,                     " & vbCrLf & _
                     "          DH_RECB_ENVI_MESG_SPB,       " & vbCrLf & _
                     "          CO_MESG_SPB,                 " & vbCrLf & _
                     "          NU_COMD_OPER,                " & vbCrLf & _
                     "          CO_SITU_MESG_SPB,            " & vbCrLf & _
                     "          CO_TEXT_XML,                 " & vbCrLf & _
                     "          HO_ENVI_MESG_SPB,            " & vbCrLf & _
                     "          CO_ULTI_SITU_PROC,           " & vbCrLf & _
                     "          CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
                     "          CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
                     "          DH_ULTI_ATLZ,                " & vbCrLf & _
                     "          NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
                     "          IN_ENTR_MANU                 " & vbCrLf
                     
            strSQL = strSQL & _
                     " FROM     A8.TB_MESG_RECB_ENVI_SPB     " & vbCrLf & _
                     " WHERE    NU_SEQU_OPER_ATIV   = " & vntNumeroSequOperAtiv & vbCrLf & _
                     " AND      CO_MESG_SPB         in ('BMA0004','BMA0005') "
            
            If xmlMesgRecebEnviSPB.loadXML(fgQueryXMLLer("MensagemRecebidaEnviada", strSQL, App.EXEName & "." & TypeName(Me))) Then
                
                Set objOperacao = CreateObject("A8LQS.clsOperacao")
                Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                
                Set xmlOperacao = objOperacao.LerXMLOperacao(vntNumeroSequOperAtiv)
                    
                Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
                Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                
                If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
                    fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
                End If
                
                Set objTipoOperacao = Nothing
                Set xmlTipoOperacao = Nothing
                
                strNumeroComando = xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
                
                Call GerenciarChamada(xmlOperacao, _
                                      xmlMesgRecebEnviSPB, _
                                      enumStatusOperacao.Cancelada, _
                                      enumStatusMensagem.MensagemCancelada, _
                                      "", _
                                      enumIndicadorSimNao.Sim, _
                                      xmlMesgRecebEnviSPB.selectSingleNode("//NU_CTRL_IF").Text)
                
                
                Set objOperacao = Nothing
            End If
            
            Set xmlMesgRecebEnviSPB = Nothing
        
        End If
        
        objRS.MoveNext
    Wend
    Exit Function
ErrorHandler:
    
    Set objMesgRecebEnviSPB = Nothing
    Set xmlMesgRecebEnviSPB = Nothing

       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0006R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0001R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0001R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objWorkflow                             As A8LQS.clsWorkflow

Dim xmlMesgBMC0001                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim intSitLancBMC                           As Integer

    On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlMesgBMC0001 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlRemessa.selectSingleNode("//SitLancBMC") Is Nothing Then
        intSitLancBMC = xmlRemessa.selectSingleNode("//SitLancBMC").Text
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Set xmlMesgBMC0001 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgBMC0001.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0001 de ida
        lngCodigoErroNegocio = 4292
        GoTo ErrorHandler
    End If

    If xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If

    'Operacao Aceita pela Camara BMC
    If intSitLancBMC = 101 Then
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.Respondida, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.AConciliarRegistro, _
                                               vbNullString, , , _
                                               False)
            
            'Atualiza dados da operacao
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            
            Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
            Call objWorkflow.Conciliacao(xmlOperacao)
            Set objWorkflow = Nothing
            
        End If

    'Operacao Rejeitada pela Camara BMC
    Else
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.MensagemRejeitada, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.Rejeitada, _
                                               vbNullString, , , _
                                               False)
            
            If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
            Else
                xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
            End If
            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        End If
    End If
    
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      "BMC0001R1", _
                                      Val(xmlMesgBMC0001.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0001.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0001.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlBMC").Text)

    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing

    Set xmlMesgBMC0001 = Nothing
    Set xmlOperacao = Nothing

    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objWorkflow = Nothing
    
    Set xmlMesgBMC0001 = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0001R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0001E
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0001E(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0001                          As MSXML2.DOMDocument40

Dim strCodMsg                               As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0001 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    Set xmlMesgBMC0001 = objMensagemEnvRec.LerXMLMensagem(udtProtocoloErroNZ.ControleRemessaNZ)
    If xmlMesgBMC0001.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0001 de ida
        lngCodigoErroNegocio = 4292
        GoTo ErrorHandler
    End If
    
    If xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgBMC0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.MensagemRejeitada, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.Rejeitada, _
                                           vbNullString, , , _
                                           False)
        
        If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
        Else
            xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
        End If
        Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
        Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    End If

    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgBMC0001.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0001.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.ErroNegocioSPB, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0001.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0001.selectSingleNode("//SG_SIST").Text)
        
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0001 = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0001 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0001E Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0002R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0002R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlMesgBMC0002                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim intSitLancBMC                           As Integer

    On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlMesgBMC0002 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlRemessa.selectSingleNode("//SitLancBMC") Is Nothing Then
        intSitLancBMC = xmlRemessa.selectSingleNode("//SitLancBMC").Text
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Set xmlMesgBMC0002 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgBMC0002.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0002 de ida
        lngCodigoErroNegocio = 4303
        GoTo ErrorHandler
    End If

    If xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If

    'Operacao Aceita pela Camara BMC
    If intSitLancBMC = 106 Then
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.Respondida, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.Cancelada, _
                                               vbNullString, , , _
                                               False)
            
            If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
            Else
                xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
            End If
            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        End If

    'Operacao Rejeitada pela Camara BMC
    ElseIf intSitLancBMC = 107 Then
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.MensagemRejeitada, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
        Call flEnviarAlertaR1(xmlMesgBMC0002)
        
        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.AConciliarRegistro, _
                                               vbNullString, , , _
                                               False)
            
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            
        End If
    
    Else
        'Situação de Lançamento BMC inválida (SitLancBMC deve ser igual a 106 ou 107)
        lngCodigoErroNegocio = 4304
        GoTo ErrorHandler
    End If

    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      "BMC0002R1", _
                                      Val(xmlMesgBMC0002.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0002.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0002.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)

    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing

    Set xmlMesgBMC0002 = Nothing
    Set xmlOperacao = Nothing

    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing

    Set xmlMesgBMC0002 = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0002R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0002E
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0002E(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0002                          As MSXML2.DOMDocument40

Dim strCodMsg                               As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0002 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    Set xmlMesgBMC0002 = objMensagemEnvRec.LerXMLMensagem(udtProtocoloErroNZ.ControleRemessaNZ)
    If xmlMesgBMC0002.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0002 de ida
        lngCodigoErroNegocio = 4303
        GoTo ErrorHandler
    End If
    
    If xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgBMC0002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.MensagemRejeitada, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgBMC0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    Call flEnviarAlertaR1(xmlMesgBMC0002)
    
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.AConciliarRegistro, _
                                           vbNullString, , , _
                                           False)
                                               
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    End If

    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgBMC0002.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0002.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.ErroNegocioSPB, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0002.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0002.selectSingleNode("//SG_SIST").Text)
        
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0002 = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0002 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0002E Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0003R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0003R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlMesgBMC0003                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40

Dim intSitLancBMC                           As Integer
Dim strXML                                  As String

    On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlMesgBMC0003 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlRemessa.selectSingleNode("//SitLancBMC") Is Nothing Then
        intSitLancBMC = xmlRemessa.selectSingleNode("//SitLancBMC").Text
    End If

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Set xmlMesgBMC0003 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgBMC0003.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0003 de ida
        lngCodigoErroNegocio = 4307
        GoTo ErrorHandler
    End If

    If xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If

    'Operacao Aceita pela Camara BMC
    If intSitLancBMC = 108 Then
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0003.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0003.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.Respondida, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0003.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.AConciliarRegistro, _
                                               vbNullString, , , _
                                               False)
        
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        strXML = objProcessoOperacao.LerTodos("RecebimentoBoleto", _
                                              CLng(xmlOperacao.selectSingleNode("//TP_OPER").Text), _
                                              enumIndicadorSimNao.Nao)
        
            If strXML <> vbNullString Then
                Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                xmlControleProcessoOperacao.loadXML strXML
                objProcessoOperacao.EnviarPrevisao xmlOperacao, xmlControleProcessoOperacao
                Set xmlControleProcessoOperacao = Nothing
            End If
        End If

    'Operacao Rejeitada pela Camara BMC
    ElseIf intSitLancBMC = 109 Then
        Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0003.selectSingleNode("//NU_CTRL_IF").Text, _
                                             xmlMesgBMC0003.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                             enumStatusMensagem.MensagemRejeitada, _
                                             vbNullString, _
                                             , _
                                             False, _
                                             xmlMesgBMC0003.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
        Call flEnviarAlertaR1(xmlMesgBMC0003)
        
        If xmlOperacao.xml <> vbNullString Then
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.Cancelada, _
                                               vbNullString, , , _
                                               False)
                                               
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        End If
    
    Else
        'Situação de Lançamento BMC inválida (SitLancBMC deve ser igual a 108 ou 109)
        lngCodigoErroNegocio = 4308
        GoTo ErrorHandler
    End If

    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      "BMC0003R1", _
                                      Val(xmlMesgBMC0003.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0003.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0003.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)

    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing

    Set xmlMesgBMC0003 = Nothing
    Set xmlOperacao = Nothing

    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing

    Set xmlMesgBMC0003 = Nothing
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0003R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0003E
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0003E(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0003                          As MSXML2.DOMDocument40

Dim strCodMsg                               As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0003 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    
    Set xmlMesgBMC0003 = objMensagemEnvRec.LerXMLMensagem(udtProtocoloErroNZ.ControleRemessaNZ)
    If xmlMesgBMC0003.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0003 de ida
        lngCodigoErroNegocio = 4307
        GoTo ErrorHandler
    End If
    
    If xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0003.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgBMC0003.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.MensagemRejeitada, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgBMC0003.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    Call flEnviarAlertaR1(xmlMesgBMC0003)
    
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.Cancelada, _
                                           vbNullString, , , _
                                           False)
                                           
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    End If

    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgBMC0003.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0003.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.ErroNegocioSPB, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0003.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0003.selectSingleNode("//SG_SIST").Text)
        
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0003 = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0003 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0003E Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0012R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0012R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0012                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0011                          As MSXML2.DOMDocument40

Dim intSitLancBMC                           As Integer

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0012 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0011 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//SitLancBMC") Is Nothing Then
        intSitLancBMC = xmlRemessa.selectSingleNode("//SitLancBMC").Text
    End If
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Set xmlMesgBMC0012 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgBMC0012.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0012 de ida
        lngCodigoErroNegocio = 4248
        GoTo ErrorHandler
    End If
    
    Set xmlMesgBMC0011 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0011", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgBMC0011.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0011 original
        lngCodigoErroNegocio = 4249
        GoTo ErrorHandler
    End If
    
    If intSitLancBMC = 301 Or intSitLancBMC = 303 Then
        Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text))
        
        If xmlOperacao.xml <> vbNullString Then
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        End If
    End If
    
    Select Case intSitLancBMC
        Case 301 'Confirmação aceita
    
            If Val(xmlMesgBMC0011.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) = enumTipoAcao.RegistroContingencia Then
                GoTo GravarMensagem
            End If
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.Registrada, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            If xmlOperacao.xml <> vbNullString Then
                Call objOperacao.AlterarStatusAcao( _
                        xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Registrada, _
                        vbNullString, , , _
                        False)
            
                Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            End If
    
        Case 302 'Rejeição aceita
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.MensagemRejeitada, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0011.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.MensagemRejeitada, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0011.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
        Case 303 'Confirmação não aceita
    
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.MensagemRejeitada, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            Call flEnviarAlertaR1(xmlMesgBMC0012)
                
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0011.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.AConciliar, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0011.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            If xmlOperacao.xml <> vbNullString Then
                Call objOperacao.AlterarStatusAcao( _
                        xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Rejeitada, _
                        vbNullString, , , _
                        False)
            
                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            End If
    
        Case 304 'Rejeição aceita
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.MensagemRejeitada, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
            Call flEnviarAlertaR1(xmlMesgBMC0012)
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgBMC0011.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.AConciliar, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgBMC0011.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
        Case Else
            
            'Situação de Lançamento BMC inválida (SitLancBMC deve ser igual a 301, 302, 303 ou 304)
            lngCodigoErroNegocio = 4247
            GoTo ErrorHandler
            
    End Select
    
GravarMensagem:
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      "BMC0012R1", _
                                      Val(xmlMesgBMC0012.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0012.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0012.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
        
ExitSub:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0012R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0012E
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0012E(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pstrHeaderNZ As String) As Boolean

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0012                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0011                          As MSXML2.DOMDocument40

Dim strCodMsg                               As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0012 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0011 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    Set xmlMesgBMC0012 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0012", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    If xmlMesgBMC0012.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0012 de ida
        lngCodigoErroNegocio = 4248
        GoTo ErrorHandler
    End If
    
    Set xmlMesgBMC0011 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0011", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgBMC0011.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0011 original
        lngCodigoErroNegocio = 4249
        GoTo ErrorHandler
    End If
    
    Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text))
    
    If xmlOperacao.xml <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.MensagemRejeitada, _
            vbNullString, _
            , _
            False, _
            xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgBMC0011.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.AConciliar, _
            vbNullString, _
            , _
            False, _
            xmlMesgBMC0011.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            
    Call flEnviarAlertaR1(xmlMesgBMC0012)
                
    If Val(xmlMesgBMC0011.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) = enumTipoAcao.RegistroContingencia Or _
       Val(xmlMesgBMC0011.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) = enumTipoAcao.DiscordanciaAdmBO Then
        GoTo GravarMensagem
    End If
    
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao( _
                xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                enumStatusOperacao.Rejeitada, _
                vbNullString, , , _
                False, _
                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
            
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
    End If

GravarMensagem:
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgBMC0012.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgBMC0012.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.ErroNegocioSPB, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgBMC0012.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgBMC0012.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
        
ExitSub:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0012E Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens CAM0002R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemCAM0002R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow

Dim xmlMesgCAM0002                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim intSitLancBMC                           As Integer

    On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Set xmlMesgCAM0002 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Set xmlMesgCAM0002 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgCAM0002.xml = vbNullString Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    If xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> vbNullString Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    'Altera Status da Mensagem Original
    Call objMensagemEnvRec.AlterarStatus(xmlMesgCAM0002.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                         xmlMesgCAM0002.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.Respondida, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgCAM0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

    'Altera Status da Operacao
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.AConciliarAceite, _
                                           vbNullString, , , _
                                           False)
    End If

            
    'Inclui mensagem
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      "CAM0002R1", _
                                      Val(xmlMesgCAM0002.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgCAM0002.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.R1, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgCAM0002.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//SG_SIST").Text, , _
                                      xmlRemessa.selectSingleNode("//NumCtrlCAM").Text)
    
    If xmlOperacao.xml <> vbNullString Then
        
        'Atualiza dados da operacao depois da atualizacao
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            
        'Verifica se ja existe BMC0005 e faz a conciliacao
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        Call objWorkflow.Conciliacao(xmlOperacao)
        Set objWorkflow = Nothing
        
    End If

    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing

    Set xmlMesgCAM0002 = Nothing
    Set xmlOperacao = Nothing

    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objWorkflow = Nothing
    
    Set xmlMesgCAM0002 = Nothing
    Set xmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCAM0002R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens CAM0002E
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemCAM0002E(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                            ByRef pstrHeaderNZ As String) As Boolean

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgCAM0002                          As MSXML2.DOMDocument40

Dim strCodMsg                               As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgCAM0002 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlRemessa.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlRemessa.selectSingleNode("//CodMsg").Text
    End If
    
    udtProtocoloErroNZAux.String = pstrHeaderNZ
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux

    Set xmlMesgCAM0002 = objMensagemEnvRec.LerXMLMensagem(udtProtocoloErroNZ.ControleRemessaNZ)
    If xmlMesgCAM0002.xml = vbNullString Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    If xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text <> "" Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If
    
    'Altera Status da Mensagem Original
    Call objMensagemEnvRec.AlterarStatus(xmlMesgCAM0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgCAM0002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.MensagemRejeitada, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgCAM0002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

    'Altera o status da operacao
    If xmlOperacao.xml <> vbNullString Then
        Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.Rejeitada, _
                                           vbNullString, , , _
                                           False)
        
        If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
        Else
            xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
        End If
        Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
        Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    End If
    
    'Inclui mensagem
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_CTRL_IF").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                      strCodMsg, _
                                      Val(xmlMesgCAM0002.selectSingleNode("//TP_BKOF").Text), _
                                      xmlMesgCAM0002.selectSingleNode("//CO_EMPR").Text, _
                                      vbNullString, _
                                      vbNullString, _
                                      enumStatusMensagem.ErroNegocioSPB, _
                                      enumIndicadorSimNao.Nao, _
                                      vbNullString, _
                                      enumLocalLiquidacao.BMC, _
                                      xmlMesgCAM0002.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                      xmlMesgCAM0002.selectSingleNode("//SG_SIST").Text)
        
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgCAM0002 = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgCAM0002 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCAM0002E Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0010
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0010(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                  Optional ByRef pstrHeaderNZ As String = vbNullString) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0012                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0011                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0010                          As MSXML2.DOMDocument40

Dim blnAchouBM0012R1                        As Boolean
Dim lngStatusMensagem                       As Long

Dim vntNumSequOperacao                      As Variant
Dim strNumConcilOperacao                    As String
Dim strCodigoVeicLegal                      As String
Dim strSiglaSistema                         As String
Dim strCodigoEmpresa                        As String
Dim strCodigoOperacao                       As String
Dim intTipoBackOffice                       As Integer
Dim strDataRegistroMensagem                 As String
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0012 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0011 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0010 = CreateObject("MSXML2.DOMDocument.4.0")
    
    intTipoBackOffice = enumTipoBackOffice.Tesouraria
    strNumConcilOperacao = "0"
    
    Set xmlMesgBMC0012 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0012R1", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    blnAchouBM0012R1 = True
    
    If xmlMesgBMC0012.xml = vbNullString Then
        blnAchouBM0012R1 = False
    ElseIf xmlMesgBMC0012.selectSingleNode("//SitLancBMC") Is Nothing Then
        blnAchouBM0012R1 = False
    ElseIf xmlMesgBMC0012.selectSingleNode("//SitLancBMC").Text <> "301" Then
        blnAchouBM0012R1 = False
    End If
        
    If Not blnAchouBM0012R1 Then
        If xmlRemessa.selectSingleNode("//NU_PROC") Is Nothing Then
            'Não foi possível encontrar a mensagem BMC0011 original
            lngCodigoErroNegocio = 4264
            GoTo ErrorHandler
        Else
            If Val(xmlRemessa.selectSingleNode("//NU_PROC").Text) > 3 Then
                If xmlMesgBMC0012.xml <> vbNullString Then
                    Call flEnviarAlertaR1(xmlMesgBMC0012)
                End If
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                GoTo GravarMensagem
            Else
                'Não foi possível encontrar a mensagem BMC0011 original
                lngCodigoErroNegocio = 4264
                GoTo ErrorHandler
            End If
        End If
    Else
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
    End If
    
    Set xmlMesgBMC0011 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0011", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgBMC0011.xml = vbNullString Then
        'Não foi possível encontrar a mensagem BMC0011 original
        lngCodigoErroNegocio = 4249
        GoTo ErrorHandler
    End If
    
    strNumConcilOperacao = xmlMesgBMC0011.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text
    
    Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text))
    
    If xmlOperacao.xml = vbNullString Then
        vntNumSequOperacao = Empty
        strCodigoVeicLegal = xmlMesgBMC0011.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlMesgBMC0011.selectSingleNode("//SG_SIST").Text
        strCodigoEmpresa = xmlMesgBMC0011.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlMesgBMC0011.selectSingleNode("//NU_CTRL_IF").Text
    Else
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        
        vntNumSequOperacao = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        strCodigoVeicLegal = xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
        strCodigoEmpresa = xmlOperacao.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
    End If
    
GravarMensagem:
    
    If pstrHeaderNZ <> vbNullString Then
        
        udtHeaderMensagemAux.String = pstrHeaderNZ
        LSet udtHeaderMensagem = udtHeaderMensagemAux

        strDataRegistroMensagem = objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                                               Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                                               vntNumSequOperacao, _
                                                               strNumConcilOperacao, _
                                                               "BMC0010", _
                                                               intTipoBackOffice, _
                                                               udtHeaderMensagem.CodigoEmpresa, _
                                                               vbNullString, _
                                                               vbNullString, _
                                                               lngStatusMensagem, _
                                                               enumIndicadorSimNao.Nao, _
                                                               vbNullString, _
                                                               enumLocalLiquidacao.BMC, _
                                                               strCodigoVeicLegal, _
                                                               strSiglaSistema, , _
                                                               xmlRemessa.selectSingleNode("//NumCtrlBMC").Text)
                                          
        Call objConciliacao.Incluir(Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                    strDataRegistroMensagem, _
                                    xmlRemessa.selectSingleNode("//DtMovto").Text, _
                                    0, _
                                    0, _
                                    vbNullString, _
                                    0, _
                                    0, _
                                    xmlRemessa.selectSingleNode("//VlrMN").Text, , _
                                    xmlRemessa.selectSingleNode("//DtLiquid").Text)
                                    
    Else
    
        If Not blnAchouBM0012R1 Then
            'Não foi possível encontrar a resposta da mensagem BMC0012 com Situação de Lançamento = 301 (Confirmação aceita pela câmara).
            lngCodigoErroNegocio = 4251
            GoTo ErrorHandler
        End If

        Set xmlMesgBMC0010 = objMensagemEnvRec.LerXMLMensagem(xmlRemessa.selectSingleNode("//NU_CTRL_IF").Text, , , _
                                                              "BMC0010")
    
        Call objMensagemEnvRec.AlterarStatus( _
                xmlMesgBMC0010.selectSingleNode("//NU_CTRL_IF").Text, _
                xmlMesgBMC0010.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                lngStatusMensagem, _
                vbNullString, _
                , _
                False, _
                xmlMesgBMC0010.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , _
                strNumConcilOperacao, , _
                Val(vntNumSequOperacao), strCodigoVeicLegal, strSiglaSistema)
                
    End If
        
    If lngStatusMensagem = enumStatusMensagem.EnviadaBUS Then
        
        strTipoMensagem = IIf(strSiglaSistema = "WZ", "2140", "140")
        
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_MESG", fgCompletaString(strTipoMensagem, "0", 9, True))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_DEST", fgCompletaString(strSiglaSistema, " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_EMPR", Format$(strCodigoEmpresa, "00000"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "DT_MESG", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        Call fgAppendNode(xmlRemessa, "SISMSG", "HO_MESG", Format$(Now, "mmss"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_RETN", "1")
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_OPER_ATIV", strCodigoOperacao)
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_VEIC_LEGA", strCodigoVeicLegal)
        
        udtProtocolo.CodigoEmpresa = Format$(strCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
        
        LSet udtProtocoloAux = udtProtocolo
        
        Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        If xmlOperacao.xml <> vbNullString Then
            
            xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//VlrMN").Text
            Call fgAppendNode(xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.Realizado)
        
        Else
            
            strTipoMensagem = IIf(strTipoMensagem = "2140", "140", "2140")
            strSiglaSistema = IIf(strSiglaSistema = "WZ", "E2", "WZ")
            
            xmlRemessa.selectSingleNode("//TP_MESG").Text = fgCompletaString(strTipoMensagem, "0", 9, True)
            xmlRemessa.selectSingleNode("//HO_MESG").Text = Format$(Now, "mmss")
            xmlRemessa.selectSingleNode("//SG_SIST_DEST").Text = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
            udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            LSet udtProtocoloAux = udtProtocolo
            Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        End If
    
    End If
    
ExitSub:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    Set xmlMesgBMC0010 = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0012 = Nothing
    Set xmlMesgBMC0011 = Nothing
    Set xmlMesgBMC0010 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0010 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------
'Processamento das mensagens STR0011 (Mensagens de cancelamento via Entrada Manual)
'-------------------------------------------------------------------------------------------

Private Function flProcessaMensagemSTR0011(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgR1                               As MSXML2.DOMDocument40
Dim xmlMesgIda                              As MSXML2.DOMDocument40
Dim xmlMesgOriginal                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim blnAchouPendente                        As Boolean
Dim strTipoMensagem                         As String
Dim strNomeTagSitLanc                       As String
Dim strSeqOperAnterior                      As String
Dim lngStatusOperacao                       As Long

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgR1 = CreateObject("MSXML2.DOMDocument.4.0")
    
    strTipoMensagem = "LTR0004R1"
    strNomeTagSitLanc = "SitLancLTR"
    lngStatusOperacao = enumStatusOperacao.Registrada
    
    'Procura mensagem R1 em condições de cancelamento
    Set xmlMesgR1 = objMensagemEnvRec.LerXMLMensagem(, , , strTipoMensagem, _
                                                           xmlRemessa.selectSingleNode("//NumCtrlSTROr").Text)
    
    blnAchouPendente = True
    
    If xmlMesgR1.xml = vbNullString Then
        blnAchouPendente = False
    ElseIf xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc) Is Nothing Then
        blnAchouPendente = False
    Else
        If Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 17 And _
           Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 19 And _
           Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 24 And _
           Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 25 Then
            blnAchouPendente = False
        End If
    End If
        
    If Not blnAchouPendente Then
        
        strTipoMensagem = "LDL0004R1"
        strNomeTagSitLanc = "SitLancLDL"
        lngStatusOperacao = enumStatusOperacao.AConciliar
        
        'Procura mensagem R1 em condições de cancelamento
        Set xmlMesgR1 = objMensagemEnvRec.LerXMLMensagem(, , , strTipoMensagem, _
                                                               xmlRemessa.selectSingleNode("//NumCtrlSTROr").Text)
        
        blnAchouPendente = True
        
        If xmlMesgR1.xml = vbNullString Then
            blnAchouPendente = False
        ElseIf xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc) Is Nothing Then
            blnAchouPendente = False
        Else
            If Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 17 And _
               Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 19 And _
               Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 24 And _
               Val(xmlMesgR1.selectSingleNode("//" & strNomeTagSitLanc).Text) <> 25 Then
                blnAchouPendente = False
            End If
        End If
            
        If Not blnAchouPendente Then
            'Não foi possível encontrar a mensagem de retorno em condições de cancelamento.
            lngCodigoErroNegocio = 4254
            GoTo ErrorHandler
        End If
    
    End If
    
    'Cancela mensagem R1
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgR1.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgR1.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.CanceladaOrigem, _
            vbNullString, _
            , _
            False, _
            xmlMesgR1.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
    'Cancela mensagem de ida
    strTipoMensagem = Replace$(strTipoMensagem, "R1", vbNullString)
    
    Set xmlMesgIda = objMensagemEnvRec.LerXMLMensagem(xmlMesgR1.selectSingleNode("//NU_CTRL_IF").Text, , , _
                                                      strTipoMensagem)
    
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgIda.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgIda.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.CanceladaOrigem, _
            vbNullString, _
            , _
            False, _
            xmlMesgIda.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
    Set xmlMesgOriginal = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Cancela mensagem original
    If Trim$(xmlMesgIda.selectSingleNode("//NU_CTRL_CAMR").Text) <> vbNullString Then
        
        strTipoMensagem = IIf(strTipoMensagem = "LTR0004", "LTR0001", "LDL0001")
        
        Set xmlMesgOriginal = objMensagemEnvRec.LerXMLMensagem(, , , strTipoMensagem, _
                                                                     xmlMesgIda.selectSingleNode("//NU_CTRL_CAMR").Text)
        
        If xmlMesgOriginal.xml <> vbNullString Then
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMesgOriginal.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMesgOriginal.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.CanceladaOrigem, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMesgOriginal.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            'Trata operações conciliadas
            Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgOriginal.selectSingleNode("//NU_CTRL_IF").Text))
            
            strSeqOperAnterior = vbNullString
            For Each objNode In xmlOperacao.selectNodes("Repeat_ConsultaConciliacao/*")
                
                If strSeqOperAnterior <> objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                    
                    strSeqOperAnterior = objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                    If Val(strSeqOperAnterior) <> 0 Then
                        Call objOperacao.AlterarStatusAcao( _
                                strSeqOperAnterior, _
                                lngStatusOperacao, _
                                vbNullString, , , _
                                False)
                    End If
                
                End If
            
            Next
        End If
        
    End If
            
    'Atualiza dados do xmlRemessa antes de devolver à entrada manual
    If Not xmlMesgOriginal.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG") Is Nothing Then
        Call fgAppendNode(xmlRemessa, "SISMSG", "NU_SEQU_CNCL_OPER_ATIV_MESG", xmlMesgOriginal.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    Else
        Call fgAppendNode(xmlRemessa, "SISMSG", "NU_SEQU_CNCL_OPER_ATIV_MESG", "0")
    End If
    
    Call fgAppendNode(xmlRemessa, "SISMSG", "CO_LOCA_LIQU", xmlMesgR1.selectSingleNode("//CO_LOCA_LIQU").Text)
    Call fgAppendNode(xmlRemessa, "SISMSG", "CO_VEIC_LEGA", xmlMesgR1.selectSingleNode("//CO_VEIC_LEGA").Text)
    Call fgAppendNode(xmlRemessa, "SISMSG", "TP_BKOF", xmlMesgR1.selectSingleNode("//TP_BKOF").Text)
    Call fgAppendNode(xmlRemessa, "SISMSG", "CO_EMPR", xmlMesgR1.selectSingleNode("//CO_EMPR").Text)
    Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST", xmlMesgR1.selectSingleNode("//SG_SIST").Text)
    
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgR1 = Nothing
    Set xmlMesgIda = Nothing
    Set xmlMesgOriginal = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgR1 = Nothing
    Set xmlMesgIda = Nothing
    Set xmlMesgOriginal = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemSTR0011 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'----------------------------------------------------------------------------------------------
'Processamento das mensagens LTR0003 e LTR0004 (Mensagens de pagamento rejeitadas e reenviadas)
'----------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLTR0003_LTR0004(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgOriginal                         As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim strSeqOperAnterior                      As String
Dim strDataRegistroMensagem                 As String

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgOriginal = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Altera Status da mensagem original para Pagamento Liberado
    Set xmlMesgOriginal = objMensagemEnvRec.LerXMLMensagem(, , , "LTR0001", _
                                                                 xmlRemessa.selectSingleNode("//NU_CTRL_CAMR").Text)
    
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgOriginal.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgOriginal.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.PagamentoLib, _
            vbNullString, _
            , _
            False, _
            xmlMesgOriginal.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
    'Trata operações conciliadas
    Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(, xmlMesgOriginal.selectSingleNode("//NU_CTRL_IF").Text))
    
    strSeqOperAnterior = vbNullString
    For Each objNode In xmlOperacao.selectNodes("Repeat_ConsultaConciliacao/*")
        
        If strSeqOperAnterior <> objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
            
            strSeqOperAnterior = objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
            If Val(strSeqOperAnterior) <> 0 Then
                Call objOperacao.AlterarStatusAcao( _
                        strSeqOperAnterior, _
                        enumStatusOperacao.PagamentoLib, _
                        vbNullString, , , _
                        False)
            End If
        
        End If
    
    Next
            
    strDataRegistroMensagem = xmlRemessa.selectSingleNode("//DataRegistroNovaMensagem").Text
    
    'Inclui mensagem de pagamento na tabela de conciliação
    Call objConciliacao.Incluir(Trim(xmlRemessa.selectSingleNode("//NU_CTRL_IF").Text), _
                                strDataRegistroMensagem, _
                                xmlRemessa.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                xmlRemessa.selectSingleNode("//VA_FINC").Text, _
                                xmlRemessa.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                "", _
                                xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text, _
                                "", _
                                xmlRemessa.selectSingleNode("//CNPJNLiqdant").Text)

    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgOriginal = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgOriginal = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLTR0003_LTR0004 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0013
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0013(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                  Optional ByRef pstrHeaderNZ As String = vbNullString) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0010                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0013                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0015                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0001                          As MSXML2.DOMDocument40

Dim lngStatusMensagem                       As Long

Dim vntNumSequOperacao                      As String
Dim strNumConcilOperacao                    As String
Dim strCodigoVeicLegal                      As String
Dim strSiglaSistema                         As String
Dim strTipoMensagem                         As String
Dim strCodigoEmpresa                        As String
Dim strCodigoOperacao                       As String
Dim intTipoBackOffice                       As Integer

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgBMC0015 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0001 = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlMesgBMC0015 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0015", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    Set xmlMesgBMC0001 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0001R1", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgBMC0015.xml <> vbNullString Then
        Call flProcessaMensagemBMC0013CAM(xmlRemessa, pstrHeaderNZ)
        Set objMensagemEnvRec = Nothing
        Set xmlMesgBMC0015 = Nothing
        Set xmlMesgBMC0001 = Nothing
        Exit Function
    ElseIf xmlMesgBMC0001.xml <> vbNullString Then
        Call flProcessaMensagemBMC0013CAM(xmlRemessa, pstrHeaderNZ)
        Set objMensagemEnvRec = Nothing
        Set xmlMesgBMC0015 = Nothing
        Set xmlMesgBMC0001 = Nothing
        Exit Function
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0010 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0013 = CreateObject("MSXML2.DOMDocument.4.0")
    
    intTipoBackOffice = enumTipoBackOffice.Tesouraria
    strNumConcilOperacao = "0"
    
    Set xmlMesgBMC0010 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0010", _
                                                                xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgBMC0010.xml = vbNullString Then
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        GoTo GravarMensagem
    Else
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
    End If
    
    strNumConcilOperacao = xmlMesgBMC0010.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text
    
    If Val(xmlMesgBMC0010.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) <> 0 Then
        Call xmlOperacao.loadXML(objConciliacao.ConsultaConciliacao(xmlMesgBMC0010.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
    End If
    
    If xmlOperacao.xml = vbNullString Then
        vntNumSequOperacao = Empty
        strCodigoVeicLegal = xmlMesgBMC0010.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlMesgBMC0010.selectSingleNode("//SG_SIST").Text
        strCodigoEmpresa = xmlMesgBMC0010.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlMesgBMC0010.selectSingleNode("//NU_CTRL_IF").Text
    Else
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        
        vntNumSequOperacao = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        strCodigoVeicLegal = xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
        strCodigoEmpresa = xmlOperacao.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
    End If
    
    Call objMensagemEnvRec.AlterarStatus( _
            xmlMesgBMC0010.selectSingleNode("//NU_CTRL_IF").Text, _
            xmlMesgBMC0010.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
            enumStatusMensagem.MensagemCancelada, _
            vbNullString, _
            , _
            False, _
            xmlMesgBMC0010.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
GravarMensagem:
    
    If pstrHeaderNZ <> vbNullString Then
        
        udtHeaderMensagemAux.String = pstrHeaderNZ
        LSet udtHeaderMensagem = udtHeaderMensagemAux
    
        Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                          Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                          vntNumSequOperacao, _
                                          strNumConcilOperacao, _
                                          "BMC0013", _
                                          intTipoBackOffice, _
                                          udtHeaderMensagem.CodigoEmpresa, _
                                          vbNullString, _
                                          vbNullString, _
                                          lngStatusMensagem, _
                                          enumIndicadorSimNao.Nao, _
                                          vbNullString, _
                                          enumLocalLiquidacao.BMC, _
                                          strCodigoVeicLegal, _
                                          strSiglaSistema, , _
                                          xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
        
    Else
    
        If xmlMesgBMC0010.xml = vbNullString Then
            'Não foi possível encontrar a mensagem BMC0010.
            lngCodigoErroNegocio = 4252
            GoTo ErrorHandler
        End If
        
        Set xmlMesgBMC0013 = objMensagemEnvRec.LerXMLMensagem(xmlRemessa.selectSingleNode("//NU_CTRL_IF").Text, , , _
                                                              "BMC0013")
    
        Call objMensagemEnvRec.AlterarStatus( _
                xmlMesgBMC0013.selectSingleNode("//NU_CTRL_IF").Text, _
                xmlMesgBMC0013.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                lngStatusMensagem, _
                vbNullString, _
                , _
                False, _
                xmlMesgBMC0013.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , _
                strNumConcilOperacao)
                
    End If
        
    If lngStatusMensagem = enumStatusMensagem.EnviadaBUS Then
        
        strTipoMensagem = IIf(strSiglaSistema = "WZ", "2140", "140")
        
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_MESG", fgCompletaString(strTipoMensagem, "0", 9, True))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_DEST", fgCompletaString(strSiglaSistema, " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_EMPR", Format$(strCodigoEmpresa, "00000"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "DT_MESG", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        Call fgAppendNode(xmlRemessa, "SISMSG", "HO_MESG", Format$(Now, "mmss"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_RETN", "2")
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_OPER_ATIV", strCodigoOperacao)
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_VEIC_LEGA", strCodigoVeicLegal)
        
        udtProtocolo.CodigoEmpresa = Format$(strCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
        
        LSet udtProtocoloAux = udtProtocolo
        
        Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        If xmlOperacao.xml <> vbNullString Then
            
            xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = xmlMesgBMC0010.selectSingleNode("//VA_FINC").Text
            Call fgAppendNode(xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
        
        Else
            
            strTipoMensagem = IIf(strTipoMensagem = "2140", "140", "2140")
            strSiglaSistema = IIf(strSiglaSistema = "WZ", "E2", "WZ")
            
            xmlRemessa.selectSingleNode("//TP_MESG").Text = fgCompletaString(strTipoMensagem, "0", 9, True)
            xmlRemessa.selectSingleNode("//HO_MESG").Text = Format$(Now, "mmss")
            xmlRemessa.selectSingleNode("//SG_SIST_DEST").Text = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
            udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            LSet udtProtocoloAux = udtProtocolo
            Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        End If
    
    End If
    
ExitSub:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0010 = Nothing
    Set xmlMesgBMC0013 = Nothing
    Set xmlMesgBMC0015 = Nothing
    Set xmlMesgBMC0001 = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0010 = Nothing
    Set xmlMesgBMC0013 = Nothing
    Set xmlMesgBMC0015 = Nothing
    Set xmlMesgBMC0001 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0013 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0013 para Operacoes de Cambio
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0013CAM(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                     Optional ByRef pstrHeaderNZ As String = vbNullString) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgBMC0015                          As MSXML2.DOMDocument40
Dim xmlMesgBMC0001                          As MSXML2.DOMDocument40
Dim xmlMesgOrig                             As MSXML2.DOMDocument40

Dim lngStatusMensagem                       As Long

Dim vntNumSequOperacao                      As String
Dim strNumConcilOperacao                    As String
Dim strCodigoVeicLegal                      As String
Dim strSiglaSistema                         As String
Dim strTipoMensagem                         As String
Dim strCodigoEmpresa                        As String
Dim strCodigoOperacao                       As String
Dim intTipoBackOffice                       As Integer

    On Error GoTo ErrorHandler
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0015 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgBMC0001 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgOrig = CreateObject("MSXML2.DOMDocument.4.0")
    
    intTipoBackOffice = enumTipoBackOffice.Tesouraria
    strNumConcilOperacao = "0"
    
    Set xmlMesgOrig = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0015", _
                                                       xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
    
    If xmlMesgOrig.xml = vbNullString Then
        Set xmlMesgOrig = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0001R1", _
                                                           xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
        If xmlMesgOrig.xml = vbNullString Then
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            GoTo GravarMensagem
        Else
            lngStatusMensagem = enumStatusMensagem.Informação
        End If
    Else
        lngStatusMensagem = enumStatusMensagem.Informação
    End If

    'Busca Dados da Operacao
    If xmlMesgOrig.xml <> vbNullString Then
        If xmlMesgOrig.selectSingleNode("//CO_MESG").Text = "BMC0015" Then
            Set xmlOperacao = objOperacao.LerXMLOperacao(, , , , , , , , , , xmlMesgOrig.selectSingleNode("//CD_ASSO_CAMB").Text)
        Else
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMesgOrig.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        End If
    End If
    
    If xmlOperacao.xml <> vbNullString Then
        
        strNumConcilOperacao = xmlMesgOrig.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text
        vntNumSequOperacao = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
        strCodigoVeicLegal = xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
        strCodigoEmpresa = xmlOperacao.selectSingleNode("//CO_EMPR").Text
        strCodigoOperacao = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
        
        'Status da Operacao A Conciliar Registro
        If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliarRegistro Then

            'Atualiza status da mensagem BMC0001 para Cancelada Camara
            Set xmlMesgBMC0001 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0001", _
                                                                  , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

            If xmlMesgBMC0001.xml <> vbNullString Then
                Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMesgBMC0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.CanceladaCamara, _
                                                     xmlMesgBMC0001.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                     , _
                                                     False, _
                                                     xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
            
            'Atualiza status da Operacao para Cancelada Camara
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.CanceladaCamara, _
                                               xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text)
                                               
        'Status da Operacao A Liberada Automatica e Enviada CAM0002
        ElseIf xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica _
        And xmlOperacao.selectSingleNode("//TP_ACAO_OPER_ATIV_EXEC").Text = enumTipoAcao.EnviadaCAM0002 Then
            
            'Atualiza status da mensagem BMC0015 para Cancelada
            Set xmlMesgBMC0015 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0015", _
                                                                  xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)

            If xmlMesgBMC0015.xml <> vbNullString Then
                Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0015.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMesgBMC0015.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.MensagemCancelada, _
                                                     xmlMesgBMC0015.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                     , _
                                                     False, _
                                                     xmlMesgBMC0015.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
            
            'Atualiza status da mensagem BMC0001 para Cancelada Camara
            Set xmlMesgBMC0001 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0001", _
                                                                  , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            
            If xmlMesgBMC0001.xml <> vbNullString Then
                Call objMensagemEnvRec.AlterarStatus(xmlMesgBMC0001.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMesgBMC0001.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.CanceladaCamara, _
                                                     xmlMesgBMC0001.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                     , _
                                                     False, _
                                                     xmlMesgBMC0001.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
            
            'Atualiza status da Operacao para Cancelada Camara
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.CanceladaCamara, _
                                               xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text)

        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            GoTo GravarMensagem
        End If
    End If
    
                
GravarMensagem:
    
    If pstrHeaderNZ <> vbNullString Then
        
        udtHeaderMensagemAux.String = pstrHeaderNZ
        LSet udtHeaderMensagem = udtHeaderMensagemAux
    
        Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                          Trim$(udtHeaderMensagem.ControleRemessaNZ), _
                                          vntNumSequOperacao, _
                                          strNumConcilOperacao, _
                                          "BMC0013", _
                                          intTipoBackOffice, _
                                          udtHeaderMensagem.CodigoEmpresa, _
                                          vbNullString, _
                                          vbNullString, _
                                          lngStatusMensagem, _
                                          enumIndicadorSimNao.Nao, _
                                          vbNullString, _
                                          enumLocalLiquidacao.BMC, _
                                          strCodigoVeicLegal, _
                                          strSiglaSistema, , _
                                          xmlRemessa.selectSingleNode("//NumCtrlBMCOr").Text)
                
    End If
        
    If lngStatusMensagem = enumStatusMensagem.Informação Then
        
        strTipoMensagem = IIf(strSiglaSistema = "WZ", "2140", "140")
        
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_MESG", fgCompletaString(strTipoMensagem, "0", 9, True))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "SG_SIST_DEST", fgCompletaString(strSiglaSistema, " ", 3, False))
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_EMPR", Format$(strCodigoEmpresa, "00000"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "DT_MESG", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
        Call fgAppendNode(xmlRemessa, "SISMSG", "HO_MESG", Format$(Now, "mmss"))
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_RETN", "2")
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_OPER_ATIV", strCodigoOperacao)
        Call fgAppendNode(xmlRemessa, "SISMSG", "CO_VEIC_LEGA", strCodigoVeicLegal)
        
        udtProtocolo.CodigoEmpresa = Format$(strCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
        
        LSet udtProtocoloAux = udtProtocolo
        
        Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        If xmlOperacao.xml <> vbNullString Then

            Call flGeraAlertaCamara("BMC0013", _
                                    intTipoBackOffice, _
                                    strCodigoVeicLegal, _
                                    0, _
                                    xmlRemessa)
            
            If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
            Else
                xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
            End If
            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            Call objFluxoProcessamento.EnviarMovimentoPJME(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        
        Else
            
            strTipoMensagem = IIf(strTipoMensagem = "2140", "140", "2140")
            strSiglaSistema = IIf(strSiglaSistema = "WZ", "E2", "WZ")
            
            xmlRemessa.selectSingleNode("//TP_MESG").Text = fgCompletaString(strTipoMensagem, "0", 9, True)
            xmlRemessa.selectSingleNode("//HO_MESG").Text = Format$(Now, "mmss")
            xmlRemessa.selectSingleNode("//SG_SIST_DEST").Text = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            udtProtocolo.TipoMensagem = fgCompletaString(strTipoMensagem, "0", 9, True)
            udtProtocolo.SiglaSistemaDestino = fgCompletaString(strSiglaSistema, " ", 3, False)
            
            LSet udtProtocoloAux = udtProtocolo
            Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlRemessa.xml, enumIdentificadorFila.BUS)
        
        End If
    
    End If
    
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0015 = Nothing
    Set xmlMesgBMC0001 = Nothing
    Set xmlMesgOrig = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objProcessoOperacao = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgBMC0015 = Nothing
    Set xmlMesgBMC0001 = Nothing
    Set xmlMesgOrig = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0013CAM Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Tratamento das mensagens de confirmação de pagamento (LTR0002R1, LTR0008R1)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemR1Confirmacao(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                 ByRef pstrHeaderNZ As String, _
                                        Optional ByRef pblnGerarAlerta As Boolean = False) As Long

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objLegado                               As A8LQS.clsLegado
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlMesgSPB                              As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeSISMSG                           As MSXML2.IXMLDOMNode

Dim strNumeroControleIF                     As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaOperacaoAnterior            As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim strMensagemRetorno                      As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim strCodigoMensagem                       As String
Dim strCodigoMensagemMae                    As String
Dim strMensagem                             As String
Dim strSituacaoLancamento                   As String
Dim lngAcaoMensagem                         As enumTipoAcao

Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim strRet                                  As String
Dim blnConcordancia                         As Boolean
Dim blnPagtoContingencia                    As Boolean
Dim blnAlterarStatus                        As Boolean

    On Error GoTo ErrorHandler
                    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
                    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagem = pxmlMensagem.selectSingleNode("//CodMsg").Text
                        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgSPB = CreateObject("MSXML2.DOMDocument.4.0")

    Call xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF)))
                    
    vntSequenciaOperacao = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)

    strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(strCodigoMensagem, 1, 7))
    
    If strCodigoMensagemMae <> vbNullString Then
        
        strMensagem = flObterMensagemMae(vntSequenciaConciliacaoOperacao, strCodigoMensagemMae)
        
        xmlMesgSPB.loadXML strMensagem
        
        For Each xmlNode In xmlMesgSPB.selectNodes("//Repeat_MensagemRecebidaEnviada/*")
        
            '-------------------------------------------------------------------------
            Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
            Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
            
            strNumeroControleIF = xmlNode.selectSingleNode("//NU_CTRL_IF").Text
             
            strRet = objHistMensagem.LerTodos(xmlNode.selectSingleNode("//NU_CTRL_IF").Text, _
                                              xmlNode.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                              xmlNode.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
             
            xmlHistorico.loadXML strRet
                                        
            If strCodigoMensagemMae = "LTR0001" Then
                
                blnConcordancia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']") Is Nothing, False, True)
            
            ElseIf strCodigoMensagemMae = "LDL0001" And _
                   Val(xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMC Then
                
                blnConcordancia = True
                
            Else
                
                blnPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
                
                If Not blnPagtoContingencia Then
                    If strCodigoMensagemMae = "LDL0001" Then
                        blnConcordancia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']") Is Nothing, False, True)
                    Else
                        blnConcordancia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0008Concordancia & "']") Is Nothing, False, True)
                    End If
                End If
            
            End If
            
            Set objHistMensagem = Nothing
            Set xmlHistorico = Nothing
            '----------------------------------------------------------------------------
        
            If blnConcordancia Then
                
                If strCodigoMensagemMae = "LTR0007" Then
                    
                    'CBLC - KIDA - 08/10/2008
                    If xmlNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.Liberada Or _
                       xmlNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.ConcordanciaLib Or _
                       xmlNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.ConcordanciaLibAuto Or _
                       (xmlNode.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusMensagem.ConcordanciaBackoffice And _
                        Val(xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CLBCAcoes) Then
                    
                        lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
                        
                        Call objMesgRecbEnviSPB.AlterarStatus(xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                              xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                              lngStatusMensagem, _
                                                              xmlNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                              xmlNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text, _
                                                              False, _
                                                              1, _
                                                              strCodigoMensagemMae)
                    End If
                Else
                    lngStatusMensagem = xmlNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
                End If
            Else
                
                lngAcaoMensagem = 0
                blnAlterarStatus = True
                
                If strCodigoMensagemMae = "LTR0007" Then
                    If blnPagtoContingencia Then
                        lngStatusMensagem = enumStatusMensagem.AConciliar
                    Else
                        lngStatusMensagem = enumStatusMensagem.Discordada
                    End If
                
                ElseIf strCodigoMensagemMae = "LDL0001" Then
                    If Not pxmlMensagem.selectSingleNode("//SitLancLDL") Is Nothing Then
                        strSituacaoLancamento = pxmlMensagem.selectSingleNode("//SitLancLDL").Text
                    End If
                    
                    Select Case Val(strSituacaoLancamento)
                        Case 6, 8
                            lngStatusMensagem = enumStatusMensagem.AConciliar
                            lngAcaoMensagem = enumTipoAcao.RejeicaoDiscordancia
                        Case 24, 25
                            blnAlterarStatus = False
                            pblnGerarAlerta = True
                        Case Else
                            lngStatusMensagem = enumStatusMensagem.Discordada
                    End Select
                    
                Else
                    lngStatusMensagem = enumStatusMensagem.Discordada
                
                End If
                
                If blnAlterarStatus Then
                    Call objMesgRecbEnviSPB.AlterarStatus(xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                          xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                          lngStatusMensagem, _
                                                          xmlNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                          xmlNode.selectSingleNode("DH_RECB_ENVI_MESG_SPB").Text, _
                                                          False, _
                                                          0, _
                                                          strCodigoMensagemMae, _
                                                          lngAcaoMensagem)
                End If
            
            End If
        Next
    End If
    
    If blnConcordancia Or blnPagtoContingencia Then
        'KIDA - BMF  - 16/07/2009
        If xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Or _
           xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD Then


            
            If strCodigoMensagem = "LTR0002R1" Or strCodigoMensagem = "LDL0003R1" Then
                lngStatusMensagem = enumStatusMensagem.Confirmada
            Else
                lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
            End If
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
        End If
    Else
        'KIDA - CBLC - 24/09/2008
        'KIDA - BMF  - 16/07/2009
        If xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Or _
           xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD Then

            
            If strCodigoMensagem = "LTR0002R1" Or strCodigoMensagem = "LDL0003R1" Then
                lngStatusMensagem = enumStatusMensagem.Confirmada
            Else
                lngStatusMensagem = enumStatusMensagem.MensagemLiquidada
            End If
            
        Else
            lngStatusMensagem = enumStatusMensagem.Discordada
        End If
    End If

    If pblnGerarAlerta Then
        lngStatusMensagem = enumStatusMensagem.MensagemPendente
    End If
    
    objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                     lngStatusMensagem, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                     , _
                                     False

    Set xmlMesgSPB = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    vntSequenciaOperacaoAnterior = vbNullString
    
    If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao, 0, strNumeroControleIF)) Then
        
        For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
            
            vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
            
            If Val(vntSequenciaOperacao) <> 0 And _
               vntSequenciaOperacaoAnterior <> vntSequenciaOperacao Then
            
                vntSequenciaOperacaoAnterior = vntSequenciaOperacao
                
                Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                
                Set xmlNodeSISMSG = xmlOperacao.selectSingleNode("//SISMSG")
                
                If Not xmlNodeSISMSG Is Nothing Then
                    xmlOperacao.documentElement.removeChild xmlNodeSISMSG
                End If
                
                If blnConcordancia Or blnPagtoContingencia Then
                    If strCodigoMensagemMae = "LTR0001" Then
                        
                        GerenciarChamada xmlOperacao, _
                                         pxmlMensagem, _
                                         lngStatusOperacao, _
                                         0, _
                                         "", _
                                         enumIndicadorSimNao.Nao, _
                                         "", _
                                         False, _
                                         False
                    
                    ElseIf strCodigoMensagemMae = "LTR0007" Then
                                                            
                        If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada Or _
                           xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica Or _
                           xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaLibAuto Or _
                           xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaLib Then
                            
                            lngStatusOperacao = enumStatusOperacao.Liquidada
                            
                            GerenciarChamada xmlOperacao, _
                                             pxmlMensagem, _
                                             lngStatusOperacao, _
                                             0, _
                                             "", _
                                             enumIndicadorSimNao.Nao, _
                                             "", _
                                             False, _
                                             True
                                             
                        ElseIf (xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada Or _
                                xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBackoffice) And _
                                xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes Then
                               
                            lngStatusOperacao = enumStatusOperacao.Liquidada
                            
                            GerenciarChamada xmlOperacao, _
                                             pxmlMensagem, _
                                             lngStatusOperacao, _
                                             0, _
                                             "", _
                                             enumIndicadorSimNao.Nao, _
                                             "", _
                                             False, _
                                             True
                               
                        Else
                            
                            If blnPagtoContingencia Then
                                lngStatusOperacao = enumStatusOperacao.Registrada
                                
                                Call objOperacao.AlterarStatusOperIndisp( _
                                                vntSequenciaOperacao, _
                                                lngStatusOperacao)
                                                
                                                            
                            End If
                            
                        End If
                    End If
                Else
                    lngStatusOperacao = enumStatusOperacao.Registrada
                    
                    If Val(strSituacaoLancamento) = 6 Or _
                       Val(strSituacaoLancamento) = 8 Then
                        lngStatusMensagem = 0
                    End If
                    
                    If Not pblnGerarAlerta Then
                        GerenciarChamada xmlOperacao, _
                                         pxmlMensagem, _
                                         lngStatusOperacao, _
                                         lngStatusMensagem, _
                                         "", _
                                         enumIndicadorSimNao.Nao, _
                                         strNumeroControleIF, _
                                         False, _
                                         True
                    End If
                End If
                
            End If
        Next
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlMesgSPB = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemR1Confirmacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das menssagens LDL1027 (IF requisita Cancelamento de lançamento LDL pendente no STR)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL1027(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

'Mensagem de Inforação da Situação de Movimentação de Colateral
'Aplica se para liquidação de Movimentação de titulos entra camaras
'Ao enviar a mensagem LDL1016 , SLCC aguarda a mensagem LDL1027 para liquidacao

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objLegado                               As A8LQS.clsLegado

Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlRetorno                              As MSXML2.DOMDocument40

Dim strXmlVeiculoLegal                      As String
Dim strNomeVeiculoLegal                     As String
Dim vntCodigoVeiculoLegal                   As Variant
Dim strSiglaSistema                         As String

Dim strMensagemLegado                       As String
Dim strMesgRecbEnviSPB                      As String
Dim strNumeroControleIF                     As String
Dim vntSequenciaOperacao                    As Variant
Dim lngStatusOperacao                       As Long

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler
    
    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlRetorno.loadXML(ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemAviso))
    
    strNumeroControleIF = xmlMensagem.selectSingleNode("//NumCtrlIFOr").Text
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strMesgRecbEnviSPB = objMesgRecbEnviSPB.LerMensagemIda(strNumeroControleIF)
     
    'Se não exitr mensagem com Numero de Controle IF Original sai fora
    If strMesgRecbEnviSPB = vbNullString Then
        Set objMesgRecbEnviSPB = Nothing
        Exit Function
    End If
     
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
        
    xmlMesgRecbEnviSPB.loadXML strMesgRecbEnviSPB
     
    'Obter Status ---------------------
    Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
    xmlStatus.loadXML flObterStatusOperacaoMensagem(xmlMensagem)
    '----------------------------------
    
    If xmlStatus.selectSingleNode("//Status_Operacao").Text <> vbNullString Then
        lngStatusOperacao = xmlStatus.selectSingleNode("//Status_Operacao").Text
    End If
    
    strSituacaoRecebida = vbNullString
    If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
        strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
    End If
    
    vntSequenciaOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    
    If vntSequenciaOperacao = 0 Then Exit Function
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
    
    If lngStatusOperacao = enumStatusOperacao.Rejeitada Then
        
        vntCodigoVeiculoLegal = xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("//SG_SIST_ORIG").Text
        
        Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
        strXmlVeiculoLegal = objVeiculoLegal.Ler(vntCodigoVeiculoLegal, strSiglaSistema)
        Set objVeiculoLegal = Nothing
        
        Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlVeiculoLegal.loadXML(strXmlVeiculoLegal)
        strNomeVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//NO_VEIC_LEGA").Text
        
        Call flGeraAlertaCamara("LDL1027", _
                                xmlVeiculoLegal.selectSingleNode("//TP_BKOF").Text, _
                                xmlVeiculoLegal.selectSingleNode("//NO_VEIC_LEGA").Text, _
                                xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, _
                                xmlMensagem, , vntSequenciaOperacao)
        
        Set xmlVeiculoLegal = Nothing
    
    End If
        
    Call objMesgRecbEnviSPB.AlterarSequenciaOperacao(xmlRetorno.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     fgDtHrStr_To_DateTime(xmlRetorno.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                                                     vntSequenciaOperacao, _
                                                     xmlRetorno.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
    If Val(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> enumStatusOperacao.Rejeitada Then
        strSituacaoRecebida = vbNullString
    End If
    
    GerenciarChamada xmlOperacao, _
                     xmlMensagem, _
                     CLng(IIf(xmlStatus.selectSingleNode("//Status_Operacao").Text = "", 0, xmlStatus.selectSingleNode("//Status_Operacao").Text)), _
                     CLng(IIf(xmlStatus.selectSingleNode("//Status_Mensagem").Text = "", 0, xmlStatus.selectSingleNode("//Status_Mensagem").Text)), _
                     IIf(xmlStatus.selectSingleNode("//Situacao_Mensagem").Text = "", "", xmlStatus.selectSingleNode("//Situacao_Mensagem").Text), _
                     IIf(lngStatusOperacao = enumStatusOperacao.Rejeitada, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao), _
                     strNumeroControleIF, _
                     , _
                     , _
                     strSituacaoRecebida
    
    Set xmlRetorno = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    Set xmlStatus = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objLegado = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlRetorno = Nothing
    Set xmlVeiculoLegal = Nothing
    Set objVeiculoLegal = Nothing
    Set objOperacao = Nothing
    Set xmlStatus = Nothing
    Set objMesgRecbEnviSPB = Nothing
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL1027 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das menssagens BMA0003 (BMA informa Operação cancelada)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0003(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim objRS                                   As ADODB.Recordset

Dim strSQL                                  As String

Dim vntCNPJNEntidadeResponsavel             As Variant
Dim vntNumSequqnciaOperacao                 As Variant

Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim strNumeroComandoOperacao                As String
Dim strNumeroControleBMAOrig                As String
Dim strNomeVeicLega                         As String
Dim strDataRegistroMensagem                 As String
Dim strMensagem                             As String
Dim strNumCtrlIF                            As String
Dim strCodigoMensagemSPB                    As String

Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngStatusMensagem                       As Long


On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strCodigoMensagemSPB = Trim(udtHeaderMensagem.CodigoMensagem)
    strNumCtrlIF = udtHeaderMensagem.ControleRemessaNZ
    
    If Not xmlMensagem.selectSingleNode("//CNPJEntRespons") Is Nothing Then
        vntCNPJNEntidadeResponsavel = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
    Else
        vntCNPJNEntidadeResponsavel = vbNullString
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumOpNegcBMA") Is Nothing Then
        strNumeroComandoOperacao = xmlMensagem.selectSingleNode("//NumOpNegcBMA").Text
    Else
        strNumeroComandoOperacao = vbNullString
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlBMAOr") Is Nothing Then
        strNumeroControleBMAOrig = xmlMensagem.selectSingleNode("//NumCtrlBMAOr").Text
    Else
        strNumeroControleBMAOrig = vbNullString
    End If
      
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNEntidadeResponsavel, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistema, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        
    fgAppendNode xmlMensagem, "SISMSG", "CO_VEIC_LEGA", strCodigoVeiculoLegal
    
    flGeraAlertaCamara xmlMensagem.selectSingleNode("//CodMsg").Text, _
                       lngTipoBackOffice, _
                       strNomeVeicLega, _
                       0, _
                       xmlMensagem

    '------------------------------------
    
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "TP_MESG", fgCompletaString(Trim("44"), "0", 9, True))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "SG_SIST_DEST", fgCompletaString(Trim(strSiglaSistema), " ", 3, False))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "CO_EMPR", Format$(lngCodigoEmpresa, "00000"))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "DT_MESG", fgDt_To_Xml(fgDataHoraServidor(DataAux)))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "HO_MESG", Format$(Now, "mmss"))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "TP_RETN", "1")
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "CO_OPER_ATIV", "")
    
    'Obter Operação Original'
    strSQL = " SELECT  B.NU_SEQU_OPER_ATIV              " & _
             "   FROM   A8.TB_MESG_RECB_ENVI_SPB    A,  " & _
             "          A8.TB_CNCL_OPER_ATIV B          " & _
             "  WHERE   A.NU_CTRL_CAMR = '" & strNumeroControleBMAOrig & "'" & _
             "    AND   A.NU_CTRL_IF = B.NU_CTRL_IF "
    
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
                
        vntNumSequqnciaOperacao = objRS!NU_SEQU_OPER_ATIV
        
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
        Set xmlOperacao = objOperacao.LerXMLOperacao(vntNumSequqnciaOperacao)
        
        If Not xmlOperacao.selectSingleNode("//CO_OPER_ATIV") Is Nothing Then
            xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
        End If
                 
        Set xmlOperacao = Nothing
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
        
        lngStatusMensagem = enumStatusMensagem.Informação
        
    Else
        
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
        
    End If
    
    objRS.Close


    '------------------------------------
    
    strMensagem = xmlMensagem.xml
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumCtrlIF), _
                                                            0, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumeroComandoOperacao, _
                                                            "", _
                                                            lngStatusMensagem, _
                                                            0, _
                                                            "", _
                                                            enumLocalLiquidacao.BMA, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            1, _
                                                            strNumeroControleBMAOrig)
    
    
    Set objMesgRecbEnviSPB = Nothing
        
    'So envia para o legado se o status igual a informação
    If lngStatusMensagem = enumStatusMensagem.Informação Then
        
        udtProtocolo.CodigoEmpresa = Format$(lngCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString(Trim(strSiglaSistema), " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(Trim("44"), "0", 9, True)
        
        LSet udtProtocoloAux = udtProtocolo
        
        '------------------------------------
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & xmlMensagem.xml, enumIdentificadorFila.BUS
        Set objProcessoOperacao = Nothing
    
    End If
    
    Exit Function
ErrorHandler:
    
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das menssagens LDL0028 (Câmara avisa Abertura)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemLDL0028(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim strMensagemEnvio                        As String
Dim strNumeroControleIF                     As String
Dim lngCodigoLocalLiquidacao                As Long
Dim lngCodigoEmpresa                        As Long
Dim strHeaderMensageNZ                      As String
Dim pstrTipoMensagem                        As String
Dim strISPBCamara                           As String

On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, enumTipoMensagemSPB.MensagemAviso)

    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
           
    Select Case Val(xmlMensagem.selectSingleNode("//ISPBLDL").Text)
        Case enumISPB.IspbBMA
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
            strISPBCamara = Format(enumISPB.IspbBMA, "00000000")
        Case enumISPB.IspbCETIP
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
            strISPBCamara = Format(enumISPB.IspbCETIP, "00000000")
        Case Else
            Exit Function
    End Select
        
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA Or _
       lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        'Envia Mensagem de consulta LDL0026
        
        pstrTipoMensagem = "LDL0026"
        
        udtProtocolo.CodigoEmpresa = Format(lngCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString("NZ", " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(pstrTipoMensagem, " ", 9, False)
        LSet udtProtocoloAux = udtProtocolo
        
        strHeaderMensageNZ = fgMontaHeaderMensageNZ(pstrTipoMensagem, lngCodigoEmpresa, strNumeroControleIF)
        
        If xmlMensagem.selectSingleNode("//CO_MESG") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_MESG", fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
        Else
            xmlMensagem.documentElement.selectSingleNode("//CO_MESG").Text = fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
        End If
        
        'Pikachu 29/08/2005
        'RATS 396 - OCor 03
        xmlMensagem.documentElement.selectSingleNode("//CO_MESG").Text = Trim$(xmlMensagem.documentElement.selectSingleNode("//CO_MESG").Text)
        
        If xmlMensagem.documentElement.selectSingleNode("TX_HEAD_NZ") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "TX_HEAD_NZ", strHeaderMensageNZ
        Else
            xmlMensagem.documentElement.selectSingleNode("TX_HEAD_NZ").Text = strHeaderMensageNZ
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("NU_CTRL_IF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "NU_CTRL_IF", strNumeroControleIF
        Else
            xmlMensagem.documentElement.selectSingleNode("NU_CTRL_IF").Text = strNumeroControleIF
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "NIVE_PREF", ""
        Else
            xmlMensagem.documentElement.selectSingleNode("NIVE_PREF").Text = ""
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "DT_SIST", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
        Else
            xmlMensagem.documentElement.selectSingleNode("DT_SIST").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("CO_ISPB_CAMR") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_ISPB_CAMR", strISPBCamara
        Else
            xmlMensagem.documentElement.selectSingleNode("CO_ISPB_CAMR").Text = strISPBCamara
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("CO_ISPB_IF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_ISPB_IF", fgObterISPBIF(lngCodigoEmpresa)
        Else
            xmlMensagem.documentElement.selectSingleNode("CO_ISPB_IF").Text = fgObterISPBIF(lngCodigoEmpresa)
        End If
                
        strMensagemEnvio = xmlMensagem.xml
                
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagemEnvio, _
                                           Trim(strNumeroControleIF), _
                                           0, _
                                           0, _
                                           pstrTipoMensagem, _
                                           0, _
                                           lngCodigoEmpresa, _
                                           "", _
                                           "", _
                                           enumStatusMensagem.EnviadaBUS, _
                                           0, _
                                           "", _
                                           lngCodigoLocalLiquidacao, _
                                           "", _
                                           "", _
                                           1, _
                                           "", _
                                           "", _
                                           0)
                
                
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & strMensagemEnvio, enumIdentificadorFila.BUS
        Set objProcessoOperacao = Nothing
    
    End If
        
    strNumeroControleIF = vbNullString
        
    If lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        'Envia Mensagem de consulta LTR0009
        
        pstrTipoMensagem = "LTR0009"
        
        udtProtocolo.CodigoEmpresa = Format(lngCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString("NZ", " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(pstrTipoMensagem, " ", 9, False)
        LSet udtProtocoloAux = udtProtocolo
        
        strHeaderMensageNZ = fgMontaHeaderMensageNZ(pstrTipoMensagem, lngCodigoEmpresa, strNumeroControleIF)
        
        If xmlMensagem.selectSingleNode("//CO_MESG") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_MESG", fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
        Else
            xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = fgCompletaString(Trim(pstrTipoMensagem), "  ", 9, False)
        End If
        
        'Pikachu 29/08/2005
        'RATS 396 - OCor 03
        xmlMensagem.documentElement.selectSingleNode("//CO_MESG").Text = Trim$(xmlMensagem.documentElement.selectSingleNode("//CO_MESG").Text)
        
        If xmlMensagem.documentElement.selectSingleNode("TX_HEAD_NZ") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "TX_HEAD_NZ", strHeaderMensageNZ
        Else
            xmlMensagem.documentElement.selectSingleNode("TX_HEAD_NZ").Text = strHeaderMensageNZ
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("NU_CTRL_IF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "NU_CTRL_IF", strNumeroControleIF
        Else
            xmlMensagem.documentElement.selectSingleNode("NU_CTRL_IF").Text = strNumeroControleIF
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "NIVE_PREF", ""
        Else
            xmlMensagem.documentElement.selectSingleNode("NIVE_PREF").Text = ""
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "DT_SIST", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
        Else
            xmlMensagem.documentElement.selectSingleNode("DT_SIST").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("CO_ISPB_CAMR") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_ISPB_CAMR", strISPBCamara
        Else
            xmlMensagem.documentElement.selectSingleNode("CO_ISPB_CAMR").Text = strISPBCamara
        End If
        
        If xmlMensagem.documentElement.selectSingleNode("CO_ISPB_IF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "CO_ISPB_IF", fgObterISPBIF(lngCodigoEmpresa)
        Else
            xmlMensagem.documentElement.selectSingleNode("CO_ISPB_IF").Text = fgObterISPBIF(lngCodigoEmpresa)
        End If
        
        strMensagemEnvio = xmlMensagem.xml
        
        Call objMesgRecbEnviSPB.IncluirBMA(strMensagemEnvio, _
                                           Trim(strNumeroControleIF), _
                                           0, _
                                           0, _
                                           pstrTipoMensagem, _
                                           0, _
                                           lngCodigoEmpresa, _
                                           "", _
                                           "", _
                                           enumStatusMensagem.EnviadaBUS, _
                                           0, _
                                           "", _
                                           lngCodigoLocalLiquidacao, _
                                           "", _
                                           "", _
                                           1, _
                                           "", _
                                           "", _
                                           0)
        
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & strMensagemEnvio, enumIdentificadorFila.BUS
        Set objProcessoOperacao = Nothing
    
    End If
    
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objMesgRecbEnviSPB = Nothing
    Set objProcessoOperacao = Nothing
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemLDL0028 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Reprocessamento das mensagens Incosistentes
'------------------------------------------------------------------------------------------

Public Function ReprocessarMensagemInconcistente(ByVal pstrNumCtrlIF As String, _
                                                 ByVal pstrDataRegistroMensagem As String, _
                                                 ByVal plngNumeroControleRepeticao As Long, _
                                                 ByVal pstrDataUltimaAtualizacao As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objRS                                   As ADODB.Recordset
Dim strSQL                                  As String
Dim strNumeroControleIF                     As String
Dim strDataRegMensagem                      As String
Dim lngNumeroControleRepeticao              As Long
Dim lngCoTextXML                            As Long
Dim lngCodigoEmpresa                        As Long
Dim lngLocalLiquidacao                      As Long

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_CTRL_IF,              " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,       " & vbCrLf & _
             "          CO_TEXT_XML,             " & vbCrLf & _
             "          DH_REGT_MESG_SPB,        " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,       " & vbCrLf & _
             "          CO_MESG_SPB,             " & vbCrLf & _
             "          CO_EMPR,                 " & vbCrLf & _
             "          CO_LOCA_LIQU             " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB " & vbCrLf & _
             " WHERE    CO_ULTI_SITU_PROC  =     " & enumStatusMensagem.MensagemInconsistente & vbCrLf & _
             "   AND    DH_REGT_MESG_SPB   =     " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbCrLf & _
             "   AND    NU_SEQU_CNTR_REPE  =     " & plngNumeroControleRepeticao & vbCrLf & _
             "   AND    DH_ULTI_ATLZ       =     " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao) & vbCrLf & _
             " ORDER BY NU_SEQU_CNTR_REPE "
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        objRS.Close
        'Registro foi alterado ou excluído por outro usuário.
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    While Not objRS.EOF
            
        strNumeroControleIF = objRS!NU_CTRL_IF
        strDataRegMensagem = fgDtHr_To_Xml(objRS!DH_REGT_MESG_SPB)
        lngNumeroControleRepeticao = objRS!NU_SEQU_CNTR_REPE
        lngCoTextXML = objRS!CO_TEXT_XML
        lngCodigoEmpresa = objRS!CO_EMPR
        lngLocalLiquidacao = objRS!CO_LOCA_LIQU
            
        xmlMensagem.loadXML fgSelectVarchar4000(lngCoTextXML)
            
        Select Case Trim(objRS!CO_MESG_SPB)
            Case "BMA0015"
                
                If flObterStatusBMA0015(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'Conta Selic câmara não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3105
                    GoTo ErrorHandler
                End If
                
            Case "CTP1002"
                
                If flObterStatusCTP1002(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'Identificador Participante câmara não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3104
                    GoTo ErrorHandler
                End If
                
            Case "LTR0001", "LTR0005R2"
                
                If flObterStatusLTR0001_LTR0005R2(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
                
            Case "BMA0002"
                
                If flObterStatuBMA0002(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
                
            Case "LDL0005R2", "LDL0009R2"
                
                If flObterStatusLDL0005R2_LDL0009R2(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
                
            Case "LTR0009R1"
                
                If flObterStatusLTR0009R1_LDL0026R1(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
                
            Case "LDL0026R1"
                
                If flObterStatusLTR0009R1_LDL0026R1(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
                
            Case "LTR0007"
                
                Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                
            Case "LTR0003R2"
            
                If Not flProcessaMensagemLTR0003R2(xmlMensagem, vbNullString) Then
                    'Não foi possível reprocessar a mensagem LTR0003R2. Operação registrada com número de controle correspondente não foi encontrada pelo sistema.
                    lngCodigoErroNegocio = 4278
                    GoTo ErrorHandler
                End If
            
            Case "LDL0001"
                
                If flObterStatusLDL0001(xmlMensagem, lngCodigoEmpresa) = enumStatusMensagem.AConciliar Then
                    If lngLocalLiquidacao <> enumLocalLiquidacao.BMC Then
                        Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                    Else
                        If flObterStatusLDL0001BMC0101(xmlMensagem, strNumeroControleIF) = enumStatusMensagem.AConciliar Then
                            Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                        Else
                            'Existe mensagem de pagamento anterior a esta, com status diferente de A Conciliar e Liquidada. Reprocessamento não permitido.
                            lngCodigoErroNegocio = 4282
                            GoTo ErrorHandler
                        End If
                    End If
                Else
                    'CNPJ não cadastrado para Veiculo Legal
                    lngCodigoErroNegocio = 3103
                    GoTo ErrorHandler
                End If
            
            Case "BMA0003"
                If flObterStatusBMA0003(xmlMensagem) = enumStatusMensagem.Informação Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'Operação original não encontrada ou não está registrada.
                    lngCodigoErroNegocio = 3117
                    GoTo ErrorHandler
                End If
            
            Case "BMC0010"
                Call fgAppendNode(xmlMensagem, "SISMSG", "NU_CTRL_IF", strNumeroControleIF)
                Call flProcessaMensagemBMC0010(xmlMensagem)
            
            Case "BMC0013"
                Call fgAppendNode(xmlMensagem, "SISMSG", "NU_CTRL_IF", strNumeroControleIF)
                Call flProcessaMensagemBMC0013(xmlMensagem)
            
            Case "BMC0101"
                If flObterStatusLDL0001BMC0101(xmlMensagem, strNumeroControleIF) = enumStatusMensagem.AConciliar Then
                    Call flReprocessarMensagem(strNumeroControleIF, pstrDataRegistroMensagem, plngNumeroControleRepeticao, pstrDataUltimaAtualizacao, Trim(objRS!CO_MESG_SPB))
                Else
                    'Existe mensagem de pagamento anterior a esta, com status diferente de A Conciliar e Liquidada. Reprocessamento não permitido.
                    lngCodigoErroNegocio = 4282
                    GoTo ErrorHandler
                End If
        
        End Select
        objRS.MoveNext
    Wend
        
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ReprocessarMensagemInconcistente Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Reprocessamento das mensagens Incosistentes
'------------------------------------------------------------------------------------------

Private Function flReprocessarMensagem(ByVal pstrNumCtrlIF As String, _
                                       ByVal pstrDataRegistroMensagem As String, _
                                       ByVal plngNumeroControleRepeticao As Long, _
                                       ByVal pstrDataUltimaAtualizacao As String, _
                                       ByVal pstrCodigoMensagem As String) As String
                                       
Dim xmlMensagem                             As MSXML2.DOMDocument40

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objRS                                   As ADODB.Recordset
Dim strSQL                                  As String
Dim strNumeroControleIF                     As String
Dim strDataRegMensagem                      As String
Dim lngNumeroControleRepeticao              As Long
Dim lngCoTextXML                            As Long
Dim lngCodigoEmpresa                        As Long
Dim vntCNPJContraparte                      As Variant
Dim strCtCstdiaAux                          As String
Dim strCtCstdia                             As String
Dim vntCNPJ_IDENT_PART                      As Variant
Dim strCodigoMensagem                       As String

Dim vntIDENT_PART_CAMR                      As Variant
Dim strNumOpLTR                             As String
Dim blnMensagemDespesa                      As Boolean

Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim lngTipoBackOffice                       As Long
Dim lngIndicadorDebtCred                    As Long
Dim strNomeVeicLega                         As String
Dim strNomeTagRepetConta                    As String
Dim lngStatusMensagem                       As Long
Dim blnProcessar                            As Boolean
Dim blnBMA0015                              As Boolean

    On Error GoTo ErrorHandler
    
    Select Case pstrCodigoMensagem
        Case "LDL0005R2", "LDL0001"
            Call flReprocessarMensagemComRepeticao(pstrNumCtrlIF, _
                                                   pstrDataRegistroMensagem, _
                                                   pstrCodigoMensagem)
            Exit Function
        Case "BMA0003"
            strSQL = " SELECT   A.NU_CTRL_IF,              " & vbCrLf & _
                     "          A.NU_SEQU_CNTR_REPE,       " & vbCrLf & _
                     "          A.CO_TEXT_XML,             " & vbCrLf & _
                     "          A.DH_REGT_MESG_SPB,        " & vbCrLf & _
                     "          A.CO_ULTI_SITU_PROC,       " & vbCrLf & _
                     "          A.CO_MESG_SPB,             " & vbCrLf & _
                     "          A.CO_EMPR,                 " & vbCrLf & _
                     "          A.NU_COMD_OPER             " & vbCrLf & _
                     "  FROM    A8.TB_MESG_RECB_ENVI_SPB A  " & vbCrLf & _
                     " WHERE    A.DH_REGT_MESG_SPB      =  " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbCrLf & _
                     "  AND     A.NU_SEQU_CNTR_REPE     =  " & plngNumeroControleRepeticao & vbCrLf & _
                     "  AND     A.DH_ULTI_ATLZ          =  " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao) & vbCrLf
    
        Case Else
            strSQL = " SELECT   A.NU_CTRL_IF,              " & vbCrLf & _
                     "          A.NU_SEQU_CNTR_REPE,       " & vbCrLf & _
                     "          A.CO_TEXT_XML,             " & vbCrLf & _
                     "          A.DH_REGT_MESG_SPB,        " & vbCrLf & _
                     "          A.CO_ULTI_SITU_PROC,       " & vbCrLf & _
                     "          A.CO_MESG_SPB,             " & vbCrLf & _
                     "          A.CO_EMPR,                 " & vbCrLf & _
                     "          A.NU_COMD_OPER,            " & vbCrLf & _
                     "          B.IN_OPER_DEBT_CRED        " & vbCrLf & _
                     "  FROM    A8.TB_MESG_RECB_ENVI_SPB A, " & vbCrLf & _
                     "          A8.TB_MESG_RECB_SPB_CNCL B  " & vbCrLf & _
                     " WHERE    A.NU_CTRL_IF        (+) = B.NU_CTRL_IF " & vbCrLf & _
                     "  AND     A.NU_SEQU_CNTR_REPE (+) = B.NU_SEQU_CNTR_REPE " & vbCrLf & _
                     "  AND     A.DH_REGT_MESG_SPB  (+) = B.DH_REGT_MESG_SPB  " & vbCrLf & _
                     "  AND     A.DH_REGT_MESG_SPB      =  " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbCrLf & _
                     "  AND     A.NU_SEQU_CNTR_REPE     =  " & plngNumeroControleRepeticao & vbCrLf & _
                     "  AND     A.DH_ULTI_ATLZ          =  " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao) & vbCrLf & _
                     " ORDER BY NU_SEQU_CNTR_REPE "
    
    End Select
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        objRS.Close
        Exit Function
    End If
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    While Not objRS.EOF
            
        strNumeroControleIF = objRS!NU_CTRL_IF
        strDataRegMensagem = fgDtHr_To_Xml(objRS!DH_REGT_MESG_SPB)
        lngNumeroControleRepeticao = objRS!NU_SEQU_CNTR_REPE
        lngCoTextXML = objRS!CO_TEXT_XML
        lngCodigoEmpresa = objRS!CO_EMPR
        
        If pstrCodigoMensagem <> "BMA0003" Then
            lngIndicadorDebtCred = IIf(IsNull(objRS!IN_OPER_DEBT_CRED), 0, objRS!IN_OPER_DEBT_CRED)
        End If
        
        strNumOpLTR = IIf(IsNull(objRS!NU_COMD_OPER), "", objRS!NU_COMD_OPER)
            
        xmlMensagem.loadXML fgSelectVarchar4000(lngCoTextXML)
        
        blnProcessar = True
        lngStatusMensagem = enumStatusMensagem.AConciliar
        strCodigoMensagem = xmlMensagem.selectSingleNode("//CodMsg").Text
        
        If lngNumeroControleRepeticao = 1 Then
            
            Select Case Trim(objRS!CO_MESG_SPB)
                Case "BMA0015"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
                    lngStatusMensagem = enumStatusMensagem.Aviso
                    strCodigoMensagem = "BMA00151"
                Case "BMA0002"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
                Case "CTP1002"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
                Case "LTR0001", "LTR0005R2"
                    
                    If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
                        vntIDENT_PART_CAMR = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
                    End If
                    
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                    
                Case "LDL0009R2", "LDL0001", "BMC0101"
                    vntCNPJ_IDENT_PART = fgObterCNPJEmpresa(lngCodigoEmpresa)
                Case "LTR0009R1"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case "LDL0026R1"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case "LTR0007"
                    
                    If lngIndicadorDebtCred = enumTipoDebitoCredito.Debito Then
                        
                        If Not xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd") Is Nothing Then
                            vntIDENT_PART_CAMR = xmlMensagem.selectSingleNode("//IdentdPartCamrDebtd").Text
                        End If
                        
                        vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text
                    
                        If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd") Is Nothing Then
                            vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd").Text
                        Else
                            vntCNPJContraparte = vbNullString
                        End If
                        
                    Else
                        
                        If xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd") Is Nothing Then
                            vntIDENT_PART_CAMR = xmlMensagem.selectSingleNode("//IdentdPartCamrCredtd").Text
                        End If
                        
                        vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdantCredtd").Text
                        
                        If Not xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd") Is Nothing Then
                            vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJNLiqdantDebtd").Text
                        Else
                            vntCNPJContraparte = vbNullString
                        End If

                    End If
                
                Case "BMA0003"
                    
                    lngStatusMensagem = enumStatusMensagem.Informação
                    blnBMA0015 = True
                    
                Case Else
                    
                    blnProcessar = False
            
            End Select
        
        Else
            Select Case Trim(objRS!CO_MESG_SPB)
                Case "BMA0002"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
                    lngStatusMensagem = enumStatusMensagem.Aviso
                Case "BMA0015"
                    
                    strCtCstdia = xmlMensagem.selectSingleNode("//CtCstdia").Text
                    blnBMA0015 = True
                    
                    lngTipoBackOffice = flObterTipoBackOfficeContaBMA0015(strCtCstdia, _
                                                                          lngCodigoEmpresa, _
                                                                          strCodigoVeiculoLegal, _
                                                                          strSiglaSistemaVeicLega)
                    
                Case "CTP1002"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
                Case "LDL0009R2"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case "LTR0009R1"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case "LDL0026R1"
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case "LDL0001"
                
                    If Trim$(strNumOpLTR) <> vbNullString Then
                        If InStr(1, strNumOpLTR, "CUSTO CETIP") > 0 Then
                            blnMensagemDespesa = True
                        Else
                            blnMensagemDespesa = False
                        End If
                    End If
                    
                    If blnMensagemDespesa Then
                        If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
                            vntIDENT_PART_CAMR = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
                        Else
                            vntIDENT_PART_CAMR = "99999999"
                        End If
                    End If
                
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case Else
                    blnProcessar = False
            End Select
        End If
        
        If blnProcessar Then
            
            If Not blnBMA0015 Then
                
                If blnMensagemDespesa Or Val(vntIDENT_PART_CAMR) <> 0 Then
                    Call fgObterCodigoVeiculoLegal(strCodigoMensagem, _
                                                   vntCNPJ_IDENT_PART, _
                                                   lngCodigoEmpresa, _
                                                   strCodigoVeiculoLegal, _
                                                   strSiglaSistemaVeicLega, _
                                                   lngTipoBackOffice, _
                                                   strNomeVeicLega, _
                                                   vntIDENT_PART_CAMR)
            
                Else
                    Call fgObterCodigoVeiculoLegal(strCodigoMensagem, _
                                                   vntCNPJ_IDENT_PART, _
                                                   lngCodigoEmpresa, _
                                                   strCodigoVeiculoLegal, _
                                                   strSiglaSistemaVeicLega, _
                                                   lngTipoBackOffice, _
                                                   strNomeVeicLega)
                
                End If
            
            End If
            
            If Trim(objRS!CO_MESG_SPB) <> "BMA0003" Then
                Call objMesgRecbEnviSPB.AlterarTipoBackOffice(strNumeroControleIF, _
                                                              lngTipoBackOffice, _
                                                              lngNumeroControleRepeticao, _
                                                              strDataRegMensagem)
                            
                Call objMesgRecbEnviSPB.AlterarVeiculoLegal(strNumeroControleIF, _
                                                            strDataRegMensagem, _
                                                            lngNumeroControleRepeticao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega)
            Else
                Call flReprocessaMensagemBMA0003(xmlMensagem)
            End If
                            
            Call objMesgRecbEnviSPB.AlterarStatus(strNumeroControleIF, _
                                                  strDataRegMensagem, _
                                                  lngStatusMensagem, _
                                                  "", _
                                                  "", _
                                                  False, _
                                                  lngNumeroControleRepeticao)
                                        
            If Trim(objRS!CO_MESG_SPB) = "LTR0007" Then
                If vntCNPJContraparte <> vbNullString Then
                    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                    Call objConciliacao.AlterarCNPJContraParte(strNumeroControleIF, _
                                                               strDataRegMensagem, _
                                                               lngNumeroControleRepeticao, _
                                                               vntCNPJContraparte)
                                                               
                    Set objConciliacao = Nothing
                End If
            
            End If
            
            Call flConciliarMensagem(strCodigoMensagem, strNumeroControleIF, lngNumeroControleRepeticao, "")
        
        End If
        
        objRS.MoveNext
    Wend
        
    objRS.Close
    
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function

ErrorHandler:
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flReprocessarMensagemBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Reprocessamento das mensagens com repetição incosistentes
'------------------------------------------------------------------------------------------

Private Function flReprocessarMensagemComRepeticao(ByVal pstrNumCtrlIF As String, _
                                                   ByVal pstrDataRegistroMensagem As String, _
                                                   ByVal pstrCodigoMensagem As String) As String
                                       
Dim xmlMensagem                             As MSXML2.DOMDocument40

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objRS                                   As ADODB.Recordset
Dim strSQL                                  As String
Dim strNumeroControleIF                     As String
Dim strDataRegMensagem                      As String
Dim lngNumeroControleRepeticao              As Long
Dim lngCoTextXML                            As Long
Dim lngCodigoEmpresa                        As Long
Dim vntCNPJ_IDENT_PART                      As Variant
Dim strCodigoMensagem                       As String
Dim vntIDENT_PART_CAMR                      As Variant
Dim strNumOpLTR                             As String
Dim blnMensagemDespesa                      As Boolean
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim lngTipoBackOffice                       As Long
Dim lngIndicadorDebtCred                    As Long
Dim strNomeVeicLega                         As String
Dim lngStatusMensagem                       As Long
Dim blnProcessar                            As Boolean

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   A.NU_CTRL_IF,              " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,       " & vbCrLf & _
             "          A.CO_TEXT_XML,             " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,        " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,       " & vbCrLf & _
             "          A.CO_MESG_SPB,             " & vbCrLf & _
             "          A.CO_EMPR,                 " & vbCrLf & _
             "          A.NU_COMD_OPER,            " & vbCrLf & _
             "          B.IN_OPER_DEBT_CRED        " & vbCrLf & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB A, " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL B  " & vbCrLf & _
             " WHERE    A.NU_CTRL_IF        (+) = B.NU_CTRL_IF " & vbCrLf & _
             " AND      A.NU_SEQU_CNTR_REPE (+) = B.NU_SEQU_CNTR_REPE " & vbCrLf & _
             " AND      A.DH_REGT_MESG_SPB  (+) = B.DH_REGT_MESG_SPB  " & vbCrLf & _
             " AND      A.DH_REGT_MESG_SPB      =  " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbCrLf & _
             " ORDER BY NU_SEQU_CNTR_REPE "
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        objRS.Close
        Exit Function
    End If
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    While Not objRS.EOF
        strNumeroControleIF = objRS!NU_CTRL_IF
        strDataRegMensagem = fgDtHr_To_Xml(objRS!DH_REGT_MESG_SPB)
        lngNumeroControleRepeticao = objRS!NU_SEQU_CNTR_REPE
        lngCoTextXML = objRS!CO_TEXT_XML
        lngCodigoEmpresa = objRS!CO_EMPR
        lngIndicadorDebtCred = IIf(IsNull(objRS!IN_OPER_DEBT_CRED), 0, objRS!IN_OPER_DEBT_CRED)
        strNumOpLTR = IIf(IsNull(objRS!NU_COMD_OPER), "", objRS!NU_COMD_OPER)
            
        xmlMensagem.loadXML fgSelectVarchar4000(lngCoTextXML)
        
        blnProcessar = True
        lngStatusMensagem = enumStatusMensagem.AConciliar
        strCodigoMensagem = xmlMensagem.selectSingleNode("//CodMsg").Text
        
        If lngNumeroControleRepeticao = 1 Then
            
            Select Case Trim(objRS!CO_MESG_SPB)
                Case "LDL0005R2", "LDL0001"
                    vntCNPJ_IDENT_PART = fgObterCNPJEmpresa(lngCodigoEmpresa)
                Case Else
                    blnProcessar = False
            End Select
        
        Else
            
            Select Case Trim(objRS!CO_MESG_SPB)
                Case "LDL0005R2", "LDL0001"
                    If Trim$(strNumOpLTR) <> vbNullString Then
                        If InStr(1, strNumOpLTR, "CUSTO CETIP") > 0 Then
                            blnMensagemDespesa = True
                        Else
                            blnMensagemDespesa = False
                        End If
                    End If
                    
                    If blnMensagemDespesa Then
                        If Not xmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
                            vntIDENT_PART_CAMR = xmlMensagem.selectSingleNode("//IdentdPartCamr").Text
                        Else
                            vntIDENT_PART_CAMR = "99999999"
                        End If
                    End If
                
                    vntCNPJ_IDENT_PART = xmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                Case Else
                    blnProcessar = False
            End Select
        
        End If
        
        If blnProcessar Then
            
            If blnMensagemDespesa Or Val(vntIDENT_PART_CAMR) <> 0 Then
                Call fgObterCodigoVeiculoLegal(strCodigoMensagem, _
                                               vntCNPJ_IDENT_PART, _
                                               lngCodigoEmpresa, _
                                               strCodigoVeiculoLegal, _
                                               strSiglaSistemaVeicLega, _
                                               lngTipoBackOffice, _
                                               strNomeVeicLega, _
                                               vntIDENT_PART_CAMR)
        
            Else
                Call fgObterCodigoVeiculoLegal(strCodigoMensagem, _
                                               vntCNPJ_IDENT_PART, _
                                               lngCodigoEmpresa, _
                                               strCodigoVeiculoLegal, _
                                               strSiglaSistemaVeicLega, _
                                               lngTipoBackOffice, _
                                               strNomeVeicLega)
            
            End If
            
            Call objMesgRecbEnviSPB.AlterarTipoBackOffice(strNumeroControleIF, _
                                                          lngTipoBackOffice, _
                                                          lngNumeroControleRepeticao, _
                                                          strDataRegMensagem)
                        
            Call objMesgRecbEnviSPB.AlterarVeiculoLegal(strNumeroControleIF, _
                                                        strDataRegMensagem, _
                                                        lngNumeroControleRepeticao, _
                                                        strCodigoVeiculoLegal, _
                                                        strSiglaSistemaVeicLega)
                            
            Call objMesgRecbEnviSPB.AlterarStatus(strNumeroControleIF, _
                                                  strDataRegMensagem, _
                                                  lngStatusMensagem, _
                                                  "", _
                                                  "", _
                                                  False, _
                                                  lngNumeroControleRepeticao)
                                        
            Call flConciliarMensagem(strCodigoMensagem, strNumeroControleIF, lngNumeroControleRepeticao, "")
        
        End If
        
        objRS.MoveNext
    Wend
        
    objRS.Close
    
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function

ErrorHandler:
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flReprocessarMensagemComRepeticao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens BMA0015 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusBMA0015(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                      ByVal plngCodigoEmpresa As Long) As Long

Dim xmlNodeTitulo                           As MSXML2.IXMLDOMNode
Dim xmlNodeConta                            As MSXML2.IXMLDOMNode

Dim strCtCstdia                             As String
Dim strCtCstdiaAux                          As String
Dim strNomeTagRepetTitulo                   As String
Dim strNomeTagRepetConta                    As String
Dim lngStatusMensagem                       As Long
Dim blnMensagemInconsistente                As Boolean

On Error GoTo ErrorHandler

    strNomeTagRepetTitulo = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1) '"Repet_BMA0015_Tit"
    strNomeTagRepetConta = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 2) '"Repet_BMA0015_Ct"

    For Each xmlNodeTitulo In pxmlMensagem.selectNodes("//" & strNomeTagRepetTitulo & "/*")
        
        For Each xmlNodeConta In xmlNodeTitulo.selectNodes(strNomeTagRepetConta & "/*")
            
            strCtCstdia = xmlNodeConta.selectSingleNode("CtCstdia").Text
            strCtCstdiaAux = fgCompletaString(xmlNodeConta.selectSingleNode("CtCstdia").Text, "0", 9, True)
            strCtCstdiaAux = Left$(strCtCstdiaAux, 4)
            
            'If fgObterCodigoVeiculoLegal(pxmlMensagem.selectSingleNode("//CodMsg").Text, _
                                         strCtCstdiaAux, _
                                         plngCodigoEmpresa, _
                                         "", _
                                         "", _
                                         0, _
                                         "") Then
                
            If flObterTipoBackOfficeContaBMA0015(strCtCstdia, plngCodigoEmpresa, "", "") <> 0 Then
                lngStatusMensagem = enumStatusMensagem.AConciliar
                
            Else
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                Exit For
            End If

        Next
        
        If lngStatusMensagem = enumStatusMensagem.MensagemInconsistente Then Exit For
        
    Next
    
    flObterStatusBMA0015 = lngStatusMensagem

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens CTP1002 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusCTP1002(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                      ByVal plngCodigoEmpresa As Long) As Long

Dim vntIdentPartCamr                        As Variant

On Error GoTo ErrorHandler

    vntIdentPartCamr = fgSelectSingleNode(pxmlMensagem, "//IdentdPartCamr").Text

    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(pxmlMensagem, "//CodMsg").Text, _
                                 vntIdentPartCamr, _
                                 plngCodigoEmpresa, _
                                 "", _
                                 "", _
                                 0) Then
        
        flObterStatusCTP1002 = enumStatusMensagem.AConciliar
        
    Else
        
        flObterStatusCTP1002 = enumStatusMensagem.MensagemInconsistente
        
    End If

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusCTP1002 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LTR0001 e LTR0005R2 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLTR0001_LTR0005R2(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                ByVal plngCodigoEmpresa As Long) As Long

Dim vntCNPJNLiqdant                        As Variant
Dim vntIDENT_PART_CAMR                     As Variant

    On Error GoTo ErrorHandler

    If Not pxmlMensagem.selectSingleNode("//IdentdPartCamr") Is Nothing Then
        vntIDENT_PART_CAMR = pxmlMensagem.selectSingleNode("//IdentdPartCamr").Text
    End If

    vntCNPJNLiqdant = pxmlMensagem.selectSingleNode("//CNPJNLiqdant").Text
                           
    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(pxmlMensagem, "//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 plngCodigoEmpresa, _
                                 "", _
                                 "", _
                                 0, _
                                 "", _
                                 vntIDENT_PART_CAMR) Then
        
        flObterStatusLTR0001_LTR0005R2 = enumStatusMensagem.AConciliar
    Else
        flObterStatusLTR0001_LTR0005R2 = enumStatusMensagem.MensagemInconsistente
    End If
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLTR0001_LTR0005R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens BMA0002 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatuBMA0002(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                     ByVal plngCodigoEmpresa As Long) As Long

Dim vntCNPJNLiqdant                         As Variant

On Error GoTo ErrorHandler

    vntCNPJNLiqdant = pxmlMensagem.selectSingleNode("//CNPJEntRespons").Text
                           
    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(pxmlMensagem, "//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 plngCodigoEmpresa, _
                                 "", _
                                 "", _
                                 0) Then
        
        flObterStatuBMA0002 = enumStatusMensagem.AConciliar
        
    Else
        
        flObterStatuBMA0002 = enumStatusMensagem.MensagemInconsistente
        
    End If

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatuBMA0002 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Montar mensagem apenas com o node de Tutulo x Conta para visualização
'------------------------------------------------------------------------------------------

Private Function flMontaMensagemNodeBMA0015(ByVal pstrCodigoMensagem As String, _
                                            ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                            ByRef pxmlNodeTitulo As MSXML2.IXMLDOMNode, _
                                            ByRef pxmlNodeConta As MSXML2.IXMLDOMNode, _
                                            ByVal pstrNomeTagReptTitulo As String, _
                                            ByVal pstrNomeTagReptConta As String) As String


Dim xmlNodeRepeticao                        As MSXML2.DOMDocument40
Dim xmlNodeTitulo                           As MSXML2.DOMDocument40
Dim xmlNodeConta                            As MSXML2.DOMDocument40
Dim strNomeTagPrincipal                     As String

On Error GoTo ErrorHandler
    
    Set xmlNodeRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlNodeTitulo = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlNodeConta = CreateObject("MSXML2.DOMDocument.4.0")
    
    strNomeTagPrincipal = pxmlMensagem.documentElement.childNodes(0).nodeName
    
    Call fgAppendNode(xmlNodeRepeticao, "", "SISMSG", "")
    Call fgAppendNode(xmlNodeRepeticao, "SISMSG", strNomeTagPrincipal, "")
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodMsg", pstrCodigoMensagem)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "NumCtrlBMA", pxmlMensagem.selectSingleNode("//NumCtrlBMA").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtLiquid", pxmlMensagem.selectSingleNode("//DtLiquid").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "ISPBIF", pxmlMensagem.selectSingleNode("//ISPBIF").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CNPJEntRespons", pxmlMensagem.selectSingleNode("//CNPJEntRespons").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "TpTitlar", pxmlMensagem.selectSingleNode("//TpTitlar").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CodTitlar", pxmlMensagem.selectSingleNode("//CodTitlar").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "CtCstdiaBMA", pxmlMensagem.selectSingleNode("//CtCstdiaBMA").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "FinlddBMA", pxmlMensagem.selectSingleNode("//FinlddBMA").Text)
    Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtHrSit", pxmlMensagem.selectSingleNode("//DtHrSit").Text)
    
    Call fgAppendNode(xmlNodeTitulo, "", pstrNomeTagReptTitulo, "")
    Call fgAppendNode(xmlNodeTitulo, pstrNomeTagReptTitulo, pxmlNodeTitulo.nodeName, "")
    Call fgAppendNode(xmlNodeTitulo, pxmlNodeTitulo.nodeName, "IdentdTitSEL", pxmlNodeTitulo.selectSingleNode("IdentdTitSEL").Text)
    Call fgAppendNode(xmlNodeTitulo, pxmlNodeTitulo.nodeName, "DtVenc", pxmlNodeTitulo.selectSingleNode("DtVenc").Text)
        
    Call fgAppendNode(xmlNodeConta, "", pstrNomeTagReptConta, "")
    Call fgAppendXML(xmlNodeConta, pstrNomeTagReptConta, pxmlNodeConta.xml)
        
    Call fgAppendXML(xmlNodeTitulo, pxmlNodeTitulo.nodeName, xmlNodeConta.xml)
    
    Call fgAppendXML(xmlNodeRepeticao, strNomeTagPrincipal, xmlNodeTitulo.xml)
    
    If Not pxmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        Call fgAppendNode(xmlNodeRepeticao, strNomeTagPrincipal, "DtMovto", pxmlMensagem.selectSingleNode("//DtMovto").Text)
    End If
    
    flMontaMensagemNodeBMA0015 = xmlNodeRepeticao.xml
    
    Set xmlNodeRepeticao = Nothing
    Set xmlNodeTitulo = Nothing
    Set xmlNodeConta = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlNodeRepeticao = Nothing
    Set xmlNodeTitulo = Nothing
    Set xmlNodeConta = Nothing


    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaMensagemNodeBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LDL0005R2 e LDL0009R2 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLDL0005R2_LDL0009R2(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                  ByVal plngCodigoEmpresa As Long) As Long

Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim strNomeTagResultLiqd                    As String
Dim lngStatusMensagem                       As Long
Dim vntCNPJNLiqdant                         As Variant

On Error GoTo ErrorHandler

    strNomeTagResultLiqd = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1) '"Repet_LDL0005R2_ResultLiqd"
    
    For Each xmlNodeResultLiqd In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
           
        If fgObterCodigoVeiculoLegal(pxmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     plngCodigoEmpresa, _
                                     "", _
                                     "", _
                                     0) Then
            
            lngStatusMensagem = enumStatusMensagem.AConciliar
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            Exit For
        End If
    
    Next
    
    flObterStatusLDL0005R2_LDL0009R2 = IIf(lngStatusMensagem = 0, enumStatusMensagem.Aviso, lngStatusMensagem)
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0005R2_LDL0009R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LTR0009R1 e LDL0026R1 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLTR0009R1_LDL0026R1(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                                  ByVal plngCodigoEmpresa As Long) As Long

Dim strNomeTagPosicao                       As String
Dim xmlNodePosicao                          As MSXML2.IXMLDOMNode

Dim vntCNPJNLiqdant                         As Variant
Dim lngStatusMensagem                       As Long
Dim blnMensagemInconsistente                As Boolean

On Error GoTo ErrorHandler

    strNomeTagPosicao = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1) 'Repet_LTR0009R1_ConsPos , Repet_LDL0026R1_ResultLiqd
    
    For Each xmlNodePosicao In pxmlMensagem.selectNodes("//" & strNomeTagPosicao & "/*")
        
        vntCNPJNLiqdant = xmlNodePosicao.selectSingleNode("CNPJNLiqdant").Text
           
        If fgObterCodigoVeiculoLegal(pxmlMensagem.selectSingleNode("//CodMsg").Text, _
                                     vntCNPJNLiqdant, _
                                     plngCodigoEmpresa, _
                                     "", _
                                     "", _
                                     0, _
                                     "") Then
            
            lngStatusMensagem = enumStatusMensagem.AConciliar
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            Exit For
        End If
    Next
    
    flObterStatusLTR0009R1_LDL0026R1 = lngStatusMensagem

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLTR0009R1_LDL0026R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LDL0001 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLDL0001(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                      ByVal plngCodigoEmpresa As Long) As Long

Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim strNomeTagResultLiqd                    As String
Dim lngStatusMensagem                       As Long
Dim vntCNPJNLiqdant                         As Variant

    On Error GoTo ErrorHandler

    strNomeTagResultLiqd = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1)
    
    If Not pxmlMensagem.selectSingleNode("//" & strNomeTagResultLiqd) Is Nothing Then
    
        For Each xmlNodeResultLiqd In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
            
            vntCNPJNLiqdant = xmlNodeResultLiqd.selectSingleNode("CNPJNLiqdant").Text
               
            If fgObterCodigoVeiculoLegal(pxmlMensagem.selectSingleNode("//CodMsg").Text, _
                                         vntCNPJNLiqdant, _
                                         plngCodigoEmpresa, _
                                         "", _
                                         "", _
                                         0) Then
                
                lngStatusMensagem = enumStatusMensagem.AConciliar
            Else
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                Exit For
            End If
        
        Next
    Else
        lngStatusMensagem = enumStatusMensagem.Aviso
    End If
    
    flObterStatusLDL0001 = lngStatusMensagem
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LDL0001 BMC e BMC0101 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLDL0001BMC0101(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                             ByVal pstrNumeroControleIF As String) As Long

Dim lngStatusMensagem                       As Long
Dim lngStatusMensagemExistente              As Long
Dim xmlMensagemExistenteBMC                 As MSXML2.DOMDocument40
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

    On Error GoTo ErrorHandler

    lngStatusMensagem = enumStatusMensagem.AConciliar
    
    If pxmlMensagem.selectSingleNode("//TpInf").Text = "D" Then
       
        Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
        Set xmlMensagemExistenteBMC = objMesgRecbEnviSPB.LerXMLMensagem(, , , _
                                                                        pxmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                                        , , , _
                                                                        pxmlMensagem.selectSingleNode("//DtMovto").Text, _
                                                                        pxmlMensagem.selectSingleNode("//TpInf").Text, _
                                                                        enumLocalLiquidacao.BMC, _
                                                                        pstrNumeroControleIF)
        
        If xmlMensagemExistenteBMC.xml <> vbNullString Then
            lngStatusMensagemExistente = Val(xmlMensagemExistenteBMC.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
            
            If lngStatusMensagemExistente <> enumStatusMensagem.MensagemLiquidada And _
               lngStatusMensagemExistente <> enumStatusMensagem.AConciliar Then
                
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            
            Else
            
                If lngStatusMensagemExistente = enumStatusMensagem.AConciliar Then
                    With xmlMensagemExistenteBMC
                        Call objMesgRecbEnviSPB.AlterarStatus(.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              .selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              enumStatusMensagem.MensagemCancelada, _
                                                              vbNullString, , _
                                                              False, _
                                                              0, _
                                                              .selectSingleNode("//CO_MESG").Text)
                    End With
                End If
                
            End If
        End If
        
    End If

    Set xmlMensagemExistenteBMC = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    flObterStatusLDL0001BMC0101 = lngStatusMensagem
    
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0001BMC0101 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Verificação e processamento do fluxo automático de liquidação de operações.
'------------------------------------------------------------------------------------------

Private Sub flConciliarOperacao(ByRef pxmlOperacao As MSXML2.DOMDocument40)

Dim objConciliacao                          As A8LQS.clsConciliacao

On Error GoTo ErrorHandler
   
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Call objConciliacao.GerenciarConciliacaoAutomaticaCETIP(pxmlOperacao)
    Set objConciliacao = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set objConciliacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'------------------------------------------------------------------------------------------
'Processa mensagem BMA0014 (BMA informa Conclusão de especificação da operação ao Participante)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0014(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                           ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

Dim vntCNPJNEntidadeResponsavel             As Variant
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim lngTipoBackOffice                       As Long
Dim strNomeVeicLega                         As String
Dim lngCodigoEmpresa                        As Long

On Error GoTo ErrorHandler
    
    Call ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInformacao)

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa

    If Not xmlMensagem.selectSingleNode("//CNPJEntRespons") Is Nothing Then
        vntCNPJNEntidadeResponsavel = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
    Else
        vntCNPJNEntidadeResponsavel = vbNullString
    End If
    
    Call fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                   vntCNPJNEntidadeResponsavel, _
                                   lngCodigoEmpresa, _
                                   strCodigoVeiculoLegal, _
                                   strSiglaSistema, _
                                   lngTipoBackOffice, _
                                   strNomeVeicLega)
        
    fgAppendNode xmlMensagem, "SISMSG", "CO_VEIC_LEGA", strCodigoVeiculoLegal
    
    flGeraAlertaCamara xmlMensagem.selectSingleNode("//CodMsg").Text, _
                       lngTipoBackOffice, _
                       strNomeVeicLega, _
                       0, _
                       xmlMensagem

    fgAppendNode xmlMensagem, "SISMSG", "DH_BACEN", xmlMensagem.selectSingleNode("//DtHrBMA").Text
    fgAppendNode xmlMensagem, "SISMSG", "DT_MESG", xmlMensagem.selectSingleNode("//DtMovto").Text

    udtProtocolo.CodigoEmpresa = Format(lngCodigoEmpresa, "00000")
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.SiglaSistemaDestino = fgCompletaString(Trim(strSiglaSistema), " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString(Trim("46"), "0", 9, True)
    
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "TP_MESG", fgCompletaString(Trim("44"), "0", 9, True))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "SG_SIST_ORIG", fgCompletaString("A8", " ", 3, False))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "SG_SIST_DEST", fgCompletaString(Trim(strSiglaSistema), " ", 3, False))
    Call fgAppendNode(xmlMensagem, xmlMensagem.documentElement.nodeName, "CO_EMPR", Format$(lngCodigoEmpresa, "00000"))
    
    LSet udtProtocoloAux = udtProtocolo
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & xmlMensagem.xml, enumIdentificadorFila.BUS
    Set objProcessoOperacao = Nothing
    
    
    Exit Function
ErrorHandler:
    
    Set objProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0014 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processa mensagem BMA0013 (BMA informa Resultado líquido de operações compromissadas genéricas)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0013(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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

Dim strNomeTagRepetConta                    As String
Dim xmlNodeConta                            As MSXML2.IXMLDOMNode

Dim strNumOpSEL                             As String
Dim strDtLiquid                             As String
Dim strCtCstdia                             As String
Dim strCtCstdiaAux                          As String
Dim intTpDeb_Cred                           As Integer
Dim vntVlrFinanc                            As Variant
Dim vntCNPJContraparte                      As Variant
Dim vntCodTitularCutd                       As Variant
Dim lngTipoNegoBMA                          As Long
Dim strNomeVeicLega                         As String
Dim strDataRetornoOperacao                  As String
Dim vntPrazoDiasRetornoOperacao             As Variant
Dim vntValorRetornoOperacao                 As Variant

On Error GoTo ErrorHandler

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strMensagem = xmlMensagem.xml
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
    lngTipoBackOffice = 0
    strSituacaoMensagemSPB = ""
    strDataRegistroMensagem = ""
    vntCodTitularCutd = 0
   
    strDataOperacao = fgSelectSingleNode(xmlMensagem, "//DtMovto").Text
    
    If Not xmlMensagem.selectSingleNode("//CNPJEntRespons") Is Nothing Then
        vntCNPJContraparte = xmlMensagem.selectSingleNode("//CNPJEntRespons").Text
    End If
    
    If Not xmlMensagem.selectSingleNode("//NumCtrlBMA") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.selectSingleNode("//NumCtrlBMA").Text
    End If
                
    If fgObterCodigoVeiculoLegal(fgSelectSingleNode(xmlMensagem, "//CodMsg").Text, _
                                 vntCNPJContraparte, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
        
        lngStatusMensagem = enumStatusMensagem.AConciliar
        
        Call flGeraAlertaCamara(strCodigoMensagemSPB, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntVlrFinanc, _
                                xmlMensagem)
        
    Else
        lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
    End If

    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    'Incluir Mensagem Pai
    lngSeqControleRepeticao = 1

    strNomeTagRepetConta = ObterNomeTagRepeticao(strCodigoMensagemSPB, 1) ' "Repet_BMA0013_Op"
            
    strMensagem = xmlMensagem.xml
    
    'Repeticao Titulos
    For Each xmlNodeConta In xmlMensagem.selectNodes("//" & strNomeTagRepetConta & "/*")
                
        strMensagem = flMontaMensagemNode(strCodigoMensagemSPB, xmlMensagem, xmlNodeConta)
        
        If Not xmlNodeConta.selectSingleNode("NumCtrlBMAEspecc") Is Nothing Then
            strNumeroControleCamara = xmlNodeConta.selectSingleNode("NumCtrlBMAEspecc").Text
        End If
        
        If Not xmlNodeConta.selectSingleNode("DtLiquid") Is Nothing Then
            strDtLiquid = xmlNodeConta.selectSingleNode("DtLiquid").Text
        End If
        
        If Not xmlNodeConta.selectSingleNode("VlrFinanc") Is Nothing Then
            vntVlrFinanc = xmlNodeConta.selectSingleNode("VlrFinanc").Text
        End If
        
        If Not xmlNodeConta.selectSingleNode("NumOpNegcBMA") Is Nothing Then
            strNumOpSEL = xmlNodeConta.selectSingleNode("NumOpNegcBMA").Text
        End If
            
        If Not xmlNodeConta.selectSingleNode("VlrFinancRet") Is Nothing Then
            vntValorRetornoOperacao = xmlNodeConta.selectSingleNode("VlrFinancRet").Text
        Else
            vntValorRetornoOperacao = 0
        End If
        
        If Not xmlNodeConta.selectSingleNode("PrzRetOp") Is Nothing Then
            vntPrazoDiasRetornoOperacao = xmlNodeConta.selectSingleNode("PrzRetOp").Text
        Else
            vntPrazoDiasRetornoOperacao = 0
        End If
        
        If Not xmlNodeConta.selectSingleNode("DtOpRet") Is Nothing Then
            strDataRetornoOperacao = xmlNodeConta.selectSingleNode("DtOpRet").Text
        Else
            strDataRetornoOperacao = ""
        End If
        
        If Not xmlNodeConta.selectSingleNode("CodTitlarCstdte") Is Nothing Then
            vntCodTitularCutd = xmlNodeConta.selectSingleNode("CodTitlarCstdte").Text
        Else
            vntCodTitularCutd = ""
        End If
        
        Select Case UCase(xmlNodeConta.selectSingleNode("NatuOp").Text)
            Case NatuOp_Compra, NatuOp_Recompra
                intTpDeb_Cred = enumTipoDebitoCredito.Credito
            Case NatuOp_Venda, NatuOp_Revenda
                intTpDeb_Cred = enumTipoDebitoCredito.Debito
        End Select
        
        If Not xmlNodeConta.selectSingleNode("CtCstdia") Is Nothing Then
            strCtCstdia = xmlNodeConta.selectSingleNode("CtCstdia").Text
        Else
            strCtCstdia = ""
        End If
        
        strCtCstdiaAux = fgCompletaString(xmlNodeConta.selectSingleNode("CtCstdia").Text, "0", 9, True)
        strCtCstdiaAux = Left$(strCtCstdiaAux, 4)
        
        strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                                Trim(strNumeroControleIF), _
                                                                0, _
                                                                0, _
                                                                strCodigoMensagemSPB, _
                                                                lngTipoBackOffice, _
                                                                lngCodigoEmpresa, _
                                                                strNumOpSEL, _
                                                                "", _
                                                                lngStatusMensagem, _
                                                                lngIndicadorEntradaManual, _
                                                                "", _
                                                                lngCodigoLocalLiquidacao, _
                                                                strCodigoVeiculoLegal, _
                                                                strSiglaSistemaVeicLega, _
                                                                lngSeqControleRepeticao, _
                                                                strNumeroControleCamara, _
                                                                "")
                
        'Salvar Dados Conciliacao
        Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                    strDataRegistroMensagem, _
                                    strDataOperacao, _
                                    0, _
                                    0, _
                                    "", _
                                    0, _
                                    0, _
                                    vntVlrFinanc, _
                                    lngSeqControleRepeticao, _
                                    strDtLiquid, _
                                    intTpDeb_Cred, _
                                    "", _
                                    vntCNPJContraparte, _
                                    "", _
                                    strCtCstdia, _
                                    "", _
                                    vntCodTitularCutd, _
                                    lngTipoNegoBMA, _
                                    0, _
                                    , , , , , , _
                                    strDataRetornoOperacao, vntPrazoDiasRetornoOperacao, vntValorRetornoOperacao _
                                    )
        
        
        If lngStatusMensagem = enumStatusMensagem.AConciliar Then
            'Chamada para a Rotina de Cociliacao automatica
            Call flConciliarMensagem(strCodigoMensagemSPB, _
                                     Trim(strNumeroControleIF), _
                                     lngSeqControleRepeticao, _
                                     strDataRegistroMensagem)
        End If
    Next
   
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0013 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processa as mensagens da BMA liquidadas no SELIC ( Ex. Termo)
'------------------------------------------------------------------------------------------

Public Function ProcessaMensagemSELIC_BMA(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                          ByVal pstrHeaderNZ As String) As String

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao

Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40

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
Dim strCodigoMensagemSPB                    As String
Dim strSituacaoMensagemSPB                  As String
Dim strAgendamento                          As String
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String
Dim strDataRegistroMensagem                 As String
Dim strDataOperacao                         As String
Dim strComandoOperacao                      As String

Dim blnGerenciarChamada                     As Boolean
Dim blnAtualizaStatusOper                   As Boolean

Dim strSituacaoRecebida                     As String

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
    
    'Obter a Operação
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

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
        '3038 - Operação já liquidada.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA Then
        
        Set objMensagem = Nothing
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
        Set objMesgRecbEnviSPB = Nothing
        Set xmlMesgRecbEnviSPB = Nothing
        
        Exit Function
    End If
    
    'Salvar Mensagem Enviada Recebida
    'Obter o mensagem de requisicao
    vntSequenciaOperacao = xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    strCodigoVeiculoLegal = xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
    lngCodigoLocalLiquidacao = xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
    lngTipoBackOffice = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//TP_BKOF").Text
    
    If Not xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        strComandoOperacao = xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text
    End If
        
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
    
    lngStatusMensagem = enumStatusMensagem.R1
        
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(strMensagem, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            0, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strComandoOperacao, _
                                                            strSituacaoMensagemSPB, _
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao)
    
        
    'Atualiza Status da Mensagem Original -------------------------------------------------------------------
    If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
        lngStatusMensagem = xmlStatus.selectSingleNode("//Status_Mensagem").Text
    Else
        lngStatusMensagem = enumStatusMensagem.Respondida
    End If
    
    objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                     lngStatusMensagem, _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                     , _
                                     , _
                                     xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text

    strSituacaoRecebida = vbNullString
    If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
        strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
    End If
    '-------------------------------------------------------------------
            
    blnGerenciarChamada = True
    
    Select Case xmlMensagem.selectSingleNode("//CodMsg").Text
        
        Case "SEL1007R1"
            Select Case xmlMensagem.selectSingleNode("//SitOpSEL").Text
                Case "ATU"
                    Call flGerenciarEnvioMensagem(xmlMensagem, strNumeroControleIF, LiberarLiquidacaoLeilaoBMA)
                    blnAtualizaStatusOper = False
                    blnGerenciarChamada = False
                    xmlStatus.selectSingleNode("//Status_Operacao").Text = 0
                Case Else
                    xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada
                    blnAtualizaStatusOper = True
            End Select
        
        Case "SEL1023R1"
            Select Case xmlMensagem.selectSingleNode("//SitOpSEL").Text
                Case "PEO"
                    xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Registrada
                    blnAtualizaStatusOper = True
                Case "CAN", "EXP", "RLF", "RST"
                    xmlStatus.selectSingleNode("//Status_Operacao").Text = enumStatusOperacao.Rejeitada
                    blnAtualizaStatusOper = True
            End Select
    End Select
                
    If blnGerenciarChamada Then
        If Val(xmlStatus.selectSingleNode("//Status_Operacao").Text) <> enumStatusOperacao.Rejeitada Then
            strSituacaoRecebida = vbNullString
        End If
        
        Call GerenciarChamada(xmlOperacao, _
                                xmlMensagem, _
                                xmlStatus.selectSingleNode("//Status_Operacao").Text, _
                                0, _
                                "", _
                                enumIndicadorSimNao.Nao, _
                                "", _
                                False, _
                                blnAtualizaStatusOper, _
                                strSituacaoRecebida)
    End If
            
    Set objMesgRecbEnviSPB = Nothing
    Set objMensagem = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set xmlOperacao = Nothing
   

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "fgProcessaMensagemSELIC_BMA Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processa Mensagem BMA0005R1 (Resposta ao Requisitante de Especificação de Lastro)
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMA0005R1(ByVal pstrNumeroComando As String, _
                                             ByVal plngStatusOperacao As Long) As Boolean

Dim objRS                                   As ADODB.Recordset
Dim strSQL                                  As String

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40

Dim objMesgRecebEnviSPB                     As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMesgRecebEnviSPB                     As MSXML2.DOMDocument40

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim lngStatusOperacao                       As Long
Dim vntNumeroSequOperAtiv                   As Variant

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_SEQU_OPER_ATIV " & _
             "   FROM   A8.TB_OPER_ATIV " & _
             "  WHERE   NU_COMD_OPER = '" & pstrNumeroComando & "'" & _
             "    AND   TP_OPER      =  " & enumTipoOperacaoLQS.CompromissadaGenericaAVista
             
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        objRS.Close
        Exit Function
    End If
    
    vntNumeroSequOperAtiv = objRS.fields("NU_SEQU_OPER_ATIV")
        
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Select Case plngStatusOperacao
        Case enumStatusOperacao.Liquidada
            lngStatusOperacao = enumStatusOperacao.Registrada
        Case enumStatusOperacao.Rejeitada
            lngStatusOperacao = enumStatusOperacao.AConciliarBMA0013
    End Select
    
    Call objOperacao.AlterarStatusOperIndisp(vntNumeroSequOperAtiv, _
                                             lngStatusOperacao)
    Set objOperacao = Nothing
       
    objRS.Close
    
    Exit Function
ErrorHandler:
    
    Set objOperacao = Nothing
      
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMA0005R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------------
'Processemnto de mensagens CTP0100R1 (Resposta ao Requisitante de Cancelamento ou estorno de operação)
'--------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemCTP0100R1(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                             ByVal pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objAlerta                               As A8LQS.clsAlerta
Dim objLegado                               As A8LQS.clsLegado
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlStatus                               As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim vntSequenciaOperacao                    As Variant
Dim lngStatusMensagemOriginal               As Long
Dim lngStatusOperacao                       As Long
Dim lngTipoAcao                             As Long

Dim vntNuSequOperAtiv                       As Variant
Dim lngStatusOperacaoOriginal               As Long

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler
                    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
                    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
                    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    If CLng(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
        '3038 - Operação já liquidada.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
           
    vntSequenciaOperacao = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
       
     'Obter Status ---------------------
     Set xmlStatus = CreateObject("MSXML2.DOMDocument.4.0")
     xmlStatus.loadXML flObterStatusOperacaoMensagem(pxmlMensagem)
     '----------------------------------
         
     If xmlStatus.selectSingleNode("//Status_Operacao").Text <> vbNullString Then
         lngStatusOperacao = xmlStatus.selectSingleNode("//Status_Operacao").Text
     Else
         lngStatusOperacao = 0
     End If
             
     'Atualiza Status da Mensagem de Requisição  -------------------------------------------------------------------
     If xmlStatus.selectSingleNode("//Status_Mensagem").Text <> vbNullString Then
         lngStatusMensagemOriginal = xmlStatus.selectSingleNode("//Status_Mensagem").Text
     Else
         lngStatusMensagemOriginal = enumStatusMensagem.MensagemLiquidada
     End If
     
     objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                      lngStatusMensagemOriginal, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , _
                                      True
    
    strSituacaoRecebida = vbNullString
    If Not xmlStatus.selectSingleNode("//Situacao_Recebida") Is Nothing Then
        strSituacaoRecebida = xmlStatus.selectSingleNode("//Situacao_Recebida").Text
    End If
    '----------------------------------------------------------------------------------------------------------
            
    'Atualiza Status da Operacao  -------------------------------------------------------------------
    If xmlStatus.selectSingleNode("//Status_Operacao").Text <> "" Then
            
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
        Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
        
        Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
        Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
            fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
        End If
        
        Set objTipoOperacao = Nothing
        Set xmlTipoOperacao = Nothing
                           
        Set xmlNode = xmlOperacao.selectSingleNode("//SISMSG[*/CodMsg!='CTP0100R1']")
        
        If Not xmlNode Is Nothing Then
            xmlOperacao.documentElement.removeChild xmlNode
        End If
                           
        If lngStatusOperacao = enumStatusOperacao.Registrada Then
                   
            Call GerenciarChamada(xmlOperacao, _
                                  pxmlMensagem, _
                                  enumStatusOperacao.Cancelada, _
                                  0, _
                                  "", _
                                  enumIndicadorSimNao.Sim, _
                                  "", _
                                  , _
                                  True)
                                  
            Select Case Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
            
                Case enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
                     enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO
                    
                    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
                    
                    xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text = xmlOperacao.selectSingleNode("//DT_OPER_ATIV_ORIG").Text
                    
                    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                    
                    Set objGestaoCaixa = Nothing
                    Set objProcessoOperacao = Nothing
                
                Case enumTipoOperacaoLQS.RegistroContratoSWAP
                    
                    'Pikachu - 14/10/2004
                    'Se existir Complemento de SWAP , cancelar tambem
                    
                    If flExisteComplementoSWAP(xmlOperacao, lngStatusMensagemOriginal, vntSequenciaOperacao) Then
                        
                        Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                        
                        Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
                        Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                        
                        If xmlTipoOperacao.loadXML(objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)) Then
                            fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.selectSingleNode("//TP_MESG_RETN_INTE").Text
                        End If
                        
                        Set objTipoOperacao = Nothing
                        Set xmlTipoOperacao = Nothing
                    
                        Call GerenciarChamada(xmlOperacao, _
                                              pxmlMensagem, _
                                              enumStatusOperacao.Cancelada, _
                                              0, _
                                              "", _
                                              enumIndicadorSimNao.Sim, _
                                              "", _
                                              , _
                                              True)
                    
                    End If
                
                Case enumTipoOperacaoLQS.OperacaoDefinitivaCETIP
                
                    If xmlOperacao.selectSingleNode("//CO_OPER_CETIP").Text = enumOperacaoCETIP_CTP0052.ResgateAntecipado Then

                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
                        
                        xmlOperacao.selectSingleNode("//DT_OPER_ATIV").Text = xmlOperacao.selectSingleNode("//DT_VENC_OPER_ATIV_ORIG").Text
                           
                        'Enviar Previsão para PJ com Data original
                        strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                        Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS)
                        
                        'Enviar Previsão para PJ Maoires Valores com Data original
                        Call objProcessoOperacao.EnviarRealPrevistoCETIP(xmlOperacao)
                        
                        Set objGestaoCaixa = Nothing

                        
                        'Enviar Previsão para Sub reserva com Data original
                        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                        strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
                        Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.SubReserva)
                        Set objSubReserva = Nothing
                         
                        Set objProcessoOperacao = Nothing
                   End If

            End Select
                                                          
        Else
            
            Select Case lngStatusOperacao
                
                Case enumStatusOperacao.Pendencia
                    
                    lngTipoAcao = enumTipoAcao.CancelamentoPendente
                
                Case enumStatusOperacao.Rejeitada, _
                     enumStatusOperacao.Expirada
                    
                    lngTipoAcao = enumTipoAcao.CancelamentoRejeitado
                
            End Select
            
            If Val(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.Rejeitada Then
                strSituacaoRecebida = vbNullString
            End If
            
            Call objOperacao.AlterarStatusOperIndisp(vntSequenciaOperacao, _
                                                     xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                     0, _
                                                     lngTipoAcao, _
                                                     strSituacaoRecebida)
            
            Select Case lngStatusOperacao
                
                Case enumStatusOperacao.Pendencia
                    
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Pendencia
                
                Case enumStatusOperacao.Rejeitada, _
                     enumStatusOperacao.Expirada
                    
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Rejeitada
                
            End Select
            
            
            Set objAlerta = CreateObject("A8LQS.clsAlerta")
            Call objAlerta.GerarAlerta(xmlOperacao, pxmlMensagem, "Rejeitada", , lngStatusOperacao)
            Set objAlerta = Nothing
                                                     
            Set objLegado = CreateObject("A8LQS.clsLegado")
            Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
            
            strMensagem = objLegado.EnviarRetorno(xmlOperacao, pxmlMensagem)
            Call objProcessoOperacao.EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS)
            
            Set objProcessoOperacao = Nothing
            Set objLegado = Nothing
                                                    
            If lngStatusOperacao = enumStatusOperacao.Rejeitada Or _
               lngStatusOperacao = enumStatusOperacao.Expirada Then
                                                                
                'Se erro mensagem CTP Com ordem de lançamento , alterar CTP1002 para A conciliar
                Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
            
                    Case enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao, _
                         enumTipoOperacaoLQS.RetiradaCustodiaConciliacao, _
                         enumTipoOperacaoLQS.TransferCustodiaConciliacao, _
                         enumTipoOperacaoLQS.OperacaoDefinitivaVendaOrdem, _
                         enumTipoOperacaoLQS.MovimExercDireitosDebenturesConciliacao, _
                         enumTipoOperacaoLQS.ConversaoPermutaValorMobiliarioConciliacao, _
                         enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP
            
                        'Alterar status a CTP1002 para A conciliar
                        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                        Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
                        If xmlCociliacao.loadXML(objConciliacao.LerTodos(, vntSequenciaOperacao)) Then
                        
                            For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
                                'Numero Controle IF CTP1002
                                strNumeroControleIF = xmlNode.selectSingleNode("NU_CTRL_IF").Text
                                
                                xmlMesgRecbEnviSPB.loadXML objMesgRecbEnviSPB.LerMensagemIda(strNumeroControleIF)
                                
                                Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.selectSingleNode("//NU_CTRL_IF").Text, _
                                                                      xmlMesgRecbEnviSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                                      enumStatusMensagem.AConciliar, _
                                                                      xmlMesgRecbEnviSPB.selectSingleNode("//DH_ULTI_ATLZ ").Text)
            
                            Next
                        
                        End If
                        
                        Set objConciliacao = Nothing
                        Set xmlCociliacao = Nothing
                End Select
            End If
        End If
        
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
    End If
        
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set objLegado = Nothing
    Set objSubReserva = Nothing
    Set objAlerta = Nothing
    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set xmlStatus = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCTP0100R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Obter o Numero de Controle Especificação original para as mensagens de Especifiação
'------------------------------------------------------------------------------------------

Public Function ObterNumeroControleEspecificacaoOriginal(ByVal strNumeroOperNegoBMA As String, _
                                                         ByVal pblnOperacaoEspecificacao As Boolean, _
                                                         ByVal plngCodigoEmpresa As Long, _
                                                         ByVal pstrCodigoVeiculoLegal As String)
                                                         
Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim strWhere                                As String
Dim strDHInicio                             As String
Dim strDHFim                                As String
                                                         
On Error GoTo ErrorHandler
    
    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")

    strSQL = " SELECT   NU_CTRL_MESG_SPB_ORIG " & _
             "   FROM   A8.TB_OPER_ATIV " & _
             "  WHERE   NU_COMD_OPER = '" & strNumeroOperNegoBMA & "'" & _
             "    AND   CO_EMPR      =  " & plngCodigoEmpresa & _
             "    AND   CO_VEIC_LEGA = '" & pstrCodigoVeiculoLegal & "'" & _
             "    AND   TP_OPER IN ( " & enumTipoOperacaoLQS.DefinitivaCobertaBMA & ", " & _
                                         enumTipoOperacaoLQS.DefinitivaDescobertaBMA & ", " & _
                                         enumTipoOperacaoLQS.OperacaoTermoCobertaBMA & ", " & _
                                         enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaEspecificaTermo & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaGenericaAVista & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaGenericaATermo & ", " & _
                                         enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA & ") " & _
            "  AND  NU_CTRL_MESG_SPB_ORIG IS NOT NULL " & _
            "  AND  DT_OPER_ATIV BETWEEN " & strDHInicio & _
            "  AND " & strDHFim & _
            " ORDER BY DH_ULTI_ATLZ DESC"

    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        ObterNumeroControleEspecificacaoOriginal = objRS!NU_CTRL_MESG_SPB_ORIG
    End If

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNumeroControleEspecificacaoOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Obter Tipo de backoffice para cada grupo de repetição da BMA0015

'Obter o tipo de backoffice atraves da conta Custodia

Public Function flObterTipoBackOfficeContaBMA0015(ByVal pstrContaCustodia As String, _
                                                  ByVal plngCodigoEmpresa As Long, _
                                                  ByRef pstrCodigoVeiculoLegal As String, _
                                                  ByRef pstrSiglaSistema As String) As Long

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim strRetorno                              As String

Dim strSQL                                  As String
Dim strCtCstdiaAux                          As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    
    strRetorno = objVeiculoLegal.LerTodos("", 0, "S", 0, "N", plngCodigoEmpresa, "", pstrContaCustodia)
    
    flObterTipoBackOfficeContaBMA0015 = 0
    
    'Não tem veiculo legal com esta conta , status incosnstente
    If strRetorno = vbNullString Then
        
        strCtCstdiaAux = fgCompletaString(pstrContaCustodia, "0", 9, True)
        strCtCstdiaAux = Left$(strCtCstdiaAux, 4) & "0000"
    
        strSQL = " SELECT   CO_VEIC_LEGA,   " & _
                 "          SG_SIST,        " & _
                 "          TP_BKOF,         " & _
                 "          NO_VEIC_LEGA " & _
                 "   FROM   A8.TB_VEIC_LEGA " & _
                 "  WHERE   SUBSTR(LPAD(CO_CNTA_CUTD_PADR_SELIC,9,0),1,8) =  '" & strCtCstdiaAux & "'"
            
        Set objRS = fgQuerySQL(strSQL)
        
        If Not objRS.EOF Then
            flObterTipoBackOfficeContaBMA0015 = objRS!TP_BKOF
            pstrCodigoVeiculoLegal = objRS!CO_VEIC_LEGA
            pstrSiglaSistema = objRS!SG_SIST
        End If
        
        objRS.Close
        
        Exit Function
    End If
    
    xmlVeiculoLegal.loadXML strRetorno
    
    'Tem mais de um tipo de backoffice , status incosnstente
    If Not xmlVeiculoLegal.selectSingleNode("//*[TP_BKOF='1']") Is Nothing And _
       Not xmlVeiculoLegal.selectSingleNode("//*[TP_BKOF='2']") Is Nothing Then
        Exit Function
    End If
        
    flObterTipoBackOfficeContaBMA0015 = Val(xmlVeiculoLegal.selectSingleNode("//TP_BKOF").Text)
    pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
    pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
    
'RATS 863
'====================================================================================================
'    Select Case xmlVeiculoLegal.selectSingleNode("//TP_BKOF").Text
'
'        Case enumTipoBackOffice.Tesouraria
'
'            If Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "PNA" Then
'                flObterTipoBackOfficeContaBMA0015 = enumTipoBackOffice.Tesouraria
'                pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
'                pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
'            End If
'
'        Case enumTipoBackOffice.FundosProprios
'
'            If Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "PLC" Then
'
'                flObterTipoBackOfficeContaBMA0015 = enumTipoBackOffice.FundosProprios
'                pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
'                pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
'
'            ElseIf Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "CL" Or _
'                   Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "CF" Then
'
'                If plngCodigoEmpresa = enumCodigoEmpresa.Santander Then
'                    pstrContaCustodia = "4300003"
'                Else
'                    pstrContaCustodia = "900003"
'                End If
'
'                strRetorno = objVeiculoLegal.LerTodos("", 0, "S", 0, "N", plngCodigoEmpresa, "", pstrContaCustodia)
'                xmlVeiculoLegal.loadXML strRetorno
'
'                flObterTipoBackOfficeContaBMA0015 = enumTipoBackOffice.Tesouraria
'                pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
'                pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
'
'            End If
'
'        Case enumTipoBackOffice.FundosTerceiros
'
'            If Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "PLC" Then
'
'                flObterTipoBackOfficeContaBMA0015 = enumTipoBackOffice.FundosTerceiros
'                pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
'                pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
'
'            ElseIf Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "CL" Or _
'                   Trim$(xmlVeiculoLegal.selectSingleNode("//TP_TITL_BMA").Text) = "CF" Then
'
'                If plngCodigoEmpresa = enumCodigoEmpresa.Santander Then
'                    pstrContaCustodia = "4300003"
'                Else
'                    pstrContaCustodia = "900003"
'                End If
'
'                strRetorno = objVeiculoLegal.LerTodos("", 0, "S", 0, "N", plngCodigoEmpresa, "", pstrContaCustodia)
'                xmlVeiculoLegal.loadXML strRetorno
'
'                flObterTipoBackOfficeContaBMA0015 = enumTipoBackOffice.Tesouraria
'                pstrCodigoVeiculoLegal = xmlVeiculoLegal.selectSingleNode("//CO_VEIC_LEGA").Text
'                pstrSiglaSistema = xmlVeiculoLegal.selectSingleNode("//SG_SIST").Text
'
'            End If
'
'    End Select
'====================================================================================================
    
    Set objVeiculoLegal = Nothing
    Set xmlVeiculoLegal = Nothing
    
    Exit Function

ErrorHandler:
    Set objVeiculoLegal = Nothing
    Set xmlVeiculoLegal = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterTipoBackOfficeContaBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Monta XML somente com os net de valor por CNPJ
Private Function flPreparaMensagemLDL0009R2(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                            ByVal plngLocalLiquidacao As Long) As MSXML2.IXMLDOMNode

Dim xmlNetCNPJ                              As MSXML2.DOMDocument40
Dim xmlDistinct                             As MSXML2.DOMDocument40
Dim xmlGrupo                                As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strCodigoMensagem                       As String
Dim strQueryXPath                           As String

Dim strNomeTagResultLiqd                    As String

On Error GoTo ErrorHandler
    
    Set xmlDistinct = CreateObject("MSXML2.DOMDocument.4.0")
    
    strCodigoMensagem = Trim$(pxmlMensagem.selectSingleNode("//CodMsg").Text)
    
    strNomeTagResultLiqd = ObterNomeTagRepeticao(strCodigoMensagem, 1)
    
    Call fgAppendNode(xmlDistinct, "", strNomeTagResultLiqd, "")
    
    'Select Disitnct CNPJNLiqdant
    '------------------------------------------------------------------------------------------------------
    For Each xmlNode In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        Set xmlGrupo = CreateObject("MSXML2.DOMDocument.4.0")
        
        If plngLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
            
            strQueryXPath = "//*[IdentdPartCamr='" & xmlNode.selectSingleNode("IdentdPartCamr").Text & "' and "
            strQueryXPath = strQueryXPath & "TpPgtoLDL='" & xmlNode.selectSingleNode("TpPgtoLDL").Text & "']"
            
            If xmlDistinct.selectSingleNode(strQueryXPath) Is Nothing Then
                
                If strCodigoMensagem = "LDL0007" Then
                    Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0007_EvtEms", "")
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "IdentdPartCamr", xmlNode.selectSingleNode("IdentdPartCamr").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "TpPgtoLDL", xmlNode.selectSingleNode("TpPgtoLDL").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "CNPJNLiqdant", xmlNode.selectSingleNode("CNPJNLiqdant").Text)
                    Call fgAppendXML(xmlDistinct, strNomeTagResultLiqd, xmlGrupo.xml)
                
                Else
                    Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0009R2_EvtEms", "")
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "IdentdPartCamr", xmlNode.selectSingleNode("IdentdPartCamr").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "TpPgtoLDL", xmlNode.selectSingleNode("TpPgtoLDL").Text)
                    Call fgAppendXML(xmlDistinct, strNomeTagResultLiqd, xmlGrupo.xml)
                End If
            End If
        Else
            strQueryXPath = "//*[CNPJNLiqdant='" & xmlNode.selectSingleNode("CNPJNLiqdant").Text & "']"
            
            If xmlDistinct.selectSingleNode(strQueryXPath) Is Nothing Then
                Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0009R2_EvtEms", "")
                Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "CNPJNLiqdant", xmlNode.selectSingleNode("CNPJNLiqdant").Text)
                Call fgAppendXML(xmlDistinct, strNomeTagResultLiqd, xmlGrupo.xml)
            End If
        End If
        
        Set xmlGrupo = Nothing
        
        'Substituir virgula(,) por ponto (.) como separador decimal
        If InStr(1, xmlNode.selectSingleNode("VlrNLiqdant").Text, ",", vbBinaryCompare) > 0 Then
            xmlNode.selectSingleNode("VlrNLiqdant").Text = Replace(xmlNode.selectSingleNode("VlrNLiqdant").Text, ",", ".", , , vbBinaryCompare)
        End If
        
    Next
    '------------------------------------------------------------------------------------------------------
    
    'Obter NET por CNPJNLiqdant
    '------------------------------------------------------------------------------------------------------
    Set xmlNetCNPJ = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlNetCNPJ, "", strNomeTagResultLiqd, "")
    For Each xmlNode In xmlDistinct.selectNodes("//" & strNomeTagResultLiqd & "/*")
        Set xmlGrupo = CreateObject("MSXML2.DOMDocument.4.0")
            
            If plngLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes Then
            
                strQueryXPath = "sum(//VlrNLiqdant[../IdentdPartCamr='" & xmlNode.selectSingleNode("IdentdPartCamr").Text & "' and ../ "
                strQueryXPath = strQueryXPath & "TpPgtoLDL='" & xmlNode.selectSingleNode("TpPgtoLDL").Text & "'])"
                
                If strCodigoMensagem = "LDL0007" Then
                    Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0007_EvtEms", "")
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "IdentdPartCamr", xmlNode.selectSingleNode("IdentdPartCamr").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "VlrNLiqdant", fgFuncaoXPath(pxmlMensagem, strQueryXPath))
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "TpPgtoLDL", xmlNode.selectSingleNode("TpPgtoLDL").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0007_EvtEms", "CNPJNLiqdant", xmlNode.selectSingleNode("CNPJNLiqdant").Text)
                Else
                    Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0009R2_EvtEms", "")
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "IdentdPartCamr", xmlNode.selectSingleNode("IdentdPartCamr").Text)
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "VlrNLiqdant", fgFuncaoXPath(pxmlMensagem, strQueryXPath))
                    Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "TpPgtoLDL", xmlNode.selectSingleNode("TpPgtoLDL").Text)
                End If
                Call fgAppendXML(xmlNetCNPJ, strNomeTagResultLiqd, xmlGrupo.xml)
            Else
                strQueryXPath = "sum(//VlrNLiqdant[../CNPJNLiqdant='" & xmlNode.selectSingleNode("CNPJNLiqdant").Text & "'])"
            
                Call fgAppendNode(xmlGrupo, "", "Grupo_LDL0009R2_EvtEms", "")
                Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "CNPJNLiqdant", xmlNode.selectSingleNode("CNPJNLiqdant").Text)
                Call fgAppendNode(xmlGrupo, "Grupo_LDL0009R2_EvtEms", "VlrNLiqdant", fgFuncaoXPath(pxmlMensagem, strQueryXPath))
                Call fgAppendXML(xmlNetCNPJ, strNomeTagResultLiqd, xmlGrupo.xml)
            End If
            
        Set xmlGrupo = Nothing
    Next
    '------------------------------------------------------------------------------------------------------
    
    
    'Substituir virgula(.) por ponto (,) como separador decimal
    '------------------------------------------------------------------------------------------------------
    For Each xmlNode In xmlNetCNPJ.selectNodes("//" & strNomeTagResultLiqd & "/*")
        If InStr(1, xmlNode.selectSingleNode("VlrNLiqdant").Text, ".", vbBinaryCompare) > 0 Then
            xmlNode.selectSingleNode("VlrNLiqdant").Text = Replace(xmlNode.selectSingleNode("VlrNLiqdant").Text, ".", ",", , , vbBinaryCompare)
        End If
    Next
    '------------------------------------------------------------------------------------------------------
    
    Set flPreparaMensagemLDL0009R2 = xmlNetCNPJ.documentElement
    
    Set xmlNetCNPJ = Nothing
    Set xmlDistinct = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlNetCNPJ = Nothing
    Set xmlDistinct = Nothing
    Set xmlGrupo = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flPreparaMensagemLDL0009R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se existe complemento para um Registro de SWAP
Private Function flExisteComplementoSWAP(ByRef pxmlOperacao As MSXML2.DOMDocument40, _
                                         ByRef plngStatusOperacao As Long, _
                                         ByRef pvntNuSequOperAtiv As Variant) As Boolean

Dim strSQL                                  As String
Dim rsQuery                                 As ADODB.Recordset
Dim strNumComando                           As String
Dim strCodigoVeiculoLegal                   As String
Dim vntIdentificadorParticipante            As Variant
Dim vntIdentificadorContraparte             As Variant
Dim strDHInicio                             As String
Dim strDHFim                                As String

On Error GoTo ErrorHandler
    
    With pxmlOperacao
        strNumComando = Trim$(.selectSingleNode("//NU_COMD_OPER").Text)
        vntIdentificadorParticipante = Val(.selectSingleNode("//CO_PARP_CAMR").Text)
        vntIdentificadorContraparte = Val(.selectSingleNode("//CO_CNPT_CAMR").Text)
    End With
    
    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
    
    strSQL = "SELECT    A.NU_SEQU_OPER_ATIV , " & _
             "          A.CO_ULTI_SITU_PROC " & _
             "  FROM    A8.TB_OPER_ATIV A,  " & _
             "          A8.TB_VEIC_LEGA B   " & _
             " Where    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA " & _
             "   AND    A.CO_CNPT_CAMR          =  " & vntIdentificadorContraparte & _
             "   AND    B.ID_PART_CAMR_CETIP    =  " & vntIdentificadorParticipante & _
             "   AND    A.NU_COMD_OPER          = '" & strNumComando & "'" & _
             "   AND    A.TP_OPER               =  " & enumTipoOperacaoLQS.RegDadosComplemContratoSWAP & _
             "   AND    DT_OPER_ATIV BETWEEN " & strDHInicio & _
             "   AND " & strDHFim
       
    Set rsQuery = fgQuerySQL(strSQL)
    
    If rsQuery.EOF Then
        flExisteComplementoSWAP = False
        pvntNuSequOperAtiv = 0
        plngStatusOperacao = 0
    Else
        flExisteComplementoSWAP = True
        pvntNuSequOperAtiv = rsQuery!NU_SEQU_OPER_ATIV
        plngStatusOperacao = rsQuery!CO_ULTI_SITU_PROC
    End If
                    
    Set rsQuery = Nothing

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteComplementoSWAP Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'------------------------------------------------------------------------------------------
'Verifica Status das mensagens BMA0003 para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusBMA0003(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Long

Dim strNumeroControleBMAOrig                As String
Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset


On Error GoTo ErrorHandler
    
    strNumeroControleBMAOrig = pxmlMensagem.selectSingleNode("//NumCtrlBMAOr").Text
    
    'Obter Operação Original'
    strSQL = " SELECT  B.NU_SEQU_OPER_ATIV              " & _
             "   FROM   A8.TB_MESG_RECB_ENVI_SPB    A,  " & _
             "          A8.TB_CNCL_OPER_ATIV B          " & _
             "  WHERE   A.NU_CTRL_CAMR = '" & strNumeroControleBMAOrig & "'" & _
             "    AND   A.NU_CTRL_IF = B.NU_CTRL_IF "
    
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        flObterStatusBMA0003 = enumStatusMensagem.Informação
    Else
        flObterStatusBMA0003 = enumStatusMensagem.MensagemInconsistente
    End If
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusBMA0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Reprocessamento das menssagens BMA0003 (BMA informa Operação cancelada)
'------------------------------------------------------------------------------------------

Private Function flReprocessaMensagemBMA0003(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean


Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

Dim objRS                                   As ADODB.Recordset
Dim strSQL                                  As String
Dim vntNumSequqnciaOperacao                 As Variant
Dim strNumeroControleBMAOrig                As String
Dim lngCodigoEmpresa                        As Long

On Error GoTo ErrorHandler

    strNumeroControleBMAOrig = xmlMensagem.selectSingleNode("//NumCtrlBMAOr").Text

    'Obter Operação Original'
    strSQL = " SELECT  B.NU_SEQU_OPER_ATIV              " & _
             "   FROM   A8.TB_MESG_RECB_ENVI_SPB    A,  " & _
             "          A8.TB_CNCL_OPER_ATIV B          " & _
             "  WHERE   A.NU_CTRL_CAMR = '" & strNumeroControleBMAOrig & "'" & _
             "    AND   A.NU_CTRL_IF = B.NU_CTRL_IF "
    
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
                
        vntNumSequqnciaOperacao = objRS!NU_SEQU_OPER_ATIV
        
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
        Set xmlOperacao = objOperacao.LerXMLOperacao(vntNumSequqnciaOperacao)
        
        If Not xmlOperacao.selectSingleNode("//CO_OPER_ATIV") Is Nothing Then
            xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = xmlOperacao.selectSingleNode("//CO_OPER_ATIV").Text
        End If
                 
        Set xmlOperacao = Nothing
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
        
        
        lngCodigoEmpresa = xmlMensagem.selectSingleNode("//CO_EMPR").Text
        
        
        udtProtocolo.CodigoEmpresa = Format$(lngCodigoEmpresa, "00000")
        udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
        udtProtocolo.SiglaSistemaDestino = fgCompletaString(Trim(xmlMensagem.selectSingleNode("//SG_SIST_DEST").Text), " ", 3, False)
        udtProtocolo.TipoMensagem = fgCompletaString(Trim("44"), "0", 9, True)
        
        LSet udtProtocoloAux = udtProtocolo
        
        '------------------------------------
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & xmlMensagem.xml, enumIdentificadorFila.BUS
        Set objProcessoOperacao = Nothing
    
    End If
    
    objRS.Close
    
    Exit Function
ErrorHandler:
    
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set objMesgRecbEnviSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flReprocessaMensagemBMA0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function VerificaCamaraAtiva(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByRef pstrCodigoMensagem As String) As Long

Dim lngCodigoLocalLiquidacao                As Long
Dim lngISPB                                 As Long
Dim blnCamaraValida                         As Boolean

On Error GoTo ErrorHandler
    
    Select Case Left$(pstrCodigoMensagem, 3)
        Case "LDL"
            If Not xmlMensagem.selectSingleNode("//ISPBLDL") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBLDL").Text)
            Else
                lngISPB = 0
            End If
        Case "LTR"
            If Not xmlMensagem.selectSingleNode("//ISPBLTR") Is Nothing Then
                lngISPB = Val(xmlMensagem.selectSingleNode("//ISPBLTR").Text)
            Else
                lngISPB = 0
            End If
        Case "CTP"
            lngISPB = enumISPB.IspbCETIP
        Case "BMA"
            lngISPB = enumISPB.IspbBMA
        Case "BMC"
            lngISPB = enumISPB.IspbBMC
        Case "SEL"
            lngISPB = enumISPB.IspbSELIC
    End Select
    
    Select Case lngISPB
        Case enumISPB.IspbBMA
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMA
        Case enumISPB.IspbCETIP
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.CETIP
        Case enumISPB.IspbCBLC
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.CLBCAcoes
        Case enumISPB.IspbBMD
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMD
        Case enumISPB.IspbBMC
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
        Case enumISPB.IspbSELIC
            lngCodigoLocalLiquidacao = enumLocalLiquidacao.SELIC
        Case Else
            lngCodigoLocalLiquidacao = 0
    End Select
    
    VerificaCamaraAtiva = lngCodigoLocalLiquidacao
        
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificaCamaraAtiva Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'------------------------------------------------------------------------------------------
'Verifica Status das mensagens LDL0007 CBLC e LDL0009R2 CBLC para reprocessamento
'------------------------------------------------------------------------------------------

Private Function flObterStatusLDL0007(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                      ByVal plngCodigoEmpresa As Long) As Long

Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim strNomeTagResultLiqd                    As String
Dim lngStatusMensagem                       As Long
Dim vntCodPartCamara                        As Variant

On Error GoTo ErrorHandler

    strNomeTagResultLiqd = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1)
    
    Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
    
    For Each xmlNodeResultLiqd In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntCodPartCamara = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
           
            If objTipoBackOffice.ObterTipoBackOfficeParticipanteCamara(vntCodPartCamara) <> 0 Then
                lngStatusMensagem = enumStatusMensagem.AConciliar
            Else
                lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
                Exit For
            End If
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            Exit For
        End If
    Next
    
    flObterStatusLDL0007 = lngStatusMensagem
    
    Set objTipoBackOffice = Nothing
    Exit Function

ErrorHandler:
        
    Set objTipoBackOffice = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0007 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Private Function flObterStatusLDL0005R2CBLC(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                           ByVal plngCodigoEmpresa As Long) As Long

Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim strNomeTagResultLiqd                    As String
Dim lngStatusMensagem                       As Long
Dim vntCodPartCamara                        As Variant

'CBLC
Dim strSQL                                  As String
Dim rsMESG                                  As ADODB.Recordset
Dim vntNumCtrlLDL                           As Variant
Dim vntVlFinanc                             As Variant

Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler

    strNomeTagResultLiqd = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1)
    
    Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
    
    For Each xmlNodeResultLiqd In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntCodPartCamara = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
            intTipoBackOffice = objTipoBackOffice.ObterTipoBackOfficeParticipanteCamara(vntCodPartCamara)
        Else
            intTipoBackOffice = enumTipoBackOffice.Tesouraria
        End If
           
        If intTipoBackOffice <> 0 Then
            lngStatusMensagem = enumStatusMensagem.AConciliar
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            Exit For
        End If
    
    Next
    
    Set objTipoBackOffice = Nothing
    
    If lngStatusMensagem = enumStatusMensagem.MensagemInconsistente Then
        flObterStatusLDL0005R2CBLC = enumStatusMensagem.MensagemInconsistente
        Exit Function
    Else
        flObterStatusLDL0005R2CBLC = lngStatusMensagem
    End If
    
    'Verifica se existe LDL0001
    If Not pxmlMensagem.selectSingleNode("//NumCtrlLDLOr") Is Nothing Then
        If pxmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text <> vbNullString Then
            vntNumCtrlLDL = pxmlMensagem.selectSingleNode("//NumCtrlLDLOr").Text
        Else
            'Se não tiver Numero Controle LDL original não temos como achar a LDL0001
            flObterStatusLDL0005R2CBLC = enumStatusMensagem.AConciliar
            Exit Function
        End If
    Else
        'Se não tiver Numero Controle LDL original não temos como achar a LDL0001
        flObterStatusLDL0005R2CBLC = enumStatusMensagem.AConciliar
        Exit Function
    End If
    
    vntVlFinanc = pxmlMensagem.selectSingleNode("//VlrLanc").Text
    
    strSQL = "  SELECT  A.NU_CTRL_IF,                   " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,            " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,             " & vbCrLf & _
             "          A.NU_CTRL_CAMR ,                " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,CO_ULTI_SITU_PROC,  " & vbCrLf & _
             "          B.VA_FINC ,                     " & vbCrLf & _
             "          B.TP_INFO_LDL                   " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A,     " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL B      " & vbCrLf & _
             "  Where   A.NU_CTRL_IF = B.NU_CTRL_IF     " & vbCrLf & _
             "    AND   A.NU_SEQU_CNTR_REPE = B.NU_SEQU_CNTR_REPE   " & vbCrLf & _
             "    AND   A.DH_REGT_MESG_SPB  = B.DH_REGT_MESG_SPB    " & vbCrLf & _
             "    AND   A.CO_MESG_SPB       = 'LDL0001'             " & vbCrLf & _
             "    AND   B.TP_INFO_LDL       = 'D'                   " & vbCrLf & _
             "    AND   B.IN_OPER_DEBT_CRED =                       " & enumTipoDebitoCredito.Credito & vbCrLf & _
             "    AND   A.NU_CTRL_CAMR      = '" & Trim$(vntNumCtrlLDL) & "'" & vbCrLf & _
             " ORDER BY  A.NU_SEQU_CNTR_REPE ASC"
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    If rsMESG.EOF Then
        'Se não tiver Numero Controle LDL original não temos como achar a LDL0001
        flObterStatusLDL0005R2CBLC = enumStatusMensagem.AConciliar
        Exit Function
    End If
    
    While Not rsMESG.EOF
        
        If rsMESG!NU_SEQU_CNTR_REPE = 1 Then
            If rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.MensagemLiquidada Or _
               rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.RecebimentoLib Then
            
                If fgVlrXml_To_Decimal(vntVlFinanc) = fgVlrXml_To_Decimal(rsMESG!VA_FINC) Then
                    
                    flObterStatusLDL0005R2CBLC = enumStatusMensagem.MensagemLiquidada
                    
                    If rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.RecebimentoLib Then
                        
                        Call flProcessaLiquidacaoLDL001CreditoCBCL(rsMESG!NU_CTRL_IF, _
                                                                   rsMESG!NU_SEQU_CNTR_REPE, _
                                                                   fgDtHr_To_Xml(rsMESG!DH_REGT_MESG_SPB), _
                                                                   IIf(IsNull(rsMESG!NU_SEQU_CNCL_OPER_ATIV_MESG), 0, rsMESG!NU_SEQU_CNCL_OPER_ATIV_MESG))
                    End If
                Else
                    flObterStatusLDL0005R2CBLC = enumStatusMensagem.AConciliar
                End If
            End If
        End If
        
        rsMESG.MoveNext
    Wend
    
    rsMESG.Close
    
    Exit Function

ErrorHandler:
    
    Set objTipoBackOffice = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0005R2CBLC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Private Sub flProcessaLiquidacaoLDL001CreditoCBCL(ByVal pstrNumeroControleIF As String, _
                                                  ByVal plngNumeroSequenciaRepeticao As Long, _
                                                  ByVal pstrDHRegistroMensagem As String, _
                                                  ByVal plngNumeroSequenciaConciliacao As Long)

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim vntSequenciaOperacao                    As Variant
Dim blnAtualizaStatusOperacao               As Boolean

'---
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim strRet                                  As String
Dim blnAcaoPagtoContingencia                As Boolean
Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
                
    '-------------------------------------------------------------------------
    'Verifica se foi Pagto Contingencia
    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
                    
    strRet = objHistMensagem.LerTodos(pstrNumeroControleIF, _
                                      pstrDHRegistroMensagem, _
                                      plngNumeroSequenciaRepeticao)
     
    xmlHistorico.loadXML strRet
                    
    blnAcaoPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
    
    Set objHistMensagem = Nothing
    Set xmlHistorico = Nothing
    '----------------------------------------------------------------------------
                
    If blnAcaoPagtoContingencia Then
        Exit Sub
    End If
                    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
        
    strRet = objMesgRecbEnviSPB.Ler(pstrNumeroControleIF, _
                                    plngNumeroSequenciaRepeticao, _
                                    pstrDHRegistroMensagem)

    
    xmlMesgRecbEnviSPB.loadXML strRet
                    
    Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                          enumStatusMensagem.MensagemLiquidada, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_RECB_ENVI_MESG_SPB").Text, _
                                          False, _
                                          0)
    
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    '-------------------------------------------------------------------------------------------------------------
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    strRet = objConciliacao.LerTodos(plngNumeroSequenciaConciliacao, 0, pstrNumeroControleIF)
    
    If xmlCociliacao.loadXML(strRet) Then
    
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
        For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
            
            vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
            
            Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
            
            Call GerenciarChamada(xmlOperacao, _
                                  xmlMensagem, _
                                  enumStatusOperacao.Liquidada, _
                                  0, _
                                  "", _
                                  enumIndicadorSimNao.Nao, _
                                  "")
        Next
        
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
    
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing

    Exit Sub
ErrorHandler:
    
    Set objHistMensagem = Nothing
    Set xmlHistorico = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaLiquidacaoLDL001CreditoCBCL Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'----------------------------------------------------------------------------------------------------------
'Processamento das mensagens BMC0112 (BMC0112 - BMC informa Demonstrativo de custos )
'----------------------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0112(ByRef xmlMensagem As MSXML2.DOMDocument40, _
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
Dim lngStatusMensagem                       As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistemaVeicLega                 As String
Dim strNomeVeicLega                         As String

Dim vntCNPJNLiqdant                         As Variant
Dim intIndDebitoCredito                     As Integer
Dim strCodigoPraca                          As String
Dim strNumCtrlBMC                           As String
Dim vntValorFinc                            As Variant

On Error GoTo ErrorHandler

    lngCodigoLocalLiquidacao = enumLocalLiquidacao.BMC
    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    lngTipoBackOffice = enumTipoBackOffice.Tesouraria
    lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    vntCNPJNLiqdant = fgObterCNPJEmpresa(lngCodigoEmpresa)
        
    vntValorFinc = xmlMensagem.selectSingleNode("//VlrTotCstBMC").Text
    strNumCtrlBMC = xmlMensagem.selectSingleNode("//NumCtrlBMC").Text
        
    lngStatusMensagem = enumStatusMensagem.AConciliar
        
    If fgObterCodigoVeiculoLegal(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                 vntCNPJNLiqdant, _
                                 lngCodigoEmpresa, _
                                 strCodigoVeiculoLegal, _
                                 strSiglaSistemaVeicLega, _
                                 lngTipoBackOffice, _
                                 strNomeVeicLega) Then
       
        Call flGeraAlertaCamara(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                lngTipoBackOffice, _
                                strNomeVeicLega, _
                                vntValorFinc, _
                                xmlMensagem)
    End If
           
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
                                                            lngStatusMensagem, _
                                                            lngIndicadorEntradaManual, _
                                                            vbNullString, _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistemaVeicLega, _
                                                            lngSeqControleRepeticao, _
                                                            strNumCtrlBMC)
    Set objMesgRecbEnviSPB = Nothing
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                strDataRegistroMensagem, _
                                xmlMensagem.selectSingleNode("//DtMovto").Text, _
                                0, _
                                0, _
                                vbNullString, _
                                0, _
                                0, _
                                vntValorFinc, _
                                lngSeqControleRepeticao)
                                
    Set objConciliacao = Nothing
    
    
    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0112 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function AtualizaStatusBMA0112(ByVal pstrChaveBMC0112 As String, _
                                      ByVal pvntSEQ_OPER_ATIV As Variant)

Dim vntChaveBMC0112                         As Variant
Dim vntSequencialConciliacaoOperacao        As Variant
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMesgRecebida                         As A8LQS.clsMensagemEnviadaRecebida

Dim strNumCtrlIF                            As String
Dim strDataHoraRegt                         As String
Dim lngSeqRepeticao                         As Long

On Error GoTo ErrorHandler
    
    vntChaveBMC0112 = Split(pstrChaveBMC0112, "|")
    
    strNumCtrlIF = Mid(vntChaveBMC0112(0), 2)
    strDataHoraRegt = vntChaveBMC0112(2)
    lngSeqRepeticao = vntChaveBMC0112(1)
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objMesgRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
            
    vntSequencialConciliacaoOperacao = objConciliacao.IncluirJustificativaConciliacaoOperacao()
    
    'Incluir Registro conciliacao com BMC00112
    Call objConciliacao.IncluirConciliacaoOperacao(vntSequencialConciliacaoOperacao, _
                                                   pvntSEQ_OPER_ATIV, _
                                                   strNumCtrlIF, _
                                                   strDataHoraRegt, _
                                                   0, _
                                                   lngSeqRepeticao)
                
    'Atualizar mensagem BMC0112 para Conciliada
    Call objMesgRecebida.AlterarStatus(strNumCtrlIF, _
                                       strDataHoraRegt, _
                                       enumStatusMensagem.Conciliada, _
                                       "", "", False)
            
    Set objConciliacao = Nothing
    Set objMesgRecebida = Nothing
            
    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set objMesgRecebida = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



'------------------------------------------------------------------------------------------
'Tratamento das mensagens de pagamento BMC0102R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemBMC0102R1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                               ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux

Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strDataOperacao                         As String
Dim strNumeroComando                        As String
Dim strCodigoMensagemMae                    As String
Dim strMensagem                             As String
Dim strNumeroControleIF                     As String
Dim strCodigoVeiculoLegal                   As String
Dim strDataRegistroMensagem                 As String
Dim strCodigoMensagemSPB                    As String
Dim strSiglaSistema                         As String
Dim vntSequenciaOperacao                    As Variant
Dim vntSequenciaConciliacaoOperacao         As Variant
Dim lngTipoBackOffice                       As Long
Dim lngCodigoEmpresa                        As Long
Dim lngIndicadorEntradaManual               As Long
Dim lngCodigoLocalLiquidacao                As Long
Dim lngStatusMensagemOriginal               As Long
Dim lngStatusOperacao                       As Long
Dim lngSeqControleRepeticao                 As Long
Dim blnSalvaMensagem                        As Boolean
Dim blnAtualizaStatusOperacao               As Boolean
Dim blnEstorno                              As Boolean

'---
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim strRet                                  As String
Dim blnAcaoPagtoContingencia                As Boolean
Dim blnGerenciarChamada                     As Boolean

On Error GoTo ErrorHandler
                    
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
                    
    strNumeroControleIF = udtHeaderMensagem.ControleRemessaNZ
    strCodigoMensagemSPB = Trim(xmlMensagem.selectSingleNode("//CodMsg").Text)
    blnSalvaMensagem = True
    blnAtualizaStatusOperacao = True
    blnEstorno = False
    
    If Not xmlMensagem.selectSingleNode("//DtMovto") Is Nothing Then
        strDataOperacao = xmlMensagem.selectSingleNode("//DtMovto").Text
    Else
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    End If
                    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMesgRecbEnviSPB.loadXML(objMesgRecbEnviSPB.LerMensagemIda(Trim(strNumeroControleIF))) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
                    
    lngSeqControleRepeticao = 1
    lngIndicadorEntradaManual = enumIndicadorSimNao.Nao
    strNumeroComando = xmlMesgRecbEnviSPB.selectSingleNode("//NU_COMD_OPER").Text
    vntSequenciaOperacao = xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    lngCodigoEmpresa = xmlMesgRecbEnviSPB.selectSingleNode("//CO_EMPR").Text
    strCodigoVeiculoLegal = xmlMesgRecbEnviSPB.selectSingleNode("//CO_VEIC_LEGA").Text
    strSiglaSistema = xmlMesgRecbEnviSPB.selectSingleNode("//SG_SIST").Text
    lngCodigoLocalLiquidacao = xmlMesgRecbEnviSPB.selectSingleNode("//CO_LOCA_LIQU").Text
    lngTipoBackOffice = Val(xmlMesgRecbEnviSPB.selectSingleNode("//TP_BKOF").Text)
    
    strDataRegistroMensagem = objMesgRecbEnviSPB.IncluirBMA(xmlMensagem.xml, _
                                                            Trim(strNumeroControleIF), _
                                                            vntSequenciaOperacao, _
                                                            vntSequenciaConciliacaoOperacao, _
                                                            strCodigoMensagemSPB, _
                                                            lngTipoBackOffice, _
                                                            lngCodigoEmpresa, _
                                                            strNumeroComando, _
                                                            "", _
                                                            enumStatusMensagem.R1, _
                                                            lngIndicadorEntradaManual, _
                                                            "", _
                                                            lngCodigoLocalLiquidacao, _
                                                            strCodigoVeiculoLegal, _
                                                            strSiglaSistema, _
                                                            lngSeqControleRepeticao)
    
    
        
     lngStatusMensagemOriginal = enumStatusMensagem.Confirmada
     
     objMesgRecbEnviSPB.AlterarStatus xmlMesgRecbEnviSPB.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                      lngStatusMensagemOriginal, _
                                      xmlMesgRecbEnviSPB.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
    
    '----------------------------------------------------------------------------------------------------------
    
    If vntSequenciaConciliacaoOperacao = 0 Then
        Set xmlMesgRecbEnviSPB = Nothing
        Set objMesgRecbEnviSPB = Nothing
        Exit Function
    End If
    
    strCodigoMensagemMae = flObterCodigoMensagemMae(Mid(strCodigoMensagemSPB, 1, 7)) ' BMC0101
    
    If strCodigoMensagemMae <> vbNullString Then
        
        strMensagem = flObterMensagemMae(vntSequenciaConciliacaoOperacao, strCodigoMensagemMae)
        
        xmlMesgRecbEnviSPB.loadXML strMensagem
        
        '-------------------------------------------------------------------------
        Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
                        
        strNumeroControleIF = xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text
        
        strRet = objHistMensagem.LerTodos(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                          xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
         
        xmlHistorico.loadXML strRet
                        
        blnAcaoPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
        
        Set objHistMensagem = Nothing
        Set xmlHistorico = Nothing
        '----------------------------------------------------------------------------
        
        If Not blnAcaoPagtoContingencia Then
            
            Call objMesgRecbEnviSPB.AlterarStatus(xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                  xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                  enumStatusMensagem.MensagemLiquidada, _
                                                  xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                  xmlMesgRecbEnviSPB.documentElement.selectSingleNode("//DH_RECB_ENVI_MESG_SPB").Text, _
                                                  False, _
                                                  1, _
                                                  strCodigoMensagemMae)
        End If
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    If Not blnAcaoPagtoContingencia Then
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao, 0, strNumeroControleIF)) Then
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
            For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
                
                vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                'Atualizar as operaçoes---------------------------
                Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                
                lngStatusOperacao = enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado
                                        
                GerenciarChamada xmlOperacao, _
                                 xmlMensagem, _
                                 lngStatusOperacao, _
                                 0, _
                                 "", _
                                 IIf(blnEstorno, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao), _
                                 "", _
                                 blnAcaoPagtoContingencia, blnAtualizaStatusOperacao
                        
            Next
            
            Set objOperacao = Nothing
            Set xmlOperacao = Nothing
        
        End If
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function

ErrorHandler:
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set objHistMensagem = Nothing
    Set xmlHistorico = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemBMC0102R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'------------------------------------------------------------------------------------------
'Processamento das mensagens CTP2002R1
'------------------------------------------------------------------------------------------

Private Function flProcessaMensagemCTP2002R1(ByVal xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMesgCTP2002                          As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgCTP2002 = CreateObject("MSXML2.DOMDocument.4.0")

    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux

    'Valida Mensagem de envio, solicitando as informações (CTP2002)
    Set xmlMesgCTP2002 = objMensagemEnvRec.LerXMLMensagem(udtHeaderMensagem.ControleRemessaNZ)
    If xmlMesgCTP2002.xml = vbNullString Then
        'Não foi possível encontrar a mensagem CTP2002 de ida
        lngCodigoErroNegocio = 4248
        GoTo ErrorHandler
    End If

    'Atualiza o status da mensagem de solicitando as informações
    Call objMensagemEnvRec.AlterarStatus(xmlMesgCTP2002.selectSingleNode("//NU_CTRL_IF").Text, _
                                         xmlMesgCTP2002.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                         enumStatusMensagem.Informação, _
                                         vbNullString, _
                                         , _
                                         False, _
                                         xmlMesgCTP2002.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

    'Incluir Mensagem CTP2002R1 para consulta
    Call objMensagemEnvRec.IncluirBMA(xmlRemessa.xml, _
                                       xmlMesgCTP2002.selectSingleNode("//NU_CTRL_IF").Text, _
                                       0, _
                                       0, _
                                       "CTP2002R1", _
                                       Val(xmlMesgCTP2002.selectSingleNode("//TP_BKOF").Text), _
                                       xmlMesgCTP2002.selectSingleNode("//CO_EMPR").Text, _
                                       vbNullString, _
                                       vbNullString, _
                                       enumStatusMensagem.R1, _
                                       enumIndicadorSimNao.Nao, _
                                       vbNullString, _
                                       enumLocalLiquidacao.CETIP, _
                                       xmlMesgCTP2002.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       xmlMesgCTP2002.selectSingleNode("//SG_SIST").Text, , _
                                       xmlRemessa.selectSingleNode("//NumCtrlCTP").Text)

Exit Function
ErrorHandler:
    Set objMensagemEnvRec = Nothing

    Set xmlMesgCTP2002 = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemCTP2002R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flObterStatusLDL0005R2BMF(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                           ByVal plngCodigoEmpresa As Long) As Long

Dim objTipoBackOffice                       As A8LQS.clsTipoBackOffice
Dim xmlNodeResultLiqd                       As MSXML2.IXMLDOMNode
Dim strNomeTagResultLiqd                    As String
Dim lngStatusMensagem                       As Long
Dim vntCodPartCamara                        As Variant

'CBLC
Dim strSQL                                  As String
Dim rsMESG                                  As ADODB.Recordset
Dim vntNumCtrlLDL                           As Variant
Dim vntVlFinanc                             As Variant

Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler

    strNomeTagResultLiqd = ObterNomeTagRepeticao(pxmlMensagem.selectSingleNode("//CodMsg").Text, 1)
    
    Set objTipoBackOffice = CreateObject("A8LQS.clsTipoBackOffice")
    
    For Each xmlNodeResultLiqd In pxmlMensagem.selectNodes("//" & strNomeTagResultLiqd & "/*")
        
        If Not xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr") Is Nothing Then
            vntCodPartCamara = xmlNodeResultLiqd.selectSingleNode("IdentdPartCamr").Text
            intTipoBackOffice = objTipoBackOffice.ObterTipoBackOfficeParticipanteCamara(vntCodPartCamara)
        Else
            intTipoBackOffice = enumTipoBackOffice.Tesouraria
        End If
           
        If intTipoBackOffice <> 0 Then
            lngStatusMensagem = enumStatusMensagem.AConciliar
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemInconsistente
            Exit For
        End If
    
    Next
    
    Set objTipoBackOffice = Nothing
    
    If lngStatusMensagem = enumStatusMensagem.MensagemInconsistente Then
        flObterStatusLDL0005R2BMF = enumStatusMensagem.MensagemInconsistente
        Exit Function
    Else
        flObterStatusLDL0005R2BMF = lngStatusMensagem
    End If
    

    
    vntVlFinanc = pxmlMensagem.selectSingleNode("//VlrLanc").Text
    
    strSQL = "  SELECT  A.NU_CTRL_IF,                   " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,            " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,             " & vbCrLf & _
             "          A.NU_CTRL_CAMR ,                " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,CO_ULTI_SITU_PROC,  " & vbCrLf & _
             "          B.VA_FINC ,                     " & vbCrLf & _
             "          B.TP_INFO_LDL                   " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A,     " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL B      " & vbCrLf & _
             "  Where   A.NU_CTRL_IF = B.NU_CTRL_IF     " & vbCrLf & _
             "    AND   A.NU_SEQU_CNTR_REPE = B.NU_SEQU_CNTR_REPE   " & vbCrLf & _
             "    AND   A.DH_REGT_MESG_SPB  = B.DH_REGT_MESG_SPB    " & vbCrLf & _
             "    AND   A.CO_MESG_SPB       = 'LDL0001'             " & vbCrLf & _
             "    AND   A.CO_LOCA_LIQU      = " & enumLocalLiquidacao.BMD & vbCrLf & _
             "    AND   B.TP_INFO_LDL       = 'D'                   " & vbCrLf & _
             "    AND   B.IN_OPER_DEBT_CRED =                       " & enumTipoDebitoCredito.Credito & vbCrLf & _
             "    AND   B.DT_OPER           = " & fgDtXML_To_Oracle(pxmlMensagem.selectSingleNode("//DtMovto").Text) & vbCrLf
    'KIDA 11/12/2009
    'RATS 696

    If vntNumCtrlLDL <> vbNullString Then
        strSQL = strSQL & "    AND   A.NU_CTRL_CAMR      = '" & Trim$(vntNumCtrlLDL) & "'" & vbCrLf
    End If
    
    strSQL = strSQL & " ORDER BY  A.NU_SEQU_CNTR_REPE ASC"
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    If rsMESG.EOF Then
        'Se não tiver Numero Controle LDL original não temos como achar a LDL0001
        flObterStatusLDL0005R2BMF = enumStatusMensagem.AConciliar
        Exit Function
    End If
    
    While Not rsMESG.EOF
        
        If rsMESG!NU_SEQU_CNTR_REPE = 1 Then
            If rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.MensagemLiquidada Or _
               rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.RecebimentoLib Then
            
                If fgVlrXml_To_Decimal(vntVlFinanc) = fgVlrXml_To_Decimal(rsMESG!VA_FINC) Then
                    
                    flObterStatusLDL0005R2BMF = enumStatusMensagem.MensagemLiquidada
                    
                    If rsMESG!CO_ULTI_SITU_PROC = enumStatusMensagem.RecebimentoLib Then
                        
                        Call flProcessaLiquidacaoLDL001CreditoCBCL(rsMESG!NU_CTRL_IF, _
                                                                   rsMESG!NU_SEQU_CNTR_REPE, _
                                                                   fgDtHr_To_Xml(rsMESG!DH_REGT_MESG_SPB), _
                                                                   IIf(IsNull(rsMESG!NU_SEQU_CNCL_OPER_ATIV_MESG), 0, rsMESG!NU_SEQU_CNCL_OPER_ATIV_MESG))
                    End If
                Else
                    flObterStatusLDL0005R2BMF = enumStatusMensagem.AConciliar
                End If
            End If
        End If
        
        rsMESG.MoveNext
    Wend
    
    rsMESG.Close
    
    Set rsMESG = Nothing
    
    Exit Function

ErrorHandler:
    Set rsMESG = Nothing
    Set objTipoBackOffice = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusLDL0005R2BMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


