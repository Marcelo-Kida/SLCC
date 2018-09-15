VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3E32CF037A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável por promover as integrações com o sistema contábil:
''      HA - Contábil Geral

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

'Gerenciamento de integração com sistema HA
Public Function Integrar(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                Optional ByRef xmlLanctoCCContabil As MSXML2.DOMDocument40 = Nothing, _
                Optional ByVal plngNumSequLancamento As Long = 0) As Boolean

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim strRegistroContabil                     As String
Dim lngTipoDebitoCreditoCC                  As Long
Dim lngTipoDebitoCreditoOriginal            As Long
Dim lngTipoMensagem                         As Long
Dim lngFormaLiquidacao                      As Long
Dim lngOperInternaExterna                   As Long

Dim lngLocalLiquidacao                      As Long
Dim intTipoContraparte                      As Integer
Dim intTipoLiquidacao                       As Integer


    On Error GoTo ErrorHandler
    
    lngTipoDebitoCreditoOriginal = Val(xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text)
    
    lngTipoMensagem = 0
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        lngTipoMensagem = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
    End If
    
    lngFormaLiquidacao = 0
    If Not xmlOperacao.selectSingleNode("//CO_FORM_LIQU") Is Nothing Then
        lngFormaLiquidacao = Val(xmlOperacao.selectSingleNode("//CO_FORM_LIQU").Text)
    End If
    
    lngOperInternaExterna = 0
    If Not xmlOperacao.selectSingleNode("//TP_IF_CRED_DEBT") Is Nothing Then
        lngOperInternaExterna = Val(xmlOperacao.selectSingleNode("//TP_IF_CRED_DEBT").Text)
    End If
    
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    lngTipoDebitoCreditoCC = objLancamentoCCContabil.ObterTipoDebitoCreditoCC(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                                              xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                                              lngTipoDebitoCreditoOriginal, _
                                                                              0, _
                                                                              xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)

    Select Case lngTipoMensagem
        Case enumTipoMensagemBUS.OperacoesCorretoras, _
             enumTipoMensagemBUS.EnvioPagDespesas
            If Not xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED") Is Nothing Then
                lngTipoDebitoCreditoCC = Val(xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED").Text)
            End If
        Case enumTipoMensagemBUS.EnvioTEDClientes
            lngTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito
                
    End Select
    
    If Not xmlOperacao.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        lngLocalLiquidacao = Val(xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(xmlOperacao.selectSingleNode("//TP_CNPT").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intTipoLiquidacao = Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    End If
    
    If lngLocalLiquidacao = enumLocalLiquidacao.CETIP And lngTipoMensagem <> enumTipoMensagemBUS.DespesasCETIP Then
        
        If intTipoContraparte = enumTipoContraparte.Cliente1 And intTipoLiquidacao = enumTipoLiquidacao.SemModalidade Then
            lngTipoDebitoCreditoCC = lngTipoDebitoCreditoOriginal
        Else
            lngTipoDebitoCreditoCC = IIf(lngTipoDebitoCreditoOriginal = enumTipoDebitoCredito.Credito, _
                                                                        enumTipoDebitoCredito.Debito, _
                                                                        enumTipoDebitoCredito.Credito)
        End If
    
    End If
        
    'if xmlOperacao.documentElement.selectSingleNode
        
    If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        Select Case Val("0" & xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.EnvioSTR0006PagDespesas, _
                 enumTipoOperacaoLQS.EnvioSTR0007PagDespesas, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesas, _
                 enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta, _
                 enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib, _
                 enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta, _
                 enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta, _
                 enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib
                
                If plngNumSequLancamento = 0 Then
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento, True)
'                    Call flMontarRegistroPZ(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, True)
                    Call fgWait(1)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, True)
'                    Call flMontarRegistroPZ(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, False)
                Else
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
   
                End If
            
            Case enumTipoOperacaoLQS.EnvioPagDespesasBoleto, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributos, _
                 enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, _
                 enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib
                 
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento, True)
                    If plngNumSequLancamento >= 2 Then
                        Call fgWait(1)
                        Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
                        Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False)
                    End If
            
            Case Else
                '''NICK - Problema em Produção - Estorno COE (IF para inverter o lançamento em caso de estorno)
                
                If (xmlOperacao.documentElement.selectSingleNode("//CD_SUB_PROD").Text = "COE" And _
                        Val(xmlOperacao.documentElement.selectSingleNode("//TP_CNAL_VEND").Text) = 1 And _
                            Val(xmlOperacao.documentElement.selectSingleNode("//TP_SOLI").Text) = 3) Then
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False, True)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento, False, True)
                Else
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento)
                    Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento)
                End If
                
        End Select
    Else
        Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC, plngNumSequLancamento)
        Call flMontarRegistro(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC, plngNumSequLancamento)

    End If
    
    If lngTipoMensagem = enumTipoMensagemBUS.OperacoesCorretoras Then
        If lngTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito And lngOperInternaExterna = 1 Then
            Call flMontarRegistroPZ(xmlOperacao, enumTipoDebitoCredito.Credito, lngTipoDebitoCreditoCC)
            Call flMontarRegistroPZ(xmlOperacao, enumTipoDebitoCredito.Debito, lngTipoDebitoCreditoCC)
        End If
    End If

    Set objLancamentoCCContabil = Nothing

    Exit Function

ErrorHandler:
    Set objLancamentoCCContabil = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Integrar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montagem da remessa para o HA
Private Function flMontarRegistro(ByRef xmlOperacao As DOMDocument40, _
                                  ByVal plngDebitoCredito As enumTipoDebitoCredito, _
                                  ByVal plngDebitoCreditoCC As enumTipoDebitoCredito, _
                                  Optional plngNumSequLancamento As Long = 0, _
                                  Optional ByVal pblnTipoOperacao As Boolean = False, _
                                  Optional ByVal pblnOperacaoEstorno As Boolean = False) As Boolean

Dim udtContabilidade                        As udtContabilidade
Dim udtContabilidadeAux                     As udtContabilidadeAux

Dim objEmpresa                              As A6A7A8.clsEmpresa
Dim objParmHistCntaCntb                     As A8LQS.clsParmHistCntaCntb
Dim strEmpresa                              As String
Dim strParmHistCntaCntb                     As String
Dim lngTipoBackOffice                       As Long
Dim xmlParmHistCntaCntb                     As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim vntValor                                As Variant
Dim strValor                                As String
Dim strObserv4                              As String
Dim strDataContabilizacao                   As String
Dim strCentroDestino                        As String

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim strParametroGeral                       As String

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMensagem                             As MSXML2.DOMDocument40
    
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim plngTipoOperacao                        As Long

Dim strComplemento                          As String

    On Error GoTo ErrorHandler

    lngTipoBackOffice = fgObterTipoBackOffice(xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    Set objParmHistCntaCntb = CreateObject("A8LQS.clsParmHistCntaCntb")
    Set xmlParmHistCntaCntb = CreateObject("MSXML2.DOMDocument.4.0")
    
    strCentroDestino = vbNullString
    
    If plngDebitoCreditoCC = enumTipoDebitoCredito.Credito Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
    
                strParmHistCntaCntb = objParmHistCntaCntb.Ler(xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                              IIf(plngDebitoCredito = enumTipoDebitoCredito.Credito, _
                                                                                      enumTipoDebitoCredito.Debito, _
                                                                                      enumTipoDebitoCredito.Credito))
    
                If strParmHistCntaCntb = vbNullString Then
                    ' 4328 - Conta e Histórico Contábil não cadastrados para o Sistema, Empresa, Local de Liquidação, Tipo de Operação e Tipo de Movimento informados.
                    lngCodigoErroNegocio = 4328
                    GoTo ErrorHandler
                End If
    
                xmlParmHistCntaCntb.loadXML strParmHistCntaCntb
                strCentroDestino = xmlParmHistCntaCntb.selectSingleNode("//CO_CENT_DEST").Text
    
            End If
        End If
    End If
    
''    If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
''        If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
''            plngTipoOperacao = CLng("0" & xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
''        Else
''            plngTipoOperacao = 0
''        End If
''    End If
    If pblnTipoOperacao Then
        
        If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
                plngTipoOperacao = CLng("0" & xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
            Else
               plngTipoOperacao = 0
            End If
        End If
        
        strParmHistCntaCntb = objParmHistCntaCntb.Ler(xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                  lngTipoBackOffice, _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                  plngDebitoCredito, _
                                                  plngTipoOperacao)
    Else
        strParmHistCntaCntb = objParmHistCntaCntb.Ler(xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                  lngTipoBackOffice, _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                  plngDebitoCredito, _
                                                  0)
    
    End If
    If strParmHistCntaCntb = vbNullString Then
        ' 4328 - Conta e Histórico Contábil não cadastrados para o Sistema, Empresa, Local de Liquidação, Tipo de Operação e Tipo de Movimento informados.
        lngCodigoErroNegocio = 4328
        GoTo ErrorHandler
    End If

    Set objParmHistCntaCntb = Nothing

    xmlParmHistCntaCntb.loadXML strParmHistCntaCntb

    strParametroGeral = fgSelectVarchar4000(0, False)
    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametroGeral.loadXML strParametroGeral

    strDataContabilizacao = vbNullString
    If Not xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO") Is Nothing Then
        If xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text <> vbNullString And _
           xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO/@OBRIG").Text = "S" Then
            strDataContabilizacao = Replace$(xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text, "-", vbNullString)
            strDataContabilizacao = Replace$(strDataContabilizacao, "/", vbNullString)
            strDataContabilizacao = Right$(strDataContabilizacao, 4) & Mid$(strDataContabilizacao, 3, 2) & Left$(strDataContabilizacao, 2)
        End If
    End If
    
    With udtContabilidade
        
        If Not xmlOperacao.selectSingleNode("//CO_BANC") Is Nothing Then
            .EMPRESA = fgCompletaString(xmlOperacao.selectSingleNode("//CO_BANC").Text, "0", 4, True)  '4
        Else
            .EMPRESA = "0000"  '4
        End If
        
        .CLAVE_DE_INTERFASE = "A8A" '3
    
        If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
            .FECHA_CONTABLE = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text '8
        Else
            .FECHA_CONTABLE = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text '8
        End If
        
        If strDataContabilizacao <> vbNullString Then
            .FECHA_CONTABLE = strDataContabilizacao '8
        End If
    
        .FECHA_DE_OPERACION = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))  '8
        .PRODUCTO = fgCompletaString("0", "0", 2, True)   '2
        .SUBPRODUCTO = fgCompletaString(" ", " ", 4, False) '4
        .GARANTIA = fgCompletaString("0", "0", 3, True) '3
        .TIPO_DE_PLAZO = " "
        .PLAZO = fgCompletaString("0", "0", 3, True) '3
        .SUBSECTOR = "0" '1
        .SECTOR_B_E = fgCompletaString("0", "0", 2, True) '2
        .CNAE = fgCompletaString("0", "0", 5, True) '5
        .EMPRESA_TUTELADA = fgCompletaString(" ", " ", 4, False) '4
        .AMBITO = fgCompletaString("0", "0", 2, True) '2
        .MOROSIDAD = " " '1
        .INVERSION = " " '1
        .OPERACION = fgCompletaString("0", "0", 3, True) '3

        If plngDebitoCreditoCC = enumTipoDebitoCredito.Credito Then
            .CODIGO_CONTABLE = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_CNTA_CRED").Text, "0", 5, True) '5
        Else
            .CODIGO_CONTABLE = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_CNTA_DEBT").Text, "0", 5, True) '5
        End If

        .DIVISA = fgCompletaString(" ", " ", 3, False) '3
        .TIPO_DE_DIVISA = " "  '1
        .TIPO_NOMINAL = fgCompletaString("0", "0", 5, True) '5
        .Filler1 = fgCompletaString(" ", " ", 5, False) '5

        Select Case .CODIGO_CONTABLE
            Case "00017"
                .VARIOS = fgCompletaString("WZ", " ", 30, False) '30
            Case "00014"
                .VARIOS = fgCompletaString("NZ", " ", 30, False) '30
            Case Else
                .VARIOS = fgCompletaString("BG", " ", 30, False) '30
        End Select
        
        If strCentroDestino = vbNullString Then
            If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
                If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.DespesasBMC Or _
                   Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Or _
                   Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EnvioTEDClientes Or _
                   Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
                    strCentroDestino = xmlParmHistCntaCntb.selectSingleNode("//CO_CENT_DEST").Text
                Else
                    strCentroDestino = "6544"
                End If
            Else
                strCentroDestino = "6544"
            End If
        End If
        
        .CENTRO_DESTINO = strCentroDestino
        
        .CLAVE_DE_AUTORIZACION = fgCompletaString(" ", " ", 6, False) '6
        .CENTRO_OPERANTE = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) '4
        .CENTRO_ORIGEN = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) '4
       
        If Not xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC") Is Nothing Then
            vntValor = Split(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text, ",", , vbBinaryCompare)
        Else
            vntValor = Split("0,0", ",", , vbBinaryCompare)
        End If
        
        If UBound(vntValor) = 0 Then
            strValor = fgCompletaString(vntValor(0), "0", 13, True) & "00"
        Else
            strValor = fgCompletaString(vntValor(0), "0", 13, True) & fgCompletaString(vntValor(1), "0", 2, False)
        End If
        
        'NICK _ ERRO DE PRODUÇÂO - Estorno COE - IF PARA INVERTER OS LANÇAMENTOS
        If pblnOperacaoEstorno Then
            plngDebitoCredito = IIf(plngDebitoCredito = enumTipoDebitoCredito.Credito, _
                                                    enumTipoDebitoCredito.Debito, _
                                                    enumTipoDebitoCredito.Credito)
        End If
        
        If plngDebitoCredito = enumTipoDebitoCredito.Credito Then
            .NUM_MOVTOS_AL_DEBE = fgCompletaString("0", "0", 7, True) '7
            .NUM_MOVTOS_AL_HABER = fgCompletaString("1", "0", 7, True) '7
            .IMPORTE_DEBE_EN_PESETAS = fgCompletaString("0", "0", 15, True) '15
            .IMPORTE_HABER_EN_PESETAS = strValor
        Else
            .NUM_MOVTOS_AL_DEBE = fgCompletaString("1", "0", 7, True) '7
            .NUM_MOVTOS_AL_HABER = fgCompletaString("0", "0", 7, True) '7
            .IMPORTE_DEBE_EN_PESETAS = strValor
            .IMPORTE_HABER_EN_PESETAS = fgCompletaString("0", "0", 15, True) '15
        End If

        .IMPORTE_DEBE_EN_DIVISA = fgCompletaString("0", "0", 15, True) '15
        .IMPORTE_HABER_EN_DIVISA = fgCompletaString("0", "0", 15, True)  '15
        .INDICADOR_DE_CORRECCION = " " '1
        .NUMERO_DE_CONTROL = fgCompletaString(" ", " ", 12, False) '12
        .CLAVE_DE_CONCEPTO = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_HIST_CNTA_CNTB").Text, " ", 3, False) '3
        .DESCRIPCION_DE_CONCEPTO = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//DE_HIST_CNTA_CNTB").Text, " ", 14, False) '14
        .TIPODE_CONCEPTO = " " '1
        .OBSERVACIONES = fgCompletaString(" ", " ", 30, False) '30
        .SANCTCCC = fgCompletaString(" ", " ", 18, False) '18
        .APLICACION_ORIGEN = fgCompletaString(" ", " ", 3, False) '3
        .APLICACION_DESTINO = fgCompletaString(" ", " ", 3, False) '3
        .OBSERVACIONES3 = fgCompletaString(" ", " ", 6, False) '6
        .RESERVAT = fgCompletaString(" ", " ", 4, False) '4
        .HACTRGEN = fgCompletaString(" ", " ", 4, False) '4
        .HAYCOCAI = " " '1
        .HAYCTORD = " " '1
        .SATINTER = fgCompletaString("0", "0", 5, True) '5
        .SACCLVOP = fgCompletaString(" ", " ", 3, False) '3
        .SACCEGES = fgCompletaString(" ", " ", 4, False) '4
        .SACAPLCP = fgCompletaString(" ", " ", 2, False) '2
        .SACCDTGT = fgCompletaString(" ", " ", 2, False) '2
        .SAYUTILI = " "  '1
        .SAYROTAC = fgCompletaString(" ", " ", 2, False) '2
        .FALTPART = fgCompletaString(" ", " ", 8, False) '8
        
        strObserv4 = .DESCRIPCION_DE_CONCEPTO
        
        If Not xmlOperacao.selectSingleNode("//CO_MESG_SPB_REGT_OPER") Is Nothing Then
            If Left$(xmlOperacao.selectSingleNode("//CO_MESG_SPB_REGT_OPER").Text, 3) = "STR" Or _
               Left$(xmlOperacao.selectSingleNode("//CO_MESG_SPB_REGT_OPER").Text, 3) = "PAG" Then
                
                Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(, , , _
                                                                            xmlOperacao.selectSingleNode("//CO_MESG_SPB_REGT_OPER").Text, _
                                                                            , , _
                                                                            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                
                If xmlMensagem.xml <> vbNullString Then
                    strObserv4 = strObserv4 & Space$(3) & _
                                 fgCompletaString(Mid$(xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text, 2), " ", 6, False) & _
                                 "A8" & Left$(xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text, 1)
                Else
                    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
                    Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
                    
                    If xmlConciliacao.xml <> vbNullString Then
                        Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(xmlConciliacao.selectSingleNode("//NU_CTRL_IF").Text, , , _
                                                                                    xmlOperacao.selectSingleNode("//CO_MESG_SPB_REGT_OPER").Text)
                        If xmlMensagem.xml <> vbNullString Then
                            strObserv4 = strObserv4 & Space$(3) & _
                                         fgCompletaString(Mid$(xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text, 2), " ", 6, False) & _
                                         "A8" & Left$(xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text, 1)
                        End If
                    End If
                    Set objConciliacao = Nothing
                End If
                
                Set objMensagemEnviadaRecebida = Nothing
            
            End If
        End If
            
        strObserv4 = fgCompletaString(strObserv4, " ", 30, False)
        .OBSERV4 = strObserv4 '30
            
        .NIO = fgCompletaString(" ", " ", 24, False) '24
        .Filler2 = fgCompletaString(" ", " ", 2, False) '2
    End With

    LSet udtContabilidadeAux = udtContabilidade
                                                                        
    IncluirContabil "HA", xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                    xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                    enumTipoLancamentoIntegracao.Normal, _
                    udtContabilidadeAux.String, _
                    plngDebitoCredito, _
                    0, plngNumSequLancamento
    
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlConciliacao = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function

ErrorHandler:

    Set xmlConciliacao = Nothing
    Set xmlMensagem = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarRegistro Function", lngCodigoErroNegocio, intNumeroSequencialErro, strComplemento)

End Function

'Montagem da remessa para o HA
Private Function flMontarRegistroPZ(ByRef xmlOperacao As DOMDocument40, _
                                    ByVal plngDebitoCredito As enumTipoDebitoCredito, _
                                    ByVal plngDebitoCreditoCC As enumTipoDebitoCredito, _
                                    Optional ByVal pblnTipoOperacao As Boolean = False) As Boolean

Dim udtContabilidade                        As udtContabilidade
Dim udtContabilidadeAux                     As udtContabilidadeAux

Dim objEmpresa                              As A6A7A8.clsEmpresa
Dim objParmHistCntaCntb                     As A8LQS.clsParmHistCntaCntb
Dim xmlParmHistCntaCntb                     As MSXML2.DOMDocument40
Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim plngTipoOperacao                        As Long


Dim strEmpresa                              As String
Dim strParmHistCntaCntb                     As String
Dim lngTipoBackOffice                       As Long
Dim vntValor                                As Variant
Dim strValor                                As String
Dim strDataContabilizacao                   As String
Dim strParametroGeral                       As String
Dim strCentroDestino                        As String

    On Error GoTo ErrorHandler

    lngTipoBackOffice = fgObterTipoBackOffice(xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    Set objParmHistCntaCntb = CreateObject("A8LQS.clsParmHistCntaCntb")
    Set xmlParmHistCntaCntb = CreateObject("MSXML2.DOMDocument.4.0")
    
    strCentroDestino = vbNullString
    
    If plngDebitoCreditoCC = enumTipoDebitoCredito.Credito Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then

                strParmHistCntaCntb = objParmHistCntaCntb.Ler(xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                              IIf(plngDebitoCredito = enumTipoDebitoCredito.Credito, _
                                                                                      enumTipoDebitoCredito.Debito, _
                                                                                      enumTipoDebitoCredito.Credito))

                If strParmHistCntaCntb = vbNullString Then
                    ' 3042 - Código do histórico de Conta Corrente inválido.
                    lngCodigoErroNegocio = 3042
                    GoTo ErrorHandler
                End If

                xmlParmHistCntaCntb.loadXML strParmHistCntaCntb
                strCentroDestino = xmlParmHistCntaCntb.selectSingleNode("//CO_CENT_DEST").Text

            End If
        End If
    End If
    
    If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
        If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
            plngTipoOperacao = CLng("0" & xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        Else
            plngTipoOperacao = 0
        End If
    End If
    
    strParmHistCntaCntb = objParmHistCntaCntb.Ler("PZ", _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                  lngTipoBackOffice, _
                                                  xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                  plngDebitoCredito, IIf(pblnTipoOperacao, plngTipoOperacao, 0))
    
    If strParmHistCntaCntb = vbNullString Then
        ' 3042 - Código do histórico de Conta Corrente inválido.
        lngCodigoErroNegocio = 3042
        GoTo ErrorHandler
    End If

    Set objParmHistCntaCntb = Nothing

    xmlParmHistCntaCntb.loadXML strParmHistCntaCntb

    strParametroGeral = fgSelectVarchar4000(0, False)
    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametroGeral.loadXML strParametroGeral

    strDataContabilizacao = vbNullString
    If Not xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO") Is Nothing Then
        If xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text <> vbNullString And _
           xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO/@OBRIG").Text = "S" Then
            strDataContabilizacao = Replace$(xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text, "-", vbNullString)
            strDataContabilizacao = Replace$(strDataContabilizacao, "/", vbNullString)
            strDataContabilizacao = Right$(strDataContabilizacao, 4) & Mid$(strDataContabilizacao, 3, 2) & Left$(strDataContabilizacao, 2)
        End If
    End If
    
    With udtContabilidade
        
        If Not xmlOperacao.selectSingleNode("//CO_BANC") Is Nothing Then
            .EMPRESA = fgCompletaString(xmlOperacao.selectSingleNode("//CO_BANC").Text, "0", 4, True)  '4
        Else
            .EMPRESA = "0000"  '4
        End If
        
        .CLAVE_DE_INTERFASE = "A8A" '3
    
        If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
            .FECHA_CONTABLE = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text '8
        Else
            .FECHA_CONTABLE = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text '8
        End If
        
        If strDataContabilizacao <> vbNullString Then
            .FECHA_CONTABLE = strDataContabilizacao '8
        End If
    
        .FECHA_DE_OPERACION = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))  '8
        .PRODUCTO = fgCompletaString("0", "0", 2, True)   '2
        .SUBPRODUCTO = fgCompletaString(" ", " ", 4, False) '4
        .GARANTIA = fgCompletaString("0", "0", 3, True) '3
        .TIPO_DE_PLAZO = " "
        .PLAZO = fgCompletaString("0", "0", 3, True) '3
        .SUBSECTOR = "0" '1
        .SECTOR_B_E = fgCompletaString("0", "0", 2, True) '2
        .CNAE = fgCompletaString("0", "0", 5, True) '5
        .EMPRESA_TUTELADA = fgCompletaString(" ", " ", 4, False) '4
        .AMBITO = fgCompletaString("0", "0", 2, True) '2
        .MOROSIDAD = " " '1
        .INVERSION = " " '1
        .OPERACION = fgCompletaString("0", "0", 3, True) '3

        If plngDebitoCreditoCC = enumTipoDebitoCredito.Credito Then
            .CODIGO_CONTABLE = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_CNTA_CRED").Text, "0", 5, True) '5
        Else
            .CODIGO_CONTABLE = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_CNTA_DEBT").Text, "0", 5, True) '5
        End If

        .DIVISA = fgCompletaString(" ", " ", 3, False) '3
        .TIPO_DE_DIVISA = " "  '1
        .TIPO_NOMINAL = fgCompletaString("0", "0", 5, True) '5
        .Filler1 = fgCompletaString(" ", " ", 5, False) '5
        .VARIOS = fgCompletaString("PZ", " ", 30, False) '30
        
        If strCentroDestino = vbNullString Then
            strCentroDestino = xmlParmHistCntaCntb.selectSingleNode("//CO_CENT_DEST").Text
        End If
        
        .CENTRO_DESTINO = strCentroDestino
        
        .CLAVE_DE_AUTORIZACION = fgCompletaString(" ", " ", 6, False) '6
        .CENTRO_OPERANTE = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) '4
        .CENTRO_ORIGEN = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) '4
       
        If Not xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC") Is Nothing Then
            vntValor = Split(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text, ",", , vbBinaryCompare)
        Else
            vntValor = Split("0,0", ",", , vbBinaryCompare)
        End If
        
        If UBound(vntValor) = 0 Then
            strValor = fgCompletaString(vntValor(0), "0", 13, True) & "00"
        Else
            strValor = fgCompletaString(vntValor(0), "0", 13, True) & fgCompletaString(vntValor(1), "0", 2, False)
        End If
                
        If plngDebitoCredito = enumTipoDebitoCredito.Credito Then
            .NUM_MOVTOS_AL_DEBE = fgCompletaString("0", "0", 7, True) '7
            .NUM_MOVTOS_AL_HABER = fgCompletaString("1", "0", 7, True) '7
            .IMPORTE_DEBE_EN_PESETAS = fgCompletaString("0", "0", 15, True) '15
            .IMPORTE_HABER_EN_PESETAS = strValor
        Else
            .NUM_MOVTOS_AL_DEBE = fgCompletaString("1", "0", 7, True) '7
            .NUM_MOVTOS_AL_HABER = fgCompletaString("0", "0", 7, True) '7
            .IMPORTE_DEBE_EN_PESETAS = strValor
            .IMPORTE_HABER_EN_PESETAS = fgCompletaString("0", "0", 15, True) '15
        End If

        .IMPORTE_DEBE_EN_DIVISA = fgCompletaString("0", "0", 15, True) '15
        .IMPORTE_HABER_EN_DIVISA = fgCompletaString("0", "0", 15, True)  '15
        .INDICADOR_DE_CORRECCION = " " '1
        .NUMERO_DE_CONTROL = fgCompletaString(" ", " ", 12, False) '12
        .CLAVE_DE_CONCEPTO = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//CO_HIST_CNTA_CNTB").Text, " ", 3, False) '3
        .DESCRIPCION_DE_CONCEPTO = fgCompletaString(xmlParmHistCntaCntb.documentElement.selectSingleNode("//DE_HIST_CNTA_CNTB").Text, " ", 14, False) '14
        .TIPODE_CONCEPTO = " " '1
        .OBSERVACIONES = fgCompletaString(" ", " ", 30, False) '30
        .SANCTCCC = fgCompletaString(" ", " ", 18, False) '18
        .APLICACION_ORIGEN = fgCompletaString(" ", " ", 3, False) '3
        .APLICACION_DESTINO = fgCompletaString(" ", " ", 3, False) '3
        .OBSERVACIONES3 = fgCompletaString(" ", " ", 6, False) '6
        .RESERVAT = fgCompletaString(" ", " ", 4, False) '4
        .HACTRGEN = fgCompletaString(" ", " ", 4, False) '4
        .HAYCOCAI = " " '1
        .HAYCTORD = " " '1
        .SATINTER = fgCompletaString("0", "0", 5, True) '5
        .SACCLVOP = fgCompletaString(" ", " ", 3, False) '3
        .SACCEGES = fgCompletaString(" ", " ", 4, False) '4
        .SACAPLCP = fgCompletaString(" ", " ", 2, False) '2
        .SACCDTGT = fgCompletaString(" ", " ", 2, False) '2
        .SAYUTILI = " "  '1
        .SAYROTAC = fgCompletaString(" ", " ", 2, False) '2
        .FALTPART = fgCompletaString(" ", " ", 8, False) '8
        .OBSERV4 = fgCompletaString(.DESCRIPCION_DE_CONCEPTO, " ", 30, False) '30
        .NIO = fgCompletaString(" ", " ", 24, False) '24
        .Filler2 = fgCompletaString(" ", " ", 2, False) '2
    End With

    LSet udtContabilidadeAux = udtContabilidade

    IncluirContabil "PZ", xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                    xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                    enumTipoLancamentoIntegracao.Normal, _
                    udtContabilidadeAux.String, _
                    plngDebitoCredito, _
                    0

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarRegistroPZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inclusão do registro de remessa para integração com sistema HA
Public Function IncluirContabil(ByVal pstrSiglaSistema As String, _
                                ByVal plngCodigoEmpresa As Long, _
                                ByVal pvntSequenciaOperacao As Variant, _
                                ByVal plngTipoLancamentoIntegracao As Long, _
                                ByVal pstrTextoXML As String, _
                                ByVal plngDebitoCredito As Long, _
                                ByVal plngIntegradoContabilidade As Long, _
                                Optional ByVal plngNumSeqLancamento As Long = 0) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Insert Into A8.TB_HIST_ENVI_INFO_CNTB(" & vbNewLine & _
             "  SG_SIST, " & vbNewLine & _
             "  CO_EMPR, " & vbNewLine & _
             "  NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "  TP_LANC_ITGR, " & vbNewLine & _
             "  DH_ENVI_CNTB, " & vbNewLine & _
             "  IN_LANC_DEBT_CRED," & vbNewLine & _
             "  TX_ITGR_CNTB," & vbNewLine & _
             "  IN_ITGR_CNTB, " & vbNewLine & _
             "  NR_SEQU_LANC  ) "

    strSQL = strSQL & vbNewLine & " Values (" & _
             "'" & pstrSiglaSistema & "', " & vbNewLine & _
             plngCodigoEmpresa & ", " & vbNewLine & _
             pvntSequenciaOperacao & ", " & vbNewLine & _
             plngTipoLancamentoIntegracao & ", " & vbNewLine & _
             fgDataHoraServidor_To_Oracle & ", " & vbNewLine & _
             plngDebitoCredito & ", " & vbNewLine & _
             "'" & pstrTextoXML & "', " & vbNewLine & _
             plngIntegradoContabilidade & ", " & vbNewLine & _
             plngNumSeqLancamento & ")"

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirContabil Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Exclusao do registro de remessa para integracao com sistema HA
Public Function ExcluirContabil(ByVal pvntSequenciaOperacao As Variant, _
                       Optional ByVal plngTipoLancamentoIntegracao As Long, _
                       Optional ByVal pstrSiglaSistema As String, _
                       Optional ByVal plngCodigoEmpresa As Long, _
                       Optional ByVal plngDebitoCredito As Long) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    

    strSQL = " DELETE FROM A8.TB_HIST_ENVI_INFO_CNTB " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & " " & vbNewLine

    If plngTipoLancamentoIntegracao <> 0 Then
        strSQL = strSQL & " AND    TP_LANC_ITGR = " & plngTipoLancamentoIntegracao & vbNewLine
    End If
    
    If pstrSiglaSistema <> "" Then
        strSQL = strSQL & " AND    SG_SIST = '" & pstrSiglaSistema & "'" & vbNewLine
    End If
    
    If plngCodigoEmpresa <> 0 Then
        strSQL = strSQL & " AND    CO_EMPR = " & plngCodigoEmpresa & vbNewLine
    End If
    
    If plngDebitoCredito <> 0 Then
        strSQL = strSQL & " AND    IN_LANC_DEBT_CRED = " & plngDebitoCredito & vbNewLine
    End If
    
    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExcluirContabil Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
