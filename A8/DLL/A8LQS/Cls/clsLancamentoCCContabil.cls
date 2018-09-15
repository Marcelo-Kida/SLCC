VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLancamentoCCContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'   Objeto responsável pelo controle dos lançamentos em Conta Corrente e Contabil

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                          As COMSVCSLib.ObjectContext

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

'Verificar a situação do lançamento de conta corrente no sistema slcc para permitir ou não o estorno
Private Function flVerificaEstornoLancamento(ByVal pvntNumeroSequencial As Variant) As Boolean

Dim strSQL                                  As String
Dim rsLancamento                            As ADODB.Recordset

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "        TP_LANC_ITGR,      " & vbNewLine & _
             "        CO_ULTI_SITU_PROC  " & vbNewLine & _
             " FROM   A8.TB_LANC_CC_CNTB " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & "   " & vbNewLine
    
    Set rsLancamento = fgQuerySQL(strSQL)
    
    With rsLancamento
        If .EOF And .BOF Then
            'Não há lançamentos a serem estornados
            flVerificaEstornoLancamento = False
        Else
            'Há lançamentos para a operação, verificar se eles podem ser extornados
            flVerificaEstornoLancamento = True
            .MoveFirst
            While Not .EOF
                If !CO_ULTI_SITU_PROC = enumStatusIntegracao.Integrado Or _
                   !CO_ULTI_SITU_PROC = enumStatusIntegracao.IntegradoCC Or _
                   !CO_ULTI_SITU_PROC = enumStatusIntegracao.EnviadoCC Or _
                   !CO_ULTI_SITU_PROC = enumStatusIntegracao.EnviadoCCAutomatico Then
                    'Impossível efetuar estorno, operação(ões) já possui(em) lançamento(s) de CC integrados ou enviados para integração.
                    lngCodigoErroNegocio = 3046
                    GoTo ErrorHandler
                End If
                .MoveNext
            Wend
        End If
    End With
    
    Set rsLancamento = Nothing

Exit Function
ErrorHandler:

    Set rsLancamento = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaEstornoLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'Promover o estorno do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function Estornar(ByVal pvntNumeroSequencial As Variant) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    If flVerificaEstornoLancamento(pvntNumeroSequencial) Then
        strSQL = " UPDATE A8.TB_LANC_CC_CNTB                                            " & vbNewLine & _
                 " SET    CO_ULTI_SITU_PROC = " & enumStatusIntegracao.Estornado & " , " & vbNewLine & _
                 "        CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "'                 , " & vbNewLine & _
                 "        CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "' , " & vbNewLine & _
                 "        DH_ULTI_ATLZ = SYSDATE                                        " & vbNewLine & _
                 " WHERE  NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & "              " & vbNewLine
                 
        fgExecuteSQL strSQL
    End If
    
    Estornar = True
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Estornar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a inclusão do registro de estorno do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function IncluirEstorno(ByRef pxmlLancamentoOriginal As MSXML2.DOMDocument40, _
                               ByVal plngDebitoCredito As Long) As Boolean
    
Dim strSQL                                  As String
    
    strSQL = " INSERT INTO A8.TB_LANC_CC_CNTB (                       " & vbNewLine & _
             "             NU_SEQU_OPER_ATIV,                         " & vbNewLine & _
             "             TP_LANC_ITGR,                              " & vbNewLine & _
             "             TP_BKOF,                                   " & vbNewLine & _
             "             CO_VEIC_LEGA,                              " & vbNewLine & _
             "             SG_SIST,                                   " & vbNewLine & _
             "             SG_SIST_CC,                                " & vbNewLine & _
             "             CO_EMPR,                                   " & vbNewLine & _
             "             TP_OPER,                                   " & vbNewLine & _
             "             CO_LOCA_LIQU,                              " & vbNewLine & _
             "             IN_LANC_DEBT_CRED,                         " & vbNewLine & _
             "             DT_OPER,                                   " & vbNewLine & _
             "             CO_BANC,                                   " & vbNewLine & _
             "             CO_AGEN,                                   " & vbNewLine & _
             "             NU_CC,                                     " & vbNewLine & _
             "             NU_COMD_OPER,                              " & vbNewLine & _
             "             VA_LANC_CC,                                " & vbNewLine & _
             "             NU_CTRL_ITGR,                              " & vbNewLine & _
             "             CO_ULTI_SITU_PROC,                         " & vbNewLine & _
             "             CO_USUA_ULTI_ATLZ,                         " & vbNewLine & _
             "             CO_ETCA_TRAB_ULTI_ATLZ,                    " & vbNewLine & _
             "             DH_ULTI_ATLZ                     )         " & vbNewLine
    
    With pxmlLancamentoOriginal.documentElement
    
    strSQL = strSQL & " VALUES (" & _
             " " & .selectSingleNode("NU_SEQU_OPER_ATIV").Text & "        ,   " & vbNewLine & _
             " 2                                                          ,   " & vbNewLine & _
             " " & .selectSingleNode("TP_BKOF").Text & "                  ,   " & vbNewLine & _
             IIf(.selectSingleNode("CO_VEIC_LEGA").Text = vbNullString, " NULL ", "'" & .selectSingleNode("CO_VEIC_LEGA").Text & "'") & "       ,   " & vbNewLine & _
             "'" & .selectSingleNode("SG_SIST").Text & "'                 ,   " & vbNewLine & _
             "'" & .selectSingleNode("SG_SIST_CC").Text & "'              ,   " & vbNewLine & _
             IIf(.selectSingleNode("CO_EMPR").Text = vbNullString, " NULL ", .selectSingleNode("CO_EMPR").Text) & "    , " & vbNewLine & _
             IIf(.selectSingleNode("TP_OPER").Text = vbNullString, " NULL ", .selectSingleNode("TP_OPER").Text) & "    , " & vbNewLine & _
             .selectSingleNode("CO_LOCA_LIQU").Text & "                   ,   " & vbNewLine & _
             plngDebitoCredito & "              ,   " & vbNewLine & _
             fgDtXML_To_Oracle(.selectSingleNode("DT_OPER").Text) & "    ,   " & vbNewLine & _
             .selectSingleNode("CO_BANC").Text & "                        ,   " & vbNewLine & _
             .selectSingleNode("CO_AGEN").Text & "                        ,   " & vbNewLine & _
             .selectSingleNode("NU_CC").Text & "                          ,   " & vbNewLine & _
             IIf(.selectSingleNode("NU_COMD_OPER").Text = vbNullString, " NULL ", .selectSingleNode("NU_COMD_OPER").Text) & " , " & vbNewLine & _
             .selectSingleNode("VA_LANC_CC").Text & "                     ,   " & vbNewLine & _
             IIf(.selectSingleNode("NU_CTRL_ITGR").Text = vbNullString, " NULL ", "'" & .selectSingleNode("NU_CTRL_ITGR").Text & "'") & " , " & vbNewLine & _
             .selectSingleNode("CO_ULTI_SITU_PROC").Text & "              ,    " & vbNewLine & _
             "'" & fgUsuarioRede & "'                                     ,    " & vbNewLine & _
             "'" & fgEstacaoTrabalhoUsuario & "'                          ,    " & vbNewLine & _
             " SYSDATE) "


    End With
    
    IncluirEstorno = fgExecuteSQL(strSQL)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirEstorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a inclusão do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function Incluir(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                        Optional ByVal plngNumSeqLancamento As Long = 0) As Boolean

Dim objHistoricoLancamento                  As A8LQS.clsHistSituacaoIntegracao
Dim objContaCorrente                        As A8LQS.clsContaCorrente

Dim strSQL                                  As String
Dim strDataOperacao                         As String
Dim lngCodigoErro                           As Long
Dim strMernsagemErro                        As String

Dim lngComandoOperacao                      As Long
Dim lngDebitoCredito                        As Long
Dim lngTipoBackOffice                       As Long

Dim vntPasso                                As Variant
Dim strSubTipoAtiv                          As String
Dim strFinalidadeTED                        As String

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim strParametroGeral                       As String
Dim blnValidarFeriadoLocal                  As Boolean
Dim blnLoteOperacoesCorretora               As Boolean

Dim lngLocalLiquidacao                      As Long
Dim intTipoContraparte                      As Integer
Dim intTipoLiquidacao                       As Integer
Dim intTipoDebitoCreditoOriginal            As Integer
Dim intCanalVenda                           As Integer
Dim strProduto                              As String
Dim lngFinalidadeTED                        As Long
Dim lngTipoOperacao                         As Long
Dim lngLayoutEntrada                        As Long

Dim strDataAtual                            As String
Dim intTipoLancIntegracao                   As Integer
Dim intSequencialLancamento                 As Integer
Dim objRegra                                As A8LQS.clsRegra
    
On Error GoTo ErrorHandler

    blnValidarFeriadoLocal = True
    strDataAtual = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    
    If Not xmlOperacao.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        lngLocalLiquidacao = Val(xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(xmlOperacao.selectSingleNode("//TP_CNPT").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intTipoLiquidacao = Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        intTipoDebitoCreditoOriginal = Val(xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val("0" & xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text)
    End If

    strProduto = "0"
    If Not xmlOperacao.selectSingleNode("//CD_SUB_PROD") Is Nothing Then
        strProduto = xmlOperacao.selectSingleNode("//CD_SUB_PROD").Text
    End If

    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        lngLayoutEntrada = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
    End If
    
    intSequencialLancamento = 0
    If Not xmlOperacao.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
        intSequencialLancamento = Val(xmlOperacao.selectSingleNode("//NR_SEQU_LANC").Text)
    End If
    
    lngTipoBackOffice = 0
    lngDebitoCredito = ObterTipoDebitoCreditoCC(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text, _
                                                lngTipoBackOffice, _
                                                lngTipoOperacao)
        
    If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        strDataOperacao = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text
    Else

        vntPasso = 10
        
        If lngTipoOperacao = enumTipoOperacaoLQS.EventosCETIP Or _
           lngTipoOperacao = enumTipoOperacaoLQS.EventosJurosSWAP Or _
           lngTipoOperacao = enumTipoOperacaoLQS.EventosJurosTERMO Then
            
            If fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text) > fgDataHoraServidor(enumFormatoDataHora.Data) Then
                If intTipoContraparte = enumTipoContraparte.Cliente1 And _
                   intTipoLiquidacao = enumTipoLiquidacao.SemModalidade Then
                    xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text = xmlOperacao.documentElement.selectSingleNode("//DT_MESG").Text
                End If
            End If
        
        End If

        strParametroGeral = fgSelectVarchar4000(0, False)
        Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlParametroGeral.loadXML(strParametroGeral)
    
        If Not xmlParametroGeral.selectSingleNode("//FERIADO_LOCAL_FUNDOS_PROPRIOS") Is Nothing Then
            If xmlParametroGeral.selectSingleNode("//VALIDAR").Text = "N" Then
                If lngTipoBackOffice = enumTipoBackOffice.FundosProprios Then
                    blnValidarFeriadoLocal = False
                End If
            End If
        End If
        
        vntPasso = 20
        If blnValidarFeriadoLocal Then
            strDataOperacao = fgDt_To_Xml(fgValidaDataUtil(33, _
                                                           Val(xmlOperacao.documentElement.selectSingleNode("//CO_AGEN").Text), _
                                                           fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text), _
                                                           lngCodigoErro, _
                                                           strMernsagemErro))
        Else
            strDataOperacao = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text
        End If
        
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        lngComandoOperacao = 0
    Else
        If IsNumeric(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text) Then
            lngComandoOperacao = Mid(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text, 1, 8)
        Else
            lngComandoOperacao = 0
        End If
    End If
    
    strSubTipoAtiv = "0"
    lngFinalidadeTED = 0
    
    If lngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
            strSubTipoAtiv = xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV").Text
        End If
    
    ElseIf lngLayoutEntrada = enumTipoMensagemBUS.EnvioTEDClientes Then
        If Not xmlOperacao.documentElement.selectSingleNode("CO_FLND_TED") Is Nothing Then
            lngFinalidadeTED = Val("0" & xmlOperacao.selectSingleNode("//CO_FLND_TED").Text)
        End If
    
    End If
    
    If lngTipoOperacao = enumTipoOperacaoLQS.EventosCETIP Then
        If intTipoContraparte = enumTipoContraparte.Cliente1 Then
            lngDebitoCredito = enumTipoDebitoCredito.Credito
        End If
    End If
        
    If lngLayoutEntrada = enumTipoMensagemBUS.EnvioTEDClientes Then
        lngDebitoCredito = enumTipoDebitoCredito.Debito
    ElseIf lngLayoutEntrada = enumTipoMensagemBUS.EnvioPagDespesas Then
        lngDebitoCredito = enumTipoDebitoCredito.Debito
    End If

    'Despesas CETIP não entram nesta regra.
    'RATS 776
    'Cassiano - 23/07/2008
    If lngLocalLiquidacao = enumLocalLiquidacao.CETIP And lngLayoutEntrada <> enumTipoMensagemBUS.DespesasCETIP Then
        
        If intTipoContraparte = enumTipoContraparte.Cliente1 And intTipoLiquidacao = enumTipoLiquidacao.SemModalidade Then
            lngDebitoCredito = intTipoDebitoCreditoOriginal
        Else
            lngDebitoCredito = IIf(intTipoDebitoCreditoOriginal = enumTipoDebitoCredito.Credito, _
                                                                  enumTipoDebitoCredito.Debito, _
                                                                  enumTipoDebitoCredito.Credito)
        End If
    
    End If
    
    intTipoLancIntegracao = enumTipoLancamentoIntegracao.Normal
    If Not xmlOperacao.documentElement.selectSingleNode("TP_LANC_ITGR") Is Nothing Then
        intTipoLancIntegracao = Val(xmlOperacao.documentElement.selectSingleNode("TP_LANC_ITGR").Text)
    End If
    
    If intTipoLancIntegracao = enumTipoLancamentoIntegracao.Estorno Then
        lngDebitoCredito = IIf(lngDebitoCredito = enumTipoDebitoCredito.Credito, _
                                                  enumTipoDebitoCredito.Debito, _
                                                  enumTipoDebitoCredito.Credito)
    End If
        
    If Not xmlOperacao.documentElement.selectSingleNode("ID_LOTE") Is Nothing Then
         If xmlOperacao.documentElement.selectSingleNode("ID_LOTE").Text <> vbNullString Then
            blnLoteOperacoesCorretora = True
         End If
    End If
        
    vntPasso = 40
    
    strSQL = " Insert Into A8.TB_LANC_CC_CNTB (" & vbNewLine & _
             "  NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "  TP_LANC_ITGR, " & vbNewLine & _
             "  TP_BKOF, " & vbNewLine & _
             "  CO_EMPR, " & vbNewLine & _
             "  CO_VEIC_LEGA, " & vbNewLine & _
             "  SG_SIST, " & vbNewLine & _
             "  SG_SIST_CC, " & vbNewLine & _
             "  TP_OPER, " & vbNewLine & _
             "  CO_LOCA_LIQU, " & vbNewLine & _
             "  IN_LANC_DEBT_CRED, " & vbNewLine & _
             "  DT_OPER, " & vbNewLine

    strSQL = strSQL & _
             "  CO_BANC, " & vbNewLine & _
             "  CO_AGEN, " & vbNewLine & _
             "  NU_CC, " & vbNewLine & _
             "  NU_COMD_OPER, " & vbNewLine & _
             "  VA_LANC_CC, " & vbNewLine & _
             "  NU_CTRL_ITGR, " & vbNewLine & _
             "  CO_ULTI_SITU_PROC, " & vbNewLine & _
             "  CO_SUB_TIPO_ATIV, " & vbNewLine & _
             "  TP_CNAL_VEND , " & vbNewLine & _
             "  DH_ULTI_ATLZ, " & vbNewLine & _
             "  CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             "  CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "  CD_FIND_TED, " & vbNewLine & _
             "  CD_SUB_PROD, " & vbNewLine & _
             "  NR_SEQU_LANC) " & vbNewLine
    
    strSQL = strSQL & " Values (" & _
             xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text & ", " & vbNewLine & _
             intTipoLancIntegracao & ", " & vbNewLine & _
             lngTipoBackOffice & ", " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text & ", " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text & "', " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "', " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("SG_SIST_CC").Text & "', " & vbNewLine & _
             lngTipoOperacao & ", " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text & ", " & vbNewLine & _
             lngDebitoCredito & ", " & vbNewLine & _
             fgDtXML_To_Oracle(strDataOperacao) & ", " & vbNewLine

    strSQL = strSQL & _
             fgVlrToDBServer(xmlOperacao.documentElement.selectSingleNode("CO_BANC").Text) & ", " & vbNewLine & _
             fgVlrToDBServer(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text) & ", " & vbNewLine & _
             fgVlrToDBServer(xmlOperacao.documentElement.selectSingleNode("NU_CC").Text) & ", " & vbNewLine & _
             IIf(lngComandoOperacao = 0, " Null ", lngComandoOperacao) & ", " & vbCrLf & _
             fgVlrToDBServer(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text) & ", " & vbNewLine & _
             " Null," & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC").Text & ", " & vbNewLine & _
             "'" & strSubTipoAtiv & "', " & vbNewLine & _
             intCanalVenda & ", " & vbNewLine & _
             " Sysdate, " & vbNewLine & _
             "'" & fgUsuarioRede & "', " & vbNewLine & _
             "'" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             lngFinalidadeTED & ", " & vbNewLine & _
             "'" & strProduto & "', " & vbNewLine & _
             intSequencialLancamento & ")"

    fgExecuteSQL strSQL

    vntPasso = 50
    Set objHistoricoLancamento = CreateObject("A8LQS.clsHistSituacaoIntegracao")
    objHistoricoLancamento.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                   intTipoLancIntegracao, _
                                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC").Text, _
                                   lngCodigoErro, _
                                   strMernsagemErro
                                   
    Set objHistoricoLancamento = Nothing

    Set objRegra = CreateObject("A8LQS.clsRegra")
    
    'FREITAS - 12/07/2010 - Alteração que verificará se a regra de Integração Conta Corrente é automática.
    If ((lngLayoutEntrada = enumTipoMensagemBUS.EnvioTEDClientes And _
        strDataOperacao = strDataAtual) Or _
        intTipoLancIntegracao = enumTipoLancamentoIntegracao.Estorno) Or _
        objRegra.VerificarRegraAutomatica(xmlOperacao, enumFuncaoSistema.IntegracaoCC) = True Then
        
        'se a operação for proveniente de lote de corretora então a integração para o BG não deve ser feita neste momento, será feita depois de acordo com o NET do lote
        If blnLoteOperacoesCorretora = False Then
        
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            Call objContaCorrente.Integrar(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumIntegrarCCContabilidade.ContaCorrente, _
                                           False, _
                                           intTipoLancIntegracao, , xmlOperacao.selectSingleNode("//VA_LANC_CC").Text, lngDebitoCredito, , , True)
                                           
            Set objContaCorrente = Nothing
    
        End If
    
    End If
    
    Set objRegra = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Linha = " & vntPasso)
    
End Function

'Promover a inclusão do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function IncluirPagDespesas(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                   ByVal plngCreditoDebito As enumTipoDebitoCredito, _
                                   ByVal plngNumSeqLancamento As Long, _
                                   ByVal plngTipoConta As enumTipoContaLancCC, _
                                   Optional ByVal pblnValorCPMF As Boolean = False) As Boolean

Dim objHistoricoLancamento                  As A8LQS.clsHistSituacaoIntegracao
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objContabilidade                        As A8LQS.clsContabilidade
Dim objRegra                                As A8LQS.clsRegra

Dim strSQL                                  As String
Dim strDataOperacao                         As String
Dim lngCodigoErro                           As Long
Dim strMernsagemErro                        As String

Dim lngComandoOperacao                      As Long
Dim lngDebitoCredito                        As Long
Dim lngTipoBackOffice                       As Long

Dim vntPasso                                As Variant
Dim strSubTipoAtiv                          As String
Dim strFinalidadeTED                        As String

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim strParametroGeral                       As String
Dim blnValidarFeriadoLocal                  As Boolean

Dim lngLocalLiquidacao                      As Long
Dim intTipoContraparte                      As Integer
Dim intTipoLiquidacao                       As Integer
Dim intTipoDebitoCreditoOriginal            As Integer
Dim intCanalVenda                           As Integer
Dim strProduto                              As String
Dim lngFinalidadeTED                        As Long
Dim strConta                                As String
Dim strAgencia                              As String
Dim strBanco                                As String
Dim strDataAtual                            As String
    
On Error GoTo ErrorHandler

    blnValidarFeriadoLocal = True
    strDataAtual = fgDt_To_Xml(fgDataHoraServidor(DataAux))
    
    If Not xmlOperacao.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        lngLocalLiquidacao = Val(xmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(xmlOperacao.selectSingleNode("//TP_CNPT").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intTipoLiquidacao = Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        intTipoDebitoCreditoOriginal = Val(xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text)
    End If
    
    'KIDA - SGC
    If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val("0" & xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text)
    End If

    'KIDA - Produto
    strProduto = "0"
    If Not xmlOperacao.selectSingleNode("//CD_SUB_PROD") Is Nothing Then
        strProduto = IIf(xmlOperacao.selectSingleNode("//CD_SUB_PROD").Text = "", "0", xmlOperacao.selectSingleNode("//CD_SUB_PROD").Text)
    End If

    
    lngTipoBackOffice = 0
    lngDebitoCredito = ObterTipoDebitoCreditoCC(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text, _
                                                lngTipoBackOffice, _
                                                xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        
    
    If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        strDataOperacao = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text
    Else

        vntPasso = 10
        
        If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosCETIP Or _
               xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJurosSWAP Or _
               xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJurosTERMO Then
                
                If fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text) > fgDataHoraServidor(enumFormatoDataHora.Data) Then
                    If intTipoContraparte = enumTipoContraparte.Cliente1 And _
                       intTipoLiquidacao = enumTipoLiquidacao.SemModalidade Then
                        xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text = xmlOperacao.documentElement.selectSingleNode("//DT_MESG").Text
                    End If
                End If
            
            End If
        End If

        strParametroGeral = fgSelectVarchar4000(0, False)
        Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlParametroGeral.loadXML(strParametroGeral)
    
        If Not xmlParametroGeral.selectSingleNode("//FERIADO_LOCAL_FUNDOS_PROPRIOS") Is Nothing Then
            If xmlParametroGeral.selectSingleNode("//VALIDAR").Text = "N" Then
                If lngTipoBackOffice = enumTipoBackOffice.FundosProprios Then
                    blnValidarFeriadoLocal = False
                End If
            End If
        End If
        
        vntPasso = 20
        If blnValidarFeriadoLocal Then
            strDataOperacao = fgDt_To_Xml(fgValidaDataUtil(33, _
                                                           Val(xmlOperacao.documentElement.selectSingleNode("//CO_AGEN").Text), _
                                                           fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text), _
                                                           lngCodigoErro, _
                                                           strMernsagemErro))
        Else
            strDataOperacao = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text
        End If
        
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        lngComandoOperacao = 0
    Else
        If IsNumeric(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text) Then
            lngComandoOperacao = Mid(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text, 1, 8)
        Else
            lngComandoOperacao = 0
        End If
    End If
    
    strSubTipoAtiv = "0"
    
    'Joabe -Integracao MV
    lngFinalidadeTED = 0
    
    If lngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
            strSubTipoAtiv = xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV").Text
        End If
    ElseIf Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioTEDClientes Then
            If Not xmlOperacao.documentElement.selectSingleNode("CO_FLND_TED") Is Nothing Then
                lngFinalidadeTED = Val("0" & xmlOperacao.selectSingleNode("//CO_FLND_TED").Text)
            End If
        End If
    End If
    
        
   ' ByVal plngCreditoDebito As enumTipoDebitoCredito, _
    '                               ByVal plngNumSeqLancamento As Long, _
     '                              ByVal plngTipoConta As enumTipoContaLancCC, _
      '                             Optional ByVal plngSeqOrigemRecurso = 0
    
    
    If pblnValorCPMF Then
        xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_CPMF").Text
    Else
        xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text
    End If
    
    lngDebitoCredito = plngCreditoDebito
    If plngTipoConta = ContaIsenta Then
            'Debito da Conta Isenta
        strBanco = xmlOperacao.documentElement.selectSingleNode("CO_BANC_DEBT_ISEN").Text
        strAgencia = xmlOperacao.documentElement.selectSingleNode("CO_AGEN_DEBT_ISEN").Text
        strConta = xmlOperacao.documentElement.selectSingleNode("NU_CONT_DEBT_ISEN").Text
           
        
    ElseIf plngTipoConta = ContaTributada Then
        strBanco = xmlOperacao.documentElement.selectSingleNode("CO_BANC_DEBT_TRIB").Text
        strAgencia = xmlOperacao.documentElement.selectSingleNode("CO_AGEN_DEBT_TRIB").Text
        strConta = xmlOperacao.documentElement.selectSingleNode("NU_CONT_DEBT_TRIB").Text
        
    ElseIf plngTipoConta = ContaContraparte Then
        strBanco = xmlOperacao.documentElement.selectSingleNode("CO_BANC_CRED").Text
        strAgencia = xmlOperacao.documentElement.selectSingleNode("CO_AGEN_CRED").Text
        strConta = xmlOperacao.documentElement.selectSingleNode("NU_CONT_CRED").Text
        
    End If
        
    vntPasso = 40
    
    strSQL = " Insert Into A8.TB_LANC_CC_CNTB (" & vbNewLine & _
             "  NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "  TP_LANC_ITGR, " & vbNewLine & _
             "  TP_BKOF, " & vbNewLine & _
             "  CO_EMPR, " & vbNewLine & _
             "  CO_VEIC_LEGA, " & vbNewLine & _
             "  SG_SIST, " & vbNewLine & _
             "  SG_SIST_CC, " & vbNewLine & _
             "  TP_OPER, " & vbNewLine & _
             "  CO_LOCA_LIQU, " & vbNewLine & _
             "  IN_LANC_DEBT_CRED, " & vbNewLine & _
             "  DT_OPER, " & vbNewLine & _
             "  CO_BANC, " & vbNewLine & _
             "  CO_AGEN, " & vbNewLine & _
             "  NU_CC, " & vbNewLine & _
             "  NU_COMD_OPER, " & vbNewLine & _
             "  VA_LANC_CC, " & vbNewLine & _
             "  NU_CTRL_ITGR, " & vbNewLine & _
             "  CO_ULTI_SITU_PROC, " & vbNewLine & _
             "  CO_SUB_TIPO_ATIV, " & vbNewLine & _
             "  TP_CNAL_VEND , " & vbNewLine & _
             "  DH_ULTI_ATLZ, " & vbNewLine & _
             "  CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             "  CO_ETCA_TRAB_ULTI_ATLZ, CD_FIND_TED, CD_SUB_PROD , NR_SEQU_LANC) "
    strSQL = strSQL & vbNewLine & " Values (" & _
             xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text & ", " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("TP_LANC_ITGR").Text & ", " & vbNewLine & _
             lngTipoBackOffice & ", " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text & ", " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text & "', " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "', " & vbNewLine & _
             "'" & xmlOperacao.documentElement.selectSingleNode("SG_SIST_CC").Text & "', " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text & ", " & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text & ", " & vbNewLine & _
             lngDebitoCredito & ", " & vbNewLine & _
             fgDtXML_To_Oracle(strDataOperacao) & ", " & vbNewLine & _
             fgVlrToDBServer(strBanco) & ", " & vbNewLine & _
             fgVlrToDBServer(strAgencia) & ", " & vbNewLine & _
             fgVlrToDBServer(strConta) & ", " & vbNewLine & _
             IIf(lngComandoOperacao = 0, " Null ", lngComandoOperacao) & ", " & vbCrLf & _
             fgVlrToDBServer(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text) & ", " & vbNewLine & _
             " Null," & vbNewLine & _
             xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC").Text & ", " & vbNewLine & _
             "'" & strSubTipoAtiv & "', " & vbNewLine & _
             intCanalVenda & ", " & vbNewLine & _
             " Sysdate, " & vbNewLine & _
             "'" & fgUsuarioRede & "', " & vbNewLine & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & lngFinalidadeTED & ", '" & strProduto & "'," & plngNumSeqLancamento & " )"

    fgExecuteSQL strSQL

    vntPasso = 50
    Set objHistoricoLancamento = CreateObject("A8LQS.clsHistSituacaoIntegracao")
    objHistoricoLancamento.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                   xmlOperacao.documentElement.selectSingleNode("TP_LANC_ITGR").Text, _
                                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC").Text, _
                                   lngCodigoErro, _
                                   strMernsagemErro, , plngNumSeqLancamento
                                   
    Set objHistoricoLancamento = Nothing
    
    Set objRegra = CreateObject("A8LQS.clsRegra")
    
    'IVAN - 17/07/2010 - Alteração que verificará se a regra de Integração Conta Corrente é automática.
    If objRegra.VerificarRegraAutomatica(xmlOperacao, enumFuncaoSistema.IntegracaoCC) = True Then
        
        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        Call objContaCorrente.Integrar(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       enumIntegrarCCContabilidade.ContaCorrente, _
                                       False, _
                                       xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text, , , , , , True)

        Set objContaCorrente = Nothing
    
    End If
        
    Set objRegra = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Linha = " & vntPasso)
    
End Function

'Promover a inclusão da consolidação do lançamento de conta corrente
Public Function IncluirConsolidacaoLanctoCC(ByVal lngSequenciaOperacao As Long, _
                                            ByVal lngSequenciaLanctoCC As Long) As Boolean

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_CNCL_OPER_LANC_CNTA_CRRT (" & vbNewLine & _
             " NR_SEQU_OPER_ATIV, " & vbNewLine & _
             " NR_SEQU_LANC_CNTA_CRRT) "
    strSQL = strSQL & vbNewLine & " Values (" & _
             lngSequenciaOperacao & ", " & vbNewLine & _
             lngSequenciaLanctoCC & ")"

    Call fgExecuteSQL(strSQL)

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirConsolidacaoLanctoCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Promover a gravação do controle de integração de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function GravarControleIntegracao(ByVal pvntSequenciaOperacao As Variant, _
                                         ByVal plngTipoLancamentoIntegracao As Long, _
                                         ByVal pstrControleIntegracao As String, _
                                         Optional ByVal plngNumSequLancamento As Long = 0) As Boolean

Dim strSQL                                  As String
Dim xmlComposicaoNet                        As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_LANC_CC_CNTB " & vbNewLine & _
             "    Set   NU_CTRL_ITGR            =   '" & pstrControleIntegracao & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate, " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "' " & vbNewLine & _
             "  Where   NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao & vbNewLine & _
             "    And   TP_LANC_ITGR            =   " & plngTipoLancamentoIntegracao & vbNewLine & _
             "    And   NR_SEQU_LANC            =   " & plngNumSequLancamento

    fgExecuteSQL strSQL

    Set xmlComposicaoNet = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlComposicaoNet.loadXML(LerTodosComposicaoNetCorretoras(pvntSequenciaOperacao))
    
    If xmlComposicaoNet.xml <> vbNullString Then
        For Each xmlNode In xmlComposicaoNet.selectNodes("Repeat_ComposicaoNetCorretoras/*")
            If xmlNode.selectSingleNode("NR_SEQU_LANC_CNTA_CRRT").Text <> pvntSequenciaOperacao Then
                
                strSQL = " Update   A8.TB_LANC_CC_CNTB " & vbNewLine & _
                         "    Set   NU_CTRL_ITGR            =   '" & pstrControleIntegracao & "', " & vbNewLine & _
                         "          DH_ULTI_ATLZ            =   Sysdate, " & vbNewLine & _
                         "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
                         "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "' " & vbNewLine & _
                         "  Where   NU_SEQU_OPER_ATIV       =   " & xmlNode.selectSingleNode("NR_SEQU_LANC_CNTA_CRRT").Text & vbNewLine & _
                         "    And   TP_LANC_ITGR            =   " & plngTipoLancamentoIntegracao & vbNewLine & _
                         "    And   NR_SEQU_LANC            =   " & plngNumSequLancamento
            
                fgExecuteSQL strSQL
                
            End If
        Next
    End If
            
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o tipo de débito/crédito do lançamento de conta corrente
Public Function ObterTipoDebitoCreditoCC(ByVal pstrVeiculoLegal As String, _
                                         ByVal pstrSiglaSistema As String, _
                                         ByVal plngTipoDebitoCredito As Long, _
                                Optional ByRef plngTipoBackOffice As Long = 0, _
                                Optional ByRef plngTipoOperacao As Long = 0) As Long

    On Error GoTo ErrorHandler

    If plngTipoBackOffice = 0 Then
        plngTipoBackOffice = fgObterTipoBackOffice(pstrVeiculoLegal, pstrSiglaSistema)
    End If

    If plngTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
       plngTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
       plngTipoOperacao = enumTipoOperacaoLQS.EventosResgate Then
        ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Credito
        Exit Function
    End If

    If plngTipoOperacao = enumTipoOperacaoLQS.DespesasCETIP Then
        ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito
        Exit Function
    End If

    If plngTipoOperacao = enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA Or _
       plngTipoOperacao = enumTipoOperacaoLQS.OperacaoTermoInternaBMA Or _
       plngTipoOperacao = enumTipoOperacaoLQS.OperacaoCompromissadaInternaBMA Then
       
        If plngTipoDebitoCredito = enumTipoDebitoCredito.Credito Then
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito
        Else
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Credito
        End If
        
        Exit Function
    End If
    '----------------------------------------------------------------------

    If plngTipoBackOffice = enumTipoBackOffice.Tesouraria Then
        If plngTipoDebitoCredito = enumTipoDebitoCredito.Credito Then
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Credito
        Else
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito
        End If
    Else
        If plngTipoDebitoCredito = enumTipoDebitoCredito.Credito Then
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Debito
        Else
            ObterTipoDebitoCreditoCC = enumTipoDebitoCredito.Credito
        End If
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoDebitoCreditoCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Promover a alteração do status do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function AlterarStatus(ByVal pvntSequenciaOperacao As Variant, _
                              ByVal pintTipoLancamentoIntegracao As Integer, _
                              ByVal pintStatusIntegracao As enumStatusIntegracao, _
                     Optional ByVal pstrDHUltimaAtualizacao As String, _
                     Optional ByVal pstrJustificativaCancel As String, _
                     Optional ByVal pblnUtilizarData As Boolean = True, _
                     Optional ByVal plngCodigoErro As Long = 0, _
                     Optional ByVal pstrDescricaoErro As String = vbNullString, _
                     Optional ByVal plngIndicadorDebitoCredito As Long = 0, _
                     Optional ByVal plngNumSeqLancamento As Long = 0) As Boolean

Dim strSQL                                  As String
Dim objHistSituacaoIntegracao               As A8LQS.clsHistSituacaoIntegracao

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_LANC_CC_CNTB                                                                 " & vbNewLine & _
             " Set      CO_ULTI_SITU_PROC       =   " & pintStatusIntegracao & ",                          " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "',                               " & vbNewLine

    If plngIndicadorDebitoCredito <> 0 Then
        strSQL = strSQL & _
             "          IN_LANC_DEBT_CRED = " & plngIndicadorDebitoCredito & ",                            " & vbNewLine
    End If
    
    If pintStatusIntegracao = enumStatusIntegracao.ErroCC Or _
       pintStatusIntegracao = enumStatusIntegracao.ErroSaldoCC Or _
       pintStatusIntegracao = enumStatusIntegracao.ErroEstornoCC Then
        strSQL = strSQL & _
             "          DH_ULTI_REPR =  sysdate , " & vbNewLine & _
             "          QT_REPR_CNTA_CRRT =  NVL(QT_REPR_CNTA_CRRT,0) + 1 ,  " & vbNewLine '

    End If
        
    strSQL = strSQL & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "',                    " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate                                                " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao & "                          " & vbNewLine & _
             " And      TP_LANC_ITGR            =   " & pintTipoLancamentoIntegracao & "                   " & vbNewLine & _
             " And      NR_SEQU_LANC            =   " & plngNumSeqLancamento & "  "

    If pblnUtilizarData Then
        strSQL = strSQL & " And      DH_ULTI_ATLZ            =   " & fgDtHrXML_To_Oracle(pstrDHUltimaAtualizacao)
    End If

    If fgExecuteSQL(strSQL) = 0 Then
        'Lançamento já integrado por outro usuário.
        lngCodigoErroNegocio = 3040
        GoTo ErrorHandler
    End If

    'Armazenar o Histórico da Situacao da Integração
    Set objHistSituacaoIntegracao = CreateObject("A8LQS.clsHistSituacaoIntegracao")
    Call objHistSituacaoIntegracao.Incluir(pvntSequenciaOperacao, _
                                           pintTipoLancamentoIntegracao, _
                                           pintStatusIntegracao, _
                                           plngCodigoErro, _
                                           pstrDescricaoErro, _
                                           pstrJustificativaCancel, plngNumSeqLancamento)
    Set objHistSituacaoIntegracao = Nothing

    Exit Function

ErrorHandler:
    Set objHistSituacaoIntegracao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o detalhe do lançamento de conta corrente
Public Function ObterDetalheLancamento(ByVal pstrFiltroXML As String) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlFiltro                               As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim strSQLAux                               As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strOperador                             As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objParamHistCC                          As A8LQS.clsParametroHistoricoCC
Dim xmlLancamento                           As MSXML2.DOMDocument40
Dim strParmCC                               As String
Dim xmlParmCC                               As MSXML2.DOMDocument40
Dim strRetorno                              As String

Dim intTipoBackOffice                       As Integer
Dim lngCodigoEmpresa                        As Long
Dim strSiglaSistema                         As String

    On Error GoTo ErrorHandler

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlFiltro.loadXML(pstrFiltroXML)

    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If


    'Filtro por Tipo Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER       IN (" & strValores & ")"
    End If

    'Filtro por Data
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & "AND A.DT_OPER between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                         objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode

    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND     A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      B.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Veiculo Legal / Sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If

    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.SG_SIST      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If

    'KIDA - SGC
    'Filtro Grupo Canal Venda
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
    End If

    '---------------------------------------------------------------------------------------------------------

    intTipoBackOffice = 99
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strWhereSegregacao = fgSegregaDados("A8.TB_LANC_CC_CNTB", _
                                                False, "A", "B", False, False, True, True, True, False)
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    strSQL = " SELECT   /*+ RULE */  'A8' AS OWNER, " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV , " & vbNewLine & _
             "          NULL AS NR_SEQU_OPER_ATIV,                                         " & vbNewLine & _
             "          NULL AS NR_SEQU_LANC_CNTA_CRRT,                                    " & vbNewLine & _
             "          A.TP_LANC_ITGR,                                                    " & vbNewLine & _
             "          DECODE(A.TP_LANC_ITGR,                                             " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Normal & ", 'Normal',          " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Estorno & ", 'Estorno')        " & vbNewLine & _
             "                 DE_LANC_ITGR,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.SG_SIST_CC,                                                      " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          A.DT_OPER,                                                         " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          C.CO_SITU_PROC,                                                    " & vbNewLine & _
             "          DECODE(A.IN_LANC_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoCredito & ", 'Estorno Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoDebito & ",  'Estorno Débito' ) " & _
                                              "IN_LANC_DEBT_CRED, " & vbNewLine

    strSQL = strSQL & _
             "          A.IN_LANC_DEBT_CRED AS     CO_IN_LANC_DEBT_CRED,                   " & vbNewLine & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.CO_BANC,                                                         " & vbNewLine & _
             "          A.CO_AGEN,                                                         " & vbNewLine & _
             "          A.NU_CC,                                                           " & vbNewLine & _
             "          A.VA_LANC_CC,                                                      " & vbNewLine & _
             "          A.VA_LANC_CC        AS VA_LANC_CC_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          NVL(E.NO_ISPB, 'Inexistente')   NO_ISPB,                           " & vbNewLine & _
             "          F.CO_HIST_CC,                                                      " & vbNewLine & _
             "          A.CO_SUB_TIPO_ATIV, A.CD_FIND_TED,                                 " & vbNewLine & _
             "          G.CO_CNTA_DEBT,                                                    " & vbNewLine & _
             "          G.CO_CNTA_CRED,                                                    " & vbNewLine & _
             "          G.CO_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          G.DE_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          H.NO_SIST,                                                          " & vbNewLine & _
             "          H.SG_SIST,                                                         " & vbNewLine & _
             "          A.TP_BKOF,                                                         " & vbNewLine & _
             "          I.DE_BKOF,                                                         " & vbNewLine
    
    '>>> Para a tabela << A8.TB_INSTITUICAO_SPB >> traz apenas os registros com código de compensação válidos

    strSQL = strSQL & _
             "          J.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          J.NO_CNPT,                                                         " & vbNewLine & _
             "          F.CD_SUB_PROD,                                                     " & vbNewLine & _
             "          J.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          DH_ULTI_REPR as DH_ULTI_REPR_CC ,                                  " & vbNewLine & _
             "          NVL(QT_REPR_CNTA_CRRT,0) as  NU_TENT_REPR_CC,                      " & vbNewLine & _
             "          A.NR_SEQU_LANC, J.CD_LOTE                                          " & vbNewLine & _
             " FROM     A8.TB_LANC_CC_CNTB              A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "         (SELECT  CO_CPEN,                                                   " & vbNewLine & _
             "                  NO_ISPB                                                    " & vbNewLine & _
             "          FROM    A8.TB_INSTITUICAO_SPB                                      " & vbNewLine & _
             "          WHERE   CO_CPEN <> 0)           E,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CC              F,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CNTA_CNTB       G,                                 " & vbNewLine & _
             "          A7.TB_SIST                      H,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 I,                                 " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 J                                  " & vbNewLine
             
    strSQL = strSQL & _
             " WHERE    A.NU_SEQU_OPER_ATIV     = J.NU_SEQU_OPER_ATIV                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_EMPR               = H.CO_EMPR                                " & vbNewLine & _
             " AND      A.SG_SIST               = H.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                " & vbNewLine & _
             " AND      A.CO_BANC               = E.CO_CPEN            (+)                 " & vbNewLine & _
             " AND      A.CO_EMPR               = F.CO_EMPR            (+)                 " & vbNewLine & _
             " AND      A.SG_SIST_CC            = F.SG_SIST            (+)                 " & vbNewLine & _
             " AND      A.TP_BKOF               = F.TP_BKOF            (+)                 " & vbNewLine & _
             " AND      A.TP_OPER               = F.TP_OPER            (+)                 " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = F.IN_LANC_DEBT_CRED  (+)                 " & vbNewLine & _
             " AND      A.CO_SUB_TIPO_ATIV      = F.CO_SUB_TIPO_ATIV   (+)                 " & vbNewLine & _
             " AND      A.CD_FIND_TED           = F.CD_FIND_TED        (+)                 " & vbNewLine & _
             " AND      A.CD_SUB_PROD           = F.CD_SUB_PROD        (+)                 " & vbNewLine & _
             " AND      A.CO_EMPR               = G.CO_EMPR            (+)                 " & vbNewLine & _
             " AND      A.TP_BKOF               = G.TP_BKOF            (+)                 " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = G.CO_LOCA_LIQU       (+)                 " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = G.IN_LANC_DEBT_CRED  (+)                 " & vbNewLine & _
             " AND      A.SG_SIST               = G.SG_SIST            (+)                 " & vbNewLine & _
             " AND      A.TP_OPER               = G.TP_OPER            (+)                 " & vbNewLine & _
             " AND      B.TP_BKOF               = I.TP_BKOF                                " & vbNewLine

    strSQL = strSQL & _
             " AND      B.TP_BKOF               = I.TP_BKOF                                " & vbNewLine & _
             " AND      D.TP_MESG_RECB_INTE     <> '50'                                      " & vbNewLine

    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If

    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If

    strSQL = strSQL & vbNewLine & " UNION ALL " & vbNewLine

    strSQL = strSQL & " SELECT   /*+ RULE */ 'A8' AS OWNER, " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV , " & vbNewLine & _
             "          J.NR_SEQU_OPER_ATIV,                                               " & vbNewLine & _
             "          J.NR_SEQU_LANC_CNTA_CRRT,                                          " & vbNewLine & _
             "          A.TP_LANC_ITGR,                                                    " & vbNewLine & _
             "          DECODE(A.TP_LANC_ITGR,                                             " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Normal & ", 'Normal',          " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Estorno & ", 'Estorno')        " & vbNewLine & _
             "                 DE_LANC_ITGR,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.SG_SIST_CC,                                                      " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          A.DT_OPER,                                                         " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          C.CO_SITU_PROC,                                                    " & vbNewLine & _
             "          DECODE(A.IN_LANC_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoCredito & ", 'Estorno Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoDebito & ",  'Estorno Débito' ) " & _
                                              "IN_LANC_DEBT_CRED, " & vbNewLine

    strSQL = strSQL & _
             "          A.IN_LANC_DEBT_CRED AS     CO_IN_LANC_DEBT_CRED,                   " & vbNewLine & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.CO_BANC,                                                         " & vbNewLine & _
             "          A.CO_AGEN,                                                         " & vbNewLine & _
             "          A.NU_CC,                                                           " & vbNewLine & _
             "          A.VA_LANC_CC,                                                      " & vbNewLine & _
             "          A.VA_LANC_CC        AS VA_LANC_CC_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          NVL(E.NO_ISPB, 'Inexistente')   NO_ISPB,                           " & vbNewLine & _
             "          F.CO_HIST_CC,                                                      " & vbNewLine & _
             "          F.CO_SUB_TIPO_ATIV, F.CD_FIND_TED,                                  " & vbNewLine & _
             "          G.CO_CNTA_DEBT,                                                    " & vbNewLine & _
             "          G.CO_CNTA_CRED,                                                    " & vbNewLine & _
             "          G.CO_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          G.DE_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          H.NO_SIST,                                                          " & vbNewLine & _
             "          H.SG_SIST,                                                         " & vbNewLine & _
             "          A.TP_BKOF,                                                         " & vbNewLine & _
             "          I.DE_BKOF,                                                         " & vbNewLine
    
    '>>> Para a tabela << A8.TB_INSTITUICAO_SPB >> traz apenas os registros com código de compensação válidos

    strSQL = strSQL & _
             "          A.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          K.NO_CNPT,                                                         " & vbNewLine & _
             "          F.CD_SUB_PROD,                                                     " & vbNewLine & _
             "          K.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          DH_ULTI_REPR as DH_ULTI_REPR_CC ,                                  " & vbNewLine & _
             "          NVL(QT_REPR_CNTA_CRRT,0) as  NU_TENT_REPR_CC,                      " & vbNewLine & _
             "          A.NR_SEQU_LANC, K.CD_LOTE                                          " & vbNewLine & _
             " FROM     A8.TB_LANC_CC_CNTB              A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "         (SELECT  CO_CPEN,                                                   " & vbNewLine & _
             "                  NO_ISPB                                                    " & vbNewLine & _
             "          FROM    A8.TB_INSTITUICAO_SPB                                      " & vbNewLine & _
             "          WHERE   CO_CPEN <> 0)           E,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CC              F,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CNTA_CNTB       G,                                 " & vbNewLine & _
             "          A7.TB_SIST                      H,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 I,                                 " & vbNewLine & _
             "          A8.TB_CNCL_OPER_LANC_CNTA_CRRT  J,                                  " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 K                                  " & vbNewLine

    strSQL = strSQL & _
             " WHERE    A.NU_SEQU_OPER_ATIV     = K.NU_SEQU_OPER_ATIV                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_EMPR               = H.CO_EMPR                                " & vbNewLine & _
             " AND      A.SG_SIST               = H.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                " & vbNewLine & _
             " AND      A.CO_BANC               = E.CO_CPEN                (+)             " & vbNewLine & _
             " AND      A.CO_EMPR               = F.CO_EMPR                (+)             " & vbNewLine & _
             " AND      A.SG_SIST_CC            = F.SG_SIST                (+)             " & vbNewLine & _
             " AND      A.TP_BKOF               = F.TP_BKOF                (+)             " & vbNewLine & _
             " AND      A.TP_OPER               = F.TP_OPER                (+)             " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = F.IN_LANC_DEBT_CRED      (+)             " & vbNewLine & _
             " AND      A.CO_SUB_TIPO_ATIV      = F.CO_SUB_TIPO_ATIV       (+)             " & vbNewLine & _
             " AND      A.CD_FIND_TED           = F.CD_FIND_TED            (+)             " & vbNewLine & _
             " AND      A.CD_SUB_PROD           = F.CD_SUB_PROD            (+)             " & vbNewLine & _
             " AND      A.CO_EMPR               = G.CO_EMPR                (+)             " & vbNewLine & _
             " AND      A.TP_BKOF               = G.TP_BKOF                (+)             " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = G.CO_LOCA_LIQU           (+)             " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = G.IN_LANC_DEBT_CRED      (+)             " & vbNewLine & _
             " AND      A.SG_SIST               = G.SG_SIST                (+)             " & vbNewLine & _
             " AND      A.TP_OPER               = G.TP_OPER                (+)             " & vbNewLine
    
    strSQL = strSQL & _
             " AND      B.TP_BKOF               = I.TP_BKOF                                " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV     = J.NR_SEQU_LANC_CNTA_CRRT (+)             " & vbNewLine & _
             " AND      D.TP_MESG_RECB_INTE     = '50'                                       " & vbNewLine

    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If

    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If

    'Verificação da Base histórica
    If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            
            strSQLAux = Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER")
            strSQLAux = Replace(strSQLAux, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
            strSQLAux = Replace(strSQLAux, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
            
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     strSQLAux
        End If
    Else
        strSQL = Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER")
        strSQL = Replace(strSQL, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
        strSQL = Replace(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
    End If
    
    ObterDetalheLancamento = fgQueryXMLLerTodos("DetalheLancamento", strSQL, App.EXEName & "." & TypeName(Me))
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlFiltro = Nothing

    Exit Function

ErrorHandler:
    Set xmlParmCC = Nothing
    Set xmlLancamento = Nothing
    Set objParamHistCC = Nothing
    Set xmlFiltro = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o detalhe do lançamento de conta corrente
Public Function ObterConsultaLancamento(ByVal pstrFiltroXML As String) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlFiltro                               As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim strSQLAux                               As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strOperador                             As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objParamHistCC                          As A8LQS.clsParametroHistoricoCC
Dim xmlLancamento                           As MSXML2.DOMDocument40
Dim strParmCC                               As String
Dim xmlParmCC                               As MSXML2.DOMDocument40
Dim strRetorno                              As String

Dim intTipoBackOffice                       As Integer
Dim lngCodigoEmpresa                        As Long
Dim strSiglaSistema                         As String

    On Error GoTo ErrorHandler

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlFiltro.loadXML(pstrFiltroXML)

    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If


    'Filtro por Tipo Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_OPER       IN (" & strValores & ")"
    End If

    'Filtro por Data
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & "AND A.DT_OPER between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                         objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode

    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND     A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      B.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Veiculo Legal / Sistema
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_VEIC_LEGA      IN (" & strValores & ")" & vbCrLf
    End If

    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.SG_SIST      IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If

    'KIDA - SGC
    'Filtro Grupo Canal Venda
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
    End If

    '---------------------------------------------------------------------------------------------------------

    intTipoBackOffice = 99
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strWhereSegregacao = fgSegregaDados("A8.TB_LANC_CC_CNTB", _
                                                False, "A", "B", False, False, True, True, True, False)
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND B.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    strSQL = " SELECT   /*+ RULE */ DISTINCT A.NU_SEQU_OPER_ATIV,  'A8' AS OWNER, " & vbNewLine & _
             "          NULL AS NR_SEQU_OPER_ATIV,                                         " & vbNewLine & _
             "          NULL AS NR_SEQU_LANC_CNTA_CRRT,                                    " & vbNewLine & _
             "          A.TP_LANC_ITGR,                                                    " & vbNewLine & _
             "          DECODE(A.TP_LANC_ITGR,                                             " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Normal & ", 'Normal',          " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Estorno & ", 'Estorno')        " & vbNewLine & _
             "                 DE_LANC_ITGR,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.SG_SIST_CC,                                                      " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          A.DT_OPER,                                                         " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          C.CO_SITU_PROC,                                                    " & vbNewLine & _
             "          DECODE(A.IN_LANC_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoCredito & ", 'Estorno Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoDebito & ",  'Estorno Débito' ) " & _
                                              "IN_LANC_DEBT_CRED, " & vbNewLine

    strSQL = strSQL & _
             "          A.IN_LANC_DEBT_CRED AS     CO_IN_LANC_DEBT_CRED,                   " & vbNewLine & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.CO_BANC,                                                         " & vbNewLine & _
             "          A.CO_AGEN,                                                         " & vbNewLine & _
             "          A.NU_CC,                                                           " & vbNewLine & _
             "          A.VA_LANC_CC,                                                      " & vbNewLine & _
             "          A.VA_LANC_CC        AS VA_LANC_CC_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          NVL(E.NO_ISPB, 'Inexistente')   NO_ISPB,                           " & vbNewLine & _
             "          F.CO_HIST_CC,                                                      " & vbNewLine & _
             "          A.CO_SUB_TIPO_ATIV, A.CD_FIND_TED,                                 " & vbNewLine & _
             "          G.CO_CNTA_DEBT,                                                    " & vbNewLine & _
             "          G.CO_CNTA_CRED,                                                    " & vbNewLine & _
             "          G.CO_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          G.DE_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          H.NO_SIST,                                                          " & vbNewLine & _
             "          H.SG_SIST,                                                         " & vbNewLine & _
             "          A.TP_BKOF,                                                         " & vbNewLine & _
             "          I.DE_BKOF,                                                         " & vbNewLine
    
    '>>> Para a tabela << A8.TB_INSTITUICAO_SPB >> traz apenas os registros com código de compensação válidos

    strSQL = strSQL & _
             "          J.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          J.NO_CNPT,                                                         " & vbNewLine & _
             "          F.CD_SUB_PROD,                                                     " & vbNewLine & _
             "          J.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          DH_ULTI_REPR as DH_ULTI_REPR_CC ,                                  " & vbNewLine & _
             "          NVL(QT_REPR_CNTA_CRRT,0) as  NU_TENT_REPR_CC,                      " & vbNewLine & _
             "          A.NR_SEQU_LANC,                                                    " & vbNewLine & _
             "          K.DH_SITU_LANC_CC, J.CD_LOTE                                       " & vbNewLine & _
             " FROM     A8.TB_LANC_CC_CNTB              A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "         (SELECT  CO_CPEN,                                                   " & vbNewLine & _
             "                  NO_ISPB                                                    " & vbNewLine & _
             "          FROM    A8.TB_INSTITUICAO_SPB                                      " & vbNewLine & _
             "          WHERE   CO_CPEN <> 0)           E,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CC              F,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CNTA_CNTB       G,                                 " & vbNewLine & _
             "          A7.TB_SIST                      H,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 I,                                 " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 J,                                 " & vbNewLine & _
             "          A8.TB_HIST_SITU_LANC_CC_CNTB    K                                  " & vbNewLine
             
    strSQL = strSQL & _
             " WHERE    A.NU_SEQU_OPER_ATIV     = J.NU_SEQU_OPER_ATIV                      " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV     = K.NU_SEQU_OPER_ATIV (+)                  " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_EMPR               = H.CO_EMPR                                " & vbNewLine & _
             " AND      A.SG_SIST               = H.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                " & vbNewLine
    
    'KIDA - 23/04/2009
    'RATS - 903
    strSQL = strSQL & _
             " AND      A.CO_BANC               = E.CO_CPEN            (+)                        " & vbNewLine & _
             " AND      A.CO_EMPR               = F.CO_EMPR            (+)                        " & vbNewLine & _
             " AND      A.SG_SIST_CC            = F.SG_SIST            (+)                        " & vbNewLine & _
             " AND      A.TP_BKOF               = F.TP_BKOF            (+)                        " & vbNewLine & _
             " AND      A.TP_OPER               = F.TP_OPER            (+)                        " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = F.IN_LANC_DEBT_CRED  (+)                        " & vbNewLine & _
             " AND      A.CO_SUB_TIPO_ATIV      = F.CO_SUB_TIPO_ATIV   (+)                        " & vbNewLine & _
             " AND      A.CD_FIND_TED           = F.CD_FIND_TED        (+)                        " & vbNewLine & _
             " AND      A.CD_SUB_PROD           = F.CD_SUB_PROD        (+)                        " & vbNewLine & _
             " AND      A.CO_EMPR               = G.CO_EMPR            (+)                        " & vbNewLine & _
             " AND      A.TP_BKOF               = G.TP_BKOF            (+)                        " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = G.CO_LOCA_LIQU       (+)                        " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = G.IN_LANC_DEBT_CRED  (+)                        " & vbNewLine & _
             " AND      A.SG_SIST               = G.SG_SIST            (+)                        " & vbNewLine & _
             " AND      A.TP_OPER               = G.TP_OPER            (+)                        " & vbNewLine & _
             " AND      B.TP_BKOF               = I.TP_BKOF                                       " & vbNewLine & _
             " AND      D.TP_MESG_RECB_INTE     <> '50'                                             " & vbNewLine & _
             " AND      DECODE(K.NU_SEQU_OPER_ATIV, NULL, SYSDATE, K.DH_SITU_LANC_CC)       = NVL((SELECT MIN(DH_SITU_LANC_CC)                    " & vbNewLine & _
             "                                                                                     FROM   A8.TB_HIST_SITU_LANC_CC_CNTB            " & vbNewLine & _
             "                                                                                     WHERE  NU_SEQU_OPER_ATIV = A.NU_SEQU_OPER_ATIV), SYSDATE)" & vbNewLine

    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If

    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If

    strSQL = strSQL & vbNewLine & " UNION ALL " & vbNewLine

    strSQL = strSQL & " SELECT   /*+ RULE */ A.NU_SEQU_OPER_ATIV, 'A8' AS OWNER, " & vbNewLine & _
             "          J.NR_SEQU_OPER_ATIV,                                               " & vbNewLine & _
             "          J.NR_SEQU_LANC_CNTA_CRRT,                                          " & vbNewLine & _
             "          A.TP_LANC_ITGR,                                                    " & vbNewLine & _
             "          DECODE(A.TP_LANC_ITGR,                                             " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Normal & ", 'Normal',          " & vbNewLine & _
             "             " & enumTipoLancamentoIntegracao.Estorno & ", 'Estorno')        " & vbNewLine & _
             "                 DE_LANC_ITGR,                                               " & vbNewLine & _
             "          A.TP_OPER,                                                         " & vbNewLine & _
             "          A.SG_SIST_CC,                                                      " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          A.DT_OPER,                                                         " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_GRUP_VEIC_LEGA,                                               " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                    " & vbNewLine & _
             "          C.CO_SITU_PROC,                                                    " & vbNewLine & _
             "          DECODE(A.IN_LANC_DEBT_CRED,                                        " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoCredito & ", 'Estorno Crédito', " & _
                                              enumTipoDebitoCreditoEstorno.EstornoDebito & ",  'Estorno Débito' ) " & _
                                              "IN_LANC_DEBT_CRED, " & vbNewLine

    strSQL = strSQL & _
             "          A.IN_LANC_DEBT_CRED AS     CO_IN_LANC_DEBT_CRED,                   " & vbNewLine & _
             "          D.TP_MESG_RECB_INTE,                                               " & vbNewLine & _
             "          D.NO_TIPO_OPER,                                                    " & vbNewLine & _
             "          D.CO_MESG_SPB_REGT_OPER,                                           " & vbNewLine & _
             "          A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                    " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.CO_BANC,                                                         " & vbNewLine & _
             "          A.CO_AGEN,                                                         " & vbNewLine & _
             "          A.NU_CC,                                                           " & vbNewLine & _
             "          A.VA_LANC_CC,                                                      " & vbNewLine & _
             "          A.VA_LANC_CC        AS VA_LANC_CC_VLRXML /* para ficar com '.' separando os decimais, no xml */," & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          NVL(E.NO_ISPB, 'Inexistente')   NO_ISPB,                           " & vbNewLine & _
             "          F.CO_HIST_CC,                                                      " & vbNewLine & _
             "          F.CO_SUB_TIPO_ATIV, F.CD_FIND_TED,                                  " & vbNewLine & _
             "          G.CO_CNTA_DEBT,                                                    " & vbNewLine & _
             "          G.CO_CNTA_CRED,                                                    " & vbNewLine & _
             "          G.CO_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          G.DE_HIST_CNTA_CNTB,                                               " & vbNewLine & _
             "          H.NO_SIST,                                                          " & vbNewLine & _
             "          H.SG_SIST,                                                         " & vbNewLine & _
             "          A.TP_BKOF,                                                         " & vbNewLine & _
             "          I.DE_BKOF,                                                         " & vbNewLine
    
    '>>> Para a tabela << A8.TB_INSTITUICAO_SPB >> traz apenas os registros com código de compensação válidos

    strSQL = strSQL & _
             "          A.TP_CNAL_VEND,                                                    " & vbNewLine & _
             "          K.NO_CNPT,                                                         " & vbNewLine & _
             "          F.CD_SUB_PROD,                                                     " & vbNewLine & _
             "          K.CD_OPER_ETRT,                                                    " & vbNewLine & _
             "          DH_ULTI_REPR as DH_ULTI_REPR_CC ,                                  " & vbNewLine & _
             "          NVL(QT_REPR_CNTA_CRRT,0) as  NU_TENT_REPR_CC,                      " & vbNewLine & _
             "          A.NR_SEQU_LANC,                                                    " & vbNewLine & _
             "          L.DH_SITU_LANC_CC, K.CD_LOTE                                       " & vbNewLine & _
             " FROM     A8.TB_LANC_CC_CNTB              A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_SITU_PROC                 C,                                 " & vbNewLine & _
             "          A8.TB_TIPO_OPER                 D,                                 " & vbNewLine & _
             "         (SELECT  CO_CPEN,                                                   " & vbNewLine & _
             "                  NO_ISPB                                                    " & vbNewLine & _
             "          FROM    A8.TB_INSTITUICAO_SPB                                      " & vbNewLine & _
             "          WHERE   CO_CPEN <> 0)           E,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CC              F,                                 " & vbNewLine & _
             "          A8.TB_PARM_HIST_CNTA_CNTB       G,                                 " & vbNewLine & _
             "          A7.TB_SIST                      H,                                 " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 I,                                 " & vbNewLine & _
             "          A8.TB_CNCL_OPER_LANC_CNTA_CRRT  J,                                 " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 K,                                 " & vbNewLine & _
             "          A8.TB_HIST_SITU_LANC_CC_CNTB    L                                  " & vbNewLine

    strSQL = strSQL & _
             " WHERE    A.NU_SEQU_OPER_ATIV     = K.NU_SEQU_OPER_ATIV                      " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV     = L.NU_SEQU_OPER_ATIV      (+)             " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA                           " & vbNewLine & _
             " AND      A.SG_SIST               = B.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_EMPR               = H.CO_EMPR                                " & vbNewLine & _
             " AND      A.SG_SIST               = H.SG_SIST                                " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC     = C.CO_SITU_PROC                           " & vbNewLine & _
             " AND      A.TP_OPER               = D.TP_OPER                                " & vbNewLine & _
             " AND      A.CO_BANC               = E.CO_CPEN                (+)             " & vbNewLine & _
             " AND      A.CO_EMPR               = F.CO_EMPR                (+)             " & vbNewLine & _
             " AND      A.SG_SIST_CC            = F.SG_SIST                (+)             " & vbNewLine & _
             " AND      A.TP_BKOF               = F.TP_BKOF                (+)             " & vbNewLine & _
             " AND      A.TP_OPER               = F.TP_OPER                (+)             " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = F.IN_LANC_DEBT_CRED      (+)             " & vbNewLine & _
             " AND      A.CO_SUB_TIPO_ATIV      = F.CO_SUB_TIPO_ATIV       (+)             " & vbNewLine & _
             " AND      A.CD_FIND_TED           = F.CD_FIND_TED            (+)             " & vbNewLine & _
             " AND      A.CD_SUB_PROD           = F.CD_SUB_PROD            (+)             " & vbNewLine & _
             " AND      A.CO_EMPR               = G.CO_EMPR                (+)             " & vbNewLine & _
             " AND      A.TP_BKOF               = G.TP_BKOF                (+)             " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU          = G.CO_LOCA_LIQU           (+)             " & vbNewLine & _
             " AND      A.IN_LANC_DEBT_CRED     = G.IN_LANC_DEBT_CRED      (+)             " & vbNewLine & _
             " AND      A.SG_SIST               = G.SG_SIST                (+)             " & vbNewLine & _
             " AND      A.TP_OPER               = G.TP_OPER                (+)             " & vbNewLine
             
    strSQL = strSQL & _
             " AND      B.TP_BKOF               = I.TP_BKOF                                       " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV     = J.NR_SEQU_LANC_CNTA_CRRT (+)                    " & vbNewLine & _
             " AND      D.TP_MESG_RECB_INTE     = '50'                                              " & vbNewLine & _
             " AND      DECODE(L.NU_SEQU_OPER_ATIV, NULL, SYSDATE, L.DH_SITU_LANC_CC) = NVL((SELECT MIN(DH_SITU_LANC_CC)                    " & vbNewLine & _
             "                                                                               FROM   A8.TB_HIST_SITU_LANC_CC_CNTB            " & vbNewLine & _
             "                                                                               WHERE  NU_SEQU_OPER_ATIV = A.NU_SEQU_OPER_ATIV), SYSDATE)" & vbNewLine

    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If

    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacao
    End If

    'Verificação da Base histórica
    If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            
            strSQLAux = Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER")
            strSQLAux = Replace(strSQLAux, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
            strSQLAux = Replace(strSQLAux, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
            
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     strSQLAux
        End If
    Else
        strSQL = Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER")
        strSQL = Replace(strSQL, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
        strSQL = Replace(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
    End If
    
    ObterConsultaLancamento = fgQueryXMLLerTodos("DetalheLancamento", strSQL, App.EXEName & "." & TypeName(Me))
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlFiltro = Nothing

    Exit Function

ErrorHandler:
    Set xmlParmCC = Nothing
    Set xmlLancamento = Nothing
    Set objParamHistCC = Nothing
    Set xmlFiltro = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterConsultaLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os lançamento de conta corrente selecionados por status
Public Function ObterLancamentosPorStatus(ByVal pstrFiltroXML As String) As String

Dim strSQL                                  As String
Dim strOwnerTabelaSegregada                 As String

Dim xmlComplementoWhere                     As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim strValores                              As String
Dim blnOcultarQuantidades                   As Boolean

Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlFiltro.loadXML(pstrFiltroXML)

    'Verifica se é necessário exibir as quantidades
    blnOcultarQuantidades = (Not _
        xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Quantidade/OcultarQuantidade") Is Nothing)

    If Not blnOcultarQuantidades Then
        'Envia dados para a segregação ----------------------------------------------------------------------------
        Set xmlComplementoWhere = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlComplementoWhere, "", "Repeat_Where", "")
        Call fgAppendNode(xmlComplementoWhere, "Repeat_Where", "Grupo_Where", "")

        If xmlFiltro.selectNodes("Repeat_Filtros/*").length > 0 Then

            'Filtro Grupo Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
                strValores = strValores & objDomNode.Text & ", "
            Next

            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_GRUP_VEIC_LEGA      IN (" & strValores & ")")
                'Inclui o atributo ALIAS para filtro específico na tabela DOMÍNIO (Segregação de Dados)
                Call fgAppendAttribute(xmlComplementoWhere, "Where", "Alias", "TabDominio_Sub")
            End If

            'Filtro por Tipo Operação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
        
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_OPER      IN (" & strValores & ")")
            End If
        
            'Filtro por Empresa
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
                strValores = strValores & objDomNode.Text & ", "
            Next

            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If

            'Filtro Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next

            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_VEIC_LEGA      IN (" & strValores & ")")
            End If

            'Filtro Local de Liquidação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next

            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_LOCA_LIQU      IN (" & strValores & ")")
            End If

            'Filtro por Data
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", _
                                    "DT_OPER BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                                                    " AND " & objDomNode.selectSingleNode("DataFim").Text)
            Next objDomNode
        
            'KIDA - SGC
            'Filtro Grupo Canal Venda
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                'strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND      IN (" & strValores & ")")
            End If
        
        End If

        'Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "IN_DISP_CONS = " & enumIndicadorSimNao.Sim)
        
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                strOwnerTabelaSegregada = fgSegregaDados("A8.TB_LANC_CC_CNTB", _
                                                         , , , , , , , , False, xmlComplementoWhere)
            Case 0
                strOwnerTabelaSegregada = fgSegregaDados("A8.TB_LANC_CC_CNTB", _
                                                         , , , , , False, False, False, False, xmlComplementoWhere)
            Case Is > 0
                strOwnerTabelaSegregada = fgSegregaDados("A8.TB_LANC_CC_CNTB", _
                                                         , , , , , False, False, False, False, xmlComplementoWhere)
                strOwnerTabelaSegregada = strOwnerTabelaSegregada & vbNewLine & _
                                          " AND TabDominio_Sub.TP_BKOF = " & intTipoBackOffice & vbNewLine
        End Select
        
        'KIDA - 30/06/2009
        'RATS - 928
        'strOwnerTabelaSegregada = Replace(strOwnerTabelaSegregada, "TabFato_Sub.CD_FIND_TED, TabFato_Sub.CD_SUB_PROD, TabFato_Sub.CO_AGEN, TabFato_Sub.CO_BANC, TabFato_Sub.CO_EMPR, TabFato_Sub.CO_ETCA_TRAB_ULTI_ATLZ, TabFato_Sub.CO_LOCA_LIQU, TabFato_Sub.CO_SUB_TIPO_ATIV, TabFato_Sub.CO_ULTI_SITU_PROC, TabFato_Sub.CO_USUA_ULTI_ATLZ, TabFato_Sub.CO_VEIC_LEGA, TabFato_Sub.DH_ULTI_ATLZ, TabFato_Sub.DH_ULTI_REPR, TabFato_Sub.DT_OPER, TabFato_Sub.IN_LANC_DEBT_CRED, TabFato_Sub.NR_SEQU_LANC, TabFato_Sub.NU_CC, TabFato_Sub.NU_COMD_OPER, TabFato_Sub.NU_CTRL_ITGR, TabFato_Sub.NU_SEQU_OPER_ATIV, TabFato_Sub.QT_REPR_CNTA_CRRT, TabFato_Sub.SG_SIST, TabFato_Sub.SG_SIST_CC, TabFato_Sub.TP_BKOF, TabFato_Sub.TP_CNAL_VEND, TabFato_Sub.TP_LANC_ITGR, TabFato_Sub.TP_OPER, TabFato_Sub.VA_LANC_CC", "distinct TabFato_Sub.co_ulti_situ_proc,TabFato_Sub.CO_AGEN,TabFato_Sub.CO_BANC,TabFato_Sub.NU_CC,TabFato_Sub.TP_OPER,TabFato_Sub.dt_oper,TabFato_Sub.tp_bkof,TabFato_Sub.tp_lanc_itgr,TabFato_Sub.sg_sist")
        
        Set xmlComplementoWhere = Nothing
        '----------------------------------------------------------------------------------------------------------
        
        'Verificação da Base histórica
        If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            If fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
                strOwnerTabelaSegregada = strOwnerTabelaSegregada & vbCrLf & _
                         " UNION ALL " & vbCrLf & _
                         Replace(strOwnerTabelaSegregada, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
            End If
        Else
            strOwnerTabelaSegregada = Replace(strOwnerTabelaSegregada, "A8.TB_LANC_CC_CNTB", "A8HIST.TB_LANC_CC_CNTB")
        End If
        
    End If

    If blnOcultarQuantidades Then
        strSQL = " SELECT   A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.DE_SITU_PROC                                                         " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC     A                                                  " & vbNewLine & _
                 " WHERE    A.TP_SITU_PROC = " & enumTipoSituacaoSistema.LancamentoContaCorrente & " " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
    Else
        strSQL = " SELECT   A.CO_SITU_PROC,                                                             " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                             " & vbNewLine & _
                 "          A.DE_SITU_PROC,                                                             " & vbNewLine & _
                 "          COUNT(B.CO_ULTI_SITU_PROC)   AS NU_QTD                                      " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC                   A,                                        " & vbNewLine & _
                 "     (" & strOwnerTabelaSegregada & ")      B                                         " & vbNewLine & _
                 " WHERE    A.CO_SITU_PROC      = B.CO_ULTI_SITU_PROC    (+)                            " & vbNewLine & _
                 " AND      A.TP_SITU_PROC      = " & enumTipoSituacaoSistema.LancamentoContaCorrente & " " & vbNewLine & _
                 " GROUP BY A.CO_SITU_PROC,                                                             " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                             " & vbNewLine & _
                 "          A.DE_SITU_PROC                                                              " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
    End If

    ObterLancamentosPorStatus = fgQueryXMLLerTodos("SituacaoLancamentos", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlFiltro = Nothing

    Exit Function

ErrorHandler:
    Set xmlComplementoWhere = Nothing
    Set xmlFiltro = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterLancamentoPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Ler todos os registro da tabela A8.TB_LANC_CC_CNTB conforme o filtro específicado
Public Function LerTodos(Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                         Optional ByVal plngTipoLancamentoIntegracao As Variant = 0, _
                         Optional ByVal pstrNumeroControleIntegracao As String = vbNullString, _
                         Optional ByVal pintCO_EMPR As Integer = 0, _
                         Optional ByVal pstrDT_OPER_YYYYMMDD As String = vbNullString, _
                         Optional ByVal pintTP_OPER As Integer = 0, _
                         Optional ByVal pstrCO_VEIC_LEGA As String = vbNullString, _
                         Optional ByVal pintCO_LOCA_LIQU As Integer = 0, _
                         Optional ByVal pintCO_BANC As Integer = 0, _
                         Optional ByVal pintCO_AGEN As Integer = 0, _
                         Optional ByVal plngNU_CC As Long = 0, _
                         Optional ByVal pintCO_ULTI_SITU_PROC As Integer = 0, _
                         Optional ByVal pstrORDER_BY As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Select   A.NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "          A.TP_LANC_ITGR, " & vbNewLine & _
             "          A.TP_BKOF, " & vbNewLine & _
             "          A.CO_EMPR, " & vbNewLine & _
             "          A.CO_VEIC_LEGA, " & vbNewLine & _
             "          A.SG_SIST, " & vbNewLine & _
             "          A.SG_SIST_CC, " & vbNewLine & _
             "          A.TP_OPER, " & vbNewLine & _
             "          A.CO_LOCA_LIQU, " & vbNewLine & _
             "          A.IN_LANC_DEBT_CRED, " & vbNewLine & _
             "          A.DT_OPER, " & vbNewLine & _
             "          A.CO_BANC, " & vbNewLine & _
             "          A.CO_AGEN, " & vbNewLine & _
             "          A.NU_CC, " & vbNewLine & _
             "          A.NU_COMD_OPER, " & vbNewLine & _
             "          A.VA_LANC_CC, " & vbNewLine & _
             "          A.NU_CTRL_ITGR, " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC, " & vbNewLine & _
             "          A.DH_ULTI_ATLZ, " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          B.TP_MESG_RECB_INTE, " & vbNewLine & _
             "          '' as DH_ULTI_REPR_CC , " & vbNewLine & _
             "          0 AS NU_TENT_REPR_CC, " & vbNewLine & _
             "          A.NR_SEQU_LANC " & vbNewLine & " "
    '"          NVL(A.QT_REPR_CC,0) AS NU_TENT_REPR_CC, "
    strSQL = strSQL & _
             " From     A8.TB_LANC_CC_CNTB A, " & vbNewLine & _
             "          A8.TB_TIPO_OPER    B " & vbNewLine & _
             " Where    A.TP_OPER = B.TP_OPER " & vbNewLine

    If pvntSequenciaOperacao > 0 Then
        strSQL = strSQL & _
             " And      NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & vbNewLine
    End If

    If plngTipoLancamentoIntegracao > 0 Then
        strSQL = strSQL & _
             " And      TP_LANC_ITGR = " & plngTipoLancamentoIntegracao & vbNewLine
    End If

    If pstrNumeroControleIntegracao <> vbNullString Then
        strSQL = strSQL & _
             " And      NU_CTRL_ITGR = '" & pstrNumeroControleIntegracao & "'" & vbNewLine
    End If

    If pintCO_EMPR > 0 Then
        strSQL = strSQL & _
             " And      CO_EMPR = " & pintCO_EMPR & vbNewLine
    End If

    If pstrDT_OPER_YYYYMMDD <> vbNullString Then
        strSQL = strSQL & _
             " And      DT_OPER = " & fgDtXML_To_Oracle(pstrDT_OPER_YYYYMMDD) & vbNewLine
    End If

    If pintTP_OPER > 0 Then
        strSQL = strSQL & _
             " And      TP_OPER = " & pintTP_OPER & vbNewLine
    End If

    If pstrCO_VEIC_LEGA <> vbNullString Then
        strSQL = strSQL & _
             " And      CO_VEIC_LEGA = '" & pstrCO_VEIC_LEGA & "'" & vbNewLine
    End If

    If pintCO_LOCA_LIQU > 0 Then
        strSQL = strSQL & _
             " And      CO_LOCA_LIQU = " & pintCO_LOCA_LIQU & vbNewLine
    End If

    If pintCO_BANC > 0 Then
        strSQL = strSQL & _
             " And      CO_BANC = " & pintCO_BANC & vbNewLine
    End If

    If pintCO_AGEN > 0 Then
        strSQL = strSQL & _
             " And      CO_AGEN = " & pintCO_AGEN & vbNewLine
    End If

    If plngNU_CC > 0 Then
        strSQL = strSQL & _
             " And      NU_CC = " & plngNU_CC & vbNewLine
    End If

    If pintCO_ULTI_SITU_PROC > 0 Then
        strSQL = strSQL & _
             " And      CO_ULTI_SITU_PROC = " & pintCO_ULTI_SITU_PROC & vbNewLine
    End If
    
    If pstrORDER_BY <> vbNullString Then
        strSQL = strSQL & _
             " ORDER BY " & pstrORDER_BY & vbNewLine
    End If

    LerTodos = fgQueryXMLLerTodos("LancamentoCCContabil", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Ler todos os registros da tabela A8.TB_CNCL_OPER_LANC_CNTA_CRRT conforme o filtro específicado
Public Function LerTodosComposicaoNetCorretoras(Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                                                Optional ByVal pvntSequenciaLanctoCC As Variant = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Select   NR_SEQU_OPER_ATIV, " & vbNewLine & _
             "          NR_SEQU_LANC_CNTA_CRRT " & vbNewLine & _
             " From     A8.TB_CNCL_OPER_LANC_CNTA_CRRT " & vbNewLine & _
             " Where    1 = 1 " & vbNewLine

    If pvntSequenciaOperacao > 0 Then
        strSQL = strSQL & _
             " And      NR_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & vbNewLine
    End If

    If pvntSequenciaLanctoCC > 0 Then
        strSQL = strSQL & _
             " And      NR_SEQU_LANC_CNTA_CRRT = " & pvntSequenciaLanctoCC & vbNewLine
    End If

    LerTodosComposicaoNetCorretoras = fgQueryXMLLerTodos("ComposicaoNetCorretoras", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosComposicaoNetCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Ler o registro da tabela A8.TB_LANC_CC_CNTB conforme o filtro específicado
Public Function Ler(ByVal pvntSequenciaOperacao As Variant, _
                    ByVal plngTipoLancamentoIntegracao As Variant, _
                    Optional ByVal plngNumSequLancamento As Long = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Select   NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "          TP_LANC_ITGR, " & vbNewLine & _
             "          TP_BKOF, " & vbNewLine & _
             "          CO_EMPR, " & vbNewLine & _
             "          CO_VEIC_LEGA, " & vbNewLine & _
             "          SG_SIST, " & vbNewLine & _
             "          SG_SIST_CC, " & vbNewLine & _
             "          TP_OPER, " & vbNewLine & _
             "          CO_LOCA_LIQU, " & vbNewLine & _
             "          IN_LANC_DEBT_CRED, " & vbNewLine & _
             "          DT_OPER, " & vbNewLine & _
             "          CO_BANC, " & vbNewLine & _
             "          CO_AGEN, " & vbNewLine & _
             "          NU_CC, " & vbNewLine & _
             "          NU_COMD_OPER, " & vbNewLine & _
             "          VA_LANC_CC, " & vbNewLine & _
             "          NU_CTRL_ITGR, " & vbNewLine & _
             "          CO_ULTI_SITU_PROC, " & vbNewLine & _
             "          DH_ULTI_ATLZ, " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ, " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
             "          TP_CNAL_VEND, " & vbNewLine & _
             "          CD_SUB_PROD, " & vbNewLine & _
             "          NR_SEQU_LANC "

    strSQL = strSQL & _
             " From     A8.TB_LANC_CC_CNTB " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV   =   " & pvntSequenciaOperacao & vbNewLine & _
             " And      TP_LANC_ITGR        =   " & plngTipoLancamentoIntegracao & vbNewLine & _
             " And      NR_SEQU_LANC        =   " & plngNumSequLancamento
            
    Ler = fgQueryXMLLer("LancamentoCCContabil", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Promover a alteração dos dados de lançamento de conta corrente

Public Function AlterarContaCorrente(ByVal pvntSequenciaOperacao As Variant, _
                                     ByVal pintTipoLancamentoIntegracao As Integer, _
                                     ByVal pvntCodigoAgencia As Variant, _
                                     ByVal pvntNumeroCC As Variant, _
                            Optional ByVal pvntVA_LANC_CC As Variant = 0, _
                            Optional ByVal pintIN_LANC_DEBT_CRED As Integer = 0) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_LANC_CC_CNTB " & vbNewLine & _
             " Set      CO_AGEN                 =   " & pvntCodigoAgencia & "," & vbNewLine & _
             "          NU_CC                   =   " & pvntNumeroCC & "," & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =  '" & fgEstacaoTrabalhoUsuario & "'," & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate, " & vbNewLine

    If pvntVA_LANC_CC > 0 Then
        strSQL = strSQL & _
                 "          VA_LANC_CC              =   " & fgVlrToDBServer(pvntVA_LANC_CC) & "," & vbNewLine
    End If

    If pintIN_LANC_DEBT_CRED > 0 Then
        strSQL = strSQL & _
                 "          IN_LANC_DEBT_CRED       =   " & pintIN_LANC_DEBT_CRED & vbNewLine
    End If

    strSQL = strSQL & _
             " Where    NU_SEQU_OPER_ATIV       =   " & pvntSequenciaOperacao & "      " & vbNewLine & _
             " And      TP_LANC_ITGR            =   " & pintTipoLancamentoIntegracao & vbNewLine

    If fgExecuteSQL(strSQL) = 0 Then
        'Lançamento já integrado por outro usuário.
        lngCodigoErroNegocio = 3040
        GoTo ErrorHandler
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Promover o estorno do lançamento de conta corrente na tabela A8.TB_LANC_CC_CNTB
Public Function Excluir(ByVal pvntNumeroSequencial As Variant) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    'KIDA - CBLC
    strSQL = " DELETE FROM A8.TB_HIST_SITU_LANC_CC_CNTB              " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & vbNewLine

    fgExecuteSQL strSQL


    strSQL = " DELETE FROM A8.TB_LANC_CC_CNTB                        " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntNumeroSequencial & vbNewLine

    fgExecuteSQL strSQL

    Excluir = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Excluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
