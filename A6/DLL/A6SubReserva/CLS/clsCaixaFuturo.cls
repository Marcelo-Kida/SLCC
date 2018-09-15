VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaFuturo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105A3F00B3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Movimento de Caixa Futuro.

Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private objContext                          As COMSVCSLib.ObjectContext

Private mlngClassDebugID                    As Long

Public Property Get ClassDebugID() As Long
    On Error GoTo ErrorHandler
    ClassDebugID = mlngClassDebugID
    Exit Property
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ClassDebugID Property", lngCodigoErroNegocio, intNumeroSequencialErro)
End Property

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pela leitura de detalhes do movimento futuro,
' e pelo retorno à camada intermediária.

Public Function ObterDetalheMovimento(ByVal pstrCodVeiculoLegal As String, _
                                      ByVal pstrSiglaSistema As String, _
                                      ByVal pstrItemCaixa As String, _
                                      ByVal pdatMovimento As Date) As String

Dim strSQL                                  As String
Dim strDataMov                              As String
Dim strOwner                                As String

    On Error GoTo ErrorHandler

    strDataMov = Format$(pdatMovimento, "yyyymmdd")
    
    If pdatMovimento < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        strOwner = "A6HIST"
    Else
        strOwner = "A6"
    End If
    
    strSQL = " SELECT   A.SG_SIST, B.NO_REDU_EMPR, C.SG_LOCA_LIQU, A.DE_TIPO_LIQU,              " & vbNewLine & _
             "          DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_LIQU_OPER, 0) AS ENTRADA,           " & vbNewLine & _
             "          DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_LIQU_OPER, 0) AS SAIDA,             " & vbNewLine & _
             "          A.DE_ATIV, A.CO_CNPJ_CNPT, A.NO_CNPT, A.DT_LIQU_OPER, A.DT_RETN_OPER    " & vbNewLine & _
             " FROM     " & strOwner & ".TB_MOVI_FUTU_OPER_ATIV       A,                        " & vbNewLine & _
             "          A8.TB_EMPRESA_HO                   B,                                   " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO             C                                    " & vbNewLine & _
             " WHERE    A.CO_EMPR      = B.CO_EMPR                                              " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU = C.CO_LOCA_LIQU                                         " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA = '" & pstrCodVeiculoLegal & "'                          " & vbNewLine & _
             " AND      A.SG_SIST      = '" & pstrSiglaSistema & "'                             " & vbNewLine & _
             " AND      A.CO_ITEM_CAIX = '" & pstrItemCaixa & "'                                " & vbNewLine & _
             " AND      A.DT_LIQU_OPER =  " & fgDtXML_To_Oracle(strDataMov) & "                 " & vbNewLine & _
             " AND      C.CO_EMPR_FUSI = 1                                                      " & vbNewLine & _
             " ORDER BY A.DT_LIQU_OPER DESC                                                     "
    
    ObterDetalheMovimento = fgQueryXMLLerTodos("DetalheMovimento", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMovimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de registros de movimento futuro, agrupados por grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterMovimentoPorGrupoVeiculoLegal(ByVal pintCodGrupoVeicLegal As Integer, _
                                                   ByVal pdatBaseMovimento As Date, _
                                                   ByVal pintQuantDiasMov As Integer, _
                                                   ByVal pintPaginacao As enumPaginacao, _
                                          Optional ByVal plngCodEmpresa As Long, _
                                          Optional ByVal pstrCodVeiculoLegal As String) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String
Dim strDataDe                               As String
Dim strDataAte                              As String

Dim xmlMovimento                            As MSXML2.DOMDocument40
Dim xmlDatasMovimento                       As MSXML2.DOMDocument40
Dim intDatas                                As Integer
Dim datDiaUtil                              As Date
    
    On Error GoTo ErrorHandler

    Set xmlMovimento = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDatasMovimento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlMovimento, "", "Repeat_MovimentoGeral", "")
    
    If pintPaginacao = Proximo Then
        datDiaUtil = pdatBaseMovimento
    Else
        datDiaUtil = fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov, pintPaginacao)
    End If
    
    Call fgAppendNode(xmlDatasMovimento, "", "DatasMovimento", "")
    Call fgAppendNode(xmlDatasMovimento, "DatasMovimento", "DiaUtil", Format$(datDiaUtil, "yyyymmdd"))
    
    For intDatas = 1 To pintQuantDiasMov
        datDiaUtil = fgAdicionarDiasUteis(datDiaUtil, 1, enumPaginacao.Proximo)
        Call fgAppendNode(xmlDatasMovimento, "DatasMovimento", "DiaUtil", Format$(datDiaUtil, "yyyymmdd"))
    Next
    
    Call fgAppendXML(xmlMovimento, "Repeat_MovimentoGeral", xmlDatasMovimento.xml)
    
    If pintPaginacao = Proximo Then
        strDataDe = Format$(pdatBaseMovimento, "yyyymmdd")
        strDataAte = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov, pintPaginacao), "yyyymmdd")
    Else
        strDataDe = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov, pintPaginacao), "yyyymmdd")
        strDataAte = Format$(pdatBaseMovimento, "yyyymmdd")
    End If
    
    strSQL = " SELECT   B.CO_VEIC_LEGA, A.NO_VEIC_LEGA, B.DT_LIQU_OPER,                                 " & vbNewLine & _
             "          NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 1, B.VA_LIQU_OPER, 0)),0) -                 " & vbNewLine & _
             "          NVL(SUM(DECODE(B.IN_MOVI_ENTR_SAID, 2, B.VA_LIQU_OPER, 0)),0) AS VA_LIQU_OPER   " & vbNewLine & _
             " FROM     A8.TB_VEIC_LEGA            A,                                                   " & vbNewLine & _
             "          A6.TB_MOVI_FUTU_OPER_ATIV  B                                                    " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA      =      B.CO_VEIC_LEGA                                       " & vbNewLine & _
             " AND      A.SG_SIST           =      B.SG_SIST                                            " & vbNewLine & _
             " AND      A.CO_GRUP_VEIC_LEGA =  " & pintCodGrupoVeicLegal & "                            " & vbNewLine
    
    If Trim(pstrCodVeiculoLegal) <> vbNullString Then
        strSQL = strSQL & _
             " AND      A.CO_VEIC_LEGA      = '" & pstrCodVeiculoLegal & "'                             " & vbNewLine
    End If
    
    If plngCodEmpresa <> 0 Then
        strSQL = strSQL & _
             " AND      A.CO_EMPR           = " & plngCodEmpresa & "                                    " & vbNewLine
    End If
    
    strSQL = strSQL & _
             " AND      B.DT_LIQU_OPER BETWEEN " & fgDtXML_To_Oracle(strDataDe) & vbNewLine & _
             " AND                             " & fgDtXML_To_Oracle(strDataAte) & vbNewLine
    
    strSQL = strSQL & _
             fgSegregaDados("A8.TB_VEIC_LEGA", _
                            False, _
                            "A", _
                            "A", _
                            False, _
                            False, _
                            True, _
                            False, _
                            True, _
                            False)
    
    'Veirificação para a base historica
    If fgDtXML_To_Date(strDataAte) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(strDataDe) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQL, "A6.TB_MOVI_FUTU_OPER_ATIV", "A6HIST.TB_MOVI_FUTU_OPER_ATIV")
        End If
    Else
        strSQL = Replace(strSQL, "A6.TB_MOVI_FUTU_OPER_ATIV", "A6HIST.TB_MOVI_FUTU_OPER_ATIV")
    End If
    
    strSQL = strSQL & _
             " GROUP BY B.CO_VEIC_LEGA, A.NO_VEIC_LEGA, B.DT_LIQU_OPER                                  " & vbNewLine & _
             " ORDER BY B.CO_VEIC_LEGA, B.DT_LIQU_OPER                                                  "
    
    ObterMovimentoPorGrupoVeiculoLegal = fgQueryXMLLerTodos("MovGrupoVeicLegal", strSQL, App.EXEName & "." & TypeName(Me))

    If ObterMovimentoPorGrupoVeiculoLegal <> vbNullString Then
        Call fgAppendXML(xmlMovimento, "Repeat_MovimentoGeral", ObterMovimentoPorGrupoVeiculoLegal)
    End If
    
    ObterMovimentoPorGrupoVeiculoLegal = xmlMovimento.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlMovimento = Nothing
    Set xmlDatasMovimento = Nothing
    
    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set xmlMovimento = Nothing
    Set xmlDatasMovimento = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMovimentoPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de registros de movimento futuro, agrupados por veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterMovimentoPorVeiculoLegal(ByVal pstrCodVeiculoLegal As String, _
                                              ByVal pstrSiglaSistema As String, _
                                              ByVal pintNivelAgrupItemCaixa As Integer, _
                                              ByVal pdatD0 As Date, _
                                              ByVal pdatBaseMovimento As Date, _
                                              ByVal pintQuantDiasMov As Integer, _
                                              ByVal pintPaginacao As enumPaginacao) As String

Dim strSQL                                  As String
Dim strD0                                   As String
Dim strDataDe                               As String
Dim strDataAte                              As String

Dim strAgrupItemCaixa                       As String
Dim intCount                                As Integer

Dim xmlMovimento                            As MSXML2.DOMDocument40
Dim xmlDatasMovimento                       As MSXML2.DOMDocument40
Dim xmlSaldoInicialMov                      As MSXML2.DOMDocument40
Dim intDatas                                As Integer
Dim datDiaUtil                              As Date
    
    On Error GoTo ErrorHandler

    Set xmlMovimento = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDatasMovimento = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlSaldoInicialMov = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlMovimento, "", "Repeat_MovimentoGeral", "")
    
    '>>>>> Elaboração do XML contendo as datas úteis do período estabelecido
    If pintPaginacao = Proximo Then
        datDiaUtil = pdatBaseMovimento
    Else
        datDiaUtil = fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov, pintPaginacao)
    End If
    
    Call fgAppendNode(xmlDatasMovimento, "", "DatasMovimento", "")
    
    If pdatBaseMovimento <> pdatD0 Then
        Call fgAppendNode(xmlDatasMovimento, "DatasMovimento", "DiaUtil", Format$(pdatD0, "yyyymmdd"))
        pintQuantDiasMov = pintQuantDiasMov - 1
    End If
    
    Call fgAppendNode(xmlDatasMovimento, "DatasMovimento", "DiaUtil", Format$(datDiaUtil, "yyyymmdd"))
    
    For intDatas = 1 To pintQuantDiasMov
        datDiaUtil = fgAdicionarDiasUteis(datDiaUtil, 1, enumPaginacao.Proximo)
        Call fgAppendNode(xmlDatasMovimento, "DatasMovimento", "DiaUtil", Format$(datDiaUtil, "yyyymmdd"))
    Next
    
    Call fgAppendXML(xmlMovimento, "Repeat_MovimentoGeral", xmlDatasMovimento.xml)
    
    '>>>>> Elaboração do XML contendo o saldo inicial do movimento de cada paginação
    If pdatBaseMovimento = pdatD0 Then
        strDataDe = Format$(pdatD0, "yyyymmdd")
        strDataAte = Format$(pdatD0, "yyyymmdd")
    Else
        strDataDe = Format$(pdatD0, "yyyymmdd")
        If pintPaginacao = enumPaginacao.Proximo Then
            strDataAte = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, 1, enumPaginacao.Anterior), "yyyymmdd")
        Else
            strDataAte = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov + 2, enumPaginacao.Anterior), "yyyymmdd")
        End If
    End If
    
    strSQL = " SELECT   NVL(SUM(DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_LIQU_OPER, 0)),0) -                 " & vbNewLine & _
             "          NVL(SUM(DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_LIQU_OPER, 0)),0) AS VA_LIQU_OPER   " & vbNewLine & _
             " FROM     A6.TB_MOVI_FUTU_OPER_ATIV  A,                                                   " & vbNewLine & _
             "          A8.TB_VEIC_LEGA            B                                                    " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA = B.CO_VEIC_LEGA                                                 " & vbNewLine & _
             " AND      A.SG_SIST      = B.SG_SIST                                                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA =      '" & pstrCodVeiculoLegal & "'                             " & vbNewLine & _
             " AND      A.SG_SIST      =      '" & pstrSiglaSistema & "'                                " & vbNewLine & _
             " AND      A.DT_LIQU_OPER BETWEEN " & fgDtXML_To_Oracle(strDataDe) & "                     " & vbNewLine & _
             " AND                             " & fgDtXML_To_Oracle(strDataAte) & "                    " & vbNewLine
             
    strSQL = strSQL & _
             fgSegregaDados("A8.TB_VEIC_LEGA", _
                            False, _
                            "B", _
                            "B", _
                            False, _
                            False, _
                            True, _
                            False, _
                            True, _
                            False)

    'Verificação para base historica
    If fgDtXML_To_Date(strDataAte) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(strDataDe) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQL, "A6.TB_MOVI_FUTU_OPER_ATIV", "A6HIST.TB_MOVI_FUTU_OPER_ATIV")
        End If
    Else
        strSQL = Replace(strSQL, "A6.TB_MOVI_FUTU_OPER_ATIV", "A6HIST.TB_MOVI_FUTU_OPER_ATIV")
    End If
        
    strSQL = strSQL & _
             " GROUP BY A.CO_VEIC_LEGA, A.SG_SIST "
    
    ObterMovimentoPorVeiculoLegal = fgQueryXMLLerTodos("SaldoInicialMovimento", strSQL, App.EXEName & "." & TypeName(Me))

    If ObterMovimentoPorVeiculoLegal <> vbNullString Then
        Call fgAppendXML(xmlMovimento, "Repeat_MovimentoGeral", ObterMovimentoPorVeiculoLegal)
    End If
    
    '>>>>> Elaboração do XML contendo itens de caixa e o respectivos movimentos
    strAgrupItemCaixa = vbNullString
    For intCount = 1 To pintNivelAgrupItemCaixa
        strAgrupItemCaixa = strAgrupItemCaixa & _
                            "|| A.CO_ITEM_CAIX_NIVE_" & Format$(intCount, "00") & " "
    Next
    
    strAgrupItemCaixa = Mid$(strAgrupItemCaixa, 4)
    strD0 = Format$(pdatD0, "yyyymmdd")
    
    If pintPaginacao = Proximo Then
        strDataDe = Format$(pdatBaseMovimento, "yyyymmdd")
        strDataAte = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov, pintPaginacao), "yyyymmdd")
    Else
        strDataDe = Format$(fgAdicionarDiasUteis(pdatBaseMovimento, pintQuantDiasMov + 1, pintPaginacao), "yyyymmdd")
        strDataAte = Format$(pdatBaseMovimento, "yyyymmdd")
    End If
    
    strSQL = " SELECT   A.DT_LIQU_OPER, " & strAgrupItemCaixa & "AS AGRUP_ITEM_CAIX,                    " & vbNewLine & _
             "          NVL(SUM(DECODE(A.IN_MOVI_ENTR_SAID, 1, A.VA_LIQU_OPER, 0)),0) -                 " & vbNewLine & _
             "          NVL(SUM(DECODE(A.IN_MOVI_ENTR_SAID, 2, A.VA_LIQU_OPER, 0)),0) AS VA_LIQU_OPER   " & vbNewLine & _
             " FROM     A6.TB_MOVI_FUTU_OPER_ATIV  A,                                                   " & vbNewLine & _
             "          A8.TB_VEIC_LEGA            B                                                    " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA = B.CO_VEIC_LEGA                                                 " & vbNewLine & _
             " AND      A.SG_SIST      = B.SG_SIST                                                      " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA =      '" & pstrCodVeiculoLegal & "'                             " & vbNewLine & _
             " AND      A.SG_SIST      =      '" & pstrSiglaSistema & "'                                " & vbNewLine & _
             " AND     (A.DT_LIQU_OPER =       " & fgDtXML_To_Oracle(strD0) & "                         " & vbNewLine & _
             " OR       A.DT_LIQU_OPER BETWEEN " & fgDtXML_To_Oracle(strDataDe) & "                     " & vbNewLine & _
             " AND                             " & fgDtXML_To_Oracle(strDataAte) & ")                   " & vbNewLine
             
    strSQL = strSQL & _
             fgSegregaDados("A8.TB_VEIC_LEGA", _
                            False, _
                            "B", _
                            "B", _
                            False, _
                            False, _
                            True, _
                            False, _
                            True, _
                            False)

    strSQL = strSQL & _
             " GROUP BY A.DT_LIQU_OPER, " & strAgrupItemCaixa & vbNewLine & _
             " ORDER BY A.DT_LIQU_OPER, " & strAgrupItemCaixa
    
    ObterMovimentoPorVeiculoLegal = fgQueryXMLLerTodos("MovAgrupItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))

    If ObterMovimentoPorVeiculoLegal <> vbNullString Then
        Call fgAppendXML(xmlMovimento, "Repeat_MovimentoGeral", ObterMovimentoPorVeiculoLegal)
    End If
    
    ObterMovimentoPorVeiculoLegal = xmlMovimento.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMovimentoPorVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
