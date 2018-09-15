VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemCaixaTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Item Caixa x Tipo Operação.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pela leitura de propriedades da tabela item de caixa x tipo de operação,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "SELECT TP_OPER,               " & vbCrLf & _
             "    '' AS TP_CAIX,            " & vbCrLf & _
             "       CO_ITEM_CAIX,          " & vbCrLf & _
             "    '' AS CO_ITEM_CAIX_OLD,   " & vbCrLf & _
             "       DH_ULTI_ATLZ,          " & vbCrLf & _
             "       TP_CNPT,               " & vbCrLf & _
             "       IN_ENTR_SAID,          " & vbCrLf & _
             "       CO_USUA_ULTI_ATLZ,     " & vbCrLf & _
             "       CO_ETCA_TRAB_ULTI_ATLZ " & vbCrLf & _
             " FROM  A6.TB_ITEM_CAIX_TIPO_OPER "

    ObterPropriedades = fgPropriedades("Grupo_ItemCaixaTipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da tabela item de caixa x tipo de operação,
' e pelo retorno à camada intermediária.

Public Function LerTodos(Optional ByVal pstrItemCaixa As String, _
                         Optional ByVal pintTipoOperacao As Integer, _
                         Optional ByVal penumTipoCaixa As enumTipoCaixa, _
                         Optional ByVal penumTipoEntradaSaida As enumTipoEntradaSaida, _
                         Optional ByVal pstrDataUltimaAtualizacao As String = vbNullString) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   A.CO_ITEM_CAIX, " & vbCrLf & _
             "   A.TP_OPER, " & vbCrLf & _
             "   A.TP_CNPT, " & vbCrLf & _
             "   A.IN_ENTR_SAID, " & vbCrLf & _
             "   A.CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   A.CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   A.DH_ULTI_ATLZ, " & vbCrLf & _
             "   B.TP_CAIX,   " & vbCrLf & _
             "   B.DE_ITEM_CAIX " & vbCrLf & _
             " FROM  A6.TB_ITEM_CAIX_TIPO_OPER   A,  " & vbCrLf & _
             "       A6.TB_ITEM_CAIX_OPER_ATIV   B   " & vbCrLf & _
             " WHERE A.CO_ITEM_CAIX  =   B.CO_ITEM_CAIX    " & vbCrLf

    If penumTipoCaixa <> 0 Then
        strSQL = strSQL & " AND   B.TP_CAIX = " & penumTipoCaixa & vbCrLf
    End If

    If Trim(pstrItemCaixa) <> "" Then
        strSQL = strSQL & " AND   A.CO_ITEM_CAIX = '" & pstrItemCaixa & "' " & vbCrLf
    ElseIf pintTipoOperacao <> 0 Then
        strSQL = strSQL & " AND   A.TP_OPER = " & pintTipoOperacao & vbCrLf
    End If

    If penumTipoEntradaSaida <> 0 Then
        strSQL = strSQL & " AND   A.IN_ENTR_SAID = " & penumTipoCaixa & vbCrLf
    End If
    
    If pstrDataUltimaAtualizacao <> "" Then
        strSQL = strSQL & " AND   A.DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pstrDataUltimaAtualizacao)
    End If
    

    strSQL = strSQL & fgSegregaDados("A6.TB_ITEM_CAIX_OPER_ATIV", _
                                     False, _
                                     "B", _
                                     "B", _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False, _
                                     False)

    strSQL = strSQL & " ORDER BY A.CO_ITEM_CAIX " & vbCrLf

    LerTodos = fgQueryXMLLerTodos("ItemCaixaTipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela leitura da tabela item de caixa x tipo de operação, pelo código do item caixa,
' e pelo retorno à camada intermediária.

Public Function Ler(ByVal pstrCodigoItemCaixa As String) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   CO_ITEM_CAIX, " & vbCrLf & _
             "   CO_ITEM_CAIX_PAI, " & vbCrLf & _
             "   TP_ITEM_CAIX, " & vbCrLf & _
             "   TP_CAIX, " & vbCrLf & _
             "   DE_ITEM_CAIX, " & vbCrLf & _
             "   DT_INIC_VIGE, " & vbCrLf & _
             "   DT_FIM_VIGE, " & vbCrLf & _
             "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   DH_ULTI_ATLZ " & vbCrLf & _
             " FROM " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
             " WHERE " & vbCrLf & _
             "   CO_ITEM_CAIX = '" & pstrCodigoItemCaixa & "' "

    Ler = fgQueryXMLLer("ItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Len(Ler) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela item de caixa x tipo de operação,
' e pelo retorno à camada intermediária.

Public Function Executar(ByRef pxmlItemCaixa As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlItemCaixa.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = SalvarTodos(pxmlItemCaixa)
        Case "Ler"
            Executar = Ler(pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text)
        Case "LerTodos"
            Executar = LerTodos(pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text, _
                                pxmlItemCaixa.selectSingleNode("//TP_OPER").Text, _
                                pxmlItemCaixa.selectSingleNode("//TP_CAIX").Text)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização em lote da tabela item de caixa x tipo de operação,
' e pelo retorno à camada intermediária.

Public Function SalvarTodos(ByVal pxmlItemCaixaTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim xmlDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    For Each xmlDomNode In pxmlItemCaixaTipoOperacao.documentElement.selectNodes("//Repeat_ItemCaixaTipoOperacao/*")
        Select Case xmlDomNode.selectSingleNode("@Operacao").Text
            Case "Incluir"
                Call flVerificarAssociacaoPorTipoBackOffice(xmlDomNode)
                Call flIncluir(xmlDomNode)
            Case "Alterar"
                Call flAlterar(xmlDomNode)
            Case "Excluir"
                Call flExcluir(xmlDomNode)
        End Select
    Next

    SalvarTodos = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão de uma nova associação item caixa x tipo operação.

Private Function flIncluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String
Dim strLerTodos                             As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A6.TB_ITEM_CAIX_TIPO_OPER(" & vbCrLf & _
             " TP_OPER, " & vbCrLf & _
             " CO_ITEM_CAIX, " & vbCrLf & _
             " TP_CNPT, " & vbCrLf & _
             " IN_ENTR_SAID, " & vbCrLf & _
             " CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             " DH_ULTI_ATLZ) " & vbCrLf & _
             " VALUES (" & vbCrLf & _
                   pxmlDomNode.selectSingleNode("TP_OPER").Text & ", " & vbCrLf & _
             "'" & pxmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "', " & vbCrLf & _
                   pxmlDomNode.selectSingleNode("TP_CNPT").Text & ", " & vbCrLf & _
                   pxmlDomNode.selectSingleNode("IN_ENTR_SAID").Text & ", " & vbCrLf & _
             "'" & fgUsuarioRede & "', " & vbCrLf & _
             "'" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             " SYSDATE)"

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela exclusão de uma associação item caixa x tipo operação.

Private Function flExcluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                   As String
Dim strLerTodos                              As String

On Error GoTo ErrorHandler
    
    strSQL = " DELETE FROM A6.TB_ITEM_CAIX_TIPO_OPER " & vbCrLf & _
             "  WHERE CO_ITEM_CAIX   = '" & pxmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "'" & _
             "   AND  TP_CNPT        =  " & pxmlDomNode.selectSingleNode("TP_CNPT").Text & _
             "   AND  IN_ENTR_SAID   =  " & pxmlDomNode.selectSingleNode("IN_ENTR_SAID").Text & _
             "   AND  DH_ULTI_ATLZ   =  " & fgDtHrXML_To_Oracle(pxmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário.
        lngCodigoErroNegocio = 4001
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela verificação de uma associação item caixa x tipo operação, por tipo de backoffice.

Private Sub flVerificarAssociacaoPorTipoBackOffice(ByVal pxmlDomNode As MSXML2.IXMLDOMNode)
                                         
Dim rsItemCaixa                             As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT A.TP_OPER                          " & vbNewLine & _
             " FROM   A6.TB_ITEM_CAIX_TIPO_OPER A,       " & vbNewLine & _
             "        A6.TB_ITEM_CAIX_OPER_ATIV B        " & vbNewLine & _
             " WHERE  A.CO_ITEM_CAIX    = B.CO_ITEM_CAIX " & vbNewLine & _
             " AND    B.TP_BKOF         = " & fgObterTipoBackOfficeUsuario & vbNewLine & _
             " AND    B.TP_CAIX         = " & pxmlDomNode.selectSingleNode("TP_CAIX").Text & vbNewLine & _
             " AND    A.TP_CNPT         = " & pxmlDomNode.selectSingleNode("TP_CNPT").Text & vbNewLine & _
             " AND    A.IN_ENTR_SAID    = " & pxmlDomNode.selectSingleNode("IN_ENTR_SAID").Text & vbNewLine & _
             " AND    A.TP_OPER         = " & pxmlDomNode.selectSingleNode("TP_OPER").Text & vbNewLine

    Set rsItemCaixa = fgQuerySQL(strSQL)

    If Not rsItemCaixa.EOF Then
        '2021 - Combinação entre Tipo de Operação, Tipo Contraparte e Tipo de Movimento, já associada a outro Item de Caixa.
        lngCodigoErroNegocio = 2021
        GoTo ErrorHandler
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarAssociacaoPorTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub


' Este método é responsável pela ALTERAÇÃO de uma associação item caixa x tipo operação.

Private Function flAlterar(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = "UPDATE A6.TB_ITEM_CAIX_TIPO_OPER" & _
             " SET CO_ITEM_CAIX     = '" & pxmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text & "'" & _
             " WHERE TP_OPER        =  " & pxmlDomNode.selectSingleNode("TP_OPER").Text & _
             "  AND  CO_ITEM_CAIX   = '" & pxmlDomNode.selectSingleNode("CO_ITEM_CAIX_OLD").Text & "'" & _
             "  AND  TP_CNPT        =  " & pxmlDomNode.selectSingleNode("TP_CNPT").Text & _
             "  AND  IN_ENTR_SAID   =  " & pxmlDomNode.selectSingleNode("IN_ENTR_SAID").Text & _
             "  AND  DH_ULTI_ATLZ   =  " & fgDtHrXML_To_Oracle(pxmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário.
        lngCodigoErroNegocio = 4001
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


