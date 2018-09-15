VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoPJTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Produto PJ x Tipo Operação.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements ObjectControl
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

' Este método é responsável pela atualização geral da tabela produto pj x tipo operação,
' e pelo retorno à camada intermediária.

Public Function Executar(ByRef pstrProdutoTipoOperacao As String) As Boolean

Dim xmlProdutoTipoOperacao                  As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strOperacao                             As String

On Error GoTo ErrorHandler

    Set xmlProdutoTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlProdutoTipoOperacao.loadXML(pstrProdutoTipoOperacao) Then
        '3 - XML Inválido.
        lngCodigoErroNegocio = 3
        GoTo ErrorHandler
    End If

    strOperacao = xmlProdutoTipoOperacao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(xmlProdutoTipoOperacao)
        Case "Excluir"
            Executar = flExcluir(xmlProdutoTipoOperacao)
        Case "Ler"
            'Executar = Ler()
        Case "LerTodos"
            Executar = LerTodos(xmlProdutoTipoOperacao.documentElement.selectSingleNode("CO_PROD").Text, _
                                xmlProdutoTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    Set xmlProdutoTipoOperacao = Nothing

    Executar = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set xmlProdutoTipoOperacao = Nothing

    Executar = False

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da tabela produto pj x tipo operação,
' e pelo retorno à camada intermediária.

Public Function LerTodos(Optional ByVal plngCodigoProduto As Long, _
                         Optional ByVal plngTipoOperacao As Long) As String

Dim strSQL                                  As String
Dim blnWhere                                As Boolean

On Error GoTo ErrorHandler

    strSQL = " Select  CO_PROD,               " & vbNewLine & _
                    " TP_OPER,                 " & vbNewLine & _
                    " DT_INIC_VIGE,           " & vbNewLine & _
                    " DT_FIM_VIGE,            " & vbNewLine & _
                    " CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
                    " CO_ETCA_TRAB_ULTI_ATLZ  " & vbNewLine & _
             " From   A6.TB_TIPO_OPER_PROD_PJ "

    blnWhere = False
    If plngTipoOperacao <> 0 Then
       strSQL = strSQL & "Where   TP_OPER   = " & plngTipoOperacao
       blnWhere = True
    End If

    If plngCodigoProduto <> 0 Then
        If blnWhere = False Then
            strSQL = strSQL & " Where CO_PROD   = " & plngCodigoProduto
        Else
            strSQL = strSQL & " And   CO_PROD   = " & plngCodigoProduto
        End If
    End If

    strSQL = strSQL & " Order By  CO_PROD "

    LerTodos = fgQueryXMLLerTodos("ProdutoPJTipoOperacao", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela leitura de propriedades da tabela produto pj x tipo operação,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "Select           A.CO_PROD,                         " & vbNewLine & _
            "                 A.TP_OPER,                          " & vbNewLine & _
            "                 A.DT_INIC_VIGE,                    " & vbNewLine & _
            "                 A.DT_FIM_VIGE,                     " & vbNewLine & _
            "                 A.CO_USUA_ULTI_ATLZ,               " & vbNewLine & _
            "                 A.CO_ETCA_TRAB_ULTI_ATLZ,          " & vbNewLine & _
            "                 A.DH_ULTI_ATLZ                     " & vbNewLine & _
            "From             A6.TB_TIPO_OPER_PROD_PJ A   "

    ObterPropriedades = fgPropriedades("Grupo_Propriedades", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela alteração de uma associação produto pj x tipo operação.

Private Function flAlterar(ByVal pxmlProdutoPJTipoOperacao As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " UPDATE " & vbCrLf & _
            "       A6.TB_TIPO_OPER_PROD_PJ " & vbCrLf & _
            " SET " & vbCrLf & _
            "       TP_OPER = " & pxmlProdutoPJTipoOperacao.selectSingleNode("TP_OPER").Text & ", " & vbCrLf & _
            "       DT_INIC_VIGE = " & fgDtXML_To_Oracle(pxmlProdutoPJTipoOperacao.selectSingleNode("DT_INIC_VIGE").Text) & ", " & vbCrLf & _
            "       DT_FIM_VIGE = " & fgDtXML_To_Oracle(pxmlProdutoPJTipoOperacao.selectSingleNode("DT_FIM_VIGE").Text) & ", " & vbCrLf & _
            "       CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "', " & vbCrLf & _
            "       CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
            "       DH_ULTI_ATLZ = SYSDATE " & vbCrLf & _
            " WHERE " & vbCrLf & _
            "       CO_PROD = " & pxmlProdutoPJTipoOperacao.selectSingleNode("CO_PROD").Text & vbCrLf & _
            " AND   TP_OPER  = " & pxmlProdutoPJTipoOperacao.selectSingleNode("TP_OPER").Text & _
            " AND   DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlProdutoPJTipoOperacao.selectSingleNode("DH_ULTI_ATLZ").Text)

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

' Este método é responsável pela exclusão de uma associação produto pj x tipo operação.

Private Function flExcluir(ByVal pxmlProdutoPJTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    With pxmlProdutoPJTipoOperacao.documentElement
        strSQL = " DELETE FROM A6.TB_TIPO_OPER_PROD_PJ " & vbCrLf & _
                 " WHERE  CO_PROD      = " & .selectSingleNode("//CO_PROD").Text & vbCrLf & _
                 " AND    TP_OPER      = " & .selectSingleNode("//TP_OPER").Text & _
                 " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("//DH_ULTI_ATLZ").Text)
    End With

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

' Este método é responsável pela inclusão de uma associação produto pj x tipo operação.

Private Function flIncluir(ByVal pxmlProdutoPJTipoOperacao As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "INSERT INTO " & vbCrLf & _
            "   A6.TB_TIPO_OPER_PROD_PJ ( " & vbCrLf & _
            "   CO_PROD, " & vbCrLf & _
            "   TP_OPER, " & vbCrLf & _
            "   DT_INIC_VIGE, " & vbCrLf & _
            "   DT_FIM_VIGE, " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
            "   DH_ULTI_ATLZ " & vbCrLf & _
            ") VALUES ( " & vbCrLf & _
            pxmlProdutoPJTipoOperacao.selectSingleNode("CO_PROD").Text & ", " & vbCrLf & _
            pxmlProdutoPJTipoOperacao.selectSingleNode("TP_OPER").Text & ", " & vbCrLf & _
            fgDtXML_To_Oracle(pxmlProdutoPJTipoOperacao.selectSingleNode("DT_INIC_VIGE").Text) & ", " & vbCrLf & _
            fgDtXML_To_Oracle(pxmlProdutoPJTipoOperacao.selectSingleNode("DT_FIM_VIGE").Text) & ", " & vbCrLf & _
            "'" & fgUsuarioRede & "', " & vbCrLf & _
            "'" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
            "SYSDATE " & vbCrLf & _
            ") " & vbCrLf

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela produto pj x tipo operação,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlProdutoPJTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim xmlDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    For Each xmlDomNode In pxmlProdutoPJTipoOperacao.documentElement.selectNodes("//Repeat_TipoOperacaoProduto/*")
        Select Case xmlDomNode.selectSingleNode("@Operacao").Text
            Case "Incluir"
                flIncluir xmlDomNode
            Case "Alterar"
                flAlterar xmlDomNode
            Case Else
                ' 7 - Operação Inválida
                lngCodigoErroNegocio = 7
                GoTo ErrorHandler
        End Select
    Next

    Salvar = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela leitura do produto pj por tipo de operação,
' e pelo retorno à camada intermediária.

Public Function ObterProdutoPJTipoOperacao(Optional ByVal plngTipoOperacao As Long = 0) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " Select    B.CO_PROD,                          " & vbNewLine & _
            "           A.TP_OPER,                           " & vbNewLine & _
            "           A.DT_INIC_VIGE,                     " & vbNewLine & _
            "           A.DT_FIM_VIGE                       " & vbNewLine & _
            " From      A6.TB_TIPO_OPER_PROD_PJ  A,         " & vbNewLine & _
            "           A8.TB_PRODUTO                   B   " & vbNewLine & _
            " Where    (B.DT_FIM_VIGE   IS  NULL            " & vbNewLine & _
            " Or        B.DT_FIM_VIGE   <   SYSDATE)        " & vbNewLine & _
            " And       B.CO_EMPR_FUSI  =   1               " & vbNewLine & _
            " And       B.CO_PROD       =   A.CO_PROD(+)    " & vbNewLine & _
            " Order By  B.CO_PROD "

    ObterProdutoPJTipoOperacao = fgQueryXMLLerTodos("Produtos", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterProdutoPJTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de produtos pj x tipos operação, agrupados por status,
' e pelo retorno à camada intermediária.

Public Function ObterProdutosTipoOperacao(ByVal plngStatusProduto As Long) As String

Dim strSQL                                  As String
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objProdutoPJTipoOperacao                As A6SubReserva.clsProdutoPJTipoOperacao
Dim strTipoOperacao                         As String
Dim strProdutoPJ                            As String

On Error GoTo ErrorHandler

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    strTipoOperacao = objTipoOperacao.LerTodos
    Set objTipoOperacao = Nothing

    Set objProdutoPJTipoOperacao = CreateObject("A6SubReserva.clsProdutoPJTipoOperacao")
    strProdutoPJ = objProdutoPJTipoOperacao.ObterProdutoPorStatus(plngStatusProduto)
    Set objProdutoPJTipoOperacao = Nothing

    ObterProdutosTipoOperacao = "<ProdutoPJTipoOperacao>" & strTipoOperacao & strProdutoPJ & "</ProdutoPJTipoOperacao>"

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterProdutosTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de produtos pj, agrupados por status e tipo operação,
' e pelo retorno à camada intermediária.

Public Function ObterProdutoPorStatus(ByVal plngStatusProduto As enumStatusProduto, _
                             Optional ByVal plngTipoOperacao As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    If plngStatusProduto = enumStatusProduto.NaoAssociado Then
        strSQL = " Select    A.CO_PROD,      " & vbNewLine & _
                "           A.DE_PROD,      " & vbNewLine & _
                "           '' AS TP_OPER,  " & vbNewLine & _
                "           '00:00:00' AS DT_INIC_VIGE, " & vbNewLine & _
                "           '00:00:00' AS DT_FIM_VIGE,  " & vbNewLine & _
                "           '00:00:00' AS DH_ULTI_ATLZ,  " & vbNewLine & _
                "           SYSDATE AS DT_BASE_DADO     " & vbNewLine & _
                " From      A8.TB_PRODUTO   A  " & vbNewLine & _
                " Where    (A.DT_FIM_VIGE   IS  NULL " & vbNewLine & _
                " Or        A.DT_FIM_VIGE    >=  SYSDATE) " & vbNewLine & _
                " And       A.CO_EMPR_FUSI   =  1 " & vbNewLine & _
                " And   A.CO_PROD   Not In (Select  Distinct(CO_PROD) " & vbNewLine & _
                "                           From    A6.TB_TIPO_OPER_PROD_PJ) "
    Else
        strSQL = " Select    A.CO_PROD,      " & vbNewLine & _
                "           A.DE_PROD,      " & vbNewLine & _
                "           B.TP_OPER,       " & vbNewLine & _
                "           B.DT_INIC_VIGE, " & vbNewLine & _
                "           B.DT_FIM_VIGE,  " & vbNewLine & _
                "           B.DH_ULTI_ATLZ,  " & vbNewLine & _
                "           SYSDATE AS DT_BASE_DADO     " & vbNewLine & _
                " From      A8.TB_PRODUTO                   A, " & vbNewLine & _
                "           A6.TB_TIPO_OPER_PROD_PJ  B  " & vbNewLine & _
                " Where    (A.DT_FIM_VIGE   IS  NULL " & vbNewLine & _
                " Or        A.DT_FIM_VIGE    >=  SYSDATE) " & vbNewLine & _
                " And       A.CO_EMPR_FUSI   =  1 " & vbNewLine
    End If

    Select Case plngStatusProduto
        Case enumStatusProduto.AssociadoSemFimVigencia
            strSQL = strSQL & " And   A.CO_PROD       =   B.CO_PROD " & vbNewLine & _
                            " And   B.DT_FIM_VIGE   IS  NULL " & vbNewLine & _
                            " And   B.TP_OPER        =   " & plngTipoOperacao & vbNewLine
        Case enumStatusProduto.AssociadoComFimVigencia
            strSQL = strSQL & " And   A.CO_PROD        =      B.CO_PROD " & vbNewLine & _
                            " And   B.DT_FIM_VIGE   IS NOT NULL " & vbNewLine & _
                            " And   B.TP_OPER         =   " & plngTipoOperacao & vbNewLine
    End Select

    strSQL = strSQL & " Order By  A.DE_PROD "

    ObterProdutoPorStatus = fgQueryXMLLerTodos("ProdutoPJ", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterProdutoPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
