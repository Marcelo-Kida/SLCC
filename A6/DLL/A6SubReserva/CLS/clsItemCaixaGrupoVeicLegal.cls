VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemCaixaGrupoVeicLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F4E1A4E006B"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Item Caixa x Grupo de Veículo Legal.

Option Explicit

Implements ObjectControl

Private objContext                          As COMSVCSLib.ObjectContext

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pela leitura de itens de caixa agrupados por grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterRelacaoItensCaixaGrupoVeicLegal(ByVal pstrXMLFiltros As String, _
                                                     ByVal pblnSegregarPorTipoBackOffice As Boolean, _
                                                     ByVal pintTipoBackOffice As enumTipoBackOffice) As String

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlRetorno                              As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim strxmlRetorno                           As String

Dim lngCodigoEmpresa                        As Long
Dim lngCodigoGrupoVeiculoLegal              As Long
Dim strCodigoVeiculoLegal                   As String
Dim strSiglaSistema                         As String

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim intTipoBackOffice                       As Integer

    On Error GoTo ErrorHandler

    Set xmlRetorno = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlRetorno, "", "Repeat_Consulta", "")
    
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    strSQL = " SELECT   A.CO_GRUP_VEIC_LEGA,                                            " & vbNewLine & _
             "          A.CO_ITEM_CAIX                                                  " & vbNewLine & _
             " FROM     A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA          A,                      " & vbNewLine & _
             "         (SELECT  A1.CO_GRUP_VEIC_LEGA                                    " & vbNewLine & _
             "          FROM    A8.TB_GRUP_VEIC_LEGA        A1,                         " & vbNewLine & _
             "                  A8.TB_VEIC_LEGA             B1                          " & vbNewLine & _
             "          WHERE   A1.CO_GRUP_VEIC_LEGA        = B1.CO_GRUP_VEIC_LEGA      "
             
    Call xmlDomFiltros.loadXML(pstrXMLFiltros)

    intTipoBackOffice = 99
    For Each xmlDomNode In xmlDomFiltros.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(xmlDomNode.Text)
    Next
    
    If intTipoBackOffice <> 99 Then
        pintTipoBackOffice = intTipoBackOffice
    End If
    
    If Not xmlDomFiltros.selectSingleNode("//GrupoVeiculoLegal") Is Nothing Then
        strSQL = strSQL & vbNewLine & _
                 " AND       B1.CO_GRUP_VEIC_LEGA = " & xmlDomFiltros.selectSingleNode("//GrupoVeiculoLegal").Text
                 
        lngCodigoGrupoVeiculoLegal = Val(xmlDomFiltros.selectSingleNode("//GrupoVeiculoLegal").Text)
    End If
    
    If Not xmlDomFiltros.selectSingleNode("//BancoLiquidante") Is Nothing Then
        strSQL = strSQL & vbNewLine & _
                 " AND       B1.CO_EMPR = " & xmlDomFiltros.selectSingleNode("//BancoLiquidante").Text
                 
        lngCodigoEmpresa = Val(xmlDomFiltros.selectSingleNode("//BancoLiquidante").Text)
    End If
    
    If pintTipoBackOffice <> 0 Then
        strSQL = strSQL & vbNewLine & _
                 " AND       B1.TP_BKOF           = " & pintTipoBackOffice
    End If
    
    If Not xmlDomFiltros.selectSingleNode("//VeiculoLegal") Is Nothing Then
        strSQL = strSQL & vbNewLine & _
                 " AND       B1.CO_VEIC_LEGA      = '" & xmlDomFiltros.selectSingleNode("//VeiculoLegal").Text & "'"
                 
        strCodigoVeiculoLegal = xmlDomFiltros.selectSingleNode("//VeiculoLegal").Text
    End If
    
    If Not xmlDomFiltros.selectSingleNode("//Sistema") Is Nothing Then
        strSQL = strSQL & vbNewLine & _
                 " AND       B1.SG_SIST      = '" & xmlDomFiltros.selectSingleNode("//Sistema").Text & "'"
                 
        strSiglaSistema = xmlDomFiltros.selectSingleNode("//Sistema").Text
    End If
    
    Select Case intTipoBackOffice
        Case 99
            strSQL = strSQL & vbNewLine & _
                     fgSegregaDados("A8.TB_VEIC_LEGA", _
                                    False, "B1", "B1", False, False, pblnSegregarPorTipoBackOffice, False, True, False)
        Case Is > 0
            strSQL = strSQL & _
                     vbNewLine & " AND B1.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    strSQL = strSQL & vbNewLine & _
             "          GROUP BY                                                        " & vbNewLine & _
             "                  A1.CO_GRUP_VEIC_LEGA)           B                       " & vbNewLine & _
             " WHERE    A.CO_GRUP_VEIC_LEGA = B.CO_GRUP_VEIC_LEGA"
    
    strSQL = strSQL & vbNewLine & _
             " ORDER BY  A.CO_GRUP_VEIC_LEGA"
    
    strxmlRetorno = fgQueryXMLLerTodos("ItensCaixa", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Trim(strxmlRetorno) <> vbNullString Then
        Call fgAppendXML(xmlRetorno, "Repeat_Consulta", strxmlRetorno)
    End If
    
    strxmlRetorno = vbNullString
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    strxmlRetorno = objVeiculoLegal.LerTodos(strCodigoVeiculoLegal, _
                                             lngCodigoGrupoVeiculoLegal, _
                                             , _
                                             IIf(intTipoBackOffice = 99, 0, intTipoBackOffice), _
                                             IIf(pintTipoBackOffice = 0, "N", "S"), _
                                             lngCodigoEmpresa, _
                                             strSiglaSistema)
    
    If Trim(strxmlRetorno) <> vbNullString Then
        Call fgAppendXML(xmlRetorno, "Repeat_Consulta", strxmlRetorno)
    End If
    
    ObterRelacaoItensCaixaGrupoVeicLegal = xmlRetorno.xml
    
    Set objVeiculoLegal = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objVeiculoLegal = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterRelacaoItensCaixaGrupoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de itens de caixa agrupados por grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterItensCaixaPorGrupoVeicLegal(ByVal plngGrupoVeiculo As Long) As String

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim strSegregacao                           As String

    On Error GoTo ErrorHandler

    strSQL = "SELECT DISTINCT A.CO_ITEM_CAIX,                            " & vbCrLf & _
             "                A.DE_ITEM_CAIX,                            " & vbCrLf & _
             "                A.CO_ITEM_CAIX_PAI ,                       " & vbCrLf & _
             "                A.TP_ITEM_CAIX,                            " & vbCrLf & _
             "                A.TP_BKOF, A.TP_CAIX                       " & vbCrLf & _
             "FROM            A6.TB_ITEM_CAIX_OPER_ATIV       A,         " & vbCrLf & _
             "                A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA  B          " & vbCrLf & _
             "WHERE         ((B.CO_GRUP_VEIC_LEGA = " & plngGrupoVeiculo & vbCrLf & _
             "AND             SUBSTR(B.CO_ITEM_CAIX, 1, 4)  = SUBSTR(A.CO_ITEM_CAIX, 1, 4))" & vbCrLf & _
             "OR             A.CO_ITEM_CAIX_PAI IS NULL)                 " & vbCrLf


    strSegregacao = fgSegregaDados("A6.TB_ITEM_CAIX_OPER_ATIV", _
                                   False, _
                                   "A", _
                                   "A", _
                                   False, _
                                   False, _
                                   True, _
                                   False, _
                                   False, _
                                   False)


    strSQL = strSQL & strSegregacao & vbCrLf


    strSQL = strSQL & _
            "ORDER BY  A.CO_ITEM_CAIX ASC "

    ObterItensCaixaPorGrupoVeicLegal = fgQueryXMLLerTodos("ItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterRelacaoItensCaixaGrupoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela de itens caixa x grupos veículo legal.

Public Function Salvar(pobjDomItemCaixaGrupoVeicLegal As Object) As Boolean

Dim xmlDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    For Each xmlDomNode In pobjDomItemCaixaGrupoVeicLegal.documentElement.selectNodes("//Repeat_ItemCaixaGrupoVeicLegal/*")
        Select Case xmlDomNode.selectSingleNode("@Operacao").Text
            Case "Incluir"
                flIncluir xmlDomNode
            Case "Excluir"
                flExcluir xmlDomNode.selectSingleNode("CO_ITEM_CAIX").Text, _
                xmlDomNode.selectSingleNode("CO_GRUP_VEIC_LEGA").Text, _
                xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
                
        End Select
    Next

    Set xmlDomNode = Nothing

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

' Este método é responsável pela leitura de itens de caixa agrupados por grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function LerTodos(Optional ByVal pstrCO_ITEM_CAIX As String, _
                         Optional ByVal pintGRUP_VEIC_LEGA As Integer) As String
                         
Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
            "   CO_ITEM_CAIX, " & vbCrLf & _
            "   CO_GRUP_VEIC_LEGA, " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
            "   DH_ULTI_ATLZ " & vbCrLf & _
            "FROM " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA " & vbCrLf
            
    If pstrCO_ITEM_CAIX <> "" And pintGRUP_VEIC_LEGA <> 0 Then
        strSQL = strSQL & _
            "WHERE " & _
            "   CO_ITEM_CAIX = '" & pstrCO_ITEM_CAIX & "' and " & vbCrLf & _
            "   CO_GRUP_VEIC_LEGA = " & pintGRUP_VEIC_LEGA & " " & vbCrLf
            
    ElseIf pstrCO_ITEM_CAIX = "" And pintGRUP_VEIC_LEGA <> 0 Then
        strSQL = strSQL & _
            "WHERE " & _
            "   CO_GRUP_VEIC_LEGA = " & pintGRUP_VEIC_LEGA & " " & vbCrLf
    
    ElseIf pstrCO_ITEM_CAIX <> "" And pintGRUP_VEIC_LEGA = 0 Then
        strSQL = strSQL & _
            "WHERE " & _
            "   CO_ITEM_CAIX = '" & pstrCO_ITEM_CAIX & "' " & vbCrLf
    End If
    
    strSQL = strSQL & _
            "ORDER BY " & vbCrLf & _
            "   CO_ITEM_CAIX " & vbCrLf

    LerTodos = fgQueryXMLLerTodos("ItemCaixaGrupoVeicLegal", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela inclusão de uma nova associação item de caixa x grupo de veículo legal.

Private Function flIncluir(ByVal pxmlItemCaixaGrupoVeicLegal As MSXML2.IXMLDOMNode) As Boolean
Dim strSQL                                   As String
On Error GoTo ErrorHandler

    strSQL = "INSERT INTO " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA ( " & vbCrLf & _
            "   CO_ITEM_CAIX, " & vbCrLf & _
            "   CO_GRUP_VEIC_LEGA, " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
            "   DH_ULTI_ATLZ " & vbCrLf & _
            ") VALUES ( '" & _
            pxmlItemCaixaGrupoVeicLegal.selectSingleNode("CO_ITEM_CAIX").Text & "', " & vbCrLf & _
            pxmlItemCaixaGrupoVeicLegal.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & ", " & vbCrLf & _
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

' Este método é responsável pela leitura de propriedades da tabela item de caixa x grupo de veículo legal,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT CO_ITEM_CAIX, " & vbCrLf & _
            "        CO_GRUP_VEIC_LEGA, " & vbCrLf & _
            "        DH_ULTI_ATLZ, " & vbCrLf & _
            "        CO_USUA_ULTI_ATLZ, " & vbCrLf & _
            "        CO_ETCA_TRAB_ULTI_ATLZ " & vbCrLf & _
            " FROM   A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA "

    ObterPropriedades = fgPropriedades("Grupo_ItemCaixaGrupoVeicLegal", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela exclusão de uma associação item de caixa x grupo de veículo legal.

Private Function flExcluir(ByVal pstrCodigoItemCaixa As String, _
                           ByVal pintGrupoVeicLegal As Integer, _
                           ByVal pstrDH_ULTI_ATLZ As String) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "DELETE FROM " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA " & vbCrLf & _
            "WHERE " & vbCrLf & _
            "   CO_ITEM_CAIX = '" & pstrCodigoItemCaixa & "' and " & vbCrLf & _
            "   CO_GRUP_VEIC_LEGA = " & pintGrupoVeicLegal & " and " & _
            "   DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pstrDH_ULTI_ATLZ)

    If fgExecuteSQL(strSQL) = 0 Then
       ' 4001 - Registro Inexistente ou excluído por outro usuário.
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

' Este método é responsável pela atualização geral da tabela item de caixa x grupo de veículo legal.

Public Function Executar(ByRef pxmlItemCaixaGrupoVeicLegal As MSXML2.DOMDocument40) As String
Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlItemCaixaGrupoVeicLegal.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir"
            Executar = Salvar(pxmlItemCaixaGrupoVeicLegal)
        Case "LerTodos"
            Executar = LerTodos(pxmlItemCaixaGrupoVeicLegal.documentElement.selectSingleNode("CO_ITEM_CAIX").Text, _
                                pxmlItemCaixaGrupoVeicLegal.documentElement.selectSingleNode("CO_GRUP_VEIC_LEGA").Text)
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
