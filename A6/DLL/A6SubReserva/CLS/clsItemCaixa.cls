VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1348C400BA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Itens de Caixa.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private Type udtCodigoItemCaixaAux
    CO_ITEM_CAIX                            As String * 16
End Type

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

' Este método é responsável pela leitura de propriedades da tabela de itens de caixa,
' e pelo retorno à camada intermediária.

Public Function ObterPropriedades() As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   CO_ITEM_CAIX, " & vbCrLf & _
             "   CO_ITEM_CAIX_PAI, " & vbCrLf & _
             "   TP_ITEM_CAIX, " & vbCrLf & _
             "   TP_BKOF, " & vbCrLf & _
             "   TP_CAIX, " & vbCrLf & _
             "   '' AS CO_NIVEL_01, " & vbCrLf & _
             "   DE_ITEM_CAIX, " & vbCrLf & _
             "   DT_INIC_VIGE, " & vbCrLf & _
             "   DT_FIM_VIGE, " & vbCrLf & _
             "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   DH_ULTI_ATLZ, " & vbCrLf & _
             "   '' as TP_VIGE " & vbCrLf & _
             "FROM " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV "

    ObterPropriedades = fgPropriedades("Grupo_ItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela leitura à tabela de itens de caixa,
' e pelo retorno à camada intermediária.

Public Function LerTodos(Optional ByVal pintTipoCaixa As enumTipoCaixa = 0, _
                         Optional ByVal pstrSomenteNivel_01 As String = "N", _
                         Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByVal pxmlFiltros As MSXML2.DOMDocument40 = Nothing) As String
                         
Dim strSQL                                  As String
Dim strSql02                                As String

Dim intTipoBackOffice                       As Integer
Dim xmlDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   CO_ITEM_CAIX, " & vbCrLf & _
             "   CO_ITEM_CAIX_PAI, " & vbCrLf & _
             "   DECODE(CO_ITEM_CAIX_PAI, NULL, NULL, SUBSTR(CO_ITEM_CAIX, 1, 4) || '000000000000') As CO_ITEM_CAIX_NIVE_01, " & vbCrLf & _
             "   TP_ITEM_CAIX, " & vbCrLf & _
             "   TP_CAIX, " & vbCrLf & _
             "   TP_BKOF, " & vbCrLf & _
             "   '' AS CO_NIVEL_01, " & vbCrLf & _
             "   DE_ITEM_CAIX, " & vbCrLf & _
             "   DT_INIC_VIGE, " & vbCrLf & _
             "   DT_FIM_VIGE, " & vbCrLf & _
             "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   DH_ULTI_ATLZ, " & vbCrLf & _
             "   '' as TP_VIGE " & vbCrLf & _
             "FROM " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf
            
    If pintTipoCaixa = enumTipoCaixa.CaixaFuturo Or pintTipoCaixa = enumTipoCaixa.CaixaSubReserva Then
        strSql02 = _
             "WHERE " & vbCrLf & _
             "   TP_CAIX = " & pintTipoCaixa & " " & vbCrLf
    ElseIf pintTipoCaixa <> 0 Then
        '2018 - Tipo de Caixa deve ser Futuro ou SubReserva
        lngCodigoErroNegocio = 2018
        GoTo ErrorHandler
    End If
        
    If pstrSomenteNivel_01 = "S" Then
        If Len(strSql02) = 0 Then
            strSql02 = _
                "WHERE " & vbCrLf & _
                "   SUBSTR(CO_ITEM_CAIX, 5, 12) = '" & String(12, "0") & "' " & vbCrLf & _
                "AND " & vbCrLf & _
                "   CO_ITEM_CAIX_PAI Is Not Null " & vbCrLf & _
                "AND " & vbCrLf & _
                "   DE_ITEM_CAIX Not In ('" & gstrCaixaSubReserva & "', '" & gstrCaixaFuturo & "') " & vbCrLf
        Else
            strSql02 = strSql02 & _
                "AND " & vbCrLf & _
                "   SUBSTR(CO_ITEM_CAIX, 5, 12) = '" & String(12, "0") & "' " & vbCrLf & _
                "AND " & vbCrLf & _
                "   CO_ITEM_CAIX_PAI Is Not Null " & vbCrLf & _
                "AND " & vbCrLf & _
                "   DE_ITEM_CAIX Not In ('" & gstrCaixaSubReserva & "', '" & gstrCaixaFuturo & "') " & vbCrLf
                
        End If
    End If
    
    If pstrIndicadorVigencia = "S" Or Len(pstrIndicadorVigencia) = 0 Then
        If Len(strSql02) = 0 Then
            strSql02 = " where "
        Else
            strSql02 = strSql02 & " and "
        End If
       
       strSql02 = strSql02 & " DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                         " AND  (DT_FIM_VIGE    IS  NULL       " & _
                         " OR    DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & vbNewLine
    End If
    
    If pxmlFiltros Is Nothing Then
        strSql02 = strSql02 & fgSegregaDados("A6.TB_ITEM_CAIX_OPER_ATIV", False, "", "", IIf(Len(strSql02) = 0, True, False), False, True, False, False, False)
    Else
        intTipoBackOffice = 99
        For Each xmlDomNode In pxmlFiltros.selectNodes("//Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(xmlDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                strSql02 = strSql02 & fgSegregaDados("A6.TB_ITEM_CAIX_OPER_ATIV", False, "", "", IIf(Len(strSql02) = 0, True, False), False, True, False, False, False)
            Case Is > 0
                strSql02 = strSql02 & IIf(Len(strSql02) = 0, " WHERE ", " AND ") & " TP_BKOF = " & intTipoBackOffice
        End Select
    End If

    strSQL = strSQL & strSql02 & _
            " ORDER BY " & vbCrLf & _
            "    CO_ITEM_CAIX "

    LerTodos = fgQueryXMLLerTodos("ItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))
    
    If LerTodos = "" Then
        flVerificaExistenciaItensPadroes
        LerTodos = fgQueryXMLLerTodos("ItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de detalhes de um determinado item de caixa,
' e pelo retorno à camada intermediária.

Public Function Ler(ByVal psCodigoItemCaixa As String) As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   CO_ITEM_CAIX, " & vbCrLf & _
             "   CO_ITEM_CAIX_PAI, " & vbCrLf & _
             "   TP_ITEM_CAIX, " & vbCrLf & _
             "   TP_CAIX, " & vbCrLf & _
             "   TP_BKOF, " & vbCrLf & _
             "   DE_ITEM_CAIX, " & vbCrLf & _
             "   DT_INIC_VIGE, " & vbCrLf & _
             "   DT_FIM_VIGE, " & vbCrLf & _
             "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   DH_ULTI_ATLZ, " & vbCrLf & _
             "   '' as TP_VIGE " & vbCrLf & _
             "FROM " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
             "WHERE " & vbCrLf & _
             "   CO_ITEM_CAIX = '" & psCodigoItemCaixa & "' "

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

' Este método é responsável pela atualização geral da tabela de itens de caixa,
' e pelo retorno à camada intermediária.

Public Function Executar(ByRef pxmlItemCaixa As MSXML2.DOMDocument40) As String
Dim strOperacao                             As String
Dim strTipoCaixa                            As enumTipoCaixa

On Error GoTo ErrorHandler

    strOperacao = pxmlItemCaixa.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Alterar"
            Executar = flAlterar(pxmlItemCaixa)
        Case "Incluir"
            Executar = flIncluir(pxmlItemCaixa)
        Case "Excluir"
            Executar = flExcluir(pxmlItemCaixa)
        Case "Ler"
            Executar = Ler(pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text)
        Case "LerTodos"
            strTipoCaixa = Val(pxmlItemCaixa.selectSingleNode("//TP_CAIX").Text)
            Executar = LerTodos(strTipoCaixa, _
                                pxmlItemCaixa.selectSingleNode("//CO_NIVEL_01").Text, _
                                pxmlItemCaixa.selectSingleNode("//TP_VIGE").Text, _
                                pxmlItemCaixa)
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
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão de um novo item de caixa.

Private Function flIncluir(ByRef pxmlItemCaixa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String
Dim strCodigoItemCaixaPai                   As String
Dim strItemCaixaFilho                       As String
Dim strDescricaoItemCaixa                   As String

On Error GoTo ErrorHandler

    strCodigoItemCaixaPai = pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/CO_ITEM_CAIX_PAI").Text
    strDescricaoItemCaixa = Trim$(pxmlItemCaixa.selectSingleNode("//DE_ITEM_CAIX").Text)
    strItemCaixaFilho = flObterCodigoItemCaixaFilho(strCodigoItemCaixaPai, _
                                                    pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/TP_ITEM_CAIX").Text, _
                                                    CInt("0" & pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/TP_CAIX").Text))
    
    pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/CO_ITEM_CAIX").Text = strItemCaixaFilho
    strCodigoItemCaixaPai = IIf(strCodigoItemCaixaPai = vbNullString, "Null", ("'" & strCodigoItemCaixaPai & "'"))
    
    flRegrasNegocioItemCaixa CInt(pxmlItemCaixa.selectSingleNode("//TP_BKOF").Text), _
                             CInt(pxmlItemCaixa.selectSingleNode("//TP_CAIX").Text), _
                             Val(pxmlItemCaixa.selectSingleNode("//TP_ITEM_CAIX").Text), _
                             pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX_PAI").Text, _
                             pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text, _
                             strDescricaoItemCaixa, _
                             pxmlItemCaixa.selectSingleNode("//DT_FIM_VIGE").Text, _
                             False
    
    strSQL = "INSERT INTO " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV ( " & vbCrLf & _
             "   CO_ITEM_CAIX, " & vbCrLf & _
             "   CO_ITEM_CAIX_PAI, " & vbCrLf & _
             "   TP_ITEM_CAIX, " & vbCrLf & _
             "   TP_CAIX, " & vbCrLf & _
             "   TP_BKOF, " & vbCrLf & _
             "   DE_ITEM_CAIX, " & vbCrLf & _
             "   DT_INIC_VIGE, " & vbCrLf & _
             "   DT_FIM_VIGE, " & vbCrLf & _
             "   CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "   CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
             "   DH_ULTI_ATLZ " & vbCrLf & _
             ") VALUES ( " & vbCrLf & _
             "   '" & strItemCaixaFilho & "', " & vbCrLf & _
             "   " & strCodigoItemCaixaPai & ", " & vbCrLf & _
             "   " & CLng("0" & pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/TP_ITEM_CAIX").Text) & ", " & vbCrLf & _
             "   " & CLng("0" & pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/TP_CAIX").Text) & ", " & vbCrLf & _
             "   " & CLng("0" & pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/TP_BKOF").Text) & ", " & vbCrLf & _
             "   '" & pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/DE_ITEM_CAIX").Text & "', " & vbCrLf & _
             fgDtHrXML_To_Oracle(pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/DT_INIC_VIGE").Text) & ", " & vbCrLf & _
             fgDtHrXML_To_Oracle(pxmlItemCaixa.selectSingleNode("//Grupo_ItemCaixa/DT_FIM_VIGE").Text) & ", " & vbCrLf & _
             "   '" & fgUsuarioRede & "', " & vbCrLf & _
             "   '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "   SYSDATE) " & vbCrLf

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é aciona a exlusão de um item de caixa, ou de um grupo de item de caixa.

Private Function flExcluir(ByRef pxmlItemCaixa As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlItemCaixa.documentElement.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar Then
        flExcluirItemCaixa pxmlItemCaixa.documentElement.selectSingleNode("CO_ITEM_CAIX").Text, _
                           pxmlItemCaixa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                           CInt(pxmlItemCaixa.documentElement.selectSingleNode("TP_BKOF").Text)
                           
    ElseIf pxmlItemCaixa.documentElement.selectSingleNode("TP_ITEM_CAIX").Text = enumTipoItemCaixa.Grupo Then
        flExcluirGrupoItemCaixa pxmlItemCaixa.documentElement.selectSingleNode("CO_ITEM_CAIX").Text, _
                                CInt(pxmlItemCaixa.documentElement.selectSingleNode("TP_BKOF").Text)
    
    Else
        '2014 - Item de Caixa deve ser de Grupo ou Elementar
        lngCodigoErroNegocio = 2014
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

' Este método é responsável pela exclusão de um item de caixa.

Private Function flExcluirItemCaixa(ByVal pstrCodigoItemCaixa As String, _
                                    ByVal pstrDH_ULTI_ATLZ As String, _
                                    ByVal pintTipoBackOffice As Integer) As Boolean
                                    
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "DELETE FROM " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
            "WHERE " & vbCrLf & _
            "   CO_ITEM_CAIX = '" & pstrCodigoItemCaixa & "' " & vbCrLf & _
            "  AND " & vbCrLf & _
            "   TP_BKOF      =   " & pintTipoBackOffice & _
            "  AND " & vbCrLf & _
            "   DH_ULTI_ATLZ  = " & fgDtHrXML_To_Oracle(pstrDH_ULTI_ATLZ)

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirItemCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela alteração de um item de caixa existente.

Private Function flAlterar(ByVal pxmlItemCaixa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    flRegrasNegocioItemCaixa CInt(pxmlItemCaixa.selectSingleNode("//TP_BKOF").Text), _
                             CInt(pxmlItemCaixa.selectSingleNode("//TP_CAIX").Text), _
                             Val(pxmlItemCaixa.selectSingleNode("//TP_ITEM_CAIX").Text), _
                             pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX_PAI").Text, _
                             pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text, _
                             pxmlItemCaixa.selectSingleNode("//DE_ITEM_CAIX").Text, _
                             pxmlItemCaixa.selectSingleNode("//DT_FIM_VIGE").Text, _
                             True
                                  
    strSQL = "UPDATE " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
            "SET " & vbCrLf & _
            "   CO_ITEM_CAIX_PAI = '" & pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX_PAI").Text & "', " & vbCrLf & _
            "   TP_ITEM_CAIX = " & CLng(pxmlItemCaixa.selectSingleNode("//TP_ITEM_CAIX").Text) & ", " & vbCrLf & _
            "   TP_CAIX = " & CLng(pxmlItemCaixa.selectSingleNode("//TP_CAIX").Text) & ", " & vbCrLf & _
            "   TP_BKOF = " & CLng(pxmlItemCaixa.selectSingleNode("//TP_BKOF").Text) & ", " & vbCrLf & _
            "   DE_ITEM_CAIX = '" & pxmlItemCaixa.selectSingleNode("//DE_ITEM_CAIX").Text & "', " & vbCrLf & _
            "   DT_INIC_VIGE = " & fgDtHrXML_To_Oracle(pxmlItemCaixa.selectSingleNode("//DT_INIC_VIGE").Text) & ", " & vbCrLf & _
            "   DT_FIM_VIGE = " & fgDtHrXML_To_Oracle(pxmlItemCaixa.selectSingleNode("//DT_FIM_VIGE").Text) & ", " & vbCrLf & _
            "   CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "', " & vbCrLf & _
            "   CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
            "   DH_ULTI_ATLZ = SYSDATE " & vbCrLf & _
            "WHERE " & vbCrLf & _
            "   CO_ITEM_CAIX = '" & pxmlItemCaixa.selectSingleNode("//CO_ITEM_CAIX").Text & "' " & vbCrLf & _
            "AND   " & vbCrLf & _
            "   DH_ULTI_ATLZ  = " & fgDtHrXML_To_Oracle(pxmlItemCaixa.selectSingleNode("//DH_ULTI_ATLZ").Text)

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

' Este método é responsável pela alteração do final de vigência de um grupo de item de caixa.

Private Sub flAlterarFimVigenciaItemGrupo(ByVal pstrCodItemCaixaPai As String, _
                                          ByVal pstrDataFinalVigencia As String)
                                         
Dim strSQL                                  As String
Dim strCodigoSelect                         As String

On Error GoTo ErrorHandler

    If pstrDataFinalVigencia = vbNullString Then
        pstrDataFinalVigencia = "NULL"
    Else
        pstrDataFinalVigencia = fgDtHrXML_To_Oracle(pstrDataFinalVigencia)
    End If
    
    strCodigoSelect = fgObterCodigoItemCaixaSelect(pstrCodItemCaixaPai)
        
    'Altera todos os filhos do Item de Grupo, o Item de Grupo será alterado no Update
    strSQL = " UPDATE   A6.TB_ITEM_CAIX_OPER_ATIV " & _
             " SET      DT_FIM_VIGE  =  " & pstrDataFinalVigencia & vbCrLf & _
             " WHERE    CO_ITEM_CAIX_PAI IS NOT NULL " & vbCrLf & _
             " AND      CO_ITEM_CAIX     LIKE  '" & strCodigoSelect & "%'" & vbCrLf & _
             " AND      CO_ITEM_CAIX     <>    '" & pstrCodItemCaixaPai & "'"

    Call fgExecuteSQL(strSQL)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarFimVigenciaItemGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

' Este método é responsável pela exclusão de um grupo de item de caixa.

Private Function flExcluirGrupoItemCaixa(ByVal pstrCodigoItemCaixa As String, _
                                         ByVal pintTipoBackOffice As Integer) As Boolean
                                         
Dim rsItemCaixa                             As ADODB.Recordset
Dim strCodigoSelect                         As String
Dim strCodigoItemCaixa                      As String
Dim strSQL                                  As String
Dim intRecordCount                          As Integer

On Error GoTo ErrorHandler

    strCodigoSelect = fgObterCodigoItemCaixaSelect(pstrCodigoItemCaixa)

    strSQL = " SELECT " & vbCrLf & _
             "           CO_ITEM_CAIX, " & vbCrLf & _
             "           DH_ULTI_ATLZ " & vbCrLf & _
             " FROM      A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
             " WHERE     " & vbCrLf & _
             "           CO_ITEM_CAIX_PAI IS NOT NULL " & vbCrLf & _
             " AND       CO_ITEM_CAIX LIKE '" & strCodigoSelect & "%'" & _
             " AND       TP_BKOF      =   " & pintTipoBackOffice & _
             " ORDER BY  CO_ITEM_CAIX  DESC"

    Set rsItemCaixa = fgQuerySQL(strSQL)

    If rsItemCaixa.EOF Then
        '4001 - Registro foi alterado ou excluído por outro usuário.
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Do Until rsItemCaixa.EOF
                        
        strCodigoItemCaixa = rsItemCaixa.fields!CO_ITEM_CAIX
                
        strSQL = " DELETE " & _
                 " FROM     A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
                 " WHERE    " & vbCrLf & _
                 "          CO_ITEM_CAIX   =  " & strCodigoItemCaixa & vbCrLf

        'Quando for o Item de Caixa Pai de Todos, Verificar a pstrDH_ULTI_ATLZ,
        'para garantir concorrência de processos
        
        intRecordCount = intRecordCount + 1
        If rsItemCaixa.RecordCount = intRecordCount Then
            strSQL = strSQL & "    AND    DH_ULTI_ATLZ   = " & fgDtHrXML_To_Oracle(fgDtHr_To_Xml(rsItemCaixa.fields!DH_ULTI_ATLZ)) & vbCrLf
        End If
        
        If fgExecuteSQL(strSQL) = 0 Then
            '4001 - Registro foi alterado ou excluído por outro usuário.
            lngCodigoErroNegocio = 4001
            GoTo ErrorHandler
        End If
        
        rsItemCaixa.MoveNext
    Loop
    
    Set rsItemCaixa = Nothing
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    Set rsItemCaixa = Nothing
    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirGrupoItemCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Este método é responsável pela leitura do código do item de caixa, a partir dos parâmetros passados,
' e pelo retorno à camada intermediária.

Public Function ObterItemCaixa(Optional ByVal plCodigoLocalLiquidacao As Long, _
                               Optional ByVal piIndicadorEntradaSaida As Integer) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
            "   CO_ITEM_CAIX " & vbCrLf & _
            "FROM " & vbCrLf & _
            "   A6.TB_ITEM_CAIX_TIPO_OPER " & vbCrLf & _
            "WHERE " & vbCrLf & _
            "   CO_LOCA_LIQU = " & plCodigoLocalLiquidacao & " AND " & vbCrLf & _
            "   IN_ENTR_SAID = " & piIndicadorEntradaSaida & " " & vbCrLf

    ObterItemCaixa = fgQueryXMLLer("CodigoItemCaixa", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterItemCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela obtenção do próximo código de item de caixa disponível, sequencialmente.

Private Function flObterCodigoItemCaixaFilho(ByVal pstrCodigoItemCaixaPai As String, _
                                             ByVal pintTipoItemCaixa As enumTipoItemCaixa, _
                                             ByVal pintTipoCaixa As enumTipoCaixa) As String

Dim rsUltimoFilho                           As ADODB.Recordset

Dim udtCodigoItemCaixaAux                   As udtCodigoItemCaixaAux
Dim strSQL                                  As String

Dim blnSomaNoPrimeiroNivel                  As Boolean
Dim intNivelItemCaixaPai                    As Integer
Dim intNivelItemCaixaFilho                  As Integer

Dim blnPrimeiroFilho                        As Boolean

On Error GoTo ErrorHandler

    If Len(pstrCodigoItemCaixaPai) > 0 Then
        'Inserindo Filhos (Se é Filho do Rótulo Padrão - Max deve ser somente por Tipo Caixa)
        strSQL = "SELECT " & vbCrLf & _
                "   DE_ITEM_CAIX " & vbCrLf & _
                "FROM " & vbCrLf & _
                "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
                "WHERE " & vbCrLf & _
                "   CO_ITEM_CAIX = " & pstrCodigoItemCaixaPai

        Set rsUltimoFilho = fgQuerySQL(strSQL)
        
        If Not rsUltimoFilho.EOF Then
            If pintTipoCaixa = CaixaSubReserva Then
                If rsUltimoFilho!DE_ITEM_CAIX = gstrCaixaSubReserva Then
                    blnSomaNoPrimeiroNivel = True
                End If
            ElseIf pintTipoCaixa = CaixaFuturo Then
                If rsUltimoFilho!DE_ITEM_CAIX = gstrCaixaFuturo Then
                    blnSomaNoPrimeiroNivel = True
                End If
            End If
        Else
            blnSomaNoPrimeiroNivel = True
        End If
        
        '(Se é Filho do Rótulo Padrão - Max deve ser somente por Tipo Caixa)
        If blnSomaNoPrimeiroNivel Then
            strSQL = "SELECT " & vbCrLf & _
                    "   MAX(CO_ITEM_CAIX) AS MaiorFilho" & vbCrLf & _
                    "   FROM " & vbCrLf & _
                    "       A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
                    "  WHERE " & vbCrLf & _
                    "        TP_CAIX = " & pintTipoCaixa & vbCrLf & _
                    "   AND  SUBSTR(CO_ITEM_CAIX, 5, 12) = '" & String(12, "0") & "' "
        Else
            strSQL = "SELECT " & vbCrLf & _
                    "   MAX(CO_ITEM_CAIX) AS MaiorFilho" & vbCrLf & _
                    "FROM " & vbCrLf & _
                    "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
                    "WHERE " & vbCrLf & _
                    "   CO_ITEM_CAIX_PAI = " & pstrCodigoItemCaixaPai
        End If
        
    Else
        'Inserção dos Rótulos Padrões de Item de Caixa (Caixa Sub-reserva e Caixa Futuro)
        strSQL = "SELECT " & vbCrLf & _
                "   MAX(CO_ITEM_CAIX) AS MaiorFilho" & vbCrLf & _
                "   FROM " & vbCrLf & _
                "       A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
                "  WHERE " & vbCrLf & _
                "        TP_CAIX = " & pintTipoCaixa & vbCrLf & _
                "   AND  SUBSTR(CO_ITEM_CAIX, 5, 12) = '" & String(12, "0") & "' "
    End If
    
    Set rsUltimoFilho = fgQuerySQL(strSQL)
    
    If pstrCodigoItemCaixaPai = "" Then
        'Inserção dos Rótulos Padrões de Item de Caixa (Caixa Sub-reserva e Caixa Futuro)
        If IsNull(rsUltimoFilho.fields!MaiorFilho) Then
            udtCodigoItemCaixaAux.CO_ITEM_CAIX = pintTipoCaixa & String(15, "0")
            flObterCodigoItemCaixaFilho = udtCodigoItemCaixaAux.CO_ITEM_CAIX
            Exit Function
        Else
            LSet udtCodigoItemCaixaAux.CO_ITEM_CAIX = rsUltimoFilho.fields!MaiorFilho
            blnSomaNoPrimeiroNivel = True
        End If
    Else
        'Inserindo Filhos (Verificar se é o primeiro filho)
        If IsNull(rsUltimoFilho.fields!MaiorFilho) Then
            LSet udtCodigoItemCaixaAux.CO_ITEM_CAIX = pstrCodigoItemCaixaPai
            blnPrimeiroFilho = True
        Else
            LSet udtCodigoItemCaixaAux.CO_ITEM_CAIX = rsUltimoFilho.fields!MaiorFilho
            blnPrimeiroFilho = False
        End If
                
    End If
    
    intNivelItemCaixaPai = fgObterNivelItemCaixa(pstrCodigoItemCaixaPai)
        
    If pintTipoItemCaixa = Grupo And intNivelItemCaixaPai = 4 Then
        '2015 - Só são permitidos 5 níveis, o último nível deve ser elementar
        lngCodigoErroNegocio = 2015
        GoTo ErrorHandler
    Else
            
        If blnSomaNoPrimeiroNivel Then
            Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, intNivelItemCaixaPai + 1, 3) = Format(Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, intNivelItemCaixaPai + 1, 3) + 1, "000")
        Else
            
            intNivelItemCaixaFilho = fgObterNivelItemCaixa(udtCodigoItemCaixaAux.CO_ITEM_CAIX)
            
            If blnPrimeiroFilho Then
                'Somar à Direito do Nível
                intNivelItemCaixaFilho = intNivelItemCaixaFilho + 1
            Else
                'Somar no Próprio Nível
                intNivelItemCaixaFilho = intNivelItemCaixaFilho
            End If
            
            Select Case intNivelItemCaixaFilho
                Case 5
                    udtCodigoItemCaixaAux.CO_ITEM_CAIX = Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 1, 13) & Format(Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 14, 3) + 1, "000") & String(15, "0")
                Case 4
                    udtCodigoItemCaixaAux.CO_ITEM_CAIX = Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 1, 10) & Format(Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 11, 3) + 1, "000") & String(15, "0")
                Case 3
                    udtCodigoItemCaixaAux.CO_ITEM_CAIX = Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 1, 7) & Format(Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 8, 3) + 1, "000") & String(15, "0")
                Case 2
                    udtCodigoItemCaixaAux.CO_ITEM_CAIX = Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 1, 4) & Format(Mid(udtCodigoItemCaixaAux.CO_ITEM_CAIX, 5, 3) + 1, "000") & String(15, "0")
                Case 1
                    Err.Raise vbObjectError, "A6SubReserva.clsItemCaixa.flObterCodigoItemCaixaFilho", "Erro não tratado"
            End Select
        End If
        
    End If
            
    flObterCodigoItemCaixaFilho = udtCodigoItemCaixaAux.CO_ITEM_CAIX
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
        
    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterCodigoItemCaixaFilho Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    

End Function

' Este método é responsável pela obtenção do código de item de caixa genérico.

Public Function fgObterItemCaixaGenerico(ByVal pintTipoCaixa As enumTipoCaixa, _
                                         ByVal pstrCodigoVeiculoLegal As String, _
                                Optional ByVal pstrSiglaSistema As String = vbNullString) As String

Dim strSQL                                  As String
Dim xmlAux                                  As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    strSQL = " SELECT CO_ITEM_CAIX              " & vbNewLine & _
             " FROM   A6.TB_ITEM_CAIX_OPER_ATIV " & vbNewLine & _
             " WHERE  TP_CAIX = " & pintTipoCaixa & vbNewLine & _
             " AND    TP_BKOF = " & fgObterTipoBackOffice(pstrCodigoVeiculoLegal, pstrSiglaSistema) & vbNewLine & _
             " AND    DE_ITEM_CAIX = '" & gstrItemGenerico & "'"

    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlAux.loadXML(fgQueryXMLLer("ItemCaixaGenerico", strSQL, App.EXEName & "." & TypeName(Me)))

    If xmlAux.xml <> vbNullString Then
        fgObterItemCaixaGenerico = xmlAux.documentElement.selectSingleNode("CO_ITEM_CAIX").Text
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsItemCaixa", "fgObterItemCaixaGenerico", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela verificação de itens de caixa padrão.

Private Sub flVerificaExistenciaItensPadroes()
Dim strSQL                                  As String
Dim rsItemCaixa                             As ADODB.Recordset
Dim xmlItemCaixa                            As MSXML2.DOMDocument40

Dim strTipoBackOffice                       As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   Count(*) AS TotalItens " & vbCrLf & _
             "FROM " & vbCrLf & _
             "   A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf
    
    strSQL = strSQL & fgSegregaDados("A6.TB_ITEM_CAIX_OPER_ATIV", False, "", "", True, False, True, False, False, False)

    Set rsItemCaixa = CreateObject("ADODB.Recordset")
    
    Set rsItemCaixa = fgQuerySQL(strSQL)
    
    If rsItemCaixa.fields!TotalItens = 0 Then
        
        Set xmlItemCaixa = CreateObject("MSXML2.DOMDocument.4.0")
        xmlItemCaixa.loadXML (ObterPropriedades)
        
        strTipoBackOffice = fgObterTipoBackOfficeUsuario
                
        With xmlItemCaixa.documentElement
            'Propriedades Comuns
            .selectSingleNode("//@Operacao").Text = "Incluir"
            .selectSingleNode("//DT_INIC_VIGE").Text = Format(fgDataHoraServidor(Data), "YYYYMMDDHHNNSS")
            .selectSingleNode("//DT_FIM_VIGE").Text = vbNullString
            .selectSingleNode("//TP_BKOF").Text = strTipoBackOffice
        
            'Cria Item de Grupo Caixa Reserva
            .selectSingleNode("//CO_ITEM_CAIX_PAI").Text = vbNullString
            .selectSingleNode("//TP_CAIX").Text = enumTipoCaixa.CaixaSubReserva
            .selectSingleNode("//DE_ITEM_CAIX").Text = gstrCaixaSubReserva
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Grupo
            
            flIncluir xmlItemCaixa
            
            'Cria Item Elementar Genérico
            .selectSingleNode("//CO_ITEM_CAIX_PAI").Text = .selectSingleNode("//CO_ITEM_CAIX").Text
            .selectSingleNode("//DE_ITEM_CAIX").Text = gstrItemGenerico
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar
            
            flIncluir xmlItemCaixa
            
            'Cria Item de Grupo Caixa Futuro
            .selectSingleNode("//CO_ITEM_CAIX_PAI").Text = vbNullString
            .selectSingleNode("//TP_CAIX").Text = enumTipoCaixa.CaixaFuturo
            .selectSingleNode("//DE_ITEM_CAIX").Text = gstrCaixaFuturo
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Grupo
            
            flIncluir xmlItemCaixa
        
            'Cria Item Elementar Genérico
            .selectSingleNode("//CO_ITEM_CAIX_PAI").Text = .selectSingleNode("//CO_ITEM_CAIX").Text
            .selectSingleNode("//DE_ITEM_CAIX").Text = gstrItemGenerico
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar
        
            flIncluir xmlItemCaixa
        
        End With
        
    End If
        
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Sub

ErrorHandler:

    If Not objContext Is Nothing Then
        ''objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaExistenciaItensPadroes Function", lngCodigoErroNegocio, intNumeroSequencialErro)
        
End Sub

' Este método é responsável pela verificação da existência de uma determinada descrição de item de caixa.

Private Function flExisteDescricaoIgualNoGrupo(ByVal pstrCodItemCaixaPai As String, _
                                               ByVal penumTipoBackOffice As enumTipoBackOffice, _
                                               ByVal penumTipoCaixa As enumTipoCaixa, _
                                               ByVal penumTipoItemCaixa As enumTipoItemCaixa, _
                                               ByVal pstrDescricao As String, _
                                               ByVal blnAlterar As Boolean, _
                                               ByVal pstrCodItemCaixa As String) As Boolean
                                         
Dim rsItemCaixa                             As ADODB.Recordset
Dim udtCodigoItemCaixaAux                   As udtCodigoItemCaixaAux
Dim strCodigoSelect                         As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strCodigoSelect = fgObterCodigoItemCaixaSelect(pstrCodItemCaixaPai)

    strSQL = "  SELECT " & vbCrLf & _
            "           CO_ITEM_CAIX " & vbCrLf & _
            "     FROM  A6.TB_ITEM_CAIX_OPER_ATIV " & vbCrLf & _
            "    WHERE " & vbCrLf & _
            "           CO_ITEM_CAIX_PAI IS NOT NULL " & vbCrLf & _
            "      AND  TP_BKOF      =   " & penumTipoBackOffice & vbCrLf & _
            "      AND  TP_CAIX      =   " & penumTipoCaixa & vbCrLf & _
            "      AND  UPPER(DE_ITEM_CAIX) =   '" & UCase$(pstrDescricao) & "'" & vbCrLf

    If blnAlterar Then
        strSQL = strSQL & " AND NOT (CO_ITEM_CAIX  = '" & pstrCodItemCaixa & "') "
    Else
        udtCodigoItemCaixaAux.CO_ITEM_CAIX = strCodigoSelect & String(15, "9")
        strSQL = strSQL & "      AND  CO_ITEM_CAIX_PAI = '" & pstrCodItemCaixaPai & "'"
    End If

    Set rsItemCaixa = fgQuerySQL(strSQL)

    If Not rsItemCaixa.EOF Then
        '2017 - Já existe um Item de Caixa com está descrição neste grupo
        lngCodigoErroNegocio = 2017
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteDescricaoIgualNoGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método é responsável pela verificação das regras de negócio para a atualização da tabela de itens de caixa.

Private Sub flRegrasNegocioItemCaixa(ByVal pTipoBackOfice As Integer, _
                                     ByVal penumTipoCaixa As enumTipoCaixa, _
                                     ByVal penumTipoItemCaixa As enumTipoItemCaixa, _
                                     ByVal pstrCodItemCaixaPai As String, _
                                     ByVal pstrCodItemCaixa As String, _
                                     ByVal pstrDesItemCaixa As String, _
                                     ByVal pstrDataFinalVigencia As String, _
                            Optional ByVal pblnAlterar As Boolean = False)

Dim rsItemCaixa                             As ADODB.Recordset
Dim strSQL                                  As String
Dim strCodigoSelect                         As String
    
On Error GoTo ErrorHandler
    
    'Consistências comuns para Item de Grupo e Item Elementar
    
    flExisteDescricaoIgualNoGrupo pstrCodItemCaixaPai, _
                                  pTipoBackOfice, _
                                  penumTipoCaixa, _
                                  penumTipoItemCaixa, _
                                  pstrDesItemCaixa, _
                                  pblnAlterar, _
                                  pstrCodItemCaixa


    If penumTipoItemCaixa = Elementar Then
        'Se Item Elementar Verificar se Item de Grupo Pai está ativo ou possui data de vigência inferior.
        strSQL = " SELECT DT_FIM_VIGE AS DT_FIM_VIGE_PAI " & vbNewLine & _
                 " FROM   A6.TB_ITEM_CAIX_OPER_ATIV      " & vbNewLine & _
                 " WHERE  CO_ITEM_CAIX     =            '" & pstrCodItemCaixaPai & "'" & vbNewLine & _
                 "   AND  TP_BKOF          =             " & pTipoBackOfice

        Set rsItemCaixa = fgQuerySQL(strSQL)
        If Not rsItemCaixa.EOF Then
            If Not IsNull(rsItemCaixa!DT_FIM_VIGE_PAI) Then
                If pstrDataFinalVigencia = vbNullString Then
                    If Not IsNull(rsItemCaixa.fields!DT_FIM_VIGE_PAI) Then  'Data de fim de Vigência do Pai
                        '2025 - Item de Caixa pertence a um Item de Grupo Desativado, não pode ser Ativado
                        lngCodigoErroNegocio = 2025
                        GoTo ErrorHandler
                    End If
                ElseIf pstrDataFinalVigencia > fgDtHr_To_Xml(rsItemCaixa.fields!DT_FIM_VIGE_PAI) Then 'Data de fim de Vigência do Pai
                    '2026 - Item de Caixa pertence a um Item de Grupo com vigência menor, a vigência do Item de Caixa deve ser  menor ou igual que a vigência do Item de Grupo a que ele pertence.
                    lngCodigoErroNegocio = 2026
                    GoTo ErrorHandler
                End If
            End If
        End If
        
        If pstrDataFinalVigencia <> vbNullString Then
            'Verifica associação do Item de Caixa Elementar com Tipo Operação
            strSQL = " SELECT CO_ITEM_CAIX              " & vbNewLine & _
                     " FROM   A6.TB_ITEM_CAIX_TIPO_OPER " & vbNewLine & _
                     " WHERE  CO_ITEM_CAIX  =           '" & pstrCodItemCaixa & "'"
    
            Set rsItemCaixa = fgQuerySQL(strSQL)
        
            If Not rsItemCaixa.EOF Then
                '2023 - Item de Caixa possui associações com Tipos de Operação. Finalização de Vigência não permitida.
                lngCodigoErroNegocio = 2023
                GoTo ErrorHandler
            End If
        End If
        
    ElseIf penumTipoItemCaixa = Grupo Then
        'Verifica associação dos Itens de Caixa Elementares pertencentes ao Item de Grupo com Tipo Operação
        strCodigoSelect = fgObterCodigoItemCaixaSelect(pstrCodItemCaixa)
        
        strSQL = " SELECT CO_ITEM_CAIX              " & vbNewLine & _
                 " FROM   A6.TB_ITEM_CAIX_TIPO_OPER " & vbNewLine & _
                 " WHERE  CO_ITEM_CAIX LIKE        '" & strCodigoSelect & "%'"
                 
        Set rsItemCaixa = fgQuerySQL(strSQL)
    
        If Not rsItemCaixa.EOF Then
            '2022 - Item de Caixa possui associações com Tipos de Operação em níveis inferiores da hierarquia. Finalização de Vigência não permitida.
            lngCodigoErroNegocio = 2022
            GoTo ErrorHandler
        End If
        
        If fgObterNivelItemCaixa(pstrCodItemCaixa) = 1 Then
            strSQL = " SELECT CO_ITEM_CAIX                   " & vbNewLine & _
                     " FROM   A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA " & vbNewLine & _
                     " WHERE  CO_ITEM_CAIX =                '" & pstrCodItemCaixa & "'"
                     
            Set rsItemCaixa = fgQuerySQL(strSQL)
        
            If Not rsItemCaixa.EOF Then
                '2024 - Item de Caixa possui associações com Grupos de Veículo Legal. Finalização de Vigência não permitida.
                lngCodigoErroNegocio = 2024
                GoTo ErrorHandler
            End If
        End If
                
        flAlterarFimVigenciaItemGrupo pstrCodItemCaixa, pstrDataFinalVigencia
        
    Else
        '2014 - Item de Caixa deve ser de Grupo ou Elementar
        lngCodigoErroNegocio = 2014
        GoTo ErrorHandler
    End If
            
    Set rsItemCaixa = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Sub
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set rsItemCaixa = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocioItemCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub
