VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsGrupoUsuarioMensSPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações de associação relativas aos
'' grupos de usuários x mensagens SPB, bem como a manipulação dos dados cadastrados

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

'Excluir associação de um grupo de usuário a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Private Function flExcluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlDomNode

        strSQL = "DELETE FROM                " & vbCrLf & _
                 "   A8.TB_GRUP_USUA_MESG_SPB " & vbCrLf & _
                 " WHERE CO_GRUP_USUA =  " & .selectSingleNode("CO_GRUP_USUA").Text & "  AND " & vbCrLf & _
                 "       CO_MESG_SPB  = '" & .selectSingleNode("CO_MESG_SPB").Text & "'  AND " & vbCrLf & _
                 "       TP_BKOF      =  " & .selectSingleNode("TP_BKOF").Text & "       AND " & vbCrLf & _
                 "       DH_ULTI_ATLZ =  " & fgDtHrXML_To_Oracle(pxmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            
    End With

    If fgExecuteSQL(strSQL) = 0 Then
        'Registro foi alterado ou excluído por outro usuário.
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

'Incluir associação de um grupo de usuário a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Private Function flIncluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_GRUP_USUA_MESG_SPB ( " & vbCrLf & _
             " CO_GRUP_USUA,                          " & vbCrLf & _
             " CO_MESG_SPB,                           " & vbCrLf & _
             " TP_BKOF,                               " & vbCrLf & _
             " TP_ACES,                               " & vbCrLf & _
             " CO_USUA_ULTI_ATLZ,                     " & vbCrLf & _
             " CO_ETCA_TRAB_ULTI_ATLZ,                " & vbCrLf & _
             " DH_ULTI_ATLZ                           " & vbCrLf & _
             " ) VALUES (                             " & vbCrLf
             
             
    With pxmlDomNode
        strSQL = strSQL & _
                 " " & CLng("0" & .selectSingleNode("CO_GRUP_USUA").Text) & " , " & vbCrLf & _
                 " '" & .selectSingleNode("CO_MESG_SPB").Text & "' ,            " & vbCrLf & _
                 " " & CLng("0" & .selectSingleNode("TP_BKOF").Text) & " ,      " & vbCrLf & _
                 " " & CLng("0" & .selectSingleNode("TP_ACES").Text) & ",       " & vbCrLf & _
                 " '" & fgUsuarioRede & "' ,                                    " & vbCrLf & _
                 " '" & fgEstacaoTrabalhoUsuario & "' ,                         " & vbCrLf & _
                 " SYSDATE )"

    End With
             

    fgExecuteSQL strSQL
    
    flIncluir = True

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

'Salvar associação de um grupo de usuário a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Private Function flSalvarTodos(ByVal pxmlItemCaixaTipoOperacao As MSXML2.DOMDocument40) As Boolean

Dim xmlDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    For Each xmlDomNode In pxmlItemCaixaTipoOperacao.documentElement.selectNodes("//Repeat_GrupoUsuarioMensSPB/*")
        Select Case xmlDomNode.selectSingleNode("@Operacao").Text
            Case "Incluir"
                flIncluir xmlDomNode
            Case "Alterar"
                flAlterar xmlDomNode
            Case "Excluir"
                flExcluir xmlDomNode
        End Select
    Next

    flSalvarTodos = True

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

'Alterar associação de um grupo de usuário a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Private Function flAlterar(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlDomNode
        strSQL = " UPDATE  A8.TB_GRUP_USUA_MESG_SPB                                                        " & vbCrLf & _
                 " SET     TP_ACES= " & CLng("0" & .selectSingleNode("TP_ACES").Text) & " ,                  " & vbCrLf & _
                 "         CO_USUA_ULTI_ATLZ = '" & fgUsuarioRede & "' ,                                     " & vbCrLf & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "' ,                     " & vbCrLf & _
                 "         DH_ULTI_ATLZ = SYSDATE                                                            " & vbCrLf & _
                 " WHERE CO_GRUP_USUA =  " & CLng("0" & .selectSingleNode("CO_GRUP_USUA").Text) & " AND      " & vbCrLf & _
                 "       CO_MESG_SPB  = '" & .selectSingleNode("CO_MESG_SPB").Text & "'             AND      " & vbCrLf & _
                 "       TP_BKOF      =  " & CLng("0" & .selectSingleNode("TP_BKOF").Text) & "      AND      " & vbCrLf & _
                 "       DH_ULTI_ATLZ =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & "   " & vbCrLf
    
    
    End With
    
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
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

'Executar operação de Inclusão,Alteração,Exclusão,LerTodos associação de um grupo de usuário a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Public Function Executar(ByRef pxmlGrupoUsuarioMensSPB As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlGrupoUsuarioMensSPB.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
    
        Case "Incluir", "Alterar", "Excluir"
            Executar = flSalvarTodos(pxmlGrupoUsuarioMensSPB)
        Case "LerTodos"
            Executar = LerTodos(pxmlGrupoUsuarioMensSPB.selectSingleNode("//CO_GRUP_USUA").Text, _
                          pxmlGrupoUsuarioMensSPB.selectSingleNode("//TP_BKOF").Text)
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

'Recuperar as associações de um grupos de usuários a mensagem SPB da tabela A8.TB_GRUP_USUA_MESG_SPB

Public Function LerTodos(Optional ByVal pstrCodGrupUsua As String = vbNullString, _
                         Optional ByVal pstrTpBkOf As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT CO_GRUP_USUA,           " & vbCrLf & _
             "        CO_MESG_SPB,            " & vbCrLf & _
             "        TP_BKOF,                " & vbCrLf & _
             "        TP_ACES,                " & vbCrLf & _
             "        DH_ULTI_ATLZ            " & vbCrLf & _
             " FROM A8.TB_GRUP_USUA_MESG_SPB  " & vbCrLf
             
    If pstrCodGrupUsua <> vbNullString Or pstrTpBkOf <> vbNullString Then
        strSQL = strSQL & " WHERE " & vbCrLf
        
        If pstrCodGrupUsua <> vbNullString Then
            strSQL = strSQL & _
                    " CO_GRUP_USUA = " & pstrCodGrupUsua & "  " & vbCrLf
        End If
        
        If pstrTpBkOf <> vbNullString Then
            If pstrCodGrupUsua <> vbNullString Then
                strSQL = strSQL & " AND "
            End If
            strSQL = strSQL & _
                     " TP_BKOF = " & pstrTpBkOf & "  " & vbCrLf
        End If
        
    End If

    LerTodos = fgQueryXMLLerTodos("SistemaMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os nomes dascolunas da tabela A8.TB_GRUP_USUA_MESG_SPB

Public Function ObterPropriedades() As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler


    strSQL = " SELECT CO_GRUP_USUA,           " & vbCrLf & _
             "        CO_MESG_SPB,            " & vbCrLf & _
             "        TP_BKOF,                " & vbCrLf & _
             "        TP_ACES                 " & vbCrLf & _
             " FROM A8.TB_GRUP_USUA_MESG_SPB  " & vbCrLf

    ObterPropriedades = fgPropriedades("Grupo_GrupoUsuarioMensSPB", strSQL, App.EXEName & "." & TypeName(Me))

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


