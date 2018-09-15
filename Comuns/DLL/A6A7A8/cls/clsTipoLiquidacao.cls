VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoLiquidacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela manutenção na tabela de Tipo de Liquidação (A8.TB_TIPO_LIQU_OPER_ATIV)

Option Explicit

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

'Recuperar registros da tabela de tipo de liquidação

Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "S") As String

Dim strSQL                                   As String
Dim blnFiltro                                As Boolean

On Error GoTo ErrorHandler
    
'    blnFiltro = CBool(pobjTipoLiquidacao.documentElement.selectSingleNode("Filtro").Text)
    strSQL = " SELECT   TP_LIQU_OPER_ATIV, " & _
             "          NO_TIPO_LIQU_OPER_ATIV, " & _
             "          CO_MESG_SPB_LIQU, " & _
             "          DT_INIC_VIGE,           " & _
             "          DT_FIM_VIGE             " & _
             " FROM   A8.TB_TIPO_LIQU_OPER_ATIV A "

    If pstrIndicadorVigencia = "S" Then
        strSQL = strSQL & " WHERE A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                          " AND  (A.DT_FIM_VIGE    IS  NULL       " & _
                          " OR   A.DT_FIM_VIGE    >= " & fgDataHoraServidor_To_Oracle & ")"
    End If
    strSQL = strSQL & " ORDER BY   A.TP_LIQU_OPER_ATIV ASC "
    
    LerTodos = fgQueryXMLLerTodos("TipoLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar os nomes das colunas da tabela de tipo de liquidação

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_LIQU_OPER_ATIV, " & _
            "           NO_TIPO_LIQU_OPER_ATIV, " & _
            "           CO_MESG_SPB_LIQU, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           '' AS TP_VIGE " & _
            " FROM      A8.TB_TIPO_LIQU_OPER_ATIV "
    
    ObterPropriedades = fgPropriedades("Grupo_TipoLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusão,Alteração,Exclusão,Ler e LerTodos da tabela de tipo de Liquidação

Public Function Executar(ByVal pobjTipoLiquidacao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = vbNullString
    
    If Not pobjTipoLiquidacao.selectSingleNode("//@Operacao") Is Nothing Then
        strOperacao = pobjTipoLiquidacao.selectSingleNode("//@Operacao").Text
    End If

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pobjTipoLiquidacao)
        Case "Ler"
            Executar = Ler(pobjTipoLiquidacao.documentElement.selectSingleNode("//CO_TIPO_LIQU").Text)
        Case "LerTodos"
            If pobjTipoLiquidacao.selectSingleNode("//TP_VIGE") Is Nothing Then
                Executar = LerTodos()
            Else
                Executar = LerTodos(pobjTipoLiquidacao.selectSingleNode("//TP_VIGE").Text)
            End If
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir um registro na tabela de tipo de liquidação

Private Function flIncluir(ByRef pobjTipoLiquidacao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_TIPO_LIQU_OPER_ATIV (" & _
             " TP_LIQU_OPER_ATIV, " & _
             " NO_TIPO_LIQU_OPER_ATIV, " & _
             " CO_MESG_SPB_LIQU, " & _
             " DT_INIC_VIGE, " & _
             " DT_FIM_VIGE   " & _
             " ) VALUES ( " & _
             CLng(pobjTipoLiquidacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) & "," & _
             "'" & pobjTipoLiquidacao.documentElement.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text & "'," & _
             "'" & pobjTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text & "'," & _
             fgDtXML_To_Oracle(pobjTipoLiquidacao.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             fgDtXML_To_Oracle(pobjTipoLiquidacao.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & ")"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela de tipo de liquidação

Private Function flAlterar(ByRef pobjTipoLiquidacao As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String
Dim datUltimaAtualizacao                     As Date

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_TIPO_LIQU_OPER_ATIV " & _
            "  SET    NO_TIPO_LIQU_OPER_ATIV    = '" & pobjTipoLiquidacao.documentElement.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text & "'," & _
            "         CO_MESG_SPB_LIQU          = '" & pobjTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text & "'," & _
            "         DT_INIC_VIGE              = " & fgDtXML_To_Oracle(pobjTipoLiquidacao.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
            "         DT_FIM_VIGE               = " & fgDtXML_To_Oracle(pobjTipoLiquidacao.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & _
            " WHERE   TP_LIQU_OPER_ATIV         = " & CLng(pobjTipoLiquidacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001-Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recuperar um registro na tabela de tipo de liquidação

Public Function Ler(ByVal plngTipoLiquidacao As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   TP_LIQU_OPER_ATIV, " & _
            "           NO_TIPO_LIQU_OPER_ATIV, " & _
            "           CO_MESG_SPB_LIQU, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE   " & _
            " FROM      A8.TB_TIPO_LIQU_OPER_ATIV " & _
            " WHERE     TP_LIQU_OPER_ATIV = " & plngTipoLiquidacao
    
    Ler = fgQueryXMLLer("Grupo_TipoLiquidacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gerenciar as operações de Inclusão,Alteração e Exclusão de um registro

Public Function Salvar(ByVal pobjTipoLiquidacao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    flRegrasNegocio pobjTipoLiquidacao
            
    If pobjTipoLiquidacao.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pobjTipoLiquidacao
    ElseIf pobjTipoLiquidacao.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pobjTipoLiquidacao
    ElseIf pobjTipoLiquidacao.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pobjTipoLiquidacao
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Regras de negocio

Private Function flRegrasNegocio(ByRef pobjTipoLiquidacao As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    If pobjTipoLiquidacao.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        
        If flVerificaTipoLiquidacao(CLng(pobjTipoLiquidacao.documentElement.selectSingleNode("CO_TIPO_LIQU").Text)) Then
            'Atributo já está cadastrado
            lngCodigoErroNegocio = 501
            GoTo ErrorHandler
        End If
        
    ElseIf pobjTipoLiquidacao.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        
    End If

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocio Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica a existencia de um tipo de liquidação

Private Function flVerificaTipoLiquidacao(ByVal plngTipoLiquidacao As Long) As Boolean

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  TP_LIQU_OPER_ATIV " & _
             "   FROM  A8.TB_TIPO_LIQU_OPER_ATIV " & _
             "  WHERE  TP_LIQU_OPER_ATIV = " & plngTipoLiquidacao
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        flVerificaTipoLiquidacao = False
    Else
        flVerificaTipoLiquidacao = True
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaTipoLiquidacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'excluir um registro na tabela de tipo de liquidação

Private Function flExcluir(ByRef pobjTipoLiquidacao As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A8.TB_TIPO_LIQU_OPER_ATIV " & _
             "  WHERE   TP_LIQU_OPER_ATIV     = " & pobjTipoLiquidacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text

    If fgExecuteSQL(strSQL) = 0 Then
        '4002-Registro foi exccluido por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


