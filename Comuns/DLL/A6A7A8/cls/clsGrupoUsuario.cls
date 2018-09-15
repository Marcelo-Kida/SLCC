VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pela manutenção do cadastro de grupo de usuários

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

'Recuperar dados da tabela A8.TB_GRUP_USUA
'Filtros :  Indicador Vigencia (S-Sim;N-Não)
'           Indicador de Segregação de dados (S-Sim;N-Não)
'           Objeto ADODB.Recordset

Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByVal pstrSegregaDados As String = "S", _
                         Optional ByRef prsGrupoUsuario As ADODB.Recordset) As String

Dim strSQL                                  As String
Dim strCondicao                             As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   CO_GRUP_USUA, " & _
            "           NO_GRUP_USUA, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ, " & _
            "           '' AS TP_VIGE " & _
            " FROM      A8.TB_GRUP_USUA "
    
    If pstrIndicadorVigencia = "S" Then
       strCondicao = " WHERE DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                       " AND (DT_FIM_VIGE    IS  NULL       " & _
                        " OR  DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
        
        strSQL = strSQL & strCondicao & vbNewLine
    
    End If
    
    If pstrSegregaDados = "S" Then
        strCondicao = fgSegregaDados("A8.TB_GRUP_USUA", _
                                      False, _
                                      "", _
                                      "", _
                                      Len(strCondicao) = 0, _
                                      False, _
                                      False, _
                                      False, _
                                      False, _
                                      True)
        strSQL = strSQL & strCondicao & vbNewLine
    
    End If
    
    strSQL = strSQL & " ORDER BY   CO_GRUP_USUA ASC "
    
    If prsGrupoUsuario Is Nothing Then
       LerTodos = fgQueryXMLLerTodos("GrupoUsuario", strSQL, App.EXEName & "." & TypeName(Me))
    Else
       Set prsGrupoUsuario = fgQuerySQL(strSQL)
       LerTodos = ""
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar as colunas da tabela A8.TB_GRUP_USUA

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_USUA, " & _
            "           NO_GRUP_USUA, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ,  " & _
            "           '' AS TP_VIGE, " & _
            "           '' AS TP_SEGR  " & _
            " FROM      A8.TB_GRUP_USUA "

    ObterPropriedades = fgPropriedades("Grupo_GrupoUsuario", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar as operações de Inclusão, Alteração, Exclusão , Ler, Lertodos

Public Function Executar(ByVal pobjGrupoUsuario As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pobjGrupoUsuario.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pobjGrupoUsuario)
        Case "Ler"
            Executar = Ler(pobjGrupoUsuario.documentElement.selectSingleNode("//CO_GRUP_USUA").Text)
        Case "LerTodos"
            Executar = LerTodos(pobjGrupoUsuario.documentElement.selectSingleNode("//TP_VIGE").Text, _
                                pobjGrupoUsuario.documentElement.selectSingleNode("//TP_SEGR").Text)
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

'Incluir um registro na tabela A8.TB_GRUP_USUA

Private Function flIncluir(ByRef pobjGrupoUsuario As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_GRUP_USUA (" & _
             " CO_GRUP_USUA, " & _
             " NO_GRUP_USUA, " & _
             " DT_INIC_VIGE, " & _
             " DT_FIM_VIGE, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( " & _
             pobjGrupoUsuario.documentElement.selectSingleNode("CO_GRUP_USUA").Text & "," & _
             "'" & pobjGrupoUsuario.documentElement.selectSingleNode("NO_GRUP_USUA").Text & "'," & _
             fgDtXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             fgDtXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
             "'" & fgObterUsuarioRede & "'," & _
             "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar um registro na tabela A8.TB_GRUP_USUA

Private Function flAlterar(ByRef pobjGrupoUsuario As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_GRUP_USUA " & _
             " SET     NO_GRUP_USUA              = '" & pobjGrupoUsuario.documentElement.selectSingleNode("NO_GRUP_USUA").Text & "'," & _
             "         DT_INIC_VIGE              = " & fgDtXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             "         DT_FIM_VIGE               = " & fgDtXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
             "         CO_USUA_ULTI_ATLZ         = '" & fgObterUsuarioRede & "'," & _
             "         CO_ETCA_TRAB_ULTI_ATLZ    = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
             "         DH_ULTI_ATLZ              =  SYSDATE " & _
             " WHERE   CO_GRUP_USUA              = " & pobjGrupoUsuario.documentElement.selectSingleNode("CO_GRUP_USUA").Text & _
             " AND     DH_ULTI_ATLZ              = " & fgDtHrXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar um registro na tabela A8.TB_GRUP_USUA

Public Function Ler(ByVal plngCO_GRUP_USUA As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_GRUP_USUA, " & _
            "           NO_GRUP_USUA, " & _
            "           DT_INIC_VIGE, " & _
            "           DT_FIM_VIGE, " & _
            "           CO_USUA_ULTI_ATLZ, " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ " & _
            " FROM      A8.TB_GRUP_USUA " & _
            " WHERE     CO_GRUP_USUA = " & plngCO_GRUP_USUA
    
    Ler = fgQueryXMLLer("Grupo_GrupoUsuario", strSQL, App.EXEName & "." & TypeName(Me))

    If Len(Ler) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar a operação de Inclusão ou Alteração de um registro

Public Function Salvar(ByVal pobjGrupoUsuario As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
   
            
    If pobjGrupoUsuario.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pobjGrupoUsuario
    ElseIf pobjGrupoUsuario.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pobjGrupoUsuario
    ElseIf pobjGrupoUsuario.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pobjGrupoUsuario
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

'Verifica a existencia de um grupo de usuario

Private Function flVerificaGrupoUsuario(ByVal plngCO_GRUP_USUA As Long) As Boolean

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  CO_GRUP_USUA " & _
             "   FROM  A8.TB_GRUP_USUA " & _
             "  WHERE  CO_GRUP_USUA = " & plngCO_GRUP_USUA
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.EOF Then
        flVerificaGrupoUsuario = False
    Else
        flVerificaGrupoUsuario = True
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaGrupoUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Excluir um registro da tabela A8.TB_GRUP_USUA

Private Function flExcluir(ByRef pobjGrupoUsuario As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A8.TB_GRUP_USUA " & _
             " WHERE   CO_GRUP_USUA = " & pobjGrupoUsuario.documentElement.selectSingleNode("CO_GRUP_USUA").Text & _
             " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pobjGrupoUsuario.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4001-Registro foi alterado ou excluido por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




