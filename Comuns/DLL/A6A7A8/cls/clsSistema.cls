VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela manutenção na tabela de sistemas (A7.TB_SIST)

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

'Recuperar resistros da tabela A7.TB_SIST desconsiderando a empresa

Public Function LerTodosIgnorarEmpresa(Optional ByVal pstrIndicadorVigencia As String = "S") As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    DISTINCT                   " & vbNewLine & _
             "           A.SG_SIST,                 " & vbNewLine & _
             "           A.NO_SIST                  " & vbNewLine & _
             " FROM      A7.TB_SIST A               " & vbNewLine
             
    If pstrIndicadorVigencia = "S" Then
        strSQL = strSQL & " WHERE A.DT_INIC_VIGE_SIST <= " & fgDataHoraServidor_To_Oracle & _
                          " AND  (A.DT_FIM_VIGE_SIST    IS  NULL       " & _
                          " OR    A.DT_FIM_VIGE_SIST  >= " & fgDataHoraServidor_To_Oracle & ")" & vbNewLine
    End If
    
    strSQL = strSQL & " ORDER BY A.SG_SIST "
    
    LerTodosIgnorarEmpresa = fgQueryXMLLerTodos("Sistema", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosIgnorarEmpresa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar resistros da tabela A7.TB_SIST

Public Function LerTodos(Optional ByVal plngEmpresa As Long, _
                         Optional ByVal pstrIndicadorVigencia As String = "S") As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " Select    A.SG_SIST,                 " & vbNewLine & _
             "           A.NO_SIST,                 " & vbNewLine & _
             "           A.CO_EMPR,                 " & vbNewLine & _
             "           B.NO_REDU_EMPR,            " & vbNewLine & _
             "           A.DT_INIC_VIGE_SIST,       " & vbNewLine & _
             "           A.DT_FIM_VIGE_SIST,        " & vbNewLine & _
             "           A.CO_USUA_ULTI_ATLZ,       " & vbNewLine & _
             "           A.CO_ETCA_TRAB_ULTI_ATLZ,  " & vbNewLine & _
             "           A.DH_ULTI_ATLZ             " & vbNewLine & _
             " From      A7.TB_SIST  A,             " & vbNewLine & _
             "           A8.TB_EMPRESA_HO B         " & vbNewLine & _
             " Where     B.CO_EMPR = A.CO_EMPR      "
             
    If plngEmpresa > 0 Then
       strSQL = strSQL & " AND A.CO_EMPR  = " & plngEmpresa
    End If

    If pstrIndicadorVigencia = "S" Then
       strSQL = strSQL & " AND   A.DT_INIC_VIGE_SIST <= " & fgDataHoraServidor_To_Oracle & _
                         " AND  (A.DT_FIM_VIGE_SIST    IS  NULL       " & _
                         " OR    TRUNC(A.DT_FIM_VIGE_SIST)  >= " & fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))) & ")" & vbNewLine
       strSQL = strSQL & " AND   B.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                         " AND  (B.DT_FIM_VIGE    IS  NULL       " & _
                         " OR    TRUNC(B.DT_FIM_VIGE)  >= " & fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))) & ")" & vbNewLine
    End If
    
    strSQL = strSQL & " Order By  A.NO_SIST                  "
    
    LerTodos = fgQueryXMLLerTodos("Sistema", strSQL, App.EXEName & "." & TypeName(Me))


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

'Obter os nomes das colunas da tabela A7.TB_SIST

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " Select    SG_SIST,                 " & vbNewLine & _
             "           NO_SIST,                 " & vbNewLine & _
             "           CO_EMPR,                 " & vbNewLine & _
             "           '' AS NO_REDU_EMPR,      " & vbNewLine & _
             "           DT_INIC_VIGE_SIST,       " & vbNewLine & _
             "           DT_FIM_VIGE_SIST,        " & vbNewLine & _
             "           CO_USUA_ULTI_ATLZ,       " & vbNewLine & _
             "           CO_ETCA_TRAB_ULTI_ATLZ,  " & vbNewLine & _
             "           DH_ULTI_ATLZ,            " & vbNewLine & _
             "           '' AS TP_VIGE            " & vbNewLine & _
             " From      A7.TB_SIST               "

    ObterPropriedades = fgPropriedades("Grupo_Sistema", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
'        objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Excluir um resistro da tabela A7.TB_SIST

Private Function flExcluir(ByRef pxmlSistema As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
   
   With pxmlSistema.documentElement
        
        strSQL = " DELETE FROM A7.TB_SIST " & vbCrLf & _
                 " WHERE   SG_SIST = '" & .selectSingleNode("SG_SIST").Text & "' " & vbCrLf & _
                 " AND     CO_EMPR =  " & .selectSingleNode("CO_EMPR").Text & vbCrLf & _
                 " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlSistema.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
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

'Gerenciar as operações de Inclusão e Alteração na tabela de sistemas

Public Function Salvar(ByVal pxmlSistema As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    'flRegrasNegocio pxmlSistema
            
    If pxmlSistema.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlSistema
    ElseIf pxmlSistema.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        If Not flSistemaEmUso(pxmlSistema.selectSingleNode("//SG_SIST").Text, _
                              pxmlSistema.selectSingleNode("//CO_EMPR").Text) Then
            flAlterar pxmlSistema
        Else
            'Sistema Não pode ser alterado pois esta em uso
            lngCodigoErroNegocio = 4149
            GoTo ErrorHandler
        End If
    ElseIf pxmlSistema.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlSistema
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

'Alterar um resistros da tabela A7.TB_SIST

Private Function flAlterar(ByRef pxmlSistema As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    With pxmlSistema.documentElement
        
        strSQL = " UPDATE A7.TB_SIST                    " & vbCrLf & _
                 " SET     SG_SIST                   = '" & .selectSingleNode("SG_SIST").Text & "'," & vbCrLf & _
                 "         NO_SIST                   = '" & .selectSingleNode("NO_SIST").Text & "'," & vbCrLf & _
                 "         DT_INIC_VIGE_SIST         = " & fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE_SIST").Text) & "," & vbCrLf & _
                 "         DT_FIM_VIGE_SIST          = " & fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE_SIST").Text) & "," & vbCrLf & _
                 "         CO_USUA_ULTI_ATLZ         = '" & fgObterUsuarioRede & "'," & vbCrLf & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ    = '" & fgObterEstacaoTrabalhoUsuario & "'," & vbCrLf & _
                 "         DH_ULTI_ATLZ              =  SYSDATE " & vbCrLf & _
                 " WHERE   SG_SIST                   = '" & .selectSingleNode("SG_SIST").Text & "'" & vbCrLf & _
                 " AND     CO_EMPR                   = " & .selectSingleNode("CO_EMPR").Text & vbCrLf & _
                 " AND     DH_ULTI_ATLZ              = " & fgDtHrXML_To_Oracle(pxmlSistema.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
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

'Incluir um resistro da tabela A7.TB_SIST

Private Function flIncluir(ByRef pxmlSistema As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlSistema.documentElement
        strSQL = " INSERT INTO A7.TB_SIST (     " & vbCrLf & _
                 " SG_SIST,                     " & vbCrLf & _
                 " NO_SIST,                     " & vbCrLf & _
                 " CO_EMPR,                     " & vbCrLf & _
                 " DT_INIC_VIGE_SIST,           " & vbCrLf & _
                 " DT_FIM_VIGE_SIST,            " & vbCrLf & _
                 " CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
                 " CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
                 "  DH_ULTI_ATLZ                " & vbCrLf & _
                 " ) VALUES (                   " & vbCrLf & _
                 "'" & .selectSingleNode("SG_SIST").Text & "'," & _
                 "'" & .selectSingleNode("NO_SIST").Text & "'," & _
                 .selectSingleNode("CO_EMPR").Text & "," & _
                 fgDtXML_To_Oracle(.selectSingleNode("DT_INIC_VIGE_SIST").Text) & "," & _
                 fgDtXML_To_Oracle(.selectSingleNode("DT_FIM_VIGE_SIST").Text) & "," & _
                 "'" & fgObterUsuarioRede & "'," & _
                 "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
                 "SYSDATE)"
    End With
    
    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recuperar um resistro da tabela A7.TB_SIST

Public Function Ler(ByVal plngEmpresa As Long, _
                    ByVal pstrSG_SIST As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   A.SG_SIST,                " & vbNewLine & _
            "           A.NO_SIST,                " & vbNewLine & _
            "           A.CO_EMPR,                " & vbNewLine & _
            "           B.NO_REDU_EMPR,           " & vbNewLine & _
            "           A.DT_INIC_VIGE_SIST,      " & vbNewLine & _
            "           A.DT_FIM_VIGE_SIST,       " & vbNewLine & _
            "           A.CO_USUA_ULTI_ATLZ,      " & vbNewLine & _
            "           A.CO_ETCA_TRAB_ULTI_ATLZ, " & vbNewLine & _
            "           A.DH_ULTI_ATLZ            " & vbNewLine & _
            " From      A7.TB_SIST    A,          " & vbNewLine & _
            "           A8.TB_EMPRESA_HO B        " & vbNewLine & _
            " Where     B.CO_EMPR = A.CO_EMPR     " & vbNewLine & _
            " AND       A.CO_EMPR = " & plngEmpresa & vbNewLine & _
            " AND       A.SG_SIST = '" & pstrSG_SIST & "' "
    
    Ler = fgQueryXMLLer("Grupo_Sistema", strSQL, App.EXEName & "." & TypeName(Me))

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

'Gerencias as operações de Inclusão,Alteração,Exclusão, Ler e Lertodos da tabela de sistemas

Public Function Executar(ByVal pxmlSistema As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String
Dim strIgnorarEmpresa                       As String

On Error GoTo ErrorHandler

    strOperacao = pxmlSistema.documentElement.selectSingleNode("//*/@Operacao").Text
    strIgnorarEmpresa = "N"
    
    If Not pxmlSistema.documentElement.selectSingleNode("//*/@IgnoraEmpresa") Is Nothing Then
        strIgnorarEmpresa = pxmlSistema.documentElement.selectSingleNode("//*/@IgnoraEmpresa").Text
    End If

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlSistema)
        Case "Ler"
            Executar = Ler(CLng("0" & pxmlSistema.documentElement.selectSingleNode("//CO_EMPR").Text), _
                           pxmlSistema.documentElement.selectSingleNode("//SG_SIST").Text)
        Case "LerTodos"
            If strIgnorarEmpresa = "N" Then
                If Not pxmlSistema.documentElement.selectSingleNode("//CO_EMPR") Is Nothing And Not pxmlSistema.documentElement.selectSingleNode("//TP_VIGE") Is Nothing Then
                    Executar = LerTodos(pxmlSistema.documentElement.selectSingleNode("//CO_EMPR").Text, _
                                        pxmlSistema.documentElement.selectSingleNode("//TP_VIGE").Text)
                Else
                    Executar = LerTodos
                End If
            Else
                Executar = LerTodosIgnorarEmpresa
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

'Vericar se existe associação de um sistema a uma regra de transporte

Private Function flSistemaEmUso(ByVal pstrSiglaSistema As String, ByVal pstrEmpresa As String) As Boolean

On Error GoTo ErrorHandler

Dim Rs                                      As Recordset

    Set Rs = fgQuerySQL("select count(*) as QTD " & vbCrLf & _
                        "from a7.tb_regr_trap_mesg a, a7.tb_regr_sist_dest b " & vbCrLf & _
                        "Where a.tp_form_mesg_said = b.tp_form_mesg_said " & vbCrLf & _
                          "and a.tp_mesg = b.tp_mesg " & vbCrLf & _
                          "and a.sg_sist_orig = b.sg_sist_orig " & vbCrLf & _
                          "and a.co_empr_orig = b.co_empr_orig " & vbCrLf & _
                          "and a.dh_inic_vige_regr_trap = b.dh_inic_vige_regr_trap " & vbCrLf & _
                          "and a.dh_inic_vige_regr_trap < trunc(sysdate) " & vbCrLf & _
                          "and (a.dt_fim_vige_regr_trap > trunc(sysdate) or a.dt_fim_vige_regr_trap is null) " & vbCrLf & _
                          "and (a.sg_sist_orig = '" & Trim$(pstrSiglaSistema) & "' or b.sg_sist_dest = '" & Trim$(pstrSiglaSistema) & "') " & vbCrLf & _
                          "and a.co_empr_orig = " & Trim$(pstrEmpresa))

    If CLng(Rs!QTD) > 0 Then
        flSistemaEmUso = True
    End If
    
    Rs.Close
    Set Rs = Nothing

    Exit Function
ErrorHandler:
    Set Rs = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flSistemaEmUso Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
