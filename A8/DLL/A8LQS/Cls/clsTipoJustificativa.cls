VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsTipoJustificativa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações relativas aos tipos de
'' justificativas para as conciliações, bem como a manipulação dos dados
'' cadastrados


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

'' Consultar todos os tipos de justificativas cadastrados para o processo de
'' conciliação 
Public Function LerTodos(Optional ByVal pstrIndicadorVigencia As String = "S", _
                         Optional ByVal pstrSegregaDados As String = "S") As String

Dim strSql                                  As String
Dim strWhereSegregacao                      As String
Dim blnWhere                                As Boolean

On Error GoTo ErrorHandler
  
    strSql = " SELECT   * " & _
             " FROM   A8.TB_TIPO_JUST_CNCL A "
    blnWhere = True
    If pstrIndicadorVigencia = "S" Then
        strSql = strSql & " WHERE A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                          " AND  (A.DT_FIM_VIGE  IS  NULL " & _
                          " OR    A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
        blnWhere = False
    End If
    
    If pstrSegregaDados = "S" Then
        strWhereSegregacao = fgSegregaDados("A8.TB_TIPO_JUST_CNCL", _
                                            False, _
                                            "A", _
                                            "A", _
                                            blnWhere, _
                                            False, _
                                            True, _
                                            False, _
                                            False, _
                                            False)
        
        strSql = strSql & strWhereSegregacao & vbNewLine
    End If
    
    strSql = strSql & " ORDER BY   A.TP_JUST_CNCL ASC "
    
    LerTodos = fgQueryXMLLerTodos("TipoJustificativa", strSql, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
'' 	TB_TIPO_JUST_CNCL
'' 
Public Function ObterPropriedades() As String

Dim strSql                                   As String

On Error GoTo ErrorHandler

    strSql = " SELECT   TP_JUST_CNCL,           " & _
            "           NO_TIPO_JUST_CNCL,      " & _
            "           TP_BKOF,                " & _
            "           DT_INIC_VIGE,           " & _
            "           DT_FIM_VIGE,            " & _
            "           CO_USUA_ULTI_ATLZ,      " & _
            "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           DH_ULTI_ATLZ,           " & _
            "           '' AS TP_VIGE,          " & _
            "           '' AS TP_SEGR           " & _
            " FROM      A8.TB_TIPO_JUST_CNCL    "

    ObterPropriedades = fgPropriedades("Grupo_TipoJustificativa", strSql, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados, tais como:
'' 	Incluir
'' 	Alterar
'' 	Excluir
'' 	Ler	e
'' 	LerTodos
'' 
Public Function Executar(ByVal pobjTipoJustificativa As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pobjTipoJustificativa.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pobjTipoJustificativa)
        Case "Ler"
            Executar = Ler(pobjTipoJustificativa.documentElement.selectSingleNode("//TP_JUST_CNCL").Text)
        Case "LerTodos"
            Executar = LerTodos(pobjTipoJustificativa.documentElement.selectSingleNode("//TP_VIGE").Text, _
                                pobjTipoJustificativa.documentElement.selectSingleNode("//TP_SEGR").Text)
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
'' Promover a inclusão de um novo tipo de justificativa 
Private Function flIncluir(ByRef pobjTipoJustificativa As MSXML2.DOMDocument40) As Boolean

Dim strSql                                   As String

On Error GoTo ErrorHandler

    strSql = " INSERT INTO A8.TB_TIPO_JUST_CNCL (" & _
             " TP_JUST_CNCL, " & _
             " TP_BKOF, " & _
             " NO_TIPO_JUST_CNCL, " & _
             " DT_INIC_VIGE, " & _
             " DT_FIM_VIGE, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( " & _
             CLng(pobjTipoJustificativa.documentElement.selectSingleNode("TP_JUST_CNCL").Text) & "," & _
             fgObterTipoBackOfficeUsuario & "," & _
             "'" & pobjTipoJustificativa.documentElement.selectSingleNode("NO_TIPO_JUST_CNCL").Text & "'," & _
             fgDtXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
             fgDtXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
             "'" & fgUsuarioRede & "'," & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSql

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Promover a alteração de um tipo de justificativa existente 
Private Function flAlterar(ByRef pobjTipoJustificativa As MSXML2.DOMDocument40) As Boolean

Dim strSql                                   As String

On Error GoTo ErrorHandler
            
    strSql = " UPDATE A8.TB_TIPO_JUST_CNCL " & _
            "  SET    NO_TIPO_JUST_CNCL       = '" & pobjTipoJustificativa.documentElement.selectSingleNode("NO_TIPO_JUST_CNCL").Text & "'," & _
            "         TP_BKOF                 = " & fgObterTipoBackOfficeUsuario & "," & _
            "         DT_INIC_VIGE            = " & fgDtXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DT_INIC_VIGE").Text) & "," & _
            "         DT_FIM_VIGE             = " & fgDtXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DT_FIM_VIGE").Text) & "," & _
            "         CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ            =  SYSDATE " & _
            " WHERE   TP_JUST_CNCL            = " & CLng(pobjTipoJustificativa.documentElement.selectSingleNode("TP_JUST_CNCL").Text) & _
            " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSql) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Consultar todos os atributos de um tipo de justificativa específico 
Public Function Ler(ByVal plngTP_JUST_CNCL As Long) As String

Dim strSql                                  As String

On Error GoTo ErrorHandler

    strSql = " SELECT   A.TP_JUST_CNCL, " & _
            "           A.NO_TIPO_JUST_CNCL, " & _
            "           A.DT_INIC_VIGE, " & _
            "           A.DT_FIM_VIGE, " & _
            "           A.CO_USUA_ULTI_ATLZ, " & _
            "           A.CO_ETCA_TRAB_ULTI_ATLZ, " & _
            "           A.DH_ULTI_ATLZ " & _
            " FROM      A8.TB_TIPO_JUST_CNCL A" & _
            " WHERE     A.TP_JUST_CNCL = " & plngTP_JUST_CNCL
            
    
    Ler = fgQueryXMLLer("Grupo_TipoJustificativa", strSql, App.EXEName & "." & TypeName(Me))
    
    If Len(Ler) = 0 Then
       '4002 - Registro Inexistente ou excluído por outro usuário.
       lngCodigoErroNegocio = 4002
       GoTo ErrorHandler
    End If

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Controlar o roteamento e execução dos métodos privados de manipulação de dados,
'' tais como:
'' 	Incluir
'' 	Alterar	e
'' 	Excluir
'' 
Public Function Salvar(ByVal pobjTipoJustificativa As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pobjTipoJustificativa.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pobjTipoJustificativa
    ElseIf pobjTipoJustificativa.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pobjTipoJustificativa
    ElseIf pobjTipoJustificativa.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pobjTipoJustificativa
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

'' Promover a exclusão de um tipo de justificativa existente 
Private Function flExcluir(ByRef pobjTipoJustificativa As MSXML2.DOMDocument40) As String

Dim strSql                                   As String

On Error GoTo ErrorHandler

    strSql = " DELETE " & strSql & _
             " FROM A8.TB_TIPO_JUST_CNCL " & _
             " WHERE   TP_JUST_CNCL    = " & CLng(pobjTipoJustificativa.documentElement.selectSingleNode("TP_JUST_CNCL").Text) & _
             " AND DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pobjTipoJustificativa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSql) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




