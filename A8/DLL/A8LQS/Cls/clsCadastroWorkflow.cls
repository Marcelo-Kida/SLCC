VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCadastroWorkflow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações do fluxo de aprovação do
'' sistema (Workflow), bem como a manipulação dos dados cadastrados. Abaixo estão
'' alguns exemplos dos eventos de workflow:
''      Concordância automática ou não
''      Liberação automática ou não
''      Conciliação automática ou não

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

'' Controlar o roteamento e execução dos métodos privados na tabela
'' (TB_PARM_FCAO_SIST), tais como:
''      Alterar
Public Function Executar(ByRef pxmlCadastroWorkflow As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlCadastroWorkflow.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Alterar"
            Executar = flAlterar(pxmlCadastroWorkflow)
            
            'Verifica se existem Regras de Exceção a serem gravadas...
            If Not pxmlCadastroWorkflow.selectSingleNode("ParamWorkflow/Repeat_ParamExcecoes") Is Nothing Then
                Executar = flSalvarTodos(pxmlCadastroWorkflow)
            End If
            
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

'' Promover a alteração de um workflow existente
Private Function flAlterar(ByVal pxmlCadastroWorkflow As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlCadastroWorkflow.documentElement
        strSQL = " UPDATE A8.TB_PARM_FCAO_SIST                                                               " & vbNewLine & _
                 " SET     IN_FCAO_SIST_AUTM         =  " & .selectSingleNode("IN_FCAO_SIST_AUTM").Text & ", " & vbNewLine & _
                 "         CO_USUA_ULTI_ATLZ         = '" & fgUsuarioRede & "',                              " & vbNewLine & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ    = '" & fgEstacaoTrabalhoUsuario & "',                   " & vbNewLine & _
                 "         DH_ULTI_ATLZ              = SYSDATE,                                              " & vbNewLine & _
                 "         TP_COND_SALD              = " & .selectSingleNode("TP_COND_SALD").Text & "         " & vbNewLine & _
                 " WHERE   TP_OPER                   = " & .selectSingleNode("TP_OPER").Text & "             " & vbNewLine & _
                 " AND     CO_FCAO_SIST              = " & .selectSingleNode("CO_FCAO_SIST").Text & "        " & vbNewLine & _
                 " AND     TP_BKOF                   = " & .selectSingleNode("TP_BKOF").Text & "             " & vbNewLine & _
                 " AND     CO_EMPR                   = " & .selectSingleNode("CO_EMPR").Text & "             " & vbNewLine & _
                 " AND     DH_ULTI_ATLZ              = " & fgDtHrXML_To_Oracle(pxmlCadastroWorkflow.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
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

'' Promover a exclusão de um workflow existente
Private Function flExcluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlDomNode
        strSQL = " DELETE                                                                    " & vbNewLine & _
                 " FROM    A8.TB_PARM_FCAO_SIST_EXCE                                         " & vbNewLine & _
                 " WHERE   TP_OPER       =  " & .selectSingleNode("TP_OPER").Text & "        " & vbNewLine & _
                 " AND     CO_FCAO_SIST  =  " & .selectSingleNode("CO_FCAO_SIST").Text & "   " & vbNewLine & _
                 " AND     TP_BKOF       =  " & .selectSingleNode("TP_BKOF").Text & "        " & vbNewLine & _
                 " AND     CO_EMPR       =  " & .selectSingleNode("CO_EMPR").Text & "        " & vbNewLine & _
                 " AND     SG_SIST       = '" & .selectSingleNode("SG_SIST").Text & "'       " & vbNewLine & _
                 " AND     CO_LOCA_LIQU  =  " & .selectSingleNode("CO_LOCA_LIQU").Text & "   " & vbNewLine & _
                 " AND     CO_GRUP_USUA  =  " & .selectSingleNode("CO_GRUP_USUA").Text & "   " & vbNewLine & _
                 " AND     DH_ULTI_ATLZ  =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
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

'' Promover a inclusão de um novo workflow
Private Function flIncluir(ByVal pxmlDomNode As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    With pxmlDomNode
        strSQL = " INSERT INTO A8.TB_PARM_FCAO_SIST_EXCE (          " & vbNewLine & _
                 "      TP_OPER,                                    " & vbNewLine & _
                 "      CO_FCAO_SIST,                               " & vbNewLine & _
                 "      TP_BKOF,                                    " & vbNewLine & _
                 "      CO_EMPR,                                    " & vbNewLine & _
                 "      SG_SIST,                                    " & vbNewLine & _
                 "      CO_LOCA_LIQU,                               " & vbNewLine & _
                 "      CO_GRUP_USUA,                               " & vbNewLine & _
                 "      CO_USUA_ULTI_ATLZ,                          " & vbNewLine & _
                 "      CO_ETCA_TRAB_ULTI_ATLZ,                     " & vbNewLine & _
                 "      DH_ULTI_ATLZ                                " & vbNewLine & _
                 " ) VALUES (                                       " & vbNewLine & _
                 "  " & .selectSingleNode("TP_OPER").Text & ",      " & vbNewLine & _
                 "  " & .selectSingleNode("CO_FCAO_SIST").Text & ", " & vbNewLine & _
                 "  " & fgObterTipoBackOfficeUsuario & ",      " & vbNewLine & _
                 "  " & .selectSingleNode("CO_EMPR").Text & ",      " & vbNewLine & _
                 " '" & .selectSingleNode("SG_SIST").Text & "',     " & vbNewLine & _
                 "  " & .selectSingleNode("CO_LOCA_LIQU").Text & ", " & vbNewLine & _
                 "  " & .selectSingleNode("CO_GRUP_USUA").Text & ", " & vbNewLine & _
                 " '" & fgUsuarioRede & "',                         " & vbNewLine & _
                 " '" & fgEstacaoTrabalhoUsuario & "',              " & vbNewLine & _
                 "      SYSDATE)"
    End With
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados de manipulação de dados
'' na tabela (TB_PARM_FCAO_SIST_EXCE), tais como:
''      Incluir
''      Excluir
Private Function flSalvarTodos(ByVal pxmlCadastroWorkflow As MSXML2.DOMDocument40) As Boolean

Dim strOperacao                             As String
Dim xmlDomNode                              As IXMLDOMNode

    On Error GoTo ErrorHandler

    For Each xmlDomNode In pxmlCadastroWorkflow.selectSingleNode("ParamWorkflow/Repeat_ParamExcecoes").childNodes
        With xmlDomNode
            strOperacao = .selectSingleNode("@Operacao").Text
            
            Select Case strOperacao
            Case "Incluir"
                Call flIncluir(xmlDomNode)
            Case "Excluir"
                Call flExcluir(xmlDomNode)
            End Select
            
        End With
    Next
    
    flSalvarTodos = True
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set xmlDomNode = Nothing
    Exit Function

ErrorHandler:
    Set xmlDomNode = Nothing
    flSalvarTodos = False
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos os atributos de um workflow específico
Public Function Ler(ByVal pintTipoOperacao As Integer, _
                    ByVal pintFuncaoSistema As Integer, _
                    ByVal pintEmpresa As Integer) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    A.TP_OPER,                             " & vbNewLine & _
             "           A.CO_FCAO_SIST,                        " & vbNewLine & _
             "           A.TP_BKOF,                             " & vbNewLine & _
             "           A.CO_EMPR,                             " & vbNewLine & _
             "           A.IN_FCAO_SIST_AUTM,                   " & vbNewLine & _
             "           A.DH_ULTI_ATLZ,                        " & vbNewLine & _
             "           B.NO_TIPO_OPER,                         " & vbNewLine & _
             "           A.TP_COND_SALD                         " & vbNewLine & _
             " FROM      A8.TB_PARM_FCAO_SIST       A,          " & vbNewLine & _
             "           A8.TB_TIPO_OPER            B           " & vbNewLine & _
             " WHERE     A.TP_OPER          = B.TP_OPER         " & vbNewLine & _
             " AND       A.TP_OPER          = " & pintTipoOperacao & vbNewLine & _
             " AND       A.CO_FCAO_SIST     = " & pintFuncaoSistema & vbNewLine & _
             " AND       A.CO_EMPR          = " & pintEmpresa

    strWhereSegregacao = fgSegregaDados("A8.TB_PARM_FCAO_SIST", _
                                        False, _
                                        "A", _
                                        "A", _
                                        False, _
                                        False, _
                                         True, _
                                        False, _
                                        False, _
                                        False)
    
    strSQL = strSQL & strWhereSegregacao
    
    Ler = fgQueryXMLLer("ParamWorkflow", strSQL, App.EXEName & "." & TypeName(Me))

    If Len(Ler) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
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

' Consultar todos os atributos de um workflow especifico
Public Function ObterParamCondicionalSaldo(ByVal pintTipoOperacao As Integer, _
                                           ByVal pintFuncaoSistema As Integer, _
                                           ByVal pintEmpresa As Integer, _
                                  Optional ByVal plngTipoBackOffice As Long) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    A.TP_COND_SALD                         " & vbNewLine & _
             " FROM      A8.TB_PARM_FCAO_SIST       A,          " & vbNewLine & _
             "           A8.TB_TIPO_OPER            B           " & vbNewLine & _
             " WHERE     A.TP_OPER          = B.TP_OPER         " & vbNewLine & _
             " AND       A.TP_OPER          = " & pintTipoOperacao & vbNewLine & _
             " AND       A.CO_FCAO_SIST     = " & pintFuncaoSistema & vbNewLine & _
             " AND       A.CO_EMPR          = " & pintEmpresa & vbNewLine & _
             " AND       A.TP_BKOF          = " & plngTipoBackOffice
    
    ObterParamCondicionalSaldo = fgQueryXMLLer("ParamWorkflow", strSQL, App.EXEName & "." & TypeName(Me))

    If Len(ObterParamCondicionalSaldo) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterParamCondicionalSaldo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos os workflows cadastrados para o processo de liquidação
Public Function LerTodos(ByVal pintEmpresa As Integer) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    A.TP_OPER,                 " & vbNewLine & _
             "           B.NO_TIPO_OPER,            " & vbNewLine & _
             "           A.CO_FCAO_SIST,            " & vbNewLine & _
             "           A.IN_FCAO_SIST_AUTM,       " & vbNewLine & _
             "           A.TP_COND_SALD             " & vbNewLine & _
             " FROM      A8.TB_PARM_FCAO_SIST   A,  " & vbNewLine & _
             "           A8.TB_TIPO_OPER        B   " & vbNewLine & _
             " WHERE     A.TP_OPER = B.TP_OPER      " & vbNewLine & _
             " AND       A.CO_EMPR = " & pintEmpresa & vbNewLine

    strWhereSegregacao = fgSegregaDados("A8.TB_PARM_FCAO_SIST", _
                                        False, _
                                        "A", _
                                        "A", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        False)
    
    strSQL = strSQL & strWhereSegregacao & vbNewLine
    strSQL = strSQL & " ORDER BY  B.NO_TIPO_OPER"

    LerTodos = fgQueryXMLLerTodos("ParamWorkflow", strSQL, App.EXEName & "." & TypeName(Me))

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

'' Consultar todos os parâmetros de exceção cadastrados para os parâmetros
'' especificados:
''      Tipo de Operação
''      Função de Sistema       e
''      Empresa
Public Function ObterExcecoesParametrizacao(ByVal pintTipoOperacao As Integer, _
                                            ByVal pintFuncaoSistema As Integer, _
                                            ByVal pintEmpresa As Integer) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    A.TP_OPER,                             " & vbNewLine & _
             "           A.CO_FCAO_SIST,                        " & vbNewLine & _
             "           A.TP_BKOF,                             " & vbNewLine & _
             "           A.CO_EMPR,                             " & vbNewLine & _
             "           A.SG_SIST,                             " & vbNewLine & _
             "           A.CO_LOCA_LIQU,                        " & vbNewLine & _
             "           A.CO_GRUP_USUA,                        " & vbNewLine & _
             "           A.DH_ULTI_ATLZ,                        " & vbNewLine & _
             "           B.SG_LOCA_LIQU,                        " & vbNewLine & _
             "           C.NO_GRUP_USUA                         " & vbNewLine & _
             " FROM      A8.TB_PARM_FCAO_SIST_EXCE  A,          " & vbNewLine & _
             "           A8.TB_LOCAL_LIQUIDACAO        B,          " & vbNewLine & _
             "           A8.TB_GRUP_USUA            C           " & vbNewLine

    strSQL = strSQL & _
             " WHERE     A.CO_LOCA_LIQU     = B.CO_LOCA_LIQU    " & vbNewLine & _
             " AND       A.CO_GRUP_USUA     = C.CO_GRUP_USUA    " & vbNewLine & _
             " AND       A.TP_OPER          = " & pintTipoOperacao & vbNewLine & _
             " AND       A.CO_FCAO_SIST     = " & pintFuncaoSistema & vbNewLine & _
             " AND       A.CO_EMPR          = " & pintEmpresa & vbNewLine & _
             " AND       B.CO_EMPR_FUSI     = 1"
             
    strWhereSegregacao = fgSegregaDados("A8.TB_PARM_FCAO_SIST", _
                                        False, _
                                        "A", _
                                        "A", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        False)
    
    If strWhereSegregacao <> vbNullString Then
        strSQL = strSQL & vbNewLine & _
            strWhereSegregacao
    End If
    
    strSQL = strSQL & vbNewLine & _
             " ORDER BY  A.SG_SIST, B.SG_LOCA_LIQU, C.NO_GRUP_USUA"

    ObterExcecoesParametrizacao = fgQueryXMLLerTodos("ParamWorkflow", strSQL, App.EXEName & "." & TypeName(Me))

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

'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
''      TB_PARM_FCAO_SIST_EXCE
Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT * FROM A8.TB_PARM_FCAO_SIST_EXCE"
    ObterPropriedades = fgPropriedades("Grupo_ParamExcecoes", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
