VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsParamSituContigencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações relativas às
'' parametrizações de sistemas em contingência, bem como a manipulação dos dados
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

'' Consultar todos os sitemas em estado de contingência cadastrados
Public Function LerTodos(Optional ByVal plngIndicadorContingencia As Long) As String
Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT DISTINCT A.SG_SIST,                 " & vbCrLf & _
             "                 B.NO_SIST,                 " & vbCrLf & _
             "                 A.IN_SIST_SITU_CNTG,       " & vbCrLf & _
             "                 A.CO_USUA_ULTI_ATLZ,       " & vbCrLf & _
             "                 A.CO_ETCA_TRAB_ULTI_ATLZ,  " & vbCrLf & _
             "                 A.DH_ULTI_ATLZ             " & vbCrLf & _
             " FROM            A8.TB_SITU_CNTG_SIST A,    " & vbCrLf & _
             "                 A7.TB_SIST B               " & vbCrLf & _
             " WHERE           B.SG_SIST = A.SG_SIST      " & vbCrLf
             
    If plngIndicadorContingencia <> 0 Then
        strSQL = strSQL & _
            " AND     A.IN_SIST_SITU_CNTG = " & plngIndicadorContingencia & vbCrLf
    End If
             
    strSQL = strSQL & _
             " ORDER BY   A.SG_SIST ASC           "
    
    LerTodos = fgQueryXMLLerTodos("ParamSituContigencia", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
'' 	TB_SITU_CNTG_SIST
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT  SG_SIST,                 " & vbNewLine & _
             "         IN_SIST_SITU_CNTG,       " & vbNewLine & _
             "         CO_USUA_ULTI_ATLZ,       " & vbNewLine & _
             "         CO_ETCA_TRAB_ULTI_ATLZ,  " & vbNewLine & _
             "         DH_ULTI_ATLZ             " & vbNewLine & _
             " FROM   A8.TB_SITU_CNTG_SIST A    "

    ObterPropriedades = fgPropriedades("Grupo_ParamSituContigencia", strSQL, App.EXEName & "." & TypeName(Me))

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
Public Function Executar(ByVal pxmlParamSituContigencia As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlParamSituContigencia.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlParamSituContigencia)
        Case "Ler"
            Executar = Ler(pxmlParamSituContigencia.documentElement.selectSingleNode("//SG_SIST").Text)
        Case "LerTodos"
            Executar = LerTodos(CLng("0" & pxmlParamSituContigencia.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text))
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
'' Promover a inclusão de uma nova parametrização de sistema em contingência
Private Function flIncluir(ByRef pxmlParamSituContigencia As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_SITU_CNTG_SIST (" & _
             " SG_SIST, " & _
             " IN_SIST_SITU_CNTG, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( '" & _
             pxmlParamSituContigencia.documentElement.selectSingleNode("SG_SIST").Text & "'," & _
             pxmlParamSituContigencia.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text & "," & _
             "'" & fgUsuarioRede & "'," & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Promover a alteração de uma parametrização de sistema em contingência existente
Private Function flAlterar(ByRef pxmlParamSituContigencia As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_SITU_CNTG_SIST " & _
            "  SET    SG_SIST            = '" & pxmlParamSituContigencia.documentElement.selectSingleNode("SG_SIST").Text & "'," & _
            "         IN_SIST_SITU_CNTG  = " & pxmlParamSituContigencia.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text & "," & _
            "         CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ            =  SYSDATE " & _
            " WHERE   SG_SIST            = '" & pxmlParamSituContigencia.documentElement.selectSingleNode("SG_SIST").Text & "'" & _
            " AND     DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlParamSituContigencia.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Consultar todos os atributos de uma parametrização de sistema em contingência
'' específica
Public Function Ler(ByVal pstrSG_SIST As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT  A.SG_SIST,                 " & _
             "         B.NO_SIST,                 " & _
             "         A.IN_SIST_SITU_CNTG,       " & _
             "         A.CO_USUA_ULTI_ATLZ,       " & _
             "         A.CO_ETCA_TRAB_ULTI_ATLZ,  " & _
             "         A.DH_ULTI_ATLZ             " & _
             " FROM    A8.TB_SITU_CNTG_SIST A,    " & _
             "         A7.TB_SIST B               " & _
             " WHERE   B.SG_SIST = A.SG_SIST      " & _
             " AND     A.SG_SIST = '" & pstrSG_SIST & "'"
    
    Ler = fgQueryXMLLer("Grupo_ParamSituContigencia", strSQL, App.EXEName & "." & TypeName(Me))
    
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
Public Function Salvar(ByVal pxmlParamSituContigencia As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlParamSituContigencia.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlParamSituContigencia
    ElseIf pxmlParamSituContigencia.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlParamSituContigencia
    ElseIf pxmlParamSituContigencia.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlParamSituContigencia
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

'' Promover a exclusão de uma parametrização de sistema em contingência existente
Private Function flExcluir(ByRef pxmlParamSituContigencia As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE " & strSQL & _
             " FROM    A8.TB_SITU_CNTG_SIST  " & vbNewLine & _
             " WHERE   SG_SIST = '" & pxmlParamSituContigencia.documentElement.selectSingleNode("SG_SIST").Text & "'" & _
             " AND DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlParamSituContigencia.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function






