VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaCorrenteCOTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1BE7DA0214"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pelos métodos de controle do parâmetro de Conta Corrente de Corretoras.

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

'Inclui uma parametrização de Conta Corrente de Corretoras.
Public Function flIncluir(ByRef pxmlParametroCCCorretora As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim strCodSubTipo                           As String

    On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_LOCA_LIQU_CNTA_COTR (  " & vbNewLine & _
             " CO_LOCA_LIQU,                     " & vbNewLine & _
             " CO_EMPR,                          " & vbNewLine & _
             " CO_AGEN_COTR,                     " & vbNewLine & _
             " NU_CC_COTR,                       " & vbNewLine & _
             " CO_UNI_ORG                        " & vbNewLine & _
             " ) VALUES (                        " & vbNewLine

    With pxmlParametroCCCorretora.documentElement
    
        strSQL = strSQL & _
            " " & .selectSingleNode("//CO_LOCA_LIQU").Text & ",   " & vbNewLine & _
            " " & .selectSingleNode("//CO_EMPR").Text & ",        " & vbNewLine & _
            " " & .selectSingleNode("//CO_AGEN_COTR").Text & ",   " & vbNewLine & _
            " " & .selectSingleNode("//NU_CC_COTR").Text & ",     " & vbNewLine & _
            " " & .selectSingleNode("//CO_UNI_ORG").Text & ")     " & vbNewLine

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Altera uma parametrização de conta corrente de corretoras
Private Function flAlterar(ByRef pxmlParametroCCCorretora As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

'    flConsiste pxmlParametroCCCorretora

    With pxmlParametroCCCorretora.documentElement
        strSQL = " UPDATE A8.TB_LOCA_LIQU_CNTA_COTR                                     " & vbNewLine & _
                 " SET    CO_UNI_ORG   = " & .selectSingleNode("//CO_UNI_ORG").Text & " " & vbNewLine & _
                 " WHERE  CO_LOCA_LIQU = " & .selectSingleNode("//CO_LOCA_LIQU").Text & " " & vbNewLine & _
                 " AND    CO_EMPR      = " & .selectSingleNode("//CO_EMPR").Text & " " & vbNewLine & _
                 " AND    CO_AGEN_COTR = " & .selectSingleNode("//CO_AGEN_COTR").Text & " " & vbNewLine & _
                 " AND    NU_CC_COTR   = " & .selectSingleNode("//NU_CC_COTR").Text & " " & vbNewLine & _
                 " AND    CO_UNI_ORG   = " & .selectSingleNode("//CO_UNI_ORG_OLD").Text

    End With
    
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

'Salva uma alteração de parametrização de conta corrente de corretora, seja de inclusão, alteração ou exclusão
Public Function Salvar(ByVal pxmlParametroCCCorretora As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlParametroCCCorretora.documentElement.selectSingleNode("//@Operacao").Text = "Incluir" Then
        flIncluir pxmlParametroCCCorretora
    ElseIf pxmlParametroCCCorretora.documentElement.selectSingleNode("//@Operacao").Text = "Alterar" Then
        flAlterar pxmlParametroCCCorretora
    ElseIf pxmlParametroCCCorretora.documentElement.selectSingleNode("//@Operacao").Text = "Excluir" Then
        flExcluir pxmlParametroCCCorretora
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

'Exclui uma parametrização de conta corrente de corretora
Private Function flExcluir(ByRef pxmlParametroCCCorretora As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlParametroCCCorretora.documentElement

        strSQL = " DELETE FROM A8.TB_LOCA_LIQU_CNTA_COTR  " & vbNewLine & _
                 " WHERE  CO_LOCA_LIQU      = " & .selectSingleNode("//CO_LOCA_LIQU").Text & " " & vbNewLine & _
                 " AND    CO_EMPR           = " & .selectSingleNode("//CO_EMPR").Text & " " & vbNewLine & _
                 " AND    CO_AGEN_COTR      = " & .selectSingleNode("//CO_AGEN_COTR").Text & " " & vbNewLine & _
                 " AND    NU_CC_COTR        = " & .selectSingleNode("//NU_CC_COTR").Text & " " & vbNewLine & _
                 " AND    CO_UNI_ORG        = " & .selectSingleNode("//CO_UNI_ORG").Text
             
    End With

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

''Obter todas as parametrizações de conta corrente de corretora
'Public Function ObterPropriedades() As String
'
'Dim strSQL                                  As String
'
'    On Error GoTo ErrorHandler
'
'    strSQL = " SELECT DISTINCT(A.CO_LOCA_LIQU),  " & vbNewLine & _
'             "        C.DE_LOCA_LIQU,            " & vbNewLine & _
'             "        A.CO_EMPR,                 " & vbNewLine & _
'             "        B.NO_REDU_EMPR,            " & vbNewLine & _
'             "        A.CO_AGEN_COTR,            " & vbNewLine & _
'             "        A.NU_CC_COTR,              " & vbNewLine & _
'             "        A.CO_UNI_ORG               " & vbNewLine & _
'             " FROM   A8.TB_LOCA_LIQU_CNTA_COTR A, A8.TB_EMPRESA_HO B, A8.TB_LOCAL_LIQUIDACAO C  " & vbNewLine & _
'             " WHERE A.CO_EMPR = B.CO_EMPR " & vbNewLine & _
'             " AND A.CO_LOCA_LIQU = C.CO_LOCA_LIQU "
'
'    ObterPropriedades = fgPropriedades("Grupo_ContaCorrenteCorretora", strSQL, App.EXEName & "." & TypeName(Me))
'
'    Exit Function
'
'ErrorHandler:
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)
'
'End Function

'Obter todas as parametrizações de conta corrente de corretora, utilizando a segregação de acesso
Public Function LerTodos() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT A.CO_LOCA_LIQU,                 " & vbNewLine & _
             "        C.DE_LOCA_LIQU,                 " & vbNewLine & _
             "        A.CO_EMPR,                      " & vbNewLine & _
             "        B.NO_REDU_EMPR,                 " & vbNewLine & _
             "        A.CO_AGEN_COTR,                 " & vbNewLine & _
             "        A.NU_CC_COTR,                   " & vbNewLine & _
             "        A.CO_UNI_ORG,                   " & vbNewLine & _
             "        A.CO_AGEN_COTR CO_AGEN_OLD,     " & vbNewLine & _
             "        A.NU_CC_COTR NU_CC_COTR_OLD,    " & vbNewLine & _
             "        A.CO_UNI_ORG  CO_UNI_ORG_OLD    " & vbNewLine & _
             " FROM   A8.TB_LOCA_LIQU_CNTA_COTR A,    " & vbNewLine & _
             "        A8.TB_EMPRESA_HO          B,    " & vbNewLine & _
             "        A8.TB_LOCAL_LIQUIDACAO    C     " & vbNewLine & _
             " WHERE  A.CO_EMPR = B.CO_EMPR           " & vbNewLine & _
             " AND    A.CO_LOCA_LIQU = C.CO_LOCA_LIQU " & vbNewLine & _
             " AND    C.CO_EMPR_FUSI = 1 "

    strSQL = strSQL & fgSegregaDados("A8.TB_LOCAL_LIQUIDACAO", _
                                     False, _
                                     "C", _
                                     "C", _
                                     False, _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False)
    
    strSQL = strSQL & " ORDER BY C.CO_LOCA_LIQU "
    
    LerTodos = fgQueryXMLLerTodos("ParametroCCCorretora", strSQL, App.EXEName & "." & TypeName(Me))

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

'Obter dados de uma parametrização de conta corrente de corretora
Public Function Ler(ByVal pstrLocalLiquidacao As String, _
                    ByVal plngCodigoEmpresa As Long, _
                    ByVal plngAgencia As Long, _
                    ByVal plngContaCorrente As Long, _
                    ByVal plngUniOrg As Long) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_LOCA_LIQU,             " & vbNewLine & _
             "          CO_EMPR,                  " & vbNewLine & _
             "          CO_AGEN_COTR,             " & vbNewLine & _
             "          NU_CC_COTR,               " & vbNewLine & _
             "          CO_UNI_ORG                " & vbNewLine & _
             " FROM     A8.TB_LOCA_LIQU_CNTA_COTR " & vbNewLine & _
             " WHERE    CO_LOCA_LIQU        =   " & pstrLocalLiquidacao & vbNewLine & _
             " AND      CO_EMPR             =   " & plngCodigoEmpresa & vbNewLine & _
             " AND      CO_AGEN_COTR        =   " & plngAgencia & vbNewLine & _
             " AND      NU_CC_COTR          =   " & plngContaCorrente & vbNewLine & _
             " AND      CO_UNI_ORG          =   " & plngUniOrg & vbNewLine
             
    Ler = fgQueryXMLLer("ParametroCCCorretora", strSQL, App.EXEName & "." & TypeName(Me))

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

'Executa uma tarefa em parametrização de conta corrente corretora, seja de leitura ou alteração
Public Function Executar(ByVal pxmlParametroCCCorretora As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlParametroCCCorretora.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir", "Alterar"
            Salvar pxmlParametroCCCorretora
        Case "Ler"
            Executar = LerTodos()
'            With pxmlParametroCCCorretora.documentElement
'                Executar = Ler(.selectSingleNode("//CO_LOCA_LIQU").Text, _
'                               CLng("0" & .selectSingleNode("//CO_EMPR").Text), _
'                               CLng("0" & .selectSingleNode("//CO_AGEN_COTR").Text), _
'                               CLng("0" & .selectSingleNode("//NU_CC_COTR").Text), _
'                               .selectSingleNode("//CO_UNI_ORG").Text)
'            End With
        Case "LerTodos"
            Executar = LerTodos()
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

'Valida dados de uma parametrização de conta corrente corretora
Private Function flConsiste(ByRef pxmlParametroCCCorretora As MSXML2.DOMDocument40) As String

On Error GoTo ErrorHandler

    With pxmlParametroCCCorretora.documentElement
         If CLng(.selectSingleNode("//CO_EMPR").Text) = enumCodigoEmpresa.Banespa Then
             If Len(Trim(.selectSingleNode("//CO_HIST_CC").Text)) > 4 Then
                ' 3048 - Tamanho máximo do Código do Histórico para o Banespa é 4 posições.
                lngCodigoErroNegocio = 3048
                GoTo ErrorHandler
             End If
         End If
    End With

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConsiste Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


