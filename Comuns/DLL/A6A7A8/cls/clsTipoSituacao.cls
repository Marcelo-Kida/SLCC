VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoSituacao"
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

'Gerenciar as operações de Inclusção,Alteração,Exclusão,Ler e LerTodos

Public Function Executar(ByVal pxmlProcessamento As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = vbNullString
    
    If Not pxmlProcessamento.selectSingleNode("//@Operacao") Is Nothing Then
        strOperacao = pxmlProcessamento.selectSingleNode("//@Operacao").Text
    End If

    Select Case strOperacao
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

'Recuperar registros da tabela de tipo de liquidação

Public Function LerTodos() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
'    blnFiltro = CBool(pobjTipoLiquidacao.documentElement.selectSingleNode("Filtro").Text)
    strSQL = " SELECT A.CO_SITU_PROC,   " & _
             "        A.DE_SITU_PROC    " & _
             " FROM   A8.TB_SITU_PROC A " & _
             " ORDER  BY A.CO_SITU_PROC "
    
    LerTodos = fgQueryXMLLerTodos("TipoSituacao", strSQL, App.EXEName & "." & TypeName(Me))

Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
