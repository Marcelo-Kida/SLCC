VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSituacaoProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 

'' Objeto responsável pelo fornecimento de informações relativas às situações de
'' processo cadastradas, tanto para operações, quanto para mensagens

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

'Executar as operações (Inclusão,Alteração,Exclusão)

Public Function Executar(ByRef pxmlEmpresa As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlEmpresa.selectSingleNode("//*/@Operacao").Text

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

'recuperar os registros da tabela de situação A8.TB_SITU_PROC

Public Function LerTodos() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_SITU_PROC,      " & vbCrLf & _
             "          TP_SITU_PROC,      " & vbCrLf & _
             "          DE_SITU_PROC,      " & vbCrLf & _
             "          TX_COMT_SITU_PROC  " & vbCrLf & _
             " FROM     A8.TB_SITU_PROC A  " & vbCrLf & _
             " ORDER BY TP_SITU_PROC,      " & vbCrLf & _
             "          DE_SITU_PROC       " & vbCrLf
             
    LerTodos = fgQueryXMLLerTodos("SituacaoProcesso", strSQL, App.EXEName & "." & TypeName(Me))

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


