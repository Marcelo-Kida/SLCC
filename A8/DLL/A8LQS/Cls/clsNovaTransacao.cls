VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNovaTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Possibilitar a criação de uma transação apartada do contexto corrente.
'Esta classe deve ser configurada como "Requires New Transaction"

Option Explicit

Private objContext                          As COMSVCSLib.ObjectContext
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

'Processa um boleto em uma nova transação
'Public Sub ProcessarBoleto(ByVal pstrQName As String, _
'                           ByRef pstrCorralationID As String)
'
'On Error GoTo ErrorHandler
'Dim objRemessa                              As A8LQS.clsRemessa
'
'    Set objRemessa = CreateObject("A8LQS.clsRemessa")
'
'    objRemessa.ReceberMensagemMQ pstrQName, _
'                                 "", _
'                                 "", _
'                                 0, _
'                                 pstrCorralationID
'
'    Set objRemessa = Nothing
'
'    If Not objContext Is Nothing Then
'        objContext.SetComplete
'    End If
'
'    Exit Sub
'ErrorHandler:
'    Set objRemessa = Nothing
'
'    If Not objContext Is Nothing Then
'        objContext.DisableCommit
'    End If
'
'    Err.Clear
'
'End Sub
