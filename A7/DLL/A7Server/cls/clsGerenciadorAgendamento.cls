VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorAgendamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela liberação de mensagens agendadas.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private ObjectContext                       As COMSVCSLib.ObjectContext
Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'Liberar uma mensagem agendada.
Public Function LiberarAgendamento(ByVal pstrNomeObjeto As String) As Boolean
    
Dim objAgendamento                          As Object

On Error GoTo ErrorHandler
    
    Set objAgendamento = CreateObject(pstrNomeObjeto)
    LiberarAgendamento = objAgendamento.LiberarAgendamento
    Set objAgendamento = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objAgendamento = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarAgendamento", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Public Function Reprocessar(ByVal pstrNomeObjeto As String) As Boolean
'
'Dim objReprocessamento                          As Object
'
'On Error GoTo ErrorHandler
'
'    Set objReprocessamento = CreateObject(pstrNomeObjeto)
'    Reprocessar = objReprocessamento.Reprocessar
'    Set objReprocessamento = Nothing
'
'    If Not ObjectContext Is Nothing Then
'        ObjectContext.SetComplete
'    End If
'
'    Exit Function
'
'ErrorHandler:
'    Set objReprocessamento = Nothing
'
'    If Not ObjectContext Is Nothing Then
'        ObjectContext.DisableCommit
'    End If
'
'    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
'    Call fgRaiseError(App.EXEName, TypeName(Me), "Reprocessar", lngCodigoErroNegocio, intNumeroSequencialErro)
'
'End Function

