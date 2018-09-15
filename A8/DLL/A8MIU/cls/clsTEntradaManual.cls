VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTEntradaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3CEA3B0014"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
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

Public Function Salvar(xmlMensagem As DOMDocument40) As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

