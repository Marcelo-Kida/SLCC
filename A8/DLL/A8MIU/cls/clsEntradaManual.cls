VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3BE04B011C"
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

Public Function LerTodosGrupoMensagem(plCodigoEmpresa As Long, psGrupoMensagem As String) As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosGrupoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodosServicos(plCodigoEmpresa As Long) As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosServicos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodosEventos(plCodigoEmpresa As Long, psGrupoMensagem As String, psServico As String) As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosEventos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerConfiguracaoMensagem(psCodigoMensagem As String) As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerConfiguracaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Salvar(objDomMensagem As DOMDocument40) As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodosEmpresa() As String

Dim objEntradaManual                        As A8Liquidacao.clsEntradaManual

On Error GoTo ErrorHandler

    Set objEntradaManual = CreateObject("A8Liquidacao.clsEntradaManual")
    LerTodosEmpresa = objEntradaManual.LerTodosEmpresa()
    Set objEntradaManual = Nothing

    Exit Function
ErrorHandler:
    
    Set objEntradaManual = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosEmpresa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodosMensagens(plCodigoEmpresa As Long, psGrupoMensagem As String, psServico As String, psEvento As String) As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosMensagens Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
