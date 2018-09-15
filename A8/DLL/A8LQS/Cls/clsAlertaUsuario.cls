VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlertaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3E5AF40313"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Componente     :
'Classe         :
'Data Criação   :
'Objetivo       :
'
'Analista       :
'
'Programador    :
'Data           :
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

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

Public Function Ler() As String
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodos() As String
    On Error GoTo ErrorHandler


    '## your code goes here...

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

Public Function Salvar() As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flIncluir() As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAlterar() As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flExcluir() As Boolean
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

