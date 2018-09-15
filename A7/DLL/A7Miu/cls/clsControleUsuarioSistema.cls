VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleUsuarioSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo Controle de Acesso Usuario

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

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

'Rotear a chamada do método LerTodos para a classe de negócio A7Server.clsControleUsuarioSistema

Public Function LerTodos(ByVal pstrUsuario As String, _
                         ByVal pstrSiglaSistema As String, _
                         ByVal pstrDataUltimoAcesso As String) As String

Dim objControleUsuario                         As A7Server.clsControleUsuarioSistema

    On Error GoTo ErrorHandler
    
    Set objControleUsuario = CreateObject("A7Server.clsControleUsuarioSistema")
    
    LerTodos = objControleUsuario.LerTodos(pstrUsuario, _
                                           pstrSiglaSistema, _
                                           pstrDataUltimoAcesso)
    Set objControleUsuario = Nothing

    Exit Function

ErrorHandler:

    Set objControleUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

