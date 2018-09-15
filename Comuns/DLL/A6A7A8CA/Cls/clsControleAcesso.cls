VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleAcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pelo controle de acesso de usuários ao sistema.

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

'Obter a Estação de Trabalho de um usuário

Public Function ObterEstacaoTrabalhoUsuario(ByVal pstrUsuarioRede As String) As String

Dim objUsuario                              As Object
Dim strEstacaoTrabalho                      As String

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    
    Call objUsuario.ObterEstacaoTrabalhoUsuario(pstrUsuarioRede, strEstacaoTrabalho, 0, 1)
    
    ObterEstacaoTrabalhoUsuario = strEstacaoTrabalho
    
    Set objUsuario = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "fgObterEstacaoTrabalhoUsuario", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


