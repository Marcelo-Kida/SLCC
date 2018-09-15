VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Interface para chamada aos metodos dos componentes de controle de acesso do Usuário

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

'LogOn de usuário no sistema

Public Function LogOn(ByVal pstrUsuarioSistema As String, _
                      ByVal pstrUsuarioRede As String, _
                      ByVal pstrEstacaoTrabalho As String, _
             Optional ByVal pblnForcaLogar As Boolean) As Boolean

Dim objPerfil                               As A6A7A8.clsUsuario

On Error GoTo ErrorHandler
    
    Set objPerfil = CreateObject("A6A7A8.clsUsuario")
        
    LogOn = objPerfil.LogOn(pstrUsuarioSistema, _
                            pstrUsuarioRede, _
                            pstrEstacaoTrabalho, _
                            pblnForcaLogar)
        
    Set objPerfil = Nothing

    Exit Function
ErrorHandler:
    
    Set objPerfil = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LogOn Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Logoff de um usuário no sistema

Public Function Logoff(ByVal pstrUsuarioRede As String) As Boolean

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    Logoff = objUsuario.Logoff(pstrUsuarioRede)
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Logoff Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Reconectar um usuário ao sistema

Public Function Reconectar(ByVal pstrUsuarioSistema As String, _
                           ByVal pstrUsuarioRede As String, _
                           ByVal pstrEstacaoTrabalho As String) As Boolean

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    Reconectar = objUsuario.AdicionaUsuario(pstrUsuarioRede, pstrUsuarioSistema, pstrEstacaoTrabalho, False)
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reconectar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o ID do Usuário
Public Function ObterUsuario(ByRef pvntCodErro As Variant, _
                             ByRef pvntMensagemErro As Variant) As String

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    ObterUsuario = objUsuario.ObterUsuario
    Set objUsuario = Nothing

Exit Function
ErrorHandler:
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Listar os usuários Logados

Public Function ListarUsuariosAtivos()

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler

    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    ListarUsuariosAtivos = objUsuario.ListarUsuariosAtivos
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ListarUsuariosAtivos Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'Obter o tipo de backoffice do usuário

Public Function ObterTipoBackOfficeUsuario() As String

Dim objControleAcesso       As A6A7A8.clsControleAcesso
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    
    ObterTipoBackOfficeUsuario = objControleAcesso.ObterTipoBackOfficeUsuario(ObterUsuario(vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objControleAcesso = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOfficeUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

