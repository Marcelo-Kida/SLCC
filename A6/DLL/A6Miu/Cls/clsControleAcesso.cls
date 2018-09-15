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

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Controle de Acesso de Usuários ao Sistema.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

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

' Este método aciona a leitura da descrição do Tipo de Backoffice do usuário logado na rede, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDescricaoTipoBackoffice(ByVal strTipoBackoffice As String) As String
    
Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    intNumeroSequencialErro = 1
    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    intNumeroSequencialErro = 2
    ObterDescricaoTipoBackoffice = objControleAcesso.ObterDescricaoTipoBackoffice(CLng(strTipoBackoffice))
    intNumeroSequencialErro = 3
    Set objControleAcesso = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set objControleAcesso = Nothing
        
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDescricaoTipoBackoffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método aciona a leitura do Tipo de Backoffice do usuário logado na rede, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterTipoBackOfficeUsuario(ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As Integer

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    intNumeroSequencialErro = 1
    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    intNumeroSequencialErro = 2
    ObterTipoBackOfficeUsuario = Val(objControleAcesso.ObterTipoBackOfficeUsuarioLogado())
    intNumeroSequencialErro = 3
    Set objControleAcesso = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set objControleAcesso = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTipoBackOfficeUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
