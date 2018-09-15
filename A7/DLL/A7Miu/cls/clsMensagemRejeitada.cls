VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo roteamento de informações provenientes da camada de interface (A7) para a camada de realização de casos de uso (A7SERVER), utilizando as classes destino:
'    A7Server.clsMensagemRejeitada
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

'Rotear a chamada do método LerTodos para a classe de negócio A7Server.clsMensagemRejeitada.
Public Function LerTodos(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objMensagemRejeitada                      As A7Server.clsMensagemRejeitada

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagemRejeitada = CreateObject("A7Server.clsMensagemRejeitada")
    
    LerTodos = objMensagemRejeitada.LerTodos(pstrDataHoraDe, _
                                             pstrDataHoraAte)
    Set objMensagemRejeitada = Nothing

    Exit Function

ErrorHandler:
    
    Set objMensagemRejeitada = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Rotear a chamada do método ObterMensagem para a classe de negócio A7Server.clsMensagemRejeitada.
Public Function ObterMensagem(ByVal plngCodCntdAtrbXml As Long, _
                              ByRef pvntCodErro As Variant, _
                              ByRef pvntMensagemErro As Variant) As String

Dim objMensagemRejeitada                      As A7Server.clsMensagemRejeitada

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagemRejeitada = CreateObject("A7Server.clsMensagemRejeitada")
    
    ObterMensagem = objMensagemRejeitada.ObterMensagem(plngCodCntdAtrbXml)
    
    Set objMensagemRejeitada = Nothing

    Exit Function

ErrorHandler:
    
    Set objMensagemRejeitada = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

