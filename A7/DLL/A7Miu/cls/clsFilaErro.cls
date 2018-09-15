VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilaErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo roteamento de informações provenientes da camada de interface (A7) para a camada de realização de casos de uso (A7SERVER), utilizando as classes destino:
'    A7Server.clsFilaErro
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

'Rotear a chamada do método ReprocessaMensagens para a classe de negócio A7Server.clsFilaErro.
Public Function ReprocessaMensagens(ByVal pstrLote As String, _
                                    ByRef pvntCodErro As Variant, _
                                    ByRef pvntMensagemErro As Variant) As String

Dim objFilaErro  As A7Server.clsFilaErro

    On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFilaErro = CreateObject("A7Server.clsFilaErro")
    
    ReprocessaMensagens = objFilaErro.ReprocessarMensagens(pstrLote)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objFilaErro = Nothing

    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objFilaErro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ReprocessaMensagem", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Rotear a chamada do método LerMensagem para a classe de negócio A7Server.clsFilaErro.
Public Function LerMensagem(ByVal pstrFiltro As String, _
                            ByRef pvntCodErro As Variant, _
                            ByRef pvntMensagemErro As Variant) As String

Dim objFilaErro                             As A7Server.clsFilaErro

    On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFilaErro = CreateObject("A7Server.clsFilaErro")
    
    LerMensagem = objFilaErro.LerMensagens(pstrFiltro)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objFilaErro = Nothing

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objFilaErro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagem", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function


