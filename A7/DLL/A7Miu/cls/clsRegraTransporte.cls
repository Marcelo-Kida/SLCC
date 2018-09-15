VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegraTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo roteamento de informações provenientes da camada de interface (A7) para a camada de realização de casos de uso (A7SERVER), utilizando as classes destino:
'    A7Server.clsRegraTransporte
'    A7Server.clsTipoMensagem

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Rotear a chamada do método ObterEventosDisponiveis para a classe de negócio A7Server.clsRegraTransporte.
Public Function ObterEventosDisponiveis(ByVal plngCodigoEmpresa As Long, _
                                        ByVal pstrSilgaSistema As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String

Dim objRegraTransporte                      As A7Server.clsRegraTransporte

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objRegraTransporte = CreateObject("A7Server.clsRegraTransporte")
    ObterEventosDisponiveis = objRegraTransporte.ObterEventosDisponiveis(plngCodigoEmpresa, pstrSilgaSistema)
    Set objRegraTransporte = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objRegraTransporte = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterEventosDisponiveis Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Rotear a chamada do método ExisteTipoMensagem para a classe de negócio A7Server.clsTipoMensagem.
Public Function ExisteTipoMensagem(ByVal pstrTipoMensagem As String, _
                                   ByVal plngTipoFormatoMensagemSaida As Long, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As Boolean

Dim objTipoMensagem                      As A7Server.clsTipoMensagem
Dim strTipoMensagem                      As String

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objTipoMensagem = CreateObject("A7Server.clsTipoMensagem")
    
    ExisteTipoMensagem = objTipoMensagem.ExisteTipoMensagem(pstrTipoMensagem, _
                                                            plngTipoFormatoMensagemSaida)
        
    Set objTipoMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objTipoMensagem = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterEventosDisponiveis Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

