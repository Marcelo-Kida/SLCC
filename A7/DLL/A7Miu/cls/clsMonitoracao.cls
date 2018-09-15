VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo roteamento de informações provenientes da camada de interface (A7) para a camada de realização de casos de uso (A7SERVER), utilizando as classes destino:
'    A7Server.clsMonitoração
'    A7Server.clsSistema

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

'Rotear a chamada do método ObterQtd para a classe de negócio A7Server.clsMonitoracao.
Public Function ObterQtd(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String, _
                         ByVal pintEmpresa As Integer, _
                         ByVal pstrSiglaSistema As String, _
                         ByVal pstrTipoMensagem As String, _
                         ByVal pstrIDMensagem As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objMonitoracao                          As A7Server.clsMonitoracao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMonitoracao = CreateObject("A7Server.clsMonitoracao")
    ObterQtd = objMonitoracao.ObterQtd(pstrDataHoraDe, _
                                       pstrDataHoraAte, _
                                       pintEmpresa, _
                                       pstrSiglaSistema, _
                                       pstrTipoMensagem, _
                                       pstrIDMensagem)
    Set objMonitoracao = Nothing

    Exit Function

ErrorHandler:
    
    Set objMonitoracao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQtd Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Rotear a chamada do método LerTodos para a classe de negócio A7Server.clsMonitoracao.
Public Function LerTodos(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String, _
                         ByVal pintEmpresa As Integer, _
                         ByVal pstrSiglaSistema As String, _
                         ByVal pstrTipoMensagem As String, _
                         ByVal pstrCodigoOcorrencia As String, _
                         ByVal pstrIDMensagem As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objMonitoracao                          As A7Server.clsMonitoracao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMonitoracao = CreateObject("A7Server.clsMonitoracao")
    LerTodos = objMonitoracao.LerTodos(pstrDataHoraDe, _
                                       pstrDataHoraAte, _
                                       pintEmpresa, _
                                       pstrSiglaSistema, _
                                       pstrTipoMensagem, _
                                       pstrCodigoOcorrencia, _
                                       pstrIDMensagem)
    Set objMonitoracao = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objMonitoracao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Rotear a chamada do método ObterDetalheMensagem para a classe de negócio A7Server.clsMonitoracao.
Public Function ObterDetalheMensagem(ByVal plCodigoMensagem As Long, _
                                     ByRef pvntCodErro As Variant, _
                                     ByRef pvntMensagemErro As Variant) As String

Dim objMonitoracao                          As A7Server.clsMonitoracao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMonitoracao = CreateObject("A7Server.clsMonitoracao")
    ObterDetalheMensagem = objMonitoracao.ObterDetalheMensagem(plCodigoMensagem)
    Set objMonitoracao = Nothing

    Exit Function

ErrorHandler:

    Set objMonitoracao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Rotear a chamada do método ObterHistoricoEvento para a classe de negócio A7Server.clsMonitoracao.
Public Function ObterHistoricoEvento(ByVal plCodigoMensagem As Long) As String

Dim objMonitoracao                          As A7Server.clsMonitoracao

    On Error GoTo ErrorHandler
    
    Set objMonitoracao = CreateObject("A7Server.clsMonitoracao")
    ObterHistoricoEvento = objMonitoracao.ObterDetalheMensagem(plCodigoMensagem)
    Set objMonitoracao = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objMonitoracao = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoEvento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Rotear a chamada do método ObterSistemasOrigem para a classe de negócio A7Server.clsSistema.
Public Function ObterSistemasOrigem(ByVal pntiEmpresa As Integer, _
                                    ByRef pvntCodErro As Variant, _
                                    ByRef pvntMensagemErro As Variant) As String

Dim objSistema                              As A6A7A8.clsSistema

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objSistema = CreateObject("A6A7A8.clsSistema")
    ObterSistemasOrigem = objSistema.LerTodos(pntiEmpresa, "S")
    Set objSistema = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objSistema = Nothing
    
    'Comentado devido ao tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSistemasOrigem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Obter Intervalo de parada do verifica Server
Public Function ObterIntervaloVerificaServer(ByRef pvntCodErro As Variant, _
                                             ByRef pvntMensagemErro As Variant) As String
    
Dim objMonitoracao                          As A7Server.clsMonitoracao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMonitoracao = CreateObject("A7Server.clsMonitoracao")
    ObterIntervaloVerificaServer = objMonitoracao.ObterIntervaloVerificaServer
    Set objMonitoracao = Nothing

Exit Function
ErrorHandler:
    
    Set objMonitoracao = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterIntervaloVerificaServer Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

Public Function TesteConectividade(ByVal pstrXMLTesteConectividade As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                              As A7Server.clsMensagem

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagem = CreateObject("A7Server.clsMensagem")
    
    Call objMensagem.EnviarTesteConexao(pstrXMLTesteConectividade)
    
    Set objMensagem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set objMensagem = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSistemasOrigem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

Public Function PostarMensagemArquivoImportado(ByVal pstrXMLTesteConectividade As String, _
                                               ByRef pvntCodErro As Variant, _
                                               ByRef pvntMensagemErro As Variant) As String

Dim objGerenciadorEnvio                     As A7Server.clsGerenciadorEnvio

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objGerenciadorEnvio = CreateObject("A7Server.clsGerenciadorEnvio")
    
    Call objGerenciadorEnvio.EnviarTesteConexao(pstrXMLTesteConectividade)
    
    Set objGerenciadorEnvio = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set objGerenciadorEnvio = Nothing
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

Public Function ObterHistoricoGEN0001(ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                              As A7Server.clsMonitoracao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagem = CreateObject("A7Server.clsMonitoracao")
    
    ObterHistoricoGEN0001 = objMensagem.ObterHistoricoGEN0001()
    
    Set objMensagem = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    Set objMensagem = Nothing
    
    'Comentado devido ao novo tratamento de erros do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoGEN0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function


'Rotear a chamada do método ObterSistemas para a classe de negócio A6A7A8.clsSistema
Public Function ObterSistemas(ByRef pvntCodErro As Variant, _
                              ByRef pvntMensagemErro As Variant) As String

Dim objSistema                              As A6A7A8.clsSistema

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objSistema = CreateObject("A6A7A8.clsSistema")
    ObterSistemas = objSistema.LerTodosIgnorarEmpresa("S")
    Set objSistema = Nothing

    Exit Function

ErrorHandler:

    Set objSistema = Nothing
    
    'Coentado devido ao novo tratamento de erro de SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSistemas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

