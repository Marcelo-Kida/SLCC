VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultaRemessaRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito às Consultas de Remessas Rejeitadas.

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

' Este método aciona a leitura de um registro do cadastro de remessas rejeitadas pelo A6, na camada de negócios,
' e retorna para a camada de interface.

Public Function Ler(ByVal strSG_SIST_ORIG_INFO As String, _
                    ByVal intTP_MESG_INTE As Integer, _
                    ByVal lngCO_EMPR As Long, _
                    ByVal lngCO_TEXT_XML_REJE As Long, _
                    ByVal strDH_REME_REJE As String, _
                    ByVal strOwner As String, _
                    ByRef pvntCodErro As Variant, _
                    ByRef pvntMensagemErro As Variant) As String

Dim objRemessaRejeitada                     As A6SubReserva.clsRemessaRejeitada

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objRemessaRejeitada = CreateObject("A6SubReserva.clsRemessaRejeitada")

    Ler = objRemessaRejeitada.Ler(strSG_SIST_ORIG_INFO, _
                                  intTP_MESG_INTE, _
                                  lngCO_EMPR, _
                                  lngCO_TEXT_XML_REJE, _
                                  strDH_REME_REJE, _
                                  strOwner)

    Set objRemessaRejeitada = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura do conteúdo XML da remessa rejeitada, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterXMLErroRemessaRejeitada(ByVal lngCO_TEXT_XML_REJE As Long, _
                                             ByRef pvntCodErro As Variant, _
                                             ByRef pvntMensagemErro As Variant) As String

Dim objRemessaRejeitada  As A6SubReserva.clsRemessaRejeitada

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objRemessaRejeitada = CreateObject("A6SubReserva.clsRemessaRejeitada")

    ObterXMLErroRemessaRejeitada = objRemessaRejeitada.ObterXMLErroRemessaRejeitada(lngCO_TEXT_XML_REJE)

    Set objRemessaRejeitada = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterXMLErroRemessaRejeitada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura ao cadastro de remessas rejeitadas pelo A6, na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodos(pxmlDocFiltros As String, _
                         pvntCodErro As Variant, _
                         pvntMensagemErro As Variant) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objRemessaRejeitada                     As A6SubReserva.clsRemessaRejeitada

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objRemessaRejeitada = CreateObject("A6SubReserva.clsRemessaRejeitada")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    LerTodos = objRemessaRejeitada.LerTodos(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objRemessaRejeitada = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
