VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultaMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
''-------------------------------------------------
' Gerado por Adrian Pretel
' Atualização em:      20-jun-2005
''-------------------------------------------------
''
' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à tela de Consulta de Movimentação.
''
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

' Este método aciona a leitura de detalhes das Movimentações de Operações a partir da chave passada como parâmetro,
' na camada de negócio, e retorna para a camada de interface.

Public Function Ler(pxmlDocFiltros As String, _
              ByRef pvntCodErro As Variant, _
              ByRef pvntMensagemErro As Variant) As String

Dim xmlFiltros      As MSXML2.DOMDocument40
Dim objMovimentacao As A8LQS.clsConsultaMovimentacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMovimentacao = CreateObject("A8LQS.clsConsultaMovimentacao")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    Ler = objMovimentacao.Ler(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objMovimentacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
Exit Function
ErrorHandler:

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de detalhes das Movimentações de Operações Rejeitadas
' a partir da chave passada como parâmetro, na camada de negócio, e retorna para a camada de interface.

Public Function LerRejeitada(pxmlDocFiltros As String, _
                       ByRef pvntCodErro As Variant, _
                       ByRef pvntMensagemErro As Variant) As String

Dim xmlFiltros      As MSXML2.DOMDocument40
Dim objMovimentacao As A8LQS.clsConsultaMovimentacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMovimentacao = CreateObject("A8LQS.clsConsultaMovimentacao")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    LerRejeitada = objMovimentacao.LerRejeitada(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objMovimentacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
Exit Function
ErrorHandler:

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerRejeitada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de detalhes das Movimentações de Integração de Conta Corrente e Contabilidade
' a partir da chave passada como parâmetro, na camada de negócio, e retorna para a camada de interface.

Public Function LerCC_HA(pxmlDocFiltros As String, _
                   ByRef pvntCodErro As Variant, _
                   ByRef pvntMensagemErro As Variant) As String

Dim xmlFiltros      As MSXML2.DOMDocument40
Dim objMovimentacao As A8LQS.clsConsultaMovimentacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMovimentacao = CreateObject("A8LQS.clsConsultaMovimentacao")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    LerCC_HA = objMovimentacao.LerCC_HA(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objMovimentacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
Exit Function
ErrorHandler:

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerCC_HA Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

