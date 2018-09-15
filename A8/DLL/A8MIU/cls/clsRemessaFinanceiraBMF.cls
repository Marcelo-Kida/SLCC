VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemessaFinanceiraBMF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao tratamento de Remessas Líquidas Financeiras BMF.

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

' Este método aciona a leitura de detalhes de remessa líquida financeira BMF, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterDetalheRemessaBMF(ByVal strXMLFiltro As String) As String

Dim objRemessa                              As A8LQS.clsRemessaFinanceiraCamara
Dim xmlFiltro                               As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlFiltro.loadXML(strXMLFiltro)

    Set objRemessa = CreateObject("A8LQS.clsRemessaFinanceiraCamara")
    ObterDetalheRemessaBMF = objRemessa.ObterDetalheRemessaCAMR(xmlFiltro)
    Set objRemessa = Nothing

    Set xmlFiltro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objRemessa = Nothing
    Set xmlFiltro = Nothing

    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheRemessaBMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'executar o processamento do arquivo da BMF
Public Function ProcessaRemessaFinanceiraBMF(ByVal pstrNomeArquivo As String, _
                                             ByRef pvntCodErro As Variant, _
                                             ByRef pvntMensagemErro As Variant) As Boolean

Dim objTRemessa                             As A8MIU.clsTRemessaFinanceiraBMF

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objTRemessa = CreateObject("A8MIU.clsTRemessaFinanceiraBMF")
    Call objTRemessa.ProcessaRemessaFinanceiraBMF(pstrNomeArquivo)
    Set objTRemessa = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objTRemessa = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessaRemessaFinanceiraBMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Ler o log das remessas já processadas
Public Function LerTodosLogRemessa(ByVal plngLocalLiquidacao As Long, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objRemessaFinanceiraCamara                As A8LQS.clsRemessaFinanceiraCamara

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objRemessaFinanceiraCamara = CreateObject("A8LQS.clsRemessaFinanceiraCamara")
    LerTodosLogRemessa = objRemessaFinanceiraCamara.LerTodosLogRemessa(plngLocalLiquidacao)
    Set objRemessaFinanceiraCamara = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objRemessaFinanceiraCamara = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterLogImportacaoRemessaBMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
