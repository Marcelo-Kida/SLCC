VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracaoSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsMonitoracaoFluxoCaixaSubReserva\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Monitoração do Fluxo do Caixa Sub-Reserva\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F149F0D0111"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Monitoração de Caixa Sub-Reserva.

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

' Este método aciona a leitura do movimento de caixa sub-reserva, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterMovimentoSubReserva(ByVal pxmlDocFiltros As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String

Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva
Dim xmlSubReserva                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
    Set xmlSubReserva = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlSubReserva.loadXML pxmlDocFiltros
    ObterMovimentoSubReserva = objMovimentoSubReserva.ObterMovimentoSubReserva(xmlSubReserva)
       
    Set objMovimentoSubReserva = Nothing
    Set xmlSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMovimentoSubReserva = Nothing
    Set xmlSubReserva = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMovimentoSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura do resumo de movimento de caixa sub-reserva, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterResumoCaixaSubReserva(ByVal pxmlDocFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String
    
Dim objResumoCaixaSubReserva                As A6SubReserva.clsMovimentoSubReserva
Dim xmlSubReserva                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objResumoCaixaSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
    Set xmlSubReserva = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlSubReserva.loadXML pxmlDocFiltros
    ObterResumoCaixaSubReserva = objResumoCaixaSubReserva.ObterResumoCaixaSubReserva(xmlSubReserva)
    
    Set objResumoCaixaSubReserva = Nothing
    Set xmlSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objResumoCaixaSubReserva = Nothing
    Set xmlSubReserva = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterResumoCaixaSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de detalhes do movimento de caixa sub-reserva, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheMovimento(ByVal pxmlDocFiltros As String, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva
Dim xmlSubReserva                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
    Set xmlSubReserva = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlSubReserva.loadXML pxmlDocFiltros
    ObterDetalheMovimento = objMovimentoSubReserva.ObterDetalheMovimento(xmlSubReserva)
       
    Set objMovimentoSubReserva = Nothing
    Set xmlSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMovimentoSubReserva = Nothing
    Set xmlSubReserva = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMovimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
