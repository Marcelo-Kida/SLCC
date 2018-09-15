VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracaoFluxoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Monitoração de Fluxo de Caixa Sub-Reserva.

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

' Este método aciona a leitura ao resumo de movimentação do caixa sub-reserva, na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodos(ByVal pxmlDocFiltros As String) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    LerTodos = objMovimentoSubReserva.ObterResumoCaixaSubReserva(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objMovimentoSubReserva = Nothing

    Exit Function

ErrorHandler:

    Set xmlFiltros = Nothing
    Set objMovimentoSubReserva = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMapaNavegacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
