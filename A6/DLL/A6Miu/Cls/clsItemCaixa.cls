VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Manutenção do Cadastro de Itens de Caixa.

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

' Este método aciona a leitura de itens de caixa agrupados por grupo de veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterRelacaoItensCaixaGrupoVeicLegal(ByVal pstrXMLDocFiltros As String, _
                                                     ByVal pblnSegregarPorTipoBackOffice As Boolean, _
                                                     ByVal pintTipoBackOffice As Integer, _
                                                     ByRef pvntCodErro As Variant, _
                                                     ByRef pvntMensagemErro As Variant) As String
    
Dim objItemCaixaGrupoVeicLegal              As A6SubReserva.clsItemCaixaGrupoVeicLegal

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objItemCaixaGrupoVeicLegal = CreateObject("A6SubReserva.clsItemCaixaGrupoVeicLegal")
    ObterRelacaoItensCaixaGrupoVeicLegal = objItemCaixaGrupoVeicLegal.ObterRelacaoItensCaixaGrupoVeicLegal(pstrXMLDocFiltros, pblnSegregarPorTipoBackOffice, pintTipoBackOffice)
    Set objItemCaixaGrupoVeicLegal = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objItemCaixaGrupoVeicLegal = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterRelacaoItensCaixaGrupoVeicLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
