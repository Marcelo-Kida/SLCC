VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracaoMovimentacaoFundos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsMonitoracaoMovimentacaoFundos\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Monitoração da Movimentação dos Fundos\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F16B3FF0105"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Monitoração de Movimentação de Fundos.

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

' Este método aciona a leitura do total de movimento de caixa agrupado por veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterTotalPorVeiculoLegal(ByVal pstrCodigoVeiculoLegal As String, _
                                          ByVal pstrSiglaSistema As String, _
                                          ByRef pvntCodErro As Variant, _
                                          ByRef pvntMensagemErro As Variant) As String
 
Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    If Len(pstrCodigoVeiculoLegal) > 0 Then
       Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
       ObterTotalPorVeiculoLegal = objMovimentoSubReserva.ObterTotalPorVeiculoLegal(pstrCodigoVeiculoLegal, _
                                                                                    pstrSiglaSistema)
    End If

    Set objMovimentoSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTotalPorVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura do total de movimento de caixa agrupado por grupo de veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterTotalPorGrupoVeiculoLegal(ByVal pintCO_GRUP_VEIC_LEGA As Integer, _
                                               ByRef pvntCodErro As Variant, _
                                               ByRef pvntMensagemErro As Variant) As String
 
Dim objMovimentoSubReserva                  As A6SubReserva.clsMovimentoSubReserva

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    If Len(pintCO_GRUP_VEIC_LEGA) > 0 Then
       Set objMovimentoSubReserva = CreateObject("A6SubReserva.clsMovimentoSubReserva")
       ObterTotalPorGrupoVeiculoLegal = objMovimentoSubReserva.ObterTotalPorGrupoVeiculoLegal(pintCO_GRUP_VEIC_LEGA)
    End If

    Set objMovimentoSubReserva = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTotalPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
