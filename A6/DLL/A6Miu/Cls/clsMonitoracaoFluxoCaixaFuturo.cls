VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracaoFluxoCaixaFuturo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsMonitoracaoFluxoCaixaFuturo\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Monitoração do Fluxo de Caixa Futuro\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : \r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F182FE50382"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Monitoração de Fluxo de Caixa Futuro.

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

' Este método aciona a leitura de detalhes do movimento de caixa futuro, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheMovimento(ByVal pstrCodVeiculoLegal As String, _
                                      ByVal pstrSiglaSistema As String, _
                                      ByVal pstrItemCaixa As String, _
                                      ByVal pdatMovimento As String, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String
    
Dim objCaixaFuturo                          As A6SubReserva.clsCaixaFuturo
Dim datDataMovimento                        As Date

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Trim(pdatMovimento) <> "" Then
       datDataMovimento = CDate(Format(pdatMovimento, "DD/MM/YYYY"))
    End If

    Set objCaixaFuturo = CreateObject("A6SubReserva.clsCaixaFuturo")
    ObterDetalheMovimento = objCaixaFuturo.ObterDetalheMovimento(pstrCodVeiculoLegal, _
                                                                 pstrSiglaSistema, _
                                                                 pstrItemCaixa, _
                                                                 datDataMovimento)
    Set objCaixaFuturo = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objCaixaFuturo = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMovimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura do movimento de caixa futuro agrupado por grupo de veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterMovimentoPorGrupoVeiculoLegal(ByVal pintCodGrupoVeicLegal As Integer, _
                                                   ByVal pdatBaseMovimento As String, _
                                                   ByVal pintQuantDiasMov As Integer, _
                                                   ByVal pintPaginacao As Integer, _
                                                   ByRef pvntCodErro As Variant, _
                                                   ByRef pvntMensagemErro As Variant, _
                                          Optional ByVal plngCodEmpresa As Long, _
                                          Optional ByVal pstrCodVeiculoLegal As String) As String
    
Dim objCaixaFuturo                          As A6SubReserva.clsCaixaFuturo
Dim datDataBaseMovimento                    As Date

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Trim(pdatBaseMovimento) <> "" Then
       datDataBaseMovimento = CDate(Format(pdatBaseMovimento, "DD/MM/YYYY"))
    End If

    Set objCaixaFuturo = CreateObject("A6SubReserva.clsCaixaFuturo")
    ObterMovimentoPorGrupoVeiculoLegal = objCaixaFuturo.ObterMovimentoPorGrupoVeiculoLegal(pintCodGrupoVeicLegal, _
                                                                                           datDataBaseMovimento, _
                                                                                           pintQuantDiasMov, _
                                                                                           pintPaginacao, _
                                                                                           plngCodEmpresa, _
                                                                                           pstrCodVeiculoLegal)
    Set objCaixaFuturo = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objCaixaFuturo = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentada devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMovimentoPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura do movimento de caixa futuro agrupado por veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterMovimentoPorVeiculoLegal(ByVal pstrCodVeiculoLegal As String, _
                                              ByVal pstrSiglaSistema As String, _
                                              ByVal pintNivelAgrupItemCaixa As Integer, _
                                              ByVal pdatD0 As String, _
                                              ByVal pdatBaseMovimento As String, _
                                              ByVal pintQuantDiasMov As Integer, _
                                              ByVal pintPaginacao As Integer, _
                                              ByRef pvntCodErro As Variant, _
                                              ByRef pvntMensagemErro As Variant) As String
    
Dim objCaixaFuturo                          As A6SubReserva.clsCaixaFuturo
Dim datDataD0                               As Date
Dim datDataBaseMovimento                    As Date

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Trim(pdatD0) <> "" Then
       datDataD0 = CDate(Format(pdatD0, "DD/MM/YYYY"))
    End If
    If Trim(pdatBaseMovimento) <> "" Then
       datDataBaseMovimento = CDate(Format(pdatBaseMovimento, "DD/MM/YYYY"))
    End If

    Set objCaixaFuturo = CreateObject("A6SubReserva.clsCaixaFuturo")
    ObterMovimentoPorVeiculoLegal = objCaixaFuturo.ObterMovimentoPorVeiculoLegal(pstrCodVeiculoLegal, _
                                                                                 pstrSiglaSistema, _
                                                                                 pintNivelAgrupItemCaixa, _
                                                                                 datDataD0, _
                                                                                 datDataBaseMovimento, _
                                                                                 pintQuantDiasMov, _
                                                                                 pintPaginacao)
    Set objCaixaFuturo = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objCaixaFuturo = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMovimentoPorVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
