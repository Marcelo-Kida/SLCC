VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoPJTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à Manutenção do Cadastro de Produto PJ x Tipo de Operação.

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

' Este método aciona a leitura ao cadastro de produtos PJ x tipos de operação, na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodos(ByVal plngCodigoProduto As Long, _
                         ByVal plngTipoOperacao As Long) As String

Dim objProdutoPJTipoOperacao                As A6SubReserva.clsProdutoPJTipoOperacao

On Error GoTo ErrorHandler

    Set objProdutoPJTipoOperacao = CreateObject("A6SubReserva.clsProdutoPJTipoOperacao")

    LerTodos = objProdutoPJTipoOperacao.LerTodos(plngCodigoProduto, plngTipoOperacao)

    Set objProdutoPJTipoOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objProdutoPJTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a leitura ao cadastro de produtos PJ x tipos de operação, agrupados por status, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterProdutoPJTipoOperacao(ByVal penumStatusProduto As Integer, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objProdutoPJTipoOperacao                As A6SubReserva.clsProdutoPJTipoOperacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objProdutoPJTipoOperacao = CreateObject("A6SubReserva.clsProdutoPJTipoOperacao")

    ObterProdutoPJTipoOperacao = objProdutoPJTipoOperacao.ObterProdutosTipoOperacao(penumStatusProduto)

    Set objProdutoPJTipoOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objProdutoPJTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterProdutoPJTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de produtos PJ agrupados por status, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterProdutoPorStatus(ByVal penumStatusProduto As Integer, _
                                      ByVal plngTipoOperacao As Long, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objProdutoPJTipoOperacao                As A6SubReserva.clsProdutoPJTipoOperacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objProdutoPJTipoOperacao = CreateObject("A6SubReserva.clsProdutoPJTipoOperacao")

    ObterProdutoPorStatus = objProdutoPJTipoOperacao.ObterProdutoPorStatus(penumStatusProduto, plngTipoOperacao)

    Set objProdutoPJTipoOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objProdutoPJTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método instancia a classe clsTProdutoPJTipoOperacao para a atualização do cadastro de
' Produto Pj x Tipo de Operação, e retorna para a camada de interface.

Public Function Executar(ByVal pstrProdutoTipoOperacao As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean

Dim objProdutoPJTipoOperacao                As A6MIU.clsTProdutoPJTipoOperacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objProdutoPJTipoOperacao = CreateObject("A6MIU.clsTProdutoPJTipoOperacao")

    objProdutoPJTipoOperacao.Executar pstrProdutoTipoOperacao

    Set objProdutoPJTipoOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objProdutoPJTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
