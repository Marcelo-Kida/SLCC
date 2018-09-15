VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTProdutoPJTipoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao Cadastro de Produtos PJ x Tipos de Operação.
' É responsável também, pela ligação da camada intermediária, à camada de negócios.

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

' Este método aciona atualizações ao cadastro de produtos PJ x tipos de operação, na camada de negócios.

Public Function Executar(ByVal pstrProdutoTipoOperacao As String) As Boolean

Dim objTProdutoPJTipoOperacao               As A6SubReserva.clsProdutoPJTipoOperacao

On Error GoTo ErrorHandler

    Set objTProdutoPJTipoOperacao = CreateObject("A6SubReserva.clsProdutoPJTipoOperacao")

    objTProdutoPJTipoOperacao.Executar pstrProdutoTipoOperacao

    Set objTProdutoPJTipoOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objTProdutoPJTipoOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
