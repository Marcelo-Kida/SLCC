VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdPartCamaraVeicLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Cadastro de Veículos Legais.

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

' Este método aciona a leitura de todos os veículos legais da base, na camada de negócio,
' e retorna para a camada de interface.

Public Function LerTodos(ByVal pstrCodigoVeiculoLegal As String, _
                         ByVal plngCodigoGrupoVeiculoLegal As Long, _
                         ByVal pblnIndicadorVigencia As String, _
                         ByVal penumTipoBackOffice As Integer, _
                         ByVal pstrSegregaDados As String) As String
                         
Dim objIdPartCamaraVeicLegal                         As A6A7A8.clsIdPartCamaraVeicLegal

On Error GoTo ErrorHandler

    Set objIdPartCamaraVeicLegal = CreateObject("A6A7A8.clsIdPartCamaraVeicLegal")
    LerTodos = objIdPartCamaraVeicLegal.LerTodos()

    Set objIdPartCamaraVeicLegal = Nothing

    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
      'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                         
End Function

' Este método aciona a leitura de um único registro da tabela de veículos legais, na camada de negócio,
' e retorna para a camada de interface.

Public Function Ler(ByVal pstrCodigoVeiculoLegal As String, _
                    ByVal pstrSiglaSistema As String, _
                    ByVal pintCodigoLocalLiquidante As Integer, _
                    ByVal pintTipoIdentificadorPartCamara As Integer, _
                    ByRef pvntCodErro As Variant, _
                    ByRef pvntMensagemErro As Variant) As String
                    
Dim objIdPartCamaraVeicLegal                         As A6A7A8.clsIdPartCamaraVeicLegal

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objIdPartCamaraVeicLegal = CreateObject("A6A7A8.clsIdPartCamaraVeicLegal")
    Ler = objIdPartCamaraVeicLegal.Ler(pstrCodigoVeiculoLegal, pstrSiglaSistema, 0, 0)

    Set objIdPartCamaraVeicLegal = Nothing

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
