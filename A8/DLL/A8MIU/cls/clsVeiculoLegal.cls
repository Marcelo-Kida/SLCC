VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVeiculoLegal"
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
                         
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    LerTodos = objVeiculoLegal.LerTodos(pstrCodigoVeiculoLegal, _
                                        plngCodigoGrupoVeiculoLegal, _
                                        pblnIndicadorVigencia, _
                                        penumTipoBackOffice, _
                                        pstrSegregaDados)

    Set objVeiculoLegal = Nothing

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
                    ByVal pstrSiglaSistema As String) As String
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal

On Error GoTo ErrorHandler

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    Ler = objVeiculoLegal.Ler(pstrCodigoVeiculoLegal, pstrSiglaSistema)

    Set objVeiculoLegal = Nothing

    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
      'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a leitura dos veículos legais, agrupados por grupo de veículo legal, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterRelacaoVeicLegalPorGrupoVeicLegal(ByVal penumTipoBackOffice As Integer) As String
                                              
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim objGrupoVeiculoLegal                    As A6A7A8.clsGrupoVeiculoLegal
Dim strxmlRetorno                           As String
    
On Error GoTo ErrorHandler
    
    Set objGrupoVeiculoLegal = CreateObject("A6A7A8.clsGrupoVeiculoLegal")
    strxmlRetorno = objGrupoVeiculoLegal.LerTodos("S", penumTipoBackOffice, "N")
    
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    strxmlRetorno = strxmlRetorno & objVeiculoLegal.LerTodos(, , "S", penumTipoBackOffice, "N")
    
    If Len(strxmlRetorno) > 0 Then
        ObterRelacaoVeicLegalPorGrupoVeicLegal = "<ObterRelacaoVeicLegalPorGrupoVeicLegal>" & _
                                                 strxmlRetorno & _
                                                 "</ObterRelacaoVeicLegalPorGrupoVeicLegal>"
    Else
        ObterRelacaoVeicLegalPorGrupoVeicLegal = vbNullString
    End If
    
    Set objVeiculoLegal = Nothing
    Set objGrupoVeiculoLegal = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:

    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterRelacaoVeicLegalPorGrupoVeicLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método aciona a leitura de detalhes do veículo legal, na camada de negócios, e retorna para a camada de interface.

Public Function ObterDetalheVeiculoLegal(ByVal pstrFiltro As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String
                                              
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
    
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    ObterDetalheVeiculoLegal = objVeiculoLegal.ObterDetalheVeiculoLegal(pstrFiltro)
    Set objVeiculoLegal = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:

    Set objVeiculoLegal = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterRelacaoVeicLegalPorGrupoVeicLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTVeiculoLegal para a alteração do código do grupo de veículo legal.

Public Function AlterarCodigoGrupo(ByVal pstrVeiculoLegal As String) As String
    
Dim objTVeiculoLegal                        As A8MIU.clsTVeiculoLegal

    On Error GoTo ErrorHandler
    
    Set objTVeiculoLegal = CreateObject("A8MIU.clsTVeiculoLegal")
    AlterarCodigoGrupo = objTVeiculoLegal.AlterarCodigoGrupo(pstrVeiculoLegal)
    Set objTVeiculoLegal = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objTVeiculoLegal = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarCodigoGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
