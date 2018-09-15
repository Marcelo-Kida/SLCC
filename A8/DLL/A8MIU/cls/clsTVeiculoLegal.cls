VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao Cadastro de Veículos Legais.
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

' Este método aciona a alteração do código do grupo de veículo legal, na camada de negócios.

Public Function AlterarCodigoGrupo(ByVal pstrVeiculoLegal As String) As String
    
Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlDomDocument                          As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlDomDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDomDocument.loadXML(pstrVeiculoLegal) Then
        fgErroLoadXML xmlDomDocument, App.EXEName, TypeName(Me), "AlterarCodigoGrupo"
    End If
    
    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    
    AlterarCodigoGrupo = objVeiculoLegal.AlterarCodigoGrupo(xmlDomDocument)
    
    Set objVeiculoLegal = Nothing
    Set xmlDomDocument = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set xmlDomDocument = Nothing
    Set objVeiculoLegal = Nothing
    
    If Not objContext Is Nothing Then
        objContext.DisableCommit
    End If
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarCodigoGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
