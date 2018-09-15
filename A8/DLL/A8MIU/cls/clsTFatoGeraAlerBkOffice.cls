VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTFatoGeraAlerBkOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes ao relacionamento entre Fatos Geradores de Alertas e o Tipo de Backoffice
' do Usuário.
' É responsável também, pela ligação da camada intermediária, à camada de negócios.

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

' Este método aciona a atualização de parametrização de fatos geradores de alerta a usuários, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function Salvar(ByVal pstrCadastroAlerta As String)

Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice
Dim objTransacao                            As A8MIU.clsTransacao

Dim xmlCadastro                             As MSXML2.DOMDocument40
Dim xmlDetalhe                              As MSXML2.DOMDocument40
Dim xmlUsuarios                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlCadastro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetalhe = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlUsuarios = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlCadastro.loadXML(pstrCadastroAlerta)
    Call xmlDetalhe.loadXML(xmlCadastro.selectSingleNode("/Repeat_CadastroAlerta").firstChild.xml)
    Call xmlUsuarios.loadXML(xmlCadastro.selectSingleNode("/Repeat_CadastroAlerta/Repeat_GrupoUsuario").xml)
    
    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    Call objFatoGeraAlerBkOffice.Executar(xmlDetalhe)
    
    If xmlUsuarios.selectNodes("/Repeat_GrupoUsuario/*").length > 0 And xmlCadastro.selectSingleNode("//*/@Operacao").Text <> "Excluir" Then
        Call objFatoGeraAlerBkOffice.SalvarGruposUsuarios(xmlUsuarios, _
                                                          xmlCadastro.selectSingleNode("//*/@Operacao").Text)
    Else
        Call objFatoGeraAlerBkOffice.SalvarGruposUsuarios(xmlDetalhe, _
                                                          xmlCadastro.selectSingleNode("//*/@Operacao").Text)
    End If
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objTransacao = Nothing
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarGrupoUsuarios Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
