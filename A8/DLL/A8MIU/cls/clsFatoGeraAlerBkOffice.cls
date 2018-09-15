VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFatoGeraAlerBkOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao relacionamento entre Fatos Geradores de Alertas e o
' Backoffice do Usuário.

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

' Este método aciona a leitura de um registro do cadastro de fatos geradores de alerta por tipo de backoffice,
' na camada de negócios, e retorna para a camada de interface.

Public Function Ler(ByVal pintCodFatoGeradorAlerta As Integer, _
                    ByVal pintTipoBackOffice As Integer, _
                    ByRef pvntCodErro As Variant, _
                    ByRef pvntMensagemErro As Variant) As String

Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    Ler = objFatoGeraAlerBkOffice.Ler(pintCodFatoGeradorAlerta, pintTipoBackOffice)
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objFatoGeraAlerBkOffice = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dos grupos de usuários associados ao fato gerador de alerta,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerGruposUsuariosAlerta(ByVal pintCodFatoGeradorAlerta As Integer, _
                                        ByVal pintTipoBackOffice As Integer, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String

Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    LerGruposUsuariosAlerta = objFatoGeraAlerBkOffice.LerGruposUsuariosAlerta(pintCodFatoGeradorAlerta, _
                                                                              pintTipoBackOffice)
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objFatoGeraAlerBkOffice = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerGruposUsuariosAlerta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dos fatos geradores de alerta a partir de um tipo de backoffice,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerTodos(ByVal pintTipoBackOffice As Integer, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    LerTodos = objFatoGeraAlerBkOffice.LerTodos(pintTipoBackOffice)
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objFatoGeraAlerBkOffice = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe transacional clsTFatoGeraAlerBkOffice para a atualização do cadastro de
' Fatos Geradores de Alerta por Tipo de Backoffice, e retorna para a camada de interface.

Public Function Salvar(ByVal pstrCadastroAlerta As String, _
                       ByRef pvntCodErro As Variant, _
                       ByRef pvntMensagemErro As Variant) As String

Dim objFatoGeraAlerBkOffice                 As A8MIU.clsTFatoGeraAlerBkOffice

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objFatoGeraAlerBkOffice = CreateObject("A8MIU.clsTFatoGeraAlerBkOffice")
    Salvar = objFatoGeraAlerBkOffice.Salvar(pstrCadastroAlerta)
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objFatoGeraAlerBkOffice = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
