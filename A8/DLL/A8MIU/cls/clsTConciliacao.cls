VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTConciliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo controle transacional nas operações
' de atualização de tabelas, referentes à Conciliação entre Operações e Mensagens.
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

' Este método aciona a conciliação de mensagens e operações, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function Conciliar(ByVal xmlLoteOperacaoMensagem As MSXML2.DOMDocument40) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Conciliar = objConciliacao.Conciliar(xmlLoteOperacaoMensagem)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Conciliar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método aciona a conciliação de mensagens e operações BMA0002, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function ConciliarCamara(ByVal pintTipo As enumTipoConciliacao, _
                                ByVal pintAcao As enumAcaoConciliacao, _
                                ByVal xmlOperacao As MSXML2.DOMDocument40, _
                                ByVal xmlMensagem As MSXML2.DOMDocument40) As String

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objC2                                   As A8LQS.clsInstituicaoISPB     'Teste
Dim xmlProcessamento                        As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objC2 = CreateObject("A8LQS.clsInstituicaoISPB")            'teste
    
    If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesRodaDolar Then
    
        Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlProcessamento.loadXML("<Repeat_Processamento>" & _
                                      "<Grupo_Mensagem>" & xmlMensagem.xml & "</Grupo_Mensagem>" & _
                                      "<Grupo_Operacao>" & xmlOperacao.xml & "</Grupo_Operacao>" & _
                                      "</Repeat_Processamento>")
                                      
        ConciliarCamara = objConciliacao.ConciliarRodaDolarPronto(pintAcao, xmlProcessamento, False)
        Set xmlProcessamento = Nothing
    
    Else
        ConciliarCamara = objConciliacao.ConciliarRegistroOperacao(pintAcao, xmlOperacao, xmlMensagem, False)
    End If
    
    Set objC2 = Nothing
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método aciona a conciliação de mensagens e operações CETIP e BMA0015, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function ConciliarCamaraLote(ByVal pintTipo As enumTipoConciliacao, _
                                    ByVal pintAcao As enumAcaoConciliacao, _
                                    ByVal xmlMensagem As MSXML2.DOMDocument40, _
                                    ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    If pintTipo = enumTipoConciliacao.Bruta Or pintTipo = enumTipoConciliacao.Bilateral Then
        ConciliarCamaraLote = objConciliacao.ConciliarBrutaBilateral(pintTipo, pintAcao, xmlMensagem, , , blnIgnoraGradeHorario)
        
    ElseIf pintTipo = enumTipoConciliacao.MultilateralTitulos Then
        ConciliarCamaraLote = objConciliacao.ConciliarBMA0015(pintTipo, pintAcao, xmlMensagem)
        
    End If
        
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    ConciliarCamaraLote = Err.Description
    
End Function

' Este método aciona a liquidação multilateral CBLC, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarMultilateralCBLC(ByVal pintAcao As enumAcaoConciliacao, _
                                         ByVal xmlProcessamento As MSXML2.IXMLDOMNode, _
                                         ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiquidarMultilateralCBLC = objConciliacao.LiquidarMultilateralCBLC(pintAcao, xmlProcessamento, blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarMultilateralCBLC = Err.Description
    
End Function

' Este método aciona a liquidação de eventos CBLC, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarEventosCBLC(ByVal pintAcao As enumAcaoConciliacao, _
                                    ByVal xmlProcessamento As MSXML2.IXMLDOMNode, _
                                    ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiquidarEventosCBLC = objConciliacao.LiquidarEventosCBLC(pintAcao, xmlProcessamento, blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarEventosCBLC = Err.Description
    
End Function

' Este método aciona a liquidação de eventos CBLC, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarMultilateralBMF(ByVal pintAcao As enumAcaoConciliacao, _
                                        ByVal xmlProcessamento As MSXML2.IXMLDOMNode, _
                                        ByVal xmlRepeticao As MSXML2.DOMDocument40, _
                                        ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiquidarMultilateralBMF = objConciliacao.LiquidarMultilateralBMF(pintAcao, _
                                                                     xmlProcessamento, _
                                                                     xmlRepeticao, _
                                                                     blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarMultilateralBMF = Err.Description
    
End Function

' Este método aciona a liquidação multilateral financeira de mensagens e operações, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarMultilateralFinanceira(ByVal plngPerfilAcesso As enumAcaoConciliacao, _
                                        ByVal pintAcao As enumAcaoConciliacao, _
                                        ByVal xmlConciliacao As MSXML2.DOMDocument40) As String

    Dim objConciliacao                          As A8LQS.clsConciliacao
    Dim objC2                                   As A8LQS.clsInstituicaoISPB     'teste

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objC2 = CreateObject("A8LQS.clsInstituicaoISPB")
    
    LiquidarMultilateralFinanceira = objConciliacao.LiquidarMultilateralFinanceira(plngPerfilAcesso, pintAcao, xmlConciliacao)
    
    Set objC2 = Nothing
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarMultilateralFinanceira = Err.Description

    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarMultilateralFinanceira Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método aciona a liquidação com corretoras de mensagens e operações, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarCorretoras(ByVal pintAcao As enumAcaoConciliacao, _
                                   ByVal xmlProcessamento As MSXML2.IXMLDOMNode, _
                                   ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiquidarCorretoras = objConciliacao.LiquidarCorretoras(pintAcao, xmlProcessamento, blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarCorretoras = Err.Description
    
End Function

' Este método aciona a conciliação de compromissadas genéricas, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function ConciliarCompromissadaGenerica(ByVal pintAcao As enumAcaoConciliacao, _
                                               ByVal xmlProcessamento As MSXML2.IXMLDOMNode, _
                                               ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    ConciliarCompromissadaGenerica = objConciliacao.ConciliarCompromissadaGenerica(pintAcao, xmlProcessamento, blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    ConciliarCompromissadaGenerica = Err.Description
    
End Function

' Este método aciona a conciliação de mensagens e operações CETIP e BMA0015, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiquidarBrutaCBLC(ByVal pintAcao As enumAcaoConciliacao, _
                                  ByVal xmlProcessamento As MSXML2.DOMDocument40, _
                                  ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiquidarBrutaCBLC = objConciliacao.LiquidarBrutaCBLC(pintAcao, xmlProcessamento, , blnIgnoraGradeHorario)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiquidarBrutaCBLC = Err.Description
    
End Function

' Este método aciona a conciliação de mensagens e operações CETIP e BMA0015, na camada de negócios.
' Também é responsável pelo controle transacional desta operação.

Public Function LiberarRodaDolarPronto(ByVal pintAcao As enumAcaoConciliacao, _
                                       ByVal xmlProcessamento As MSXML2.DOMDocument40) As String

Dim objConciliacao                          As A8LQS.clsConciliacao

    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    LiberarRodaDolarPronto = objConciliacao.ConciliarRodaDolarPronto(pintAcao, xmlProcessamento, False)
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    Set objConciliacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.DisableCommit
    End If
    
    'Controla o erro, especificamente neste processo, via o próprio retorno da função,
    'pois a rotina chamadora deverá continuar o seu processamento mesmo se ocorrer erros.
    LiberarRodaDolarPronto = Err.Description
    
End Function

