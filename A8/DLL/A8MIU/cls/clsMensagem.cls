VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao tratamento de Mensagens.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const COD_OP_ERRO_NEGOCIO_ESPEC     As Long = 3013
Private Const COD_MSG_ERRO_NEGOCIO_ESPEC    As Long = 3014

Private Const COD_TIPO_ERRO_NEGOCIO         As Integer = 1
Private Const COD_TIPO_ERRO_INTERNO         As Integer = 2

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

' Este método aciona a leitura de grupos de usuários associados a mensagens SPB, na camada de negócios,
' e retorna para a camada de interface.

Public Function LerTodosGrupoMensagem(ByVal pblnMensagemTratadaSLCC As Boolean, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
        
    pvntCodErro = 0
    pvntMensagemErro = ""
        
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    LerTodosGrupoMensagem = objMensagem.LerTodosGrupoMensagem(pblnMensagemTratadaSLCC)

    Set objMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosGrupoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura dos dados sobre a conciliação de uma mensagem passada como parâmetro,
' na camada de negócios, e retorna para a camada de interface.
Public Function ObterConciliacaoMensagem(ByVal xmlDocFiltros As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsConciliacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim strNumeroIf                             As String
Dim strDataRegistro                         As String

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    strNumeroIf = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_NumeroCtrlIF/NumeroCtrlIF").Text
    strDataRegistro = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_DataRegistro/DataRegistro").Text
    
    Set objOperacao = CreateObject("A8LQS.clsConciliacao")
    ObterConciliacaoMensagem = objOperacao.ConsultaConciliacao(, strNumeroIf, strDataRegistro)
    
    Set objOperacao = Nothing
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterConciliacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura do histórico das situações de uma mensagem passada como parâmetro,
' na camada de negócios, e retorna para a camada de interface.
Public Function ObterHistoricoMensagem(ByVal strXMLFiltro As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

Dim strNumeroIf                             As String
Dim strDataRegistro                         As String
Dim strNumSequCntrRepe                      As String

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDomFiltro.loadXML strXMLFiltro

    strNumeroIf = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_NumeroCtrlIF/NumeroCtrlIF").Text
    strDataRegistro = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_DataRegistro/DataRegistro").Text
    strNumSequCntrRepe = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_NumeroSequenciaControleRepeticao/NumeroSequenciaControleRepeticao").Text
    
    Set xmlDomFiltro = Nothing

    Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    ObterHistoricoMensagem = objHistoricoMensagem.LerTodos(strNumeroIf, strDataRegistro, strNumSequCntrRepe)
    Set objHistoricoMensagem = Nothing

Exit Function
ErrorHandler:

    Set xmlDomFiltro = Nothing
    Set objHistoricoMensagem = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de serviços associados à mensagem passada como parâmetro,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerTodosServico(ByVal plngSqGrupo As Long, _
                                ByVal pblnMensagemTratadaSLCC As Boolean, _
                                ByRef pvntCodErro As Variant, _
                                ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    LerTodosServico = objMensagem.LerTodosServicos(plngSqGrupo, pblnMensagemTratadaSLCC)

    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosServico Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de eventos associados à mensagem,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerTodosEvento(ByVal plngSqServico As Long, _
                               ByVal pblnMensagemTratadaSLCC As Boolean, _
                               ByRef pvntCodErro As Variant, _
                               ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    LerTodosEvento = objMensagem.LerTodosEventos(plngSqServico, pblnMensagemTratadaSLCC)

    Set objMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosEvento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de mensagens associadas aos parâmetros passados como filtro,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerTodosMensagem(ByVal plngSequenciaGrupo As Long, _
                                 ByVal plngSequenciaServico As Long, _
                                 ByVal plngSequenciaEvento As Long, _
                                 ByVal pblnMensagemTratadaSLCC As Boolean, _
                                 ByVal pstrCdGrupoMesg As String, _
                                 ByRef pvntCodErro As Variant, _
                                 ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
        
    pvntCodErro = 0
    pvntMensagemErro = ""
        
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    LerTodosMensagem = objMensagem.LerTodosMensagem(plngSequenciaGrupo, _
                                                    plngSequenciaServico, _
                                                    plngSequenciaEvento, _
                                                    pblnMensagemTratadaSLCC, _
                                                    pstrCdGrupoMesg)

    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de atributos da mensagem,
' na camada de negócios, e retorna para a camada de interface.

Public Function LerMensagem(ByVal pstrCodigoMensagem As String, _
                            ByVal plngTipoMensagem As Long, _
                            ByVal plngCodigoEmpresa As Long, _
                            ByRef pvntCodErro As Variant, _
                            ByRef pvntMensagemErro As Variant, _
                            ByVal plngTipoOperacao As Long) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    LerMensagem = objMensagem.LerMensagem(pstrCodigoMensagem, plngTipoMensagem, plngCodigoEmpresa, plngTipoOperacao)
    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de domínios de tabelas do PJ e do A8,
' na camada de negócios, e retorna para a camada de interface.
Public Function LerTodosDominioTabela(ByVal plngCodigoEmpresa As Long, _
                                      ByVal pstrNomeTabela As String, _
                                      ByVal pvntCodigo As Variant, _
                                      ByVal pstrDescricao As String, _
                                      ByVal pstrMensagem As String, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagemSLCC

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMensagem = CreateObject("A8LQS.clsMensagemSLCC")
    
    LerTodosDominioTabela = objMensagem.LerTodosDominioTabela(plngCodigoEmpresa, _
                                                              pstrNomeTabela, _
                                                              pstrMensagem, _
                                                              pvntCodigo, _
                                                              pstrDescricao)

    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "FiltroDominioTabela Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTMensagem para o envio de mensagens, na camada de negócios,
' e retorna para a camada de interface.

Public Function EnviarMensagem(ByVal pstrMensagem As String, _
                               ByVal plngTipoMensagem As Integer, _
                               ByVal plngCodigoEmpresa As Long, _
                               ByVal pintTipoBackOffice As Integer, _
                               ByVal plngCodigoLocalLiqu As Long, _
                               ByVal pstrCodigoVeiculoLegal As String, _
                               ByVal pstrSiglaSistema As String, _
                               ByVal pstrAgendamento As String, _
                               ByVal pstrChaveBMC0112 As String, _
                               ByRef pvntCodErro As Variant, _
                               ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8MIU.clsTMensagem
Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = CreateObject("A8MIU.clsTMensagem")
    
    xmlMensagem.loadXML pstrMensagem
    Select Case plngTipoMensagem

        Case enumMensagemEntradaManual.TratadaSLCC
            EnviarMensagem = objMensagem.EnviarMensagem(xmlMensagem, pstrChaveBMC0112)
        Case enumMensagemEntradaManual.NaoTratadaSLCC
            EnviarMensagem = objMensagem.EnviarMensagemNaoTratada(xmlMensagem, _
                                                                  plngCodigoEmpresa, _
                                                                  pintTipoBackOffice, _
                                                                  plngCodigoLocalLiqu, _
                                                                  pstrCodigoVeiculoLegal, _
                                                                  pstrSiglaSistema, _
                                                                  pstrAgendamento)

    End Select
    
    Set xmlMensagem = Nothing
    Set objMensagem = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    
    Set xmlMensagem = Nothing
    Set objMensagem = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "FiltroDominioTabela Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de uma determinada mensagem na camada de negócios, formata o retorno em HTML para a
' exibição na tela, e retorna para a camada de interface.
Public Function ObterMensagemHTML(ByVal xmlDocFiltros As String, _
                                  ByRef pvntCodErro As Variant, _
                                  ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim objDomFiltro                            As MSXML2.DOMDocument40
Dim lngSequencial                           As Long
Dim lngCodigoEmpresa                        As Long
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not objDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    lngSequencial = Val(objDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Sequencial/Sequencial").Text)
    lngCodigoEmpresa = Val(objDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Empresa/Empresa").Text)
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterMensagemHTML = objMensagem.ObterMensagemHTML(lngSequencial)
    
    Set objMensagem = Nothing
    Set objDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set objDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagemHTML Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
Public Function ObterXMLMensagem(ByVal plngSequencial As Long, _
                                 ByRef pvntCodErro As Variant, _
                                 ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
    
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterXMLMensagem = objMensagem.ObterXMLMensagem(plngSequencial)
    Set objMensagem = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objMensagem = Nothing
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
    
End Function
' Este método aciona a leitura de uma mensagem por operação, na camada de negócios, formata o retorno em HTML para a
' exibição na tela, e retorna para a camada de interface.
Public Function ObterMensagemHTMLPorOperacao(ByVal xmlDocFiltros As String, _
                                             ByRef pvntCodErro As Variant, _
                                             ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim objDomFiltro                            As MSXML2.DOMDocument40
Dim lngSequencial                           As Long
Dim lngCodigoEmpresa                        As Long
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not objDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    lngSequencial = Val(objDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Sequencial/Sequencial").Text)
    lngCodigoEmpresa = Val(objDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_Empresa/Empresa").Text)
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    ObterMensagemHTMLPorOperacao = objMensagem.ObterMensagemHTMLPorOperacao(lngSequencial)
    
    Set objMensagem = Nothing
    Set objDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set objDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagemHTML Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a criação de um número controle IF na camada de negócios.
Public Function ObterNumeroControleIF(ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterNumeroControleIF = objMensagem.ObterNumeroControleIF
    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNumeroControleIF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de uma determinada mensagem na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheMensagem(ByVal xmlDocFiltros As String, _
                                     ByRef pvntCodErro As Variant, _
                                     ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    ObterDetalheMensagem = objMensagemSLCC.ObterDetalheMensagem(xmlDomFiltro)
    Set objMensagemSLCC = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objMensagemSLCC = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dados necessários para a conciliação de mensagens, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterNetMensagemConciliacao(ByVal xmlDocFiltros As String, _
                                            ByRef pvntCodErro As Variant, _
                                            ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterNetMensagemConciliacao = objMensagem.ObterNetMensagemConciliacao(xmlDomFiltro)
    Set objMensagem = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetMensagemConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTMensagem para a atualização de dados de agendamento, na camada de negócios,
' e retorna para a camada de interface.

Public Function AlterarAgendamento(ByVal xmlDocFiltros As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As Boolean

Dim objTMensagem                            As A8MIU.clsTMensagem
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strNumeroControleIF                     As String
Dim strDTRegistroMensagemSPB                As String
Dim strHorarioAgendamento                   As String
Dim strDHUltimaAtualizacao                  As String
Dim lngCodigoMensagemXML                    As Long
Dim strCodigoMensagem                       As String
Dim intStatusMensagem                       As enumStatusMensagem

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    For Each objDomNode In xmlDomFiltro.documentElement.childNodes
        'Captura parâmetros
        strNumeroControleIF = objDomNode.selectSingleNode("NumeroControleIF").Text
        strDTRegistroMensagemSPB = objDomNode.selectSingleNode("DTRegistroMensagemSPB").Text
        strHorarioAgendamento = objDomNode.selectSingleNode("HorarioAgendamento").Text
        strDHUltimaAtualizacao = objDomNode.selectSingleNode("DHUltimaAtualizacao").Text
        lngCodigoMensagemXML = Val(objDomNode.selectSingleNode("CodigoMensagemXML").Text)
        strCodigoMensagem = objDomNode.selectSingleNode("CodigoMensagem").Text
        intStatusMensagem = objDomNode.selectSingleNode("StatusMensagem").Text
    Next
    
    'Chama o processo de confirmação
    Set objTMensagem = CreateObject("A8MIU.clsTMensagem")
    Call objTMensagem.AlterarAgendamento(strNumeroControleIF, _
                                         strDTRegistroMensagemSPB, _
                                         strHorarioAgendamento, _
                                         strDHUltimaAtualizacao, _
                                         lngCodigoMensagemXML, _
                                         strCodigoMensagem, _
                                         intStatusMensagem)
    Set objTMensagem = Nothing
        
    Set xmlDomFiltro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTMensagem = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de mensagens agrupadas por status, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterMensagensPorStatus(ByVal xmlDocFiltros As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    ObterMensagensPorStatus = objMensagemSLCC.ObterMensagensPorStatus(xmlDomFiltro)
    Set objMensagemSLCC = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objMensagemSLCC = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de mensagens associadas a uma outra passada como parâmetro, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterMensagensAsssociadas(ByVal strXMLFiltros As String, _
                                          ByRef pvntCodErro As Variant, _
                                          ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim strNumeroControleIF                     As String
Dim strDataRegistroMensagem                 As String
Dim strNumSequCntrRepe                      As String

Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomFiltro.loadXML(strXMLFiltros)
    
    strNumeroControleIF = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_NumeroCtrlIF/NumeroCtrlIF").Text
    strDataRegistroMensagem = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_DataRegistro/DataRegistro").Text
    strNumSequCntrRepe = xmlDomFiltro.selectSingleNode("/Repeat_Filtros/Grupo_NumeroSequenciaControleRepeticao/NumeroSequenciaControleRepeticao").Text
    If Val(strNumSequCntrRepe) = 0 Then strNumSequCntrRepe = "1"
    
    Set xmlDomFiltro = Nothing

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterMensagensAsssociadas = objMensagem.ObterMensagensAssociadas(strNumeroControleIF, strDataRegistroMensagem, strNumSequCntrRepe)
    Set objMensagem = Nothing

Exit Function
ErrorHandler:

    Set objMensagem = Nothing
    Set xmlDomFiltro = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensAssociadas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de grupos de mensagens, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterMensagensPorGrupo(ByVal xmlDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    ObterMensagensPorGrupo = objMensagemSLCC.ObterMensagensPorGrupo(xmlDomFiltro)
    Set objMensagemSLCC = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objMensagemSLCC = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensPorGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura da quantidade conciliada em determinada mensagem, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterQtdConciliadaMensagem(ByVal xmlDocFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim strNumeroControleIF                     As String
Dim strDataRegistro                         As String

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Trim(xmlDocFiltros) <> vbNullString Then
        Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
        If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
            '100 - Documento XML Inválido.
            lngCodigoErroNegocio = 100
            GoTo ErrorHandler
        End If
        
        strNumeroControleIF = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_NumeroCtrlIF/NumeroCtrlIF").Text
        strDataRegistro = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_DataRegistro/DataRegistro").Text
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    ObterQtdConciliadaMensagem = objConciliacao.ObterQtdConciliadaMensagem(strNumeroControleIF, strDataRegistro)
    
    Set objConciliacao = Nothing
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQtdConciliadaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura do código ISPB IF de uma determinada empresa, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterISPBIF(ByVal plngCodigoEmpresa As Long, _
                            ByRef pvntCodErro As Variant, _
                            ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterISPBIF = objMensagem.ObterISPBIF(plngCodigoEmpresa)
    Set objMensagem = Nothing
    
Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterISPBIF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de uma mensagem, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheMensagemCamara(ByVal xmlDocFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    ObterDetalheMensagemCamara = objMensagemSLCC.ObterDetalheMensagemCamara(xmlDomFiltro)
    Set objMensagemSLCC = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objMensagemSLCC = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMensagemCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTMensagem para o reenvio de mensagens, na camada de negócios,
' e retorna para a camada de interface.

Public Function Reenviar(ByVal pstrNumeroControleIF As String, _
                         ByVal pstrDHUltimaAtualizacao As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean

Dim objTMensagem                            As A8MIU.clsTMensagem

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTMensagem = CreateObject("A8MIU.clsTMensagem")
    Reenviar = objTMensagem.Reenviar(pstrNumeroControleIF, pstrDHUltimaAtualizacao)
    Set objTMensagem = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objTMensagem = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de domínios SPB para o nome de TAG informado, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDominioSPB(ByVal pstrNomeTag As String, _
                                ByRef pvntCodErro As Variant, _
                                ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    ObterDominioSPB = objMensagemSLCC.ObterDominioSPB(pstrNomeTag)
    Set objMensagemSLCC = Nothing
    
Exit Function
ErrorHandler:
    Set objMensagemSLCC = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDominioSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dados necessários para a conciliação de mensagens, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterNetLiquidacaoFisica(ByVal xmlDocFiltros As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    ObterNetLiquidacaoFisica = objMensagem.ObterNetLiquidacaoFisica(xmlDomFiltro)
    Set objMensagem = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetLiquidacaoFisica Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsEntradaManual para a gravação do parâmetro de entrada manual, na camada de negócios,
' e retorna para a camada de interface.

Public Sub GravarParametroEntradaManual(ByVal plngCodigoEmpresa As Long, _
                                        ByVal plngTipoOperacao As Long, _
                                        ByVal pstrXMLParm As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant)

Dim objControleAcesso                       As A8MIU.clsControleAcessDado
Dim objEntradaManual                        As A8LQS.clsEntradaManual
Dim lngTipoBackOffice                       As Long
     
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objControleAcesso = CreateObject("A8MIU.clsControleAcessDado")
    lngTipoBackOffice = CLng(objControleAcesso.ObterTipoBackOfficeUsuario(pvntCodErro, pvntMensagemErro))
    
    If pvntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing

    Set objEntradaManual = CreateObject("A8LQS.clsEntradaManual")
    objEntradaManual.GravarParametroEntradaManual lngTipoBackOffice, _
                                                  plngCodigoEmpresa, _
                                                  plngTipoOperacao, _
                                                  pstrXMLParm
        
    Set objEntradaManual = Nothing

Exit Sub
ErrorHandler:
    Set objControleAcesso = Nothing
    Set objEntradaManual = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "GravarParametroEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Sub
                                                                                    
' Este método aciona a leitura de parâmetros de entrada manual, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterParametroEntadaManual(ByVal plngCodigoEmpresa As Long, _
                                           ByVal plngTipoOperacao As Long, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String
                                           
Dim objControleAcesso                       As A8MIU.clsControleAcessDado
Dim objEntradaManual                        As A8LQS.clsEntradaManual
Dim lngTipoBackOffice                       As Long
     
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objControleAcesso = CreateObject("A8MIU.clsControleAcessDado")
    lngTipoBackOffice = CLng(objControleAcesso.ObterTipoBackOfficeUsuario(pvntCodErro, pvntMensagemErro))
    
    If pvntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleAcesso = Nothing

    Set objEntradaManual = CreateObject("A8LQS.clsEntradaManual")
    ObterParametroEntadaManual = objEntradaManual.ObterParametroEntradaManual(lngTipoBackOffice, _
                                                                              plngCodigoEmpresa, _
                                                                              plngTipoOperacao)
    
    Set objEntradaManual = Nothing

Exit Function
ErrorHandler:
    Set objControleAcesso = Nothing
    Set objEntradaManual = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "GravarParametroEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTMensagem para o reprocessamento de mensagens inconsistentes, na camada de negócios,
' e retorna para a camada de interface.
' Monta também um XML com o resultado do processamento em lote, a ser exibido para o usuário.
Public Function ReprocessarMensagemInconcistente(ByVal pxmlGerenciar As String, _
                                                 ByRef pvntCodErro As Variant, _
                                                 ByRef pvntMensagemErro As Variant) As String

Dim objTMensagem                            As A8MIU.clsTMensagem

Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40
Dim xmlLotes                                As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim strNumeroControleIF                     As String
Dim strDTRegistroMensagemSPB                As String
Dim lngNumeControleRepeticao                As Long
Dim strDHUltimaAtualizacao                  As String
Dim strCodigoMensagem                       As String

Dim strxmlErro                              As String
Dim lngCodErro                              As Long
Dim intTipoErro                             As Integer
Dim strErro                                 As String

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    'Monta XML para controle de erros
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Monta XML para controle de resultados
    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")
    
    Set xmlLotes = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlLotes.loadXML(pxmlGerenciar) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    For Each xmlNode In xmlLotes.documentElement.childNodes
        
        strNumeroControleIF = xmlNode.selectSingleNode("NU_CTRL_IF").Text
        strDTRegistroMensagemSPB = xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text
        strDHUltimaAtualizacao = xmlNode.selectSingleNode("DH_ULTI_ATLZ").Text
        lngNumeControleRepeticao = xmlNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text
        strCodigoMensagem = xmlNode.selectSingleNode("CO_MESG").Text
        
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", enumTipoConfirmacao.MENSAGEM, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "NumeroControleIF", strNumeroControleIF, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoMensagem", strCodigoMensagem, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "DTRegistroMensagemSPB", strDTRegistroMensagemSPB, "Repeat_ControleErro")
        
        Set objTMensagem = CreateObject("A8MIU.clsTMensagem")
        strxmlErro = objTMensagem.ReprocessarMensagemInconsistente(strNumeroControleIF, _
                                                                   strDTRegistroMensagemSPB, _
                                                                   lngNumeControleRepeticao, _
                                                                   strDHUltimaAtualizacao)
        Set objTMensagem = Nothing
        
        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)
            
            'Verifica se ocorreu um erro de negócio específico
            If intTipoErro = COD_TIPO_ERRO_NEGOCIO And _
              (lngCodErro = COD_OP_ERRO_NEGOCIO_ESPEC Or lngCodErro = COD_MSG_ERRO_NEGOCIO_ESPEC) Then
                Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 1, "Repeat_ControleErro")
            Else
                Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")
            End If
            
            strErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text
            
            'Incrementa o controle de erros com o STATUS do processo (ERRO)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoErro", intTipoErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", strErro, "Repeat_ControleErro")
        Else
            lngCodErro = 0
            
            'Incrementa o controle de erros com o STATUS do processo (OK)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", "OK", "Repeat_ControleErro")
        End If
    Next
    
    ReprocessarMensagemInconcistente = xmlControleResultado.xml
    
    Set xmlLotes = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objTMensagem = Nothing
    Set xmlLotes = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Confirmar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

