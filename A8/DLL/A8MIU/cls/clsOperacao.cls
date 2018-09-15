VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao tratamento de Operações de Sistemas Legados.

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

' Este método aciona a leitura das operações agrupadas por tipo de operação, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterOperacoesPorTipoOperacao(ByVal xmlDocFiltros As String, _
                                              ByRef pvntCodErro As Variant, _
                                              ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
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
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterOperacoesPorTipoOperacao = objOperacao.ObterOperacoesPorTipoOperacao(xmlDomFiltro)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorTipoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura das operações agrupadas por status, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterOperacoesPorStatus(ByVal xmlDocFiltros As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
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
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterOperacoesPorStatus = objOperacao.ObterOperacoesPorStatus(xmlDomFiltro)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterOperacoesPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de domínios associados à operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDominiosCompromissoOperacao(ByRef pvntCodErro As Variant, _
                                                 ByRef pvntMensagemErro As Variant) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim strXMLDoc                               As String
Dim xmlSaida                                As MSXML2.DOMDocument40
Dim xmlDados                                As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    Set xmlSaida = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDados = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlSaida, "", "Grupo_Dados", ""
    
    strXMLDoc = objMensagemSLCC.ObterDominioAtributo("TP_CPRO_OPER_ATIV")
    If xmlDados.loadXML(strXMLDoc) Then
        fgAppendAttribute xmlDados, "Repeat_DominioAtributo", "NO_ATRIBUTO", "TP_CPRO_OPER_ATIV"
    End If
    
    fgAppendXML xmlSaida, "Grupo_Dados", xmlDados.xml
    
    strXMLDoc = objMensagemSLCC.ObterDominioAtributo("TP_CPRO_RETN_OPER_ATIV")
    If xmlDados.loadXML(strXMLDoc) Then
        fgAppendAttribute xmlDados, "Repeat_DominioAtributo", "NO_ATRIBUTO", "TP_CPRO_RETN_OPER_ATIV"
    End If
    
    fgAppendXML xmlSaida, "Grupo_Dados", xmlDados.xml
    
    ObterDominiosCompromissoOperacao = xmlSaida.xml
    
    Set objMensagemSLCC = Nothing
    Set xmlSaida = Nothing
    Set xmlDados = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    
    Set objMensagemSLCC = Nothing
    Set xmlSaida = Nothing
    Set xmlDados = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDominiosCompromissoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de detalhes de operações com corretoras, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterDetalheCorretoras(ByVal strXmlDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim vntNumeroSequencia                      As Variant
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    xmlFiltro.loadXML strXmlDocFiltros
    vntNumeroSequencia = xmlFiltro.selectSingleNode("//SequenciaOperacao").Text
    Set xmlFiltro = Nothing
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterDetalheCorretoras = objOperacao.ObterDetalheCorretoras(vntNumeroSequencia)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set xmlFiltro = Nothing
    Set objOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function


' Este método aciona a leitura de detalhes de operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheOperacao(ByVal xmlDocFiltros As String, _
                                     ByRef pvntCodErro As Variant, _
                                     ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
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
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterDetalheOperacao = objOperacao.ObterDetalheOperacao(xmlDomFiltro)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de detalhes de operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDadosConsultaOperacao(ByVal xmlDocFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomFiltro.loadXML(xmlDocFiltros)
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterDadosConsultaOperacao = objOperacao.ObterDadosConsultaOperacao(xmlDomFiltro)
    Set objOperacao = Nothing
    
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set xmlDomFiltro = Nothing
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de dados necessários à conciliação financeira, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterNetOperacaoConciliacaoFinanceira(ByVal xmlDocFiltros As String, _
                                                      ByVal strCampoIdentContraparte As String, _
                                                      ByRef pvntCodErro As Variant, _
                                                      ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
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
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterNetOperacaoConciliacaoFinanceira = objOperacao.ObterNetOperacaoConciliacaoFinanceira(xmlDomFiltro, strCampoIdentContraparte)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetOperacaoConciliacaoFinanceira Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de dados necessários à conciliação física, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterNetOperacaoConciliacaoFisica(ByVal xmlDocFiltros As String, _
                                                  ByRef pvntCodErro As Variant, _
                                                  ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
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
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    ObterNetOperacaoConciliacaoFisica = objOperacao.ObterNetOperacaoConciliacaoFisica(xmlDomFiltro)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetOperacaoConciliacaoFisica Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de mensagens internas, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterMensagensInternas(ByVal xmlDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacaoMensagemInterna
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    vntSequenciaOperacao = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_SequenciaOperacao/SequenciaOperacao").Text
    If vntSequenciaOperacao <> vbNullString Then
        Set objOperacao = CreateObject("A8LQS.clsOperacaoMensagemInterna")
        ObterMensagensInternas = objOperacao.LerTodos(vntSequenciaOperacao)
        Set objOperacao = Nothing
    End If
    
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensInternas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de mensagens associadas às operações, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterMensagensAssociadas(ByVal xmlDocFiltros As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsMensagem
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    vntSequenciaOperacao = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_SequenciaOperacao/SequenciaOperacao").Text
    
    Set objOperacao = CreateObject("A8LQS.clsMensagem")
    ObterMensagensAssociadas = objOperacao.LerMensagensPorOperacao(vntSequenciaOperacao)
    
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensAssociadas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura do histórico das situações das operações, na camada de negócios,
' e retorna para a camada de interface.
Public Function ObterSituacoesOperacao(ByVal xmlDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsHistSituacaoOperacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    vntSequenciaOperacao = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_SequenciaOperacao/SequenciaOperacao").Text
    
    Set objOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    ObterSituacoesOperacao = objOperacao.LerTodos(vntSequenciaOperacao)
    
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSituacoesOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de dados sobre a conciliação de operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterConciliacaoOperacao(ByVal xmlDocFiltros As String, _
                                         ByRef pvntCodErro As Variant, _
                                         ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsConciliacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    vntSequenciaOperacao = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_SequenciaOperacao/SequenciaOperacao").Text
    
    Set objOperacao = CreateObject("A8LQS.clsConciliacao")
    ObterConciliacaoOperacao = objOperacao.ConsultaConciliacao(vntSequenciaOperacao)
    
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterConciliacaoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para o entorno de operações,
' e retorna para a camada de interface.

Public Function Estornar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatus As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDHUltimaAtualizacao As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean
    
Dim objTOperacao                            As A8MIU.clsTOperacao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
    Estornar = objTOperacao.Estornar(pvntSequenciaOperacao, pintStatus, plngNumeroComando, pstrDHUltimaAtualizacao)
    Set objTOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objTOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Estornar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método instancia a classe clsTOperacao para o cancelamento de operações,
' e retorna para a camada de interface.

Public Function Cancelar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pintStatus As Integer, _
                         ByVal plngNumeroComando As Long, _
                         ByVal pstrDHUltimaAtualizacao As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean
    
Dim objTOperacao                            As A8MIU.clsTOperacao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
    Cancelar = objTOperacao.Cancelar(pvntSequenciaOperacao, pintStatus, plngNumeroComando, pstrDHUltimaAtualizacao)
    Set objTOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objTOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Cancelar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de uma única operação, na camada de negócios,
' e retorna para a camada de interface.

Public Function Ler(ByVal pvntSequenciaOperacao As Variant) As String
    
Dim objOperacao                             As A8LQS.clsOperacao
    
    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Ler = objOperacao.Ler(pvntSequenciaOperacao)
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Cancelar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método instancia a classe clsTOperacao para o reenvio de operações,
' e retorna para a camada de interface.

Public Function Reenviar(ByVal pvntSequenciaOperacao As Variant, _
                         ByVal pstrDHUltimaAtualizacao As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean
    
Dim objTOperacao                            As A8MIU.clsTOperacao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
    Reenviar = objTOperacao.Reenviar(pvntSequenciaOperacao, pstrDHUltimaAtualizacao)
    Set objTOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objTOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de dados sobre a conciliação de uma única operação, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterConciliacao(ByVal pvntNumeroSequencia As Variant, _
                                 ByVal pstrNumeroControleIF As String, _
                                 ByVal pstrDataRegistroMensagemSPB As String, _
                                 ByVal pstrDataJustificativa As String, _
                                 ByVal plngTipoOperacao As Long, _
                                 ByVal plngEmpresa As Long, _
                                 ByVal plngCodigoSituProc As Long, _
                                 ByRef pvntCodErro As Variant, _
                                 ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsConciliacao

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objOperacao = CreateObject("A8LQS.clsConciliacao")
    ObterConciliacao = objOperacao.ConsultaConciliacao(pvntNumeroSequencia, _
                                                       pstrNumeroControleIF, _
                                                       pstrDataRegistroMensagemSPB, _
                                                       pstrDataJustificativa, _
                                                       plngTipoOperacao, _
                                                       plngEmpresa, _
                                                       plngCodigoSituProc)
    Set objOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para a liquidação via contingência de operações,
' e retorna para a camada de interface.

Public Function LiquidarViaContingencia(ByVal xmlLoteOperacao As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String

Dim objTOperacao                            As A8MIU.clsTOperacao

Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40
Dim xmlDocument                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim intTipoConfirmacao                      As Integer
Dim vntSequenciaOperacao                    As Variant
Dim intStatus                               As Integer  '<-- Conteúdo: enumStatusOperacao ou enumStatusMensagem
Dim strDHUltimaAtualizacao                  As String

Dim strxmlErro                              As String
Dim intTipoErro                             As Integer
Dim strErro                                 As String
Dim lngCodErro                              As Long

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""

    'Monta XML para controle de erros
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Monta XML para controle de resultados
    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")
    
    Set xmlDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDocument.loadXML(xmlLoteOperacao) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    For Each objDomNode In xmlDocument.documentElement.childNodes
        'Captura parâmetros
        intTipoConfirmacao = objDomNode.selectSingleNode("TipoConfirmacao").Text
        
        If Not objDomNode.selectSingleNode("Operacao") Is Nothing Then
            vntSequenciaOperacao = objDomNode.selectSingleNode("Operacao").Text
        End If

        intStatus = objDomNode.selectSingleNode("Status").Text
        strDHUltimaAtualizacao = objDomNode.selectSingleNode("DHUltimaAtualizacao").Text
        
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", intTipoConfirmacao, "Repeat_ControleErro")
        
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vntSequenciaOperacao, "Repeat_ControleErro")
        
        Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
        strxmlErro = objTOperacao.LiquidarViaContingencia(vntSequenciaOperacao, intStatus, strDHUltimaAtualizacao)
        Set objTOperacao = Nothing
        
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
    
    LiquidarViaContingencia = xmlControleResultado.xml
    
    Set xmlDocument = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTOperacao = Nothing
    Set xmlDocument = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para a baixa via contingência de operações,
' e retorna para a camada de interface.

Public Function BaixarViaContingencia(ByVal xmlLoteOperacao As String, _
                                      ByRef pvntCodErro As Variant, _
                                      ByRef pvntMensagemErro As Variant) As String

Dim objTOperacao                            As A8MIU.clsTOperacao

Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40
Dim xmlDocument                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim intTipoConfirmacao                      As Integer
Dim vntSequenciaOperacao                    As Variant
Dim intStatus                               As Integer  '<-- Conteúdo: enumStatusOperacao ou enumStatusMensagem
Dim strDHUltimaAtualizacao                  As String

Dim strxmlErro                              As String
Dim intTipoErro                             As Integer
Dim strErro                                 As String
Dim lngCodErro                              As Long

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    'Monta XML para controle de erros
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Monta XML para controle de resultados
    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")
    
    Set xmlDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDocument.loadXML(xmlLoteOperacao) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    For Each objDomNode In xmlDocument.documentElement.childNodes
        'Captura parâmetros
        intTipoConfirmacao = objDomNode.selectSingleNode("TipoConfirmacao").Text
        
        If Not objDomNode.selectSingleNode("Operacao") Is Nothing Then
            vntSequenciaOperacao = objDomNode.selectSingleNode("Operacao").Text
        End If

        intStatus = objDomNode.selectSingleNode("Status").Text
        strDHUltimaAtualizacao = objDomNode.selectSingleNode("DHUltimaAtualizacao").Text
        
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", intTipoConfirmacao, "Repeat_ControleErro")
        
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vntSequenciaOperacao, "Repeat_ControleErro")
        
        Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
        strxmlErro = objTOperacao.BaixarViaContingencia(vntSequenciaOperacao, intStatus, strDHUltimaAtualizacao)
        Set objTOperacao = Nothing
        
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
    
    BaixarViaContingencia = xmlControleResultado.xml
    
    Set xmlDocument = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTOperacao = Nothing
    Set xmlDocument = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para a alteração de tipo de compromisso de operações,
' e retorna para a camada de interface.

Public Function AlterarTipoCompromisso(ByVal strDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As Boolean

Dim objOperacao                             As A8MIU.clsTOperacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

Dim lngNumeroSequencia                      As Long
Dim strTipoCompromisso                      As String
Dim strTipoCompromissoRet                   As String
Dim strDataUltimaAtualizacao                As String
Dim strAcaoAnterior                         As String
Dim intStatusOperacao                       As Long

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(strDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    With xmlDomFiltro.documentElement
        lngNumeroSequencia = CLng(.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
        strTipoCompromisso = .selectSingleNode("TP_CPRO_OPER_ATIV").Text
        strTipoCompromissoRet = .selectSingleNode("TP_CPRO_RETN_OPER_ATIV").Text
        strDataUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
        strAcaoAnterior = .selectSingleNode("TX_CNTD_ANTE_ACAO").Text
        intStatusOperacao = Int(Val(.selectSingleNode("CO_ULTI_SITU_PROC").Text))
        
    End With
    Set xmlDomFiltro = Nothing
    
    Set objOperacao = CreateObject("A8MIU.clsTOperacao")
    
    AlterarTipoCompromisso = objOperacao.AlterarTipoCompromisso(lngNumeroSequencia, _
                                                                strDataUltimaAtualizacao, _
                                                                strTipoCompromisso, _
                                                                strTipoCompromissoRet, _
                                                                intStatusOperacao, _
                                                                strAcaoAnterior)
    
    Call objOperacao.Confirmar(lngNumeroSequencia, intStatusOperacao, strDataUltimaAtualizacao, "")
                                 
    
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objOperacao = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarTipoCompromisso Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para a alteração de agendamento de operações,
' e retorna para a camada de interface.

Public Function AlterarAgendamento(ByVal xmlDocFiltros As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As Boolean

Dim objTOperacao                            As A8MIU.clsTOperacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim vntSequenciaOperacao                    As Variant
Dim strHorarioAgendamento                   As String
Dim strDHUltimaAtualizacao                  As String
Dim lngLocalLiquidacao                      As Long
Dim strCodigoMensagem                       As String
Dim intStatusOperacao                       As enumStatusOperacao

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
        vntSequenciaOperacao = objDomNode.selectSingleNode("Operacao").Text
        strHorarioAgendamento = objDomNode.selectSingleNode("HorarioAgendamento").Text
        strDHUltimaAtualizacao = objDomNode.selectSingleNode("DHUltimaAtualizacao").Text
        lngLocalLiquidacao = Val(objDomNode.selectSingleNode("LocalLiquidacao").Text)
        strCodigoMensagem = objDomNode.selectSingleNode("CodigoMensagem").Text
        intStatusOperacao = objDomNode.selectSingleNode("StatusOperacao").Text
    Next
    
    Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
    Call objTOperacao.AlterarAgendamento(vntSequenciaOperacao, _
                                         strHorarioAgendamento, _
                                         strDHUltimaAtualizacao, _
                                         lngLocalLiquidacao, _
                                         strCodigoMensagem, _
                                         intStatusOperacao)
    Set objTOperacao = Nothing
        
    Set xmlDomFiltro = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTOperacao = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura da quantidade conciliada de operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterQtdConciliadaOperacao(ByVal xmlDocFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant

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
        
        vntSequenciaOperacao = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_SequenciaOperacao/SequenciaOperacao").Text
    Else
        vntSequenciaOperacao = 0
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    ObterQtdConciliadaOperacao = objConciliacao.ObterQtdConciliadaOperacao(vntSequenciaOperacao)
    
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQtdConciliadaOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de valores agrupados por grupo de veículo legal de operações, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterValoresPorGrupoVeiculoLegal(ByVal xmlDocFiltros As String, _
                                                 ByRef pvntCodErro As Variant, _
                                                 ByRef pvntMensagemErro As Variant) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlDomFiltro                            As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    ObterValoresPorGrupoVeiculoLegal = objOperacao.ObterValoresPorGrupoVeiculoLegal(xmlDomFiltro)
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
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresPorGrupoVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTOperacao para o ajuste de valor de operações,
' e retorna para a camada de interface.

Public Function AjustarValor(ByVal pvntSequenciaOperacao As Variant, _
                             ByVal pstrDataUltimaAtualizacao As String, _
                             ByVal pvntValorNovo As Variant, _
                             ByVal pstrJustificativa As String, _
                             ByRef pvntCodErro As Variant, _
                             ByRef pvntMensagemErro As Variant) As Boolean

Dim objTOperacao       As A8MIU.clsTOperacao
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objTOperacao = CreateObject("A8MIU.clsTOperacao")
    AjustarValor = objTOperacao.AjustarValor(pvntSequenciaOperacao, _
                                             pstrDataUltimaAtualizacao, _
                                             pvntValorNovo, _
                                             pstrJustificativa)
    
    Set objTOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objTOperacao = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AjustarValor Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

'Obter Intervalo de parada do verifica Server
Public Function ObterIntervaloVerificaServer(ByRef pvntCodErro As Variant, _
                                             ByRef pvntMensagemErro As Variant) As String
    
Dim objAlerta                        As A8LQS.clsAlerta
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    ObterIntervaloVerificaServer = objAlerta.ObterIntervaloVerificaServer
    Set objAlerta = Nothing

Exit Function
ErrorHandler:
    
    Set objAlerta = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterIntervaloVerificaServer Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

