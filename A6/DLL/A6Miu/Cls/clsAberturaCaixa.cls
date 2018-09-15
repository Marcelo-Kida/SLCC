VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCaixaSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente         : BJMIU\r\nClasse             : clsAberturaCaixa\r\nData Criação       : 24/07/2003\r\nObjetivo           : Iniciar a classes não transasionais para a tela de Abertura do Caixa\r\n                   :\r\nAnalista           : Carlos Fortes\r\n\r\nProgramador        : \r\nData               : 24/07/2003\r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1998840003"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito às Movimentações de Caixa.

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

' Este método aciona a leitura dos valores para abertura de itens de caixa, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterValoresAbertura(pxmlDocFiltros As String, _
                                     ByRef pvntCodErro As Variant, _
                                     ByRef pvntMensagemErro As Variant) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    ObterValoresAbertura = objCaixaSubReserva.ObterValoresAbertura(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objCaixaSubReserva = Nothing

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresAbertura Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura dos valores para fechamento de itens de caixa, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterValoresFechamento(pxmlDocFiltros As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")

    If Not xmlFiltros.loadXML(pxmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If

    ObterValoresFechamento = objCaixaSubReserva.ObterValoresFechamento(xmlFiltros)

    Set xmlFiltros = Nothing
    Set objCaixaSubReserva = Nothing

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresFechamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método instancia a classe clsTCaixaSubReserva para a abertura de itens de caixa,
' e retorna para a camada de interface.
' É responsável também pelo controle do resultado do processamento em lote.

Public Function AbrirCaixa(ByVal pxmlDocAbrirCaixa As String, _
                           ByRef pvntCodErro As Variant, _
                           ByRef pvntMensagemErro As Variant) As String

Dim objTAberturaCaixa                       As A6MIU.clsTCaixaSubReserva
Dim xmlDocAbrirCaixa                        As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

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
    
    Set xmlDocAbrirCaixa = CreateObject("MSXML2.DOMDocument.4.0")

    xmlDocAbrirCaixa.loadXML pxmlDocAbrirCaixa

    For Each xmlDomNode In xmlDocAbrirCaixa.documentElement.childNodes
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text, "Repeat_ControleErro")
        
        Set objTAberturaCaixa = CreateObject("A6MIU.clsTCaixaSubReserva")
        strxmlErro = objTAberturaCaixa.AbrirCaixa(xmlDomNode)
        Set objTAberturaCaixa = Nothing
    
        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)
            
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")
            
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
    
    AbrirCaixa = xmlControleResultado.xml
    
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    Set xmlDocAbrirCaixa = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTAberturaCaixa = Nothing
    Set xmlDocAbrirCaixa = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    If Not objContext Is Nothing Then
        'objContext.DisableCommit
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "AbrirCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método instancia a classe clsTCaixaSubReserva para o fechamento de itens de caixa,
' e retorna para a camada de interface.
' É responsável também pelo controle do resultado do processamento em lote.

Public Function FecharCaixa(ByVal pxmlDocFecharCaixa As String, _
                            ByRef pvntCodErro As Variant, _
                            ByRef pvntMensagemErro As Variant) As String

Dim objTCaixaSubReserva                     As A6MIU.clsTCaixaSubReserva
Dim xmlDocFecharCaixa                       As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

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
    
    Set xmlDocFecharCaixa = CreateObject("MSXML2.DOMDocument.4.0")

    xmlDocFecharCaixa.loadXML pxmlDocFecharCaixa

    For Each xmlDomNode In xmlDocFecharCaixa.documentElement.childNodes
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", xmlDomNode.selectSingleNode("NO_VEIC_LEGA").Text, "Repeat_ControleErro")
        
        Set objTCaixaSubReserva = CreateObject("A6MIU.clsTCaixaSubReserva")
        strxmlErro = objTCaixaSubReserva.FecharCaixa(xmlDomNode.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                     xmlDomNode.selectSingleNode("SG_SIST").Text)
        Set objTCaixaSubReserva = Nothing
    
        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)
            
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")
            
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
    
    FecharCaixa = xmlControleResultado.xml
    
    Set xmlDocFecharCaixa = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set xmlDocFecharCaixa = Nothing
    Set objTCaixaSubReserva = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "FecharCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura da posição atual de determinado item de caixa, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterPosicaoCaixaSubReserva(ByVal pstrCodigoVeiculoLegal As String, _
                                            ByVal pstrSiglaSistema As String, _
                                            ByVal pstrDataCaixaSubReserva As String, _
                                            ByVal pblnRetornaPosicaoDisponivel As Boolean, _
                                            ByRef pvntCodErro As Variant, _
                                            ByRef pvntMensagemErro As Variant) As String

Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva
Dim xmlPosCaixa                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")
    Set xmlPosCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlPosCaixa.loadXML(objCaixaSubReserva.ObterPosicaoCaixaSubReserva(pstrCodigoVeiculoLegal, _
                                                                            pstrSiglaSistema, _
                                                                            pstrDataCaixaSubReserva, _
                                                                            pblnRetornaPosicaoDisponivel))
    ObterPosicaoCaixaSubReserva = xmlPosCaixa.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set objCaixaSubReserva = Nothing
    Set xmlPosCaixa = Nothing
    
    Exit Function

ErrorHandler:
    Set objCaixaSubReserva = Nothing
    Set xmlPosCaixa = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPosicaoCaixaSubReserva Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura do histórico da posição de itens de caixa, na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterHistoricoPosicaoCaixa(ByVal pxmlFiltros As String, _
                                           ByRef pvntCodErro As Variant, _
                                           ByRef pvntMensagemErro As Variant) As String

On Error GoTo ErrorHandler

Dim objCaixaSubReserva                      As A6SubReserva.clsCaixaSubReserva

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set objCaixaSubReserva = CreateObject("A6SubReserva.clsCaixaSubReserva")
    ObterHistoricoPosicaoCaixa = objCaixaSubReserva.ObterHistoricoPosicaoCaixa(pxmlFiltros)
    Set objCaixaSubReserva = Nothing

    Exit Function
ErrorHandler:
    Set objCaixaSubReserva = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoPosicaoCaixa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

