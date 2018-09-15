VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito à manutenção e fluxo de Conta Corrente / Contábil.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const COD_ERRO_NEGOCIO_ESPEC        As Long = 3040

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

' Este método instancia a classe transcional clsTContaCorrente para a atualização dos lançamentos de conta-corrente,
' controla o processamento dos lançamentos em lote, monta um XML com o resultado deste processamento,
' e retorna para a camada de interface.
Public Function Gerenciar(ByVal pxmlDocGerenciar As String, _
                          ByRef pvntCodErro As Variant, _
                          ByRef pvntMensagemErro As Variant) As String

Dim objTContaCorrente                       As A8MIU.clsTContaCorrente

Dim xmlDocGerenciar                         As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

Dim vntSequeciaOperacao                     As Variant
Dim intTipoIntegracao                       As Integer
Dim intStatusIntegracao                     As enumStatusIntegracao
Dim strDHUltimaAtualizacao                  As String
Dim strJustificativaCancel                  As String
Dim blnIgnoraGradeHorario                   As Boolean
Dim lngTipoLancamentoIntegracao             As Long
Dim strJustificativa                        As String
Dim strVaLancCC                             As String
Dim strInLancDebCred                        As String
Dim strNetOperacoes                         As String
Dim lngNumSeqLancamento                     As Long

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
    
    Set xmlDocGerenciar = CreateObject("MSXML2.DOMDocument.4.0")

    xmlDocGerenciar.loadXML pxmlDocGerenciar

    For Each xmlDomNode In xmlDocGerenciar.documentElement.childNodes
        vntSequeciaOperacao = xmlDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
        intTipoIntegracao = xmlDomNode.selectSingleNode("TP_LANC_ITGR").Text
        intStatusIntegracao = xmlDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
        strDHUltimaAtualizacao = xmlDomNode.selectSingleNode("DH_ULTI_ATLZ").Text
        
        If Not xmlDomNode.selectSingleNode("TipoLancamentoIntegracao") Is Nothing Then
            lngTipoLancamentoIntegracao = xmlDomNode.selectSingleNode("TipoLancamentoIntegracao").Text
        Else
            lngTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal
        End If
        
        If Not xmlDomNode.selectSingleNode("IgnoraGradeHorario") Is Nothing Then
            blnIgnoraGradeHorario = Val(xmlDomNode.selectSingleNode("IgnoraGradeHorario").Text)
        End If
        
        If Not xmlDomNode.selectSingleNode("TX_JUST_CANC") Is Nothing Then
            strJustificativaCancel = xmlDomNode.selectSingleNode("TX_JUST_CANC").Text
        End If
        
        If Not xmlDomNode.selectSingleNode("TX_JUST_CANC") Is Nothing Then
            strJustificativa = xmlDomNode.selectSingleNode("TX_JUST_CANC").Text
        Else
            strJustificativa = ""
        End If
        
        'Cesar 07/05/2007 - Conta Corrente Corretoras
        If Not xmlDomNode.selectSingleNode("VA_LANC_CC") Is Nothing Then
            strVaLancCC = xmlDomNode.selectSingleNode("VA_LANC_CC").Text
        Else
            strVaLancCC = vbNullString
        End If
        
        If Not xmlDomNode.selectSingleNode("IN_LANC_DEBT_CRED") Is Nothing Then
            strInLancDebCred = xmlDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text
        Else
            strInLancDebCred = vbNullString
        End If
        
        If Not xmlDomNode.selectSingleNode("NetOperacoes") Is Nothing Then
            strNetOperacoes = xmlDomNode.selectSingleNode("NetOperacoes").Text
        Else
            strNetOperacoes = vbNullString
        End If
        
        If Not xmlDomNode.selectSingleNode("NR_SEQU_LANC") Is Nothing Then
            lngNumSeqLancamento = CLng("0" & xmlDomNode.selectSingleNode("NR_SEQU_LANC").Text)
        Else
            lngNumSeqLancamento = 0
        End If
        
        'Incrementa o controle de erros com a OPERAÇÃO/MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", enumTipoConfirmacao.Integracao, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vntSequeciaOperacao, "Repeat_ControleErro")
        
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "NU_SEQU_OPER_ATIV", vntSequeciaOperacao, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoLancamentoIntegracao", intTipoIntegracao, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TP_LANC_ITGR", intTipoIntegracao, "Repeat_ControleErro")
        
        Set objTContaCorrente = CreateObject("A8MIU.clsTContaCorrente")
        
        strxmlErro = objTContaCorrente.Gerenciar(vntSequeciaOperacao, _
                                                 intTipoIntegracao, _
                                                 intStatusIntegracao, _
                                                 strDHUltimaAtualizacao, _
                                                 strJustificativaCancel, _
                                                 blnIgnoraGradeHorario, _
                                                 lngTipoLancamentoIntegracao, _
                                                 strJustificativa, _
                                                 strVaLancCC, _
                                                 strInLancDebCred, _
                                                 strNetOperacoes, _
                                                 lngNumSeqLancamento)
                                                 
        Set objTContaCorrente = Nothing
    
        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)
            
            'Verifica se ocorreu um erro de negócio específico
            If intTipoErro = COD_TIPO_ERRO_NEGOCIO And lngCodErro = COD_ERRO_NEGOCIO_ESPEC Then
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
    
    Gerenciar = xmlControleResultado.xml
    
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    Set xmlDocGerenciar = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objTContaCorrente = Nothing
    Set xmlDocGerenciar = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "Gerenciar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de um lançamento de conta-corrente na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalheLancamento(ByVal pstrFiltroXML As String, _
                                       ByRef pvntCodErro As Variant, _
                                       ByRef pvntMensagemErro As Variant) As String
    
Dim objContaCorrente                        As A8LQS.clsLancamentoCCContabil
    
On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objContaCorrente = CreateObject("A8LQS.clsLancamentoCCContabil")
    ObterDetalheLancamento = objContaCorrente.ObterDetalheLancamento(pstrFiltroXML)
    Set objContaCorrente = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objContaCorrente = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de um lançamento de conta-corrente na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterConsultaLancamento(ByVal pstrFiltroXML As String, _
                                        ByRef pvntCodErro As Variant, _
                                        ByRef pvntMensagemErro As Variant) As String
    
Dim objContaCorrente                        As A8LQS.clsLancamentoCCContabil
    
    On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objContaCorrente = CreateObject("A8LQS.clsLancamentoCCContabil")
    ObterConsultaLancamento = objContaCorrente.ObterConsultaLancamento(pstrFiltroXML)
    Set objContaCorrente = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objContaCorrente = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterConsultaLancamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

' Este método aciona a leitura de lançamentos de conta-corrente agrupados por status na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterLancamentosPorStatus(ByVal pstrFiltroXML As String, _
                                          ByRef pvntCodErro As Variant, _
                                          ByRef pvntMensagemErro As Variant) As String
    
Dim objContaCorrente                        As A8LQS.clsLancamentoCCContabil
    
On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objContaCorrente = CreateObject("A8LQS.clsLancamentoCCContabil")
    ObterLancamentosPorStatus = objContaCorrente.ObterLancamentosPorStatus(pstrFiltroXML)
    Set objContaCorrente = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objContaCorrente = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterLancamentosPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function
