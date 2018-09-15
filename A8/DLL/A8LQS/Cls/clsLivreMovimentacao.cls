VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLivreMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Objeto responsável pelo recebimento da Livre Movimentação e controlar o envio dessas informações para os sistemas PJ.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objMQAX200                          As A6A7A8CA.clsMQAX200

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

'Controlar o fluxo do recebimento do regsitro do legado

Public Function GerenciarRecebimento(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                            Optional ByRef strError As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objLegado                               As A8LQS.clsLegado
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    strError = vbNullString

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    Set objLegado = CreateObject("A8LQS.clsLegado")

    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                 xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3120 - Operação original não localizada no SLCC.
        lngCodigoErroNegocio = 3120
        GoTo ErrorHandler
    End If

    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.LiquidacaoOperacaoTermoBMA And _
        xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.LiquidacaoCompromissadaEspecificaVolta Then
        '3121 - Tipo de Operação original inválido para Livre Movimentação.
        lngCodigoErroNegocio = 3121
        GoTo ErrorHandler
    End If

    If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaBackofficePrevia And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Registrada And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.RegistradaAutomatica And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaBackoffice And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.DiscordanciaBackoffice And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaAdmArea And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.DiscordanciaAdmArea And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.PagamentoLib And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Pendencia And _
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Pendencia Then
        '3119 - Status da operação inválido para recebimento do registro de Livre Movimentação.
        lngCodigoErroNegocio = 3119
        GoTo ErrorHandler
    End If

    'Trocar o produto original da operação pelo o que dever ser estornado
    flConverterProduto xmlRemessa, xmlOperacao, True

    strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False


    'Trocar o produto original da operação pelo o que dever ser enviado
    flConverterProduto xmlRemessa, xmlOperacao, False

    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    'Incluir no historico da operação
    'Armazenar a mensagem original
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       enumTipoSolicitacao.LivreMovimentacao, _
                                       xmlRemessa

    'Retorno para o Legado
    fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SIST_PROC", enumStatusOperacao.Liquidada
    fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    strMensagem = objLegado.EnviarRetorno(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set xmlOperacao = Nothing
    Set objLegado = Nothing

    GerenciarRecebimento = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Converte as informacões para envio do Estorno e do Novo Valor

Private Function flConverterProduto(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                    ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    ByVal pblnEstorno As Boolean) As Boolean

    On Error GoTo ErrorHandler

    xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("VA_OPER_ATIV").Text

    If pblnEstorno Then
        If Not xmlOperacao.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = xmlRemessa.documentElement.selectSingleNode("CO_PROD_ESTO").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("TP_CONT") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("TP_CONT").Text = xmlRemessa.documentElement.selectSingleNode("TP_CONT_ESTO").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SEGM") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_SEGM").Text = xmlRemessa.documentElement.selectSingleNode("CO_SEGM_ESTO").Text
        End If
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_EVEN_FINC") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_EVEN_FINC").Text = xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC_ESTO").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("CO_INDX") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_INDX").Text = xmlRemessa.documentElement.selectSingleNode("CO_INDX_ESTO").Text
        End If
    Else
        If Not xmlOperacao.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("TP_CONT") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("TP_CONT").Text = xmlRemessa.documentElement.selectSingleNode("TP_CONT").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SEGM") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_SEGM").Text = xmlRemessa.documentElement.selectSingleNode("CO_SEGM").Text
        End If
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_EVEN_FINC") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_EVEN_FINC").Text = xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC").Text
        End If
    
        If Not xmlOperacao.documentElement.selectSingleNode("CO_INDX") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("CO_INDX").Text = xmlRemessa.documentElement.selectSingleNode("CO_INDX").Text
        End If
    End If

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConverterProduto Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
