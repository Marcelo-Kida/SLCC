VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : clsOperacaoDefinitriva\r\nData Criação   : \r\nObjetivo       : Gerenciar e Controlar os fluxos operacionais das operações definitivas.\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F2946630040"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável pelo recebimento dos Boletos e Operações dos sistemas
'' legados e controlar o envio dessas informações para os sistemas PJ, A6 e
'' Mensageria.

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

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de complementação de uma operação
'-------------------------------------------------------------------------------------------------

Private Function flComplementarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                               Optional ByRef pstrErros As String) As Boolean

Dim strOperacao                             As String
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim blnComplementar                         As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Se for pagamento de redesconto tem de obter o código de veículo legal da mensagem original
    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Or _
        xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.ConversaoRedesconto Then
        objOperacao.ComplOperacaoComRDCOriginal xmlRemessa
    End If

    'Verifica se a Entrada é Manual
    If xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        'O Codigo da Operação é unico por Sistema
        
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemBUS.PagamentoRedesconto Then
            strOperacao = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
        Else
            strOperacao = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                               , , , , , , , , , , _
                                               xmlRemessa.documentElement.selectSingleNode("CO_IDEF_LAST").Text)
        End If
    Else
        strOperacao = vbNullString
    End If

    'Verifica se Inclui ou Complementa a Operação
    If Trim(strOperacao) <> vbNullString Then
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        xmlOperacao.loadXML strOperacao

        'Aqui testar se é compromissada incluir .... - Carlos
        If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.CompromissadaIda Or _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.Redesconto Or _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Or _
           xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Then
            
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Inicial Then
                blnComplementar = True
            Else
                blnComplementar = False
            End If
        
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroLiquidacaoMultilateralCBLC Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroLiquidacaoBrutaCBLC Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroLiquidacaoEventoCBLC Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.TransferenciasBMC Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EnvioTEDClientes Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EnvioPagDespesas Then
                
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Then
                blnComplementar = True
            Else
                'Somente operações com situação EM SER podem ser complementadas
                lngCodigoErroNegocio = 3123
                GoTo ErrorHandler
            End If
        
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.LiquidacaoMultilateralBMC Then
                
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.AConciliar Then
                blnComplementar = True
            Else
                'Somente operações com situação A CONCILIAR podem ser complementadas
                lngCodigoErroNegocio = 4236
                GoTo ErrorHandler
            End If
        
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Or _
               CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Inicial Then
                blnComplementar = True
            Else
                'Somente operações com situação EM SER podem ser complementadas
                lngCodigoErroNegocio = 3123
                GoTo ErrorHandler
            End If
        
        'KIDA - CCR
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EmissaoOperacaoCCR Or _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.NegociacaoOperacaoCCR Then
            
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Or _
               CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Inicial Or _
               CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.PendenteDeAceite Or _
               CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.PendenteDeRegistro Then
                blnComplementar = True
            End If
        
        Else
            
            'Se a operação for encontrada e não estiver com status inicial, rejeitar
            If CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.Inicial Then
                'Somente operações com situação inicial podem ser complementadas
                lngCodigoErroNegocio = 3034
                GoTo ErrorHandler
            End If
            blnComplementar = True
        
        End If

        If blnComplementar Then
            objOperacao.ComplementarOperacao xmlRemessa, _
                                             xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text, _
                                             CLng(xmlOperacao.documentElement.selectSingleNode("Grupo_Operacao/TP_OPER").Text)
        Else
            objOperacao.IncluirOperacao xmlRemessa, enumStatusOperacao.EmSer
        End If
    Else
        If xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then

            'Carlos 10/08/2004 Inicio - ******************************************************************************
            'Quando tipo de solicitacao igual a 2(complementacao) executa esta rotina para aguardar a inclusao do tipo 1
            'para estes layout so vem tipo solicitacao 2
            If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RetornoCompromissada And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.VinculoDesvinculoTransf And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.TransferenciaLDL_BMA And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.LiquidacaoEventosBMA And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.EspecificacaoOperacoesBMA And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.LiquidacaoFisicaOperacaoBMA And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.EventoJurosCETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroLiquidacaoBrutaCBLC And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroLiquidacaoEventoCBLC And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.PgtoRedesconto And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroOperacaoesCETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.LancamentoPU_CETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.AlteracaoDadosContaCorrente And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroLiquidacaoMultilateralBMF And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.TransferenciasBMC And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.LiquidacaoMultilateralBMC And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.RegistroOperacoesRodaDolar And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.EnvioTEDClientes And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.OperacoesComCorretorasCETIP And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.EnvioPagDespesas And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.LancamentoContaCorrenteBG And _
               xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.ConsultaOperacaoCCR And xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text <> enumTipoMensagemLQS.ConsultaLimitesImportacaoCCR Then
                
                'KIDA - CCR
                
                If xmlRemessa.documentElement.selectSingleNode("NU_PROC") Is Nothing Then
                    'Mensagem de complementação sem inclusão
                    lngCodigoErroNegocio = 3110
                    GoTo ErrorHandler
                Else
                    If CLng(xmlRemessa.documentElement.selectSingleNode("NU_PROC").Text) < 3 Then
                        'Mensagem de complementação sem inclusão
                        lngCodigoErroNegocio = 3110
                        GoTo ErrorHandler
                    End If
                End If
            End If
            'Carlos 10/08/2004 - Fim ******************************************************************************

            flProcessarBoleto xmlRemessa, enumStatusOperacao.EmSer
        Else
            flProcessarBoleto xmlRemessa, enumStatusOperacao.ManualEmSer
        End If
    End If

    'Verificar os fluxos alternativos
    If Not flFluxosAlternativos(xmlRemessa) Then
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacoesRodaDolar Then
        Exit Function
    End If
    
    If xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        flGerenciarChamadas xmlRemessa, "RecebimentoOperacao"
    Else
        If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
            If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.NegociacaoOperacaoCCR _
            Or xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacaoInterbancaria _
            Or xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacaoArbitragem _
            Or xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.IFInformaLiquidacaoInterbancaria Then
                flGerenciarChamadas xmlRemessa, "RecebimentoOperacao"
            End If
        End If
    End If

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flComplementarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de processamento de um boleto

Private Function flProcessarBoleto(ByRef xmlBoleto As MSXML2.DOMDocument40, _
                                   ByRef plngStatusOperacao As Long) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim strMensagem                             As String
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim vntSequenciaOperacao                    As Variant
Dim strXML                                  As String
Dim lngTipoOperacao                         As Long
Dim lngTipoLiquidacao                       As Long
Dim blnAlteracaoEvento                      As Boolean

    On Error GoTo ErrorHandler

    If plngStatusOperacao = enumStatusOperacao.Inativa Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Call objOperacao.IncluirOperacao(xmlBoleto, plngStatusOperacao)
        Set objOperacao = Nothing
        Exit Function
    End If
    
    vntSequenciaOperacao = 0
    lngTipoOperacao = 0
    blnAlteracaoEvento = False
    If Not flTermoEventosDespesas(xmlBoleto, _
                                  plngStatusOperacao, _
                                  vntSequenciaOperacao, _
                                  lngTipoOperacao) Then
        Exit Function
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Se for pagamento de redesconto tem de obter o código de veículo legal da mensagem original
    If xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Or _
       xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.ConversaoRedesconto Then
        objOperacao.ComplOperacaoComRDCOriginal xmlBoleto
    End If

    If vntSequenciaOperacao = 0 Then
        objOperacao.IncluirOperacao xmlBoleto, _
                                    plngStatusOperacao
    Else
        objOperacao.ComplementarOperacao xmlBoleto, _
                                         vntSequenciaOperacao, _
                                         lngTipoOperacao
        blnAlteracaoEvento = True
    End If
    Set objOperacao = Nothing

    If xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroOperacoesRodaDolar Then
        Call flConciliarRodaDolar(xmlBoleto)
        Exit Function
    End If
    
    'Por ser entrada manual não envia a previsão neste momento, é enviado na confirmação
    If Not xmlBoleto.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        
        If Not xmlBoleto.selectSingleNode("//TP_OPER") Is Nothing Then
            If xmlBoleto.selectSingleNode("//TP_OPER").Text <> enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then
                Exit Function
            End If
        End If
        
    End If

    'Carlos - Gerenciamento de Envio de Mensagens
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    'If Not xmlBoleto.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '    lngTipoLiquidacao = CLng(xmlBoleto.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    'End If
    strXML = LerTodos("RecebimentoBoleto", _
                      CLng(xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text), _
                      enumIndicadorSimNao.Nao, , _
                      lngTipoLiquidacao)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    'Alterado 04/03/2003 para envio de previsão para o PJ e A6 somente na inclusão de boleto - Carlos
    If xmlBoleto.documentElement.selectSingleNode("TP_SOLI").Text = enumTipoSolicitacao.Inclusao Or _
       xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.VinculoDesvinculoTransferencia Or _
       xmlBoleto.documentElement.selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.SELIC Then
        'Se a operação pagamento de redesconto e já teve uma operação de conversão de redesconto que esteja com o
        'status Liquidada não envia PJ - Carlos 06/04/2004
        If xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Then
            'Testa se a previsão do PJ já enviada  - Carlos 06/04/2004
            If CLng(xmlBoleto.documentElement.selectSingleNode("PREV_PJ_ENV").Text) = enumIndicadorSimNao.Nao Then
                EnviarPrevisao xmlBoleto, xmlControleProcessoOperacao
            End If
        Else
            If xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.EventosSelic Or _
               xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.DespesasSelic Then
                If blnAlteracaoEvento = False Then
                    EnviarPrevisao xmlBoleto, xmlControleProcessoOperacao
                End If
            'KIDA - CBLC
            ElseIf xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.RegistroLiquidacaoBrutaCBLC Then
                
                If fgDtXML_To_Date(xmlBoleto.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    EnviarPrevisao xmlBoleto, xmlControleProcessoOperacao
                End If
            Else
                EnviarPrevisao xmlBoleto, xmlControleProcessoOperacao
            End If
        End If
    ElseIf xmlBoleto.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto And _
        xmlBoleto.documentElement.selectSingleNode("TP_SOLI").Text = enumTipoSolicitacao.Inclusao Then
        'Testa se a previsão do PJ já enviada  - Carlos 06/04/2004
        If CLng(xmlBoleto.documentElement.selectSingleNode("PREV_PJ_ENV").Text) = enumIndicadorSimNao.Nao Then
            EnviarPrevisao xmlBoleto, xmlControleProcessoOperacao
        End If
    
    End If

    'Nos casos de Termo na Data Liquidacao, Despesas e Eventos Selic virá inclusão e não complementação
    If (plngStatusOperacao = enumStatusOperacao.EmSer Or _
        plngStatusOperacao = enumStatusOperacao.AConciliar) And _
       (xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosAmortização Or _
        xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJuros Or _
        xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosResgate) Then
        
        flGerenciarChamadas xmlBoleto, "RecebimentoOperacao"
    
    ElseIf (plngStatusOperacao = enumStatusOperacao.EmSer Or _
            plngStatusOperacao = enumStatusOperacao.AConciliar) And _
            xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DespesaSelic Then
        
        flGerenciarChamadas xmlBoleto, "RecebimentoOperacao"
    
    ElseIf plngStatusOperacao = enumStatusOperacao.AConciliar And _
            (xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
            xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta) Then
        
        flGerenciarChamadas xmlBoleto, "RecebimentoOperacao"
    
    ElseIf xmlBoleto.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DespesasCETIP Then
        
        flGerenciarChamadas xmlBoleto, "RecebimentoOperacao"
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessarBoleto Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica se a funcionalidade de Conciliação está automático ou não no workflow

Private Function flConciliarRodaDolar(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objContingencia                         As A8LQS.clsContingencia
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objAlerta                               As A8LQS.clsAlerta
Dim lngJustificativa                        As enumJustificativa

    On Error GoTo ErrorHandler

    flConciliarRodaDolar = False

    Set objContingencia = CreateObject("A8LQS.clsContingencia")
    Set objAlerta = CreateObject("A8LQS.clsAlerta")
    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    
    If objContingencia.VerificarSistContingencia(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) Then
        
        Call objHistSituacaoOperacao.AlterarJustificativa(xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                          enumStatusOperacao.AConciliar, _
                                                          enumJustificativa.SistemaEmContingencia)
        GoTo ExitSub
    
    Else
        
        If objConciliacao.ConciliarRodaDolarPronto(enumAcaoConciliacao.BOConcordar, xmlRemessa, True) <> vbNullString Then
            Call objAlerta.GerarAlerta(xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoConciliadaAutomatica)
            GoTo ExitSub
        End If
    
    End If
    
    flConciliarRodaDolar = True

ExitSub:
    Set objContingencia = Nothing
    Set objAlerta = Nothing
    Set objHistSituacaoOperacao = Nothing
    Set objConciliacao = Nothing

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarRodaDolar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de previsão para os sistemas A6 e PJ
'-------------------------------------------------------------------------------------------------

Public Function EnviarPrevisao(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                               ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.SELIC Then
        EnviarPrevisaoBMACetip xmlOperacao, xmlControleProcessoOperacao
        Exit Function
    End If

    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        lngIndDebtCredOrig = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    End If
    
    If Not xmlOperacao.selectSingleNode("//CO_PROD") Is Nothing Then
        lngCodProdOrig = xmlOperacao.selectSingleNode("//CO_PROD").Text
    End If

    blnCliente1 = fgVerificarCliente1(xmlOperacao)

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoInformalCompromissada Then
                
                strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            
            ElseIf xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.Termo Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
                    
                    If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
                    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    
                    If blnCliente1 Then
                        fgConverterProduto1 xmlOperacao, True
                        strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        'debito/Credito e Produto Original
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    End If
                End If
            Else
                If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
                strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                'debito/Credito e Produto Original
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras Then
                    fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Sim
                    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Nao
                End If
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                        strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, True)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        'debito/Credito e Produto Original
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    End If
                End If

                If blnCliente1 Then
                    fgConverterProduto1 xmlOperacao
                    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    
                    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                        If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                            fgConverterProduto1 xmlOperacao, True, True
                            strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, True)
                            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                            'debito/Credito e Produto Original
                            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                        End If
                    End If
                End If

                If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Definitiva Or _
                    CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaIda Or _
                    CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVolta Or _
                    CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                    flEnviarRealizadoPrevisto xmlOperacao
                End If
            End If
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim Then
        
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.Termo Then
            If Not xmlOperacao.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
                If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
                strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                'debito/Credito e Produto Original
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
        
                If blnCliente1 Then
                    fgConverterProduto1 xmlOperacao, True
                    strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
        Else
            If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
            strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            'debito/Credito e Produto Original
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Or _
               xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoInformalCompromissada Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                    strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
            
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgConverterProduto1 xmlOperacao
                strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                'debito/Credito e Produto Original
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        fgConverterProduto1 xmlOperacao, True, True
                        strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, True)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    End If
                End If
            End If
        End If
    End If

    'debito/Credito e Produto Original
    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    End If
    
    If Not xmlOperacao.selectSingleNode("//CO_PROD") Is Nothing Then
        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
    End If

    Set objGestaoCaixa = Nothing
    Set objSubReserva = Nothing

    Exit Function

ErrorHandler:
    Set objGestaoCaixa = Nothing
    Set objSubReserva = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de previsão de operações de redesconto para os sistemas A6 e PJ
'-------------------------------------------------------------------------------------------------

Public Function EnviarPrevisaoRDSC(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String

    On Error GoTo ErrorHandler

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")

    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, True)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, True)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva

    Set objGestaoCaixa = Nothing
    Set objSubReserva = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objGestaoCaixa = Nothing
    Set objSubReserva = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisaoRDSC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de previsão para os sistemas A6 e PJ para operações Cetip e BMA
'-------------------------------------------------------------------------------------------------

Public Function EnviarPrevisaoBMACetip(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                       ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim strMensagemME                           As String

Dim intTipoSolicitacao                      As Integer
Dim lngStatusOperacao                       As Long
Dim lngTipoMensagem                         As Long
Dim blnEnviarD0                             As Boolean
Dim blnEnviarVencimento                     As Boolean
Dim lngTipoOperacao                         As Long

    On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            'If xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
            '    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.nao
            'End If
        Else
            Exit Function
        End If
    End If

    blnEnviarD0 = True
    blnEnviarVencimento = True
    
    If Not xmlOperacao.selectSingleNode("//TP_SOLI") Is Nothing Then
        intTipoSolicitacao = Val(xmlOperacao.selectSingleNode("//TP_SOLI").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        lngStatusOperacao = Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        lngTipoMensagem = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
    End If
    
    If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
        If intTipoSolicitacao = enumTipoSolicitacao.Inclusao Then
            blnEnviarVencimento = False
        ElseIf intTipoSolicitacao = enumTipoSolicitacao.Complementacao Or _
               intTipoSolicitacao = enumTipoSolicitacao.Cancelamento Then
            blnEnviarD0 = False
        End If
    End If
    
    If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoDefinitivaCETIP Then
        If xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            blnEnviarD0 = False
        End If
    End If
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            
            If blnEnviarD0 Then
                strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, False
            End If
            
            If lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
                
                strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                
                flAntecipacaoResgateContratoDerivativo xmlOperacao, "PJ", True

            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.OperacaoDefinitivaCETIP Then
                
                If Not xmlOperacao.documentElement.selectSingleNode("CO_OPER_CETIP") Is Nothing Then
                    If xmlOperacao.documentElement.selectSingleNode("CO_OPER_CETIP").Text = enumOperacaoCETIP_CTP0052.ResgateAntecipado Then
                        If Not xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                            flAntecipacaoResgateContratoDerivativo xmlOperacao, "PJ", True
                        End If
                    End If
                End If

            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.DefinitivaCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.DefinitivaDescobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.DespesasBMC Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0105Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0106Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0108Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0109Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0121Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0134Clientes Then
                    
                flEnviarRealizadoPrevisto xmlOperacao
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoleto Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributos Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib Then
            
                flEnviarRealizadoPrevisto xmlOperacao
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacoesBMC Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica Then
                
                flEnviarRealizadoPrevisto xmlOperacao
                
                'Enviar Previsao PJ - Moeda Estrangeira
                strMensagemME = objGestaoCaixa.EnviarPrevisaoME(xmlOperacao)
                EnviarMensagemMQ strMensagemME, enumIdentificadorFila.BUS, False, False
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais Then
            
                flEnviarRealizadoPrevisto xmlOperacao
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or lngTipoOperacao = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
            Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
            Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                
                flEnviarRealizadoPrevisto xmlOperacao
            
            End If
            
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
                
                If blnEnviarD0 Then
                    EnviarRealPrevistoCETIP xmlOperacao
                End If
            
                If lngTipoMensagem = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
                   lngTipoMensagem = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
                    
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG") Is Nothing Then
                        If Not xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                        
                            strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                            Call EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
                            Call EnviarRealPrevistoCETIP(xmlOperacao, xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                        
                        End If
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                       lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
                    
                    If blnEnviarVencimento Then
                
                        If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                            
                            strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text)
                            Call EnviarMensagemMQ(strMensagem, enumIdentificadorFila.BUS, False, False)
                            Call EnviarRealPrevistoCETIP(xmlOperacao, xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text)
                        
                        End If
                    
                    End If
                
                End If
            
            End If

            Set objGestaoCaixa = Nothing

        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim Then
        
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        
        If lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
            
            fgAlterarDadosA6Volta xmlOperacao, "DT_LIQU_OPER_ATIV", "DT_OPER_ATIV", True
            strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, False)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV", "DT_LIQU_OPER_ATIV", True
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV_RETN", "DT_OPER_ATIV", False
            strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, True)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV", "DT_OPER_ATIV_RETN", False
        
        Else
            
            If blnEnviarD0 Then
                strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            End If
            
            If lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
               lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                
                flAntecipacaoResgateContratoDerivativo xmlOperacao, "A6", True
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.OperacaoDefinitivaCETIP Then
                
                If Not xmlOperacao.documentElement.selectSingleNode("CO_OPER_CETIP") Is Nothing Then
                    If xmlOperacao.documentElement.selectSingleNode("CO_OPER_CETIP").Text = enumOperacaoCETIP_CTP0052.ResgateAntecipado Then
                        'Estornar a Previsão do PJ na Data Vencimento Original
                        If Not xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                            flAntecipacaoResgateContratoDerivativo xmlOperacao, "A6", True
                        End If
                    End If
                End If

            End If
            
            If lngTipoMensagem = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
               lngTipoMensagem = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
                
                If Not xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG") Is Nothing Then
                    If Not xmlOperacao.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                    
                        strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                        Call EnviarMensagemMQ(strMensagem, enumIdentificadorFila.SubReserva)
                        
                    End If
                End If
            
            ElseIf lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
                
                If blnEnviarVencimento Then
                    
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text)
                        Call EnviarMensagemMQ(strMensagem, enumIdentificadorFila.SubReserva)
                    End If
                        
                End If
            
            End If
        
        End If
        
        Set objSubReserva = Nothing
    
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ_ME").Text = enumIndicadorSimNao.Sim Then
    
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    
        If lngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoInterbancaria Or lngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoArbitragem Then
    
            'Enviar Previsao PJ - Moeda Estrangeira
            strMensagemME = objGestaoCaixa.EnviarPrevisaoME(xmlOperacao)
            EnviarMensagemMQ strMensagemME, enumIdentificadorFila.BUS, True, False
    
        End If
    
        Set objGestaoCaixa = Nothing
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisaoBMACetip Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de estorno de previsão para os sistemas A6 e PJ
'-------------------------------------------------------------------------------------------------

Public Function EstornarPrevisao(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                 ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim xmlBoleto                               As MSXML2.DOMDocument40
Dim strOperacao                             As String
Dim blnCliente1                             As Boolean

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

Dim blnEnviarInformacaoPJ                   As Boolean

On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.SELIC Then
        EstornarPrevisaoBMACetip xmlOperacao, xmlControleProcessoOperacao
        Exit Function
    End If

    lngIndDebtCredOrig = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    lngCodProdOrig = xmlOperacao.selectSingleNode("//CO_PROD").Text

    blnCliente1 = fgVerificarCliente1(xmlOperacao)
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            
            If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            
            blnEnviarInformacaoPJ = True
            If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
                If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.VinculoDesvinculoTransferencia Then
                    If Not xmlOperacao.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                        If Trim$(xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text) = "YS" Then
                            blnEnviarInformacaoPJ = False
                        End If
                    End If
                End If
            End If
            
            If blnEnviarInformacaoPJ Then
                strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            End If
            
            'debito/Credito e Produto Original
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras And blnEnviarInformacaoPJ Then
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Sim
                strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Nao
            End If

            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    
                    If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                    strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

                End If
            End If
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgConverterProduto1 xmlOperacao
                strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                'debito/Credito e Produto Original
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        fgConverterProduto1 xmlOperacao, True, True
                        strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao, True)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        'debito/Credito e Produto Original
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    End If
                End If
            End If
            Set objGestaoCaixa = Nothing
            If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Definitiva Or _
                CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaIda Or _
                CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVolta Or _
                CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                flEstornarRealizadoPrevisto xmlOperacao
            End If
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.Termo Then
            If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            
            If blnEnviarInformacaoPJ Then
                strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            End If
            
            'debito/Credito e Produto Original
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                    strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    'debito/Credito e Produto Original
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
            
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgConverterProduto1 xmlOperacao
                strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                'debito/Credito e Produto Original
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        fgConverterProduto1 xmlOperacao, True, True
                        strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, True)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                        'debito/Credito e Produto Original
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                    End If
                End If
            End If
            Set objSubReserva = Nothing
        
        Else
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            
            strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, True)
            Call EnviarMensagemMQ(strMensagem, enumIdentificadorFila.SubReserva)
        
            If blnCliente1 Then
                Call fgDebitoCreditoCliente1(xmlOperacao)
                Call fgConverterProduto1(xmlOperacao, True, , False)
                strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            End If
        
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

            Set objSubReserva = Nothing
        
        End If
    
    End If

    'debito/Credito e Produto Original
    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarPrevisao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de envio de estorno de previsão para os sistemas A6 e PJ para operações Cetip e BMA

Public Function EstornarPrevisaoBMACetip(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                         ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String

Dim blnEnviarPrevisao                       As Boolean
Dim blnEnviarPrevisaoVencimento             As Boolean
Dim lngStatusOperacao                       As Long
Dim strValorComplementacao                  As String
Dim strMensagemME                           As String

Dim intTipoSolicitacao                      As Integer
Dim lngTipoMensagem                         As Long
Dim lngTipoOperacao                         As Long
    
    On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            'If xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = 0 Then
            '    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.nao
            'End If
        Else
            Exit Function
        End If
    End If

    blnEnviarPrevisao = True
    
    If Not xmlOperacao.selectSingleNode("//TP_SOLI") Is Nothing Then
        intTipoSolicitacao = Val(xmlOperacao.selectSingleNode("//TP_SOLI").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        lngStatusOperacao = Val(xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        lngTipoMensagem = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
    End If
    
    If Not xmlOperacao.selectSingleNode("//CodMsg") Is Nothing Then
        If xmlOperacao.selectSingleNode("//CodMsg").Text = "CTP0100R1" Then
            blnEnviarPrevisao = False
            xmlControleProcessoOperacao.selectSingleNode("//IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
        End If
    End If
                    
    If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP And _
       intTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
        blnEnviarPrevisao = False
        xmlControleProcessoOperacao.selectSingleNode("//IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
    End If
    
    'KIDA - 27/05/2009
    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
            xmlControleProcessoOperacao.selectSingleNode("//IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
        End If
    End If
    
                    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            
            'KIDA - CBLC
            If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoBrutaCBLC Then
                                                    
                If fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                    If blnEnviarPrevisao Then
                        strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    End If
                End If
                                    
            Else
                Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                If blnEnviarPrevisao Then
                    strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                End If
            End If
            
            If lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
               lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
                
                strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                
                flAntecipacaoResgateContratoDerivativo xmlOperacao, "PJ", False
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.DefinitivaCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.DefinitivaDescobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.DespesasBMC Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0006Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0007Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0008Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0009Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0025Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioTEDSTR0034Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0105Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0106Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0108Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0109Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0121Clientes Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPAG0134Clientes Then
                   
                flEstornarRealizadoPrevisto xmlOperacao
            
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoleto Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributos Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesas Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib Then
                
                flEstornarRealizadoPrevisto xmlOperacao
                
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais Then
                
                flEstornarRealizadoPrevisto xmlOperacao
                
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacoesBMC Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica Then
                
                flEstornarRealizadoPrevisto xmlOperacao
                
                'Estornar Previsao PJ - Moeda Estrangeira
                strMensagemME = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
                EnviarMensagemMQ strMensagemME, enumIdentificadorFila.BUS, False, False
                
            ElseIf lngTipoOperacao = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
                Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                Or lngTipoOperacao = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
                Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
   
                flEstornarRealizadoPrevisto xmlOperacao

            End If

            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
                
                If blnEnviarPrevisao Then
                    flEstornarRealPrevistoCETIP xmlOperacao
                End If
            
                If lngTipoMensagem = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
                   lngTipoMensagem = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
                    
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG") Is Nothing Then
                        
                        strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        
                        Call flEstornarRealPrevistoCETIP(xmlOperacao, xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                    
                    End If
                
                ElseIf lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                       lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
                    
                    If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                        
                        blnEnviarPrevisaoVencimento = True
                        
                        If (lngStatusOperacao = enumStatusOperacao.Cancelada Or _
                            lngStatusOperacao = enumStatusOperacao.CanceladaOrigem) And _
                            intTipoSolicitacao = enumTipoSolicitacao.Complementacao Then
                            
                            blnEnviarPrevisaoVencimento = False
                        
                        End If
                        
                        If Not xmlOperacao.selectSingleNode("//NAO_ENVIAR_VENCIMENTO") Is Nothing Then
                            blnEnviarPrevisaoVencimento = False
                        End If
                        
                        If blnEnviarPrevisaoVencimento Then
                            
                            strValorComplementacao = vbNullString
                            If Not xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_COMPL") Is Nothing Then
                                strValorComplementacao = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_COMPL").Text
                            End If
                            
                            strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao, , _
                                                                          xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text, _
                                                                          strValorComplementacao)
                            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        
                            Call flEstornarRealPrevistoCETIP(xmlOperacao, _
                                                             xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text, _
                                                             strValorComplementacao)
                        
                        End If
                        
                    End If
                
                End If
            End If

            Set objGestaoCaixa = Nothing
        End If
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim Then
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        If lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
           lngTipoOperacao = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
            
            fgAlterarDadosA6Volta xmlOperacao, "DT_LIQU_OPER_ATIV", "DT_OPER_ATIV", True
            strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV", "DT_LIQU_OPER_ATIV", True
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV_RETN", "DT_OPER_ATIV", False
            strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgAlterarDadosA6Volta xmlOperacao, "DT_OPER_ATIV", "DT_OPER_ATIV_RETN", False
        
        Else
            
            'KIDA - CBLC
            If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoBrutaCBLC Then
                                                    
                If fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                    If blnEnviarPrevisao Then
                        strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    End If
                End If
            Else
                If blnEnviarPrevisao Then
                    strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                End If
            End If
            
            If lngTipoMensagem = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
               lngTipoMensagem = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
                
                If Not xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG") Is Nothing Then
                    strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, , xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                End If
            
            ElseIf lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                   lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
                
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    
                    If blnEnviarPrevisaoVencimento Then
                        strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao, , _
                                                                     xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text, _
                                                                     strValorComplementacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                    End If
                
                End If
            
            End If
                
            If lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
               lngTipoOperacao = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                flAntecipacaoResgateContratoDerivativo xmlOperacao, "A6", False
            End If
            
        End If
        Set objSubReserva = Nothing
    
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ_ME").Text = enumIndicadorSimNao.Sim Then
    
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    
        If lngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoInterbancaria Or lngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoArbitragem Then
    
            'Estornar Previsao PJ - Moeda Estrangeira
            strMensagemME = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
            EnviarMensagemMQ strMensagemME, enumIdentificadorFila.BUS, False, False
    
        End If
    
        Set objGestaoCaixa = Nothing
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarPrevisaoBMACetip Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de realizado para o sistema PJ e do realizado solicitado para o sistema A6
'-------------------------------------------------------------------------------------------------

Private Function flEnviarRealizado(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                   ByRef xmlControleProcessoOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

On Error GoTo ErrorHandler

    lngCodProdOrig = 0

    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        lngIndDebtCredOrig = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    Else
        lngIndDebtCredOrig = 0
    End If
    
    If Not xmlOperacao.selectSingleNode("//CO_PROD") Is Nothing Then
        lngCodProdOrig = xmlOperacao.selectSingleNode("//CO_PROD").Text
    Else
        lngCodProdOrig = 0
    End If

    blnCliente1 = fgVerificarCliente1(xmlOperacao)

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
        
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            
            If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
            strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            
            'debito/Credito e Produto Original
            If lngIndDebtCredOrig <> 0 Then
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            End If
            
            If lngCodProdOrig <> 0 Then
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            End If
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras Then
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Sim
                strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Nao
            End If
            
            If blnCliente1 Then
                fgConverterProduto1 xmlOperacao
                strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                'debito/Credito e Produto Original
                If lngIndDebtCredOrig <> 0 Then
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                End If
                If lngCodProdOrig <> 0 Then
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
            Set objGestaoCaixa = Nothing
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlOperacao)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        'debito/Credito e Produto Original
        If lngIndDebtCredOrig <> 0 Then
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
        End If
        If lngCodProdOrig <> 0 Then
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
        End If
        
        If blnCliente1 Then
            fgConverterProduto1 xmlOperacao
            strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            'debito/Credito e Produto Original
            If lngIndDebtCredOrig <> 0 Then
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            End If
            If lngCodProdOrig <> 0 Then
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            End If
        End If
        Set objSubReserva = Nothing
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizado Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de realizado previsto para o sistema PJ
'-------------------------------------------------------------------------------------------------

Private Function flEnviarRealizadoPrevisto(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

On Error GoTo ErrorHandler

    If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
    Else
        xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
    End If

    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        lngIndDebtCredOrig = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    End If
    lngCodProdOrig = xmlOperacao.selectSingleNode("//CO_PROD").Text

    blnCliente1 = fgVerificarCliente1(xmlOperacao)

    If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
        If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True, False
        
        'debito/Credito e Produto Original
        If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
           xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
        End If
        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
        
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
            If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                
                'debito/Credito e Produto Original
                If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                End If
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            End If
        End If
        
        If blnCliente1 Then
            fgConverterProduto1 xmlOperacao
            strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            'debito/Credito e Produto Original
            If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            End If
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    fgConverterProduto1 xmlOperacao, True, True
                    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    'debito/Credito e Produto Original
                    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    End If
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
        End If
        Set objGestaoCaixa = Nothing
    End If

    'debito/Credito e Produto Original
    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    End If
    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

    xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Nao

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoPrevisto Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de envio de realizado previsto das operações CETIP para o sistema PJ

Public Function EnviarRealPrevistoCETIP(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                               Optional ByVal strDataVencimento As String = vbNullString) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
    Else
        xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
    End If

    If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao, , strDataVencimento)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
        Set objGestaoCaixa = Nothing
    End If

    xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Nao
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealPrevistoCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio de estorno de realizado previsto para o sistema PJ
'-------------------------------------------------------------------------------------------------

Private Function flEstornarRealizadoPrevisto(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

On Error GoTo ErrorHandler

    If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
    Else
        xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
    End If

    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        lngIndDebtCredOrig = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    End If
    lngCodProdOrig = xmlOperacao.selectSingleNode("//CO_PROD").Text

    blnCliente1 = fgVerificarCliente1(xmlOperacao)

    If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
        If blnCliente1 Then fgDebitoCreditoCliente1 xmlOperacao
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
        'debito/Credito e Produto Original
        If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
            xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
        End If
        xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
        
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
            If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                If blnCliente1 Then fgConverterProduto1 xmlOperacao, True
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao, True)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                'debito/Credito e Produto Original
                If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                    xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                End If
                xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            End If
        End If
        
        If blnCliente1 Then
            fgConverterProduto1 xmlOperacao
            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            'debito/Credito e Produto Original
            If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
            End If
            xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
            
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Then
                If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN") Is Nothing Then
                    fgConverterProduto1 xmlOperacao, True, True
                    strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao, True)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                    'debito/Credito e Produto Original
                    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
                    End If
                    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
                End If
            End If
        End If
        Set objGestaoCaixa = Nothing
    End If

    xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Nao

    'debito/Credito e Produto Original
    If Not xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    End If
    xmlOperacao.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarRealizadoPrevisto Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Controlar o fluxo de envio de estorno de realizado previsto para o sistema PJ

Private Function flEstornarRealPrevistoCETIP(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    Optional ByVal strDataVencimento As String = vbNullString, _
                                    Optional ByVal strValorComplementacao As String = vbNullString) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    If xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim
    Else
        xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
    End If

    If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao, , strDataVencimento, strValorComplementacao)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objGestaoCaixa = Nothing
    End If

    xmlOperacao.documentElement.selectSingleNode("ENVIAR_PREV").Text = enumIndicadorSimNao.Nao

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEstornarRealPrevistoCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo do recebimento do regsitro do legado
'-------------------------------------------------------------------------------------------------

Public Function GerenciarRecebimentoOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                    Optional ByRef strError As String) As Boolean
 
Dim objContingencia                         As A8LQS.clsContingencia
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode
Dim strOperacao                             As String
Dim lngTipoOperacao                         As Long
Dim strCodigoMensagem                       As String
Dim strMensagemRetornoLegado                As String
Dim lngLocalLiquidacao                      As Long
Dim intLayouEntrada                         As Integer
Dim strLeitura                              As String

Dim strRemessaAux                           As String

Dim objRS                                   As ADODB.Recordset

    On Error GoTo ErrorHandler

    strError = vbNullString
    
    Set objRS = Nothing
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    intLayouEntrada = 0
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        intLayouEntrada = Val(xmlRemessa.selectSingleNode("//TP_MESG").Text)
    End If
    
    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        
        If xmlRemessa.selectSingleNode("//NET_ENTRADA_MANUAL") Is Nothing Then
            If Not fgValidarRemessa(xmlRemessa, strError) Then
                GerenciarRecebimentoOperacao = False
                Exit Function
            End If
        End If
                
        If intLayouEntrada = enumTipoMensagemBUS.OperacoesCorretoras Then
            Call flAlteracaoRemessaCorretoras(xmlRemessa)
        End If
        
        Set objContingencia = CreateObject("A8LQS.clsContingencia")
        If objContingencia.VerificarSistContingencia(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) Then
            fgAppendNode xmlRemessa, "MESG", "IN_SIST_SITU_CNTG", enumIndicadorSimNao.Sim
        Else
            fgAppendNode xmlRemessa, "MESG", "IN_SIST_SITU_CNTG", enumIndicadorSimNao.Nao
        End If
        Set objContingencia = Nothing
        
        'KIDA - Entrada manual CETIP 21
        If Not xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            lngLocalLiquidacao = Val(xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text)
            If lngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
                Call flAlteracaoRemessaCETIP(xmlRemessa)
            End If
        End If
        
        'CCR
        If Not xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            lngLocalLiquidacao = Val(xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text)
            If lngLocalLiquidacao = enumLocalLiquidacao.CCR Then
                Call flAlteracaoRemessaCCR(xmlRemessa)
            End If
        End If
        
        Select Case intLayouEntrada
            Case enumTipoMensagemBUS.RegistroContratoSWAP_CETIP21
            
              If Not xmlRemessa.selectSingleNode("//CO_CNPJ_CNPT_CETIP21") Is Nothing Then
                 strRemessaAux = Replace$(xmlRemessa.xml, "CO_CNPJ_CNPT_CETIP21", "CO_CNPJ_CNPT")
                 xmlRemessa.loadXML strRemessaAux
              End If
        End Select
    
    Else

        If Not xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            lngLocalLiquidacao = Val(xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text)
            If lngLocalLiquidacao = enumLocalLiquidacao.CETIP Then
                Call flAlteracaoRemessaCETIP(xmlRemessa)
            End If
        End If
        
        'CCR
        If Not xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
            lngLocalLiquidacao = Val(xmlRemessa.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text)
            If lngLocalLiquidacao = enumLocalLiquidacao.CCR Then
                Call flAlteracaoRemessaCCR(xmlRemessa)
            End If
            
            If Not fgValidarRemessa(xmlRemessa, strError) Then
                GerenciarRecebimentoOperacao = False
                Exit Function
            End If

        End If
        
        Select Case intLayouEntrada
            
            Case enumTipoMensagemBUS.LancamentoPU_CETIP
                
                For Each objNode In xmlRemessa.selectNodes("//REPE_INFO_PU_FATR_SWAP/*")
                    
                    If Not objNode.selectSingleNode("TP_REFE") Is Nothing Then
                        If Not objNode.selectSingleNode("VA_PU_FATR") Is Nothing Then
                            If objNode.selectSingleNode("TP_REFE").Text = "F" Then
                                fgAppendNode xmlRemessa, "MESG", "VA_FATR_ATLZ", objNode.selectSingleNode("VA_PU_FATR").Text
                            Else
                                fgAppendNode xmlRemessa, "MESG", "VA_PU_NEGC", objNode.selectSingleNode("VA_PU_FATR").Text
                            End If
                        End If
                    End If
                    
                Next
        
            Case enumTipoMensagemBUS.OperacoesCorretoras
                
                Call flAlteracaoRemessaCorretoras(xmlRemessa)
        
            Case enumTipoMensagemBUS.EnvioTEDClientes
                
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
                fgAppendNode xmlRemessa, "MESG", "CO_AGEN", xmlRemessa.selectSingleNode("//CO_AGEN_DEBT").Text
                fgAppendNode xmlRemessa, "MESG", "CO_BANC", xmlRemessa.selectSingleNode("//CO_BANC_DEBT").Text
                fgAppendNode xmlRemessa, "MESG", "NU_CC", xmlRemessa.selectSingleNode("//NU_CONT_DEBT").Text
                fgAppendNode xmlRemessa, "MESG", "VA_LANC_CC", xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text
                fgAppendNode xmlRemessa, "MESG", "TP_CNTA_CRED", xmlRemessa.selectSingleNode("//TP_CONT_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "NU_CNTA_CRED", xmlRemessa.selectSingleNode("//NU_CONT_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "NO_TITU_1", xmlRemessa.selectSingleNode("//NO_TITU_CRED_1").Text
                fgAppendNode xmlRemessa, "MESG", "CO_FIND_COBE", xmlRemessa.selectSingleNode("//CO_FLND_TED").Text
                fgAppendNode xmlRemessa, "MESG", "CO_IDEF_ENVI_CORR", "0"
                    
                Select Case Val(xmlRemessa.documentElement.selectSingleNode("CO_MESG_STR_PAG").Text)
                
                    Case enumTEDStrPag.pagTransfContaClienteIF, _
                         enumTEDStrPag.pagTransfContasDiferentesTitula, _
                         enumTEDStrPag.pagTransfContasMesmaTitula, _
                         enumTEDStrPag.pagTransfEnvolvContasInvestimento, _
                         enumTEDStrPag.pagTransfIFContaCliente, _
                         enumTEDStrPag.pagTransfReservasBancDepositoJudi
                        
                        fgAppendNode xmlRemessa, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.PAG
                             
                    Case enumTEDStrPag.strTransfContaClienteIF, _
                         enumTEDStrPag.strTransfContasDiferentesTitula, _
                         enumTEDStrPag.strTransfContasMesmaTitula, _
                         enumTEDStrPag.strTransfEnvolvContasInvestimento, _
                         enumTEDStrPag.strTransfIFContaCliente, _
                         enumTEDStrPag.strTransfReservasBancDepositoJudi
                         
                        fgAppendNode xmlRemessa, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.SSTR
                         
                End Select
                
                If Not xmlRemessa.selectSingleNode("//ID_DEPO_JUDI") Is Nothing Then
                    xmlRemessa.selectSingleNode("//ID_DEPO_JUDI").Text = Format$(Val(xmlRemessa.selectSingleNode("//ID_DEPO_JUDI").Text), String$(18, "0"))
                End If
            
            Case enumTipoMensagemBUS.EnvioPagDespesas
            
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
                fgAppendNode xmlRemessa, "MESG", "CO_AGEN", xmlRemessa.selectSingleNode("//CO_AGEN_DEBT_TRIB").Text
                fgAppendNode xmlRemessa, "MESG", "CO_BANC", xmlRemessa.selectSingleNode("//CO_BANC_DEBT_TRIB").Text
                fgAppendNode xmlRemessa, "MESG", "NU_CC", xmlRemessa.selectSingleNode("//NU_CONT_DEBT_TRIB").Text
                fgAppendNode xmlRemessa, "MESG", "VA_LANC_CC", xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text
                fgAppendNode xmlRemessa, "MESG", "VA_LANC_CC_CPMF", xmlRemessa.selectSingleNode("//VA_OPER_ATIV_CPMF").Text
                fgAppendNode xmlRemessa, "MESG", "TP_CNTA_CRED", xmlRemessa.selectSingleNode("//TP_CONT_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "NU_CNTA_CRED", xmlRemessa.selectSingleNode("//NU_CONT_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "NO_TITU_1", xmlRemessa.selectSingleNode("//NO_TITU_CRED_1").Text
                fgAppendNode xmlRemessa, "MESG", "CO_FIND_COBE", xmlRemessa.selectSingleNode("//CO_FLND_TED").Text
                fgAppendNode xmlRemessa, "MESG", "CO_IDEF_ENVI_CORR", "0"
                fgAppendNode xmlRemessa, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.SSTR
                fgAppendNode xmlRemessa, "MESG", "CO_CNPJ_CNPT", xmlRemessa.selectSingleNode("//CO_CNPJ_CPF_CRED_1").Text
                fgAppendNode xmlRemessa, "MESG", "NO_CNPT", xmlRemessa.selectSingleNode("//NO_TITU_CRED_1").Text
                fgAppendNode xmlRemessa, "MESG", "CO_ISPB_BANC_LIQU_CNPT", xmlRemessa.selectSingleNode("//CO_ISPB_IF_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "CO_AGEN_COTR", xmlRemessa.selectSingleNode("//CO_AGEN_CRED").Text
                fgAppendNode xmlRemessa, "MESG", "NU_CC_COTR", xmlRemessa.selectSingleNode("//NU_CONT_CRED").Text
                
                If xmlRemessa.selectSingleNode("//CO_MESG_STR_PAG") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_MESG_STR_PAG", "0"
                End If
                
                If Val("0" & xmlRemessa.selectSingleNode("//CO_MESG_STR_PAG").Text) = 7 Then
                    fgAppendNode xmlRemessa, "MESG", "TP_IF_CRED_DEBT", enumTipoDebitoCredito.Debito
                End If
            
            Case enumTipoMensagemBUS.LancamentoContaCorrenteBG
                
                fgAppendNode xmlRemessa, "MESG", "VA_OPER_ATIV", xmlRemessa.selectSingleNode("//VA_LANC_CC").Text
            
            Case enumTipoMensagemBUS.LiquidacaoFisicaOperacaoBMA
                
                fgAppendNode xmlRemessa, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.BMA
            
            Case enumTipoMensagemBUS.RegistroLiquidacaoEventoCBLC
                
                If Not xmlRemessa.selectSingleNode("//TP_PAGTO_LDL") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "TP_PGTO_LDL", xmlRemessa.selectSingleNode("//TP_PAGTO_LDL").Text
                End If

                If Not xmlRemessa.selectSingleNode("//CO_PARP_NEGO") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_PARP_CAMR", xmlRemessa.selectSingleNode("//CO_PARP_NEGO").Text
                End If
            
            Case enumTipoMensagemBUS.RegistroLiquidacaoBrutaCBLC
                
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Bruta

                If Not xmlRemessa.selectSingleNode("//CO_PARP_NEGO") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_PARP_CAMR", xmlRemessa.selectSingleNode("//CO_PARP_NEGO").Text
                End If

            Case enumTipoMensagemBUS.RegistroLiquidacaoMultilateralCBLC
                
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral

                If Not xmlRemessa.selectSingleNode("//CO_PARP_NEGO") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_PARP_CAMR", xmlRemessa.selectSingleNode("//CO_PARP_NEGO").Text
                End If

            Case enumTipoMensagemBUS.RegistroLiquidacaoMultilateralBMF
                
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral
                
                If Not xmlRemessa.selectSingleNode("//CO_PARP_NEGO") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_PARP_CAMR", xmlRemessa.selectSingleNode("//CO_PARP_NEGO").Text
                End If

            Case enumTipoMensagemBUS.LiquidacaoMultilateralBMC
                
                fgAppendNode xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral

            Case enumTipoMensagemLQS.AlteracaoDadosContaCorrente
                
                Call flAlterarDadosContaCorrente(xmlRemessa)
                Exit Function

            Case enumTipoMensagemBUS.TermoLiquidacao
                
                If xmlRemessa.selectSingleNode("//TP_LIQU") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "TP_LIQU", "02"
                Else
                    If Trim(xmlRemessa.selectSingleNode("//TP_LIQU").Text) = vbNullString Then
                        xmlRemessa.selectSingleNode("//TP_LIQU").Text = "02"
                    End If
                End If

            Case enumTipoMensagemBUS.CompromissadaIda
                
                If xmlRemessa.selectSingleNode("//NU_PROC") Is Nothing Then
                    If xmlRemessa.selectSingleNode("//CO_IDEF_LAST") Is Nothing Then
                        fgAppendNode xmlRemessa, "MESG", "CO_IDEF_LAST", xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & _
                                                                         xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text & _
                                                                         xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text
                    Else
                        If Trim$(xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text) = vbNullString Then
                            xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text = xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & _
                                                                                 xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text & _
                                                                                 xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text
                        End If
                    End If
                End If
            
            Case enumTipoMensagemBUS.OperacaoCompromissadaCETIP
                
                If xmlRemessa.selectSingleNode("//NU_PROC") Is Nothing Then
                    If xmlRemessa.selectSingleNode("//CO_IDEF_LAST") Is Nothing Then
                        fgAppendNode xmlRemessa, "MESG", "CO_IDEF_LAST", xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP").Text & _
                                                                         xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text & _
                                                                         xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text
                    Else
                        If Trim$(xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text) = vbNullString Then
                            xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text = xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC_CETIP").Text & _
                                                                                 xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text & _
                                                                                 xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text
                        End If
                    End If
                End If
            
            Case enumTipoMensagemBUS.Redesconto
                
                If xmlRemessa.selectSingleNode("//CO_IDEF_LAST") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "CO_IDEF_LAST", xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & _
                                                                     xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text
                Else
                    If Trim(xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text) = vbNullString Then
                        xmlRemessa.selectSingleNode("//CO_IDEF_LAST").Text = xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text & _
                                                                             xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text
                    End If
                End If
            
            Case enumTipoMensagemBUS.DespesasBMC
            
                If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
                Else
                    xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoEntradaSaida.Saida
                End If
            
            Case enumTipoMensagemBUS.RegistroContratoSWAP_CETIP21
            
                If Not xmlRemessa.selectSingleNode("//CO_CNPJ_CNPT_CETIP21") Is Nothing Then
                   strRemessaAux = Replace$(xmlRemessa.xml, "CO_CNPJ_CNPT_CETIP21", "CO_CNPJ_CNPT")
                   xmlRemessa.loadXML strRemessaAux
                   For Each objNode In xmlRemessa.selectNodes("//GRUP_INFO_CNTR")
                      If Not objNode.selectSingleNode("CO_CNPJ_CNPT") Is Nothing Then
                          If Len(Trim(objNode.selectSingleNode("CO_CNPJ_CNPT").Text)) > 14 Then
                              objNode.selectSingleNode("CO_CNPJ_CNPT").Text = Right(Trim(objNode.selectSingleNode("CO_CNPJ_CNPT").Text), 14)
                          End If
                      End If
                   Next
                End If
            
        End Select


        If Not xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER") Is Nothing Then
            If IsNumeric(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text) Then
                If Len(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text) < 16 Or Left(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text, 1) = "0" Then
                    xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text = Val(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text)
                End If
            End If
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG") Is Nothing Then
            If IsNumeric(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG").Text) Then
                If Len(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG").Text) < 16 Or Left(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG").Text, 1) = "0" Then
                    xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG").Text = Val(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER_ORIG").Text)
                End If
            End If
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//TP_TRAF_LDL") Is Nothing Then
            If IsNumeric(xmlRemessa.documentElement.selectSingleNode("//TP_TRAF_LDL").Text) Then
                xmlRemessa.documentElement.selectSingleNode("//TP_TRAF_LDL").Text = Format$(xmlRemessa.documentElement.selectSingleNode("//TP_TRAF_LDL").Text, "00")
            End If
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//CO_CNTR_SISB") Is Nothing Then
            If IsNumeric(xmlRemessa.documentElement.selectSingleNode("//CO_CNTR_SISB").Text) Then
                xmlRemessa.documentElement.selectSingleNode("//CO_CNTR_SISB").Text = Val(xmlRemessa.documentElement.selectSingleNode("//CO_CNTR_SISB").Text)
            End If
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//NR_IDEF_NEGO_BMC") Is Nothing Then
            If IsNumeric(xmlRemessa.documentElement.selectSingleNode("//NR_IDEF_NEGO_BMC").Text) Then
                xmlRemessa.documentElement.selectSingleNode("//NR_IDEF_NEGO_BMC").Text = Val(xmlRemessa.documentElement.selectSingleNode("//NR_IDEF_NEGO_BMC").Text)
            End If
        End If
    
        If xmlRemessa.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
            Call fgAppendNode(xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", 0)
        Else
            If Trim$(xmlRemessa.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED").Text) = vbNullString Then
                xmlRemessa.documentElement.selectSingleNode("//IN_OPER_DEBT_CRED").Text = 0
            End If
        End If
    
        'Quando a mensagem = 76, deixar o campo TP_LIQU_OPER_ATIV = null
        'Aline/Ivan - 29/08/2011
        If intLayouEntrada <> enumTipoMensagemBUS.RegistroContratoTermoCETIP Then
            If xmlRemessa.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
                Call fgAppendNode(xmlRemessa, "MESG", "TP_LIQU_OPER_ATIV", 0)
            Else
                If Trim$(xmlRemessa.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = vbNullString Then
                    xmlRemessa.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text = 0
                End If
            End If
        End If
    
        If xmlRemessa.documentElement.selectSingleNode("//CO_FORM_LIQU") Is Nothing Then
            Call fgAppendNode(xmlRemessa, "MESG", "CO_FORM_LIQU", enumFormaLiquidacao.Contabil)
        Else
            If Trim$(xmlRemessa.documentElement.selectSingleNode("//CO_FORM_LIQU").Text) = vbNullString Then
                xmlRemessa.documentElement.selectSingleNode("//CO_FORM_LIQU").Text = enumFormaLiquidacao.Contabil
            End If
        End If
    
        Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
        objTipoOperacao.ObterTipoOperacaoPorMsgBUS xmlRemessa, lngTipoOperacao, strCodigoMensagem, strMensagemRetornoLegado

        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        strLeitura = objWorkflow.ObterInatividadeFluxoRecebimento(lngTipoOperacao, _
                                                                  , _
                                                                  lngLocalLiquidacao)
        Set objWorkflow = Nothing

        If strLeitura <> vbNullString Then
            Call flProcessarBoleto(xmlRemessa, enumStatusOperacao.Inativa)
            Exit Function
        End If
        
        If intLayouEntrada <> enumTipoMensagemBUS.PagamentoRedesconto Then

            Set objRS = CreateObject("ADODB.Recordset")
            
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                         , , , , , _
                                                         objRS, , , _
                                                         False)
    
            If xmlOperacao.Text <> vbNullString Then
                If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) <> lngTipoOperacao Then
                    If xmlOperacao.documentElement.selectSingleNode("SG_SIST_CETIP") Is Nothing Then
                        '3037 - Identificador da Operação já existe.
                        lngCodigoErroNegocio = 3037
                        GoTo ErrorHandler
                    Else
                        If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Inicial Then
                            '3037 - Identificador da Operação já existe.
                            lngCodigoErroNegocio = 3037
                            GoTo ErrorHandler
                        End If
                    End If
                End If
            End If
        
        End If

    End If
        

    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text)
        Case enumTipoSolicitacao.Inclusao
            
            flProcessarBoleto xmlRemessa, enumStatusOperacao.Inicial

        Case enumTipoSolicitacao.Complementacao, enumTipoSolicitacao.Confirmacao
            
            If Not xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR_BMC") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "VA_MOED_ESTR", xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR_BMC").Text
                Else
                    xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR").Text = xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR_BMC").Text
                End If
            End If
            
            If Not xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC_BMC") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "QT_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC_BMC").Text
                Else
                    xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text = xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC_BMC").Text
                End If
                If xmlRemessa.documentElement.selectSingleNode("QT_MINI_ATIV_MERC") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "QT_MINI_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC_BMC").Text
                Else
                    xmlRemessa.documentElement.selectSingleNode("QT_MINI_ATIV_MERC").Text = xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC_BMC").Text
                End If
            End If
            
            If Not xmlRemessa.documentElement.selectSingleNode("TP_REQU_BMC") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("TP_REQU") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "TP_REQU", xmlRemessa.documentElement.selectSingleNode("TP_REQU_BMC").Text
                Else
                    xmlRemessa.documentElement.selectSingleNode("TP_REQU").Text = xmlRemessa.documentElement.selectSingleNode("TP_REQU_BMC").Text
                End If
            End If
            
            If Not xmlRemessa.documentElement.selectSingleNode("TP_TRAF") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "TP_TRAF_LDL", xmlRemessa.documentElement.selectSingleNode("TP_TRAF").Text
                Else
                    xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text = xmlRemessa.documentElement.selectSingleNode("TP_TRAF").Text
                End If
            End If
            
            If xmlOperacao.xml <> vbNullString Then
                If Not objRS Is Nothing Then
                    If objRS.RecordCount > 1 Then
                        objRS.MoveFirst
                        While Not objRS.EOF
                            If Val(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text) <> objRS!NU_SEQU_OPER_ATIV And _
                               Val(xmlOperacao.documentElement.selectSingleNode("TP_SOLI").Text) = enumTipoSolicitacao.Inclusao Then
                                
                                Call objOperacao.AlterarStatusAcao(objRS!NU_SEQU_OPER_ATIV, _
                                                                   enumStatusOperacao.Rejeitada, _
                                                                   vbNullString, _
                                                                   , _
                                                                   enumTipoAcao.RejeicaoPorDuplicidade, _
                                                                   False)
                            
                            End If
                            objRS.MoveNext
                        Wend
                    End If
                End If
            End If
            
            flComplementarOperacao xmlRemessa, strError

        Case enumTipoSolicitacao.Cancelamento

            If xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
                    fgAppendNode xmlRemessa, "MESG", "TP_OPER", lngTipoOperacao
                Else
                    xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = lngTipoOperacao
                End If
            End If

            flCancelarOperacao xmlRemessa, strError

        Case enumTipoSolicitacao.CancelamentoComMensagem
            flCancelarOperacaoComMensagem xmlRemessa

        Case enumTipoSolicitacao.Alteracao
            flAlterarOperacao xmlRemessa

        Case enumTipoSolicitacao.CancelamentoPorLastro
            
            If xmlRemessa.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "TP_OPER", lngTipoOperacao
            Else
                xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = lngTipoOperacao
            End If

            flCancelarPorLastro xmlRemessa
            
        Case enumTipoSolicitacao.Reativacao
            
            If intLayouEntrada = enumTipoMensagemLQS.RegistroOperacoesBMC Then
                flProcessaReativacao xmlRemessa
            End If

    End Select
    
    If Not xmlRemessa.selectSingleNode("//CO_CHAV_BMC") Is Nothing Then
        Call flAlterarStatusBMC0112(xmlRemessa.selectSingleNode("//CO_CHAV_BMC").Text, _
                                    xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    
    End If
    
    GerenciarRecebimentoOperacao = True

    Set objTipoOperacao = Nothing

    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        If Not objMQAX200 Is Nothing Then
            objMQAX200.MQEnd
        End If
        Set objMQAX200 = Nothing
    End If

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objTipoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing


    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de execução por operação
'-------------------------------------------------------------------------------------------------

Private Function flGerenciarChamadas(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                     ByVal pstrFuncionalidade As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objRegra                                As A8LQS.clsRegra
Dim objContingencia                         As A8LQS.clsContingencia
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacaoAux          As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim dtmDataOperacao                         As Date
Dim blnCliente1                             As Boolean
Dim lngTipoLiquidacao                       As Long
Dim lngCodigoRetornoVerificacao             As Long

Dim blnEnviarInformacaoPJ                   As Boolean

On Error GoTo ErrorHandler

    'Tratar BMA e CETIP e outras
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.SELIC Then
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        objWorkflow.GerenciarChamadas xmlRemessa, pstrFuncionalidade
        Set objWorkflow = Nothing
        Exit Function
    End If

    'Carlos - Gerenciamento de chamada das rotinas
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    'If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '    lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    'End If
    strXML = LerTodos(pstrFuncionalidade, _
                      CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                      enumIndicadorSimNao.Nao, , _
                      lngTipoLiquidacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
        'na entrada manual.
        If Not xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
        If xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
            End If
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
            End If
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If
        End If
        Else
            If pstrFuncionalidade = "Confirmacao" Then
                Set xmlControleProcessoOperacaoAux = CreateObject("MSXML2.DOMDocument.4.0")
                'If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                '    lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
                'End If
                strXML = LerTodos("RecebimentoBoleto", _
                                  CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                  enumIndicadorSimNao.Nao, , _
                                  lngTipoLiquidacao)

                If strXML = "" Then
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
                Else
                    xmlControleProcessoOperacaoAux.loadXML strXML

                    'se entrada manual envia a previsão na confirmação
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = xmlControleProcessoOperacaoAux.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text
                    xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = xmlControleProcessoOperacaoAux.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text
                End If
            End If
        End If
        If xmlRemessa.documentElement.selectSingleNode("IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim Then
            EnviarPrevisao xmlRemessa, xmlControleProcessoOperacao
        End If
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CONF").Text = enumIndicadorSimNao.Sim Then
        Set objRegra = CreateObject("A8LQS.clsRegra")
        'Verifica regra para confirmação automática
        If objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Confirmar, lngCodigoRetornoVerificacao) Then
            'Verifica para Despesas e Eventos, caso seja confirmação automática não alterar o Status para Concordância.
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.DespesaSelic And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.EventosAmortização And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.EventosJuros And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.EventosResgate Then
                Set objContingencia = CreateObject("A8LQS.clsContingencia")
                If objContingencia.VerificarSistContingencia(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text) Then
                    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                    objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                 xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                 enumJustificativa.SistemaEmContingencia
                    Set objHistSituacaoOperacao = Nothing
                    Exit Function
                Else
                    'Qdo de entrada manual sempre parar na entrada manual
                    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
                        Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                        objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                     xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                     enumJustificativa.EntradaManual
                        Set objHistSituacaoOperacao = Nothing
                        Exit Function
                    End If

                    'Para Operações Definitivas e Compromissadas a data pode ser retroativa, com isso para a confirmação
                    If Not xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                        dtmDataOperacao = fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text)
                    Else
                        dtmDataOperacao = fgDataHoraServidor(enumFormatoDataHora.Data)
                    End If

                    If dtmDataOperacao = fgDataHoraServidor(enumFormatoDataHora.Data) Then
                        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                            enumStatusOperacao.ConcordanciaAutomatica
                        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica
                    Else
                        Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                        objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                     xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                     enumJustificativa.OperacaoRetroativa
                        Set objHistSituacaoOperacao = Nothing
                        Exit Function
                    End If
                End If
                Set objContingencia = Nothing
            End If
        Else
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         IIf(lngCodigoRetornoVerificacao <> 0, _
                                                                lngCodigoRetornoVerificacao, _
                                                                enumJustificativa.RegraWorkflow)
            Set objHistSituacaoOperacao = Nothing
            Exit Function
        End If
        Set objRegra = Nothing
    End If

    If fgVerificarCliente1(xmlRemessa) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    'Nestas operações é enviado somente após a conciliação
    'Inclusao de Despesa - Carlos 06/04/2004
    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.TermoDataLiquidacaoCerta And _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta And _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.DespesaSelic And _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.EventosSelic Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
            If xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
                If blnCliente1 Then
                    fgDebitoCreditoCliente1 xmlRemessa
                End If
                Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                
                blnEnviarInformacaoPJ = True
                If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
                    If Val(xmlRemessa.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.VinculoDesvinculoTransferencia Then
                        If Not xmlRemessa.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                            If Trim$(xmlRemessa.selectSingleNode("//SG_SIST_ORIG").Text) = "YS" Then
                                blnEnviarInformacaoPJ = False
                            End If
                        End If
                    End If
                End If
                
                If blnEnviarInformacaoPJ Then
                    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                End If
                
                If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras And blnEnviarInformacaoPJ Then
                    fgCriarTAGTransferencia xmlRemessa, enumIndicadorSimNao.Sim
                    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                    fgCriarTAGTransferencia xmlRemessa, enumIndicadorSimNao.Nao
                End If
                If blnCliente1 Then
                    'irá retornar o debito/credito original
                    fgDebitoCreditoCliente1 xmlRemessa
                    fgConverterProduto1 xmlRemessa
                    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                    fgConverterProduto1 xmlRemessa
                End If
                Set objGestaoCaixa = Nothing
            End If
        End If
    End If

    'Somente para Operação Compromissada de Volta, entra para conciliar se a Data da Operação for igual Data de Vencimento
    If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVolta Then
        If Not fgDatasIguaisCompromissada(xmlRemessa) Then
            xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CNCL").Text = enumIndicadorSimNao.Nao
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CNCL").Text = enumIndicadorSimNao.Sim Then
        Set objRegra = CreateObject("A8LQS.clsRegra")
        'Verifica regra para conciliação automática
        If objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Conciliar, lngCodigoRetornoVerificacao) Then
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            If Not objConciliacao.ConciliarOperacao(xmlRemessa) Then
                Set objConciliacao = Nothing
                Exit Function
            End If
            Set objConciliacao = Nothing
        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar, _
                                                IIf(lngCodigoRetornoVerificacao <> 0, _
                                                        lngCodigoRetornoVerificacao, _
                                                        enumJustificativa.RegraWorkflow)
            Exit Function
        End If
        Set objRegra = Nothing
    End If

    'Termo Liquidacao e Eventos Selic envia neste momento e não após a Liberação.
    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EventosAmortização Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EventosJuros Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EventosResgate Then
            'Função especifica de termo liquidação e eventos, pois não tem liberação para a mesma
            LiquidarTermoVoltaEventos xmlRemessa, xmlRemessa
            Exit Function
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
            LiquidarComprVoltaConciliacao xmlRemessa
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_LIBE").Text = enumIndicadorSimNao.Sim Then
        'Alterado tipo de ação para liberação
        'xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text = enumTipoSolicitacao.Liberar
        '************************************************************************************************
        'Inibido a função de enviar para fila, para teste com o BUS gerando a mensagem - Carlos 25/09/2003
        'EnviarMensagemMQ xmlRemessa.xml, enumIdentificadorFila.Liberacao
        objOperacao.OperacaoDisponivelConsulta xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               enumIndicadorSimNao.Nao
        flLiberarOperacao xmlRemessa, False
        '************************************************************************************************
    End If

    Set objOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamadas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de execução do reenvio de uma operação
'-------------------------------------------------------------------------------------------------

Public Function GerenciarReenvio(ByVal strOperacao As String) As Boolean

Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim udtProtocoloSize                        As udtProtocoloAux
Dim lngTipoLiquidacao                       As Long
Dim strProtocolo                            As String
Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlOperacao.loadXML Mid(strOperacao, Len(udtProtocoloSize) + 1)
    strProtocolo = Mid(strOperacao, 1, Len(udtProtocoloSize))

    If xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        Exit Function
    End If

    'Carlos - Gerenciamento de Envio de Mensagens
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    'If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '    lngTipoLiquidacao = CLng(xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    'End If
    
    strXML = LerTodos("Reenvio", _
                      CLng(xmlOperacao.selectSingleNode("//TP_OPER").Text), _
                      enumIndicadorSimNao.Nao, , _
                      lngTipoLiquidacao)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    EnviarPrevisao xmlOperacao, xmlControleProcessoOperacao
    flEnviarRealizado xmlOperacao, xmlControleProcessoOperacao


    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CCR Then
        If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then
            flEnviarPrivistoMEPJ xmlOperacao
        End If
    End If

    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
        'Transformar o Tipo de Débito/Crédito para Mensageria
        'RATS 399 - Ocor 01
        For Each xmlNode In xmlOperacao.selectNodes("//IN_OPER_DEBT_CRED")
            xmlNode.Text = IIf(CLng(xmlNode.Text) = enumTipoDebitoCredito.Debito, "D", "C")
        Next
    End If

    'Mensageria CETIP
    'Alteração de envio de mensagem conforme planilha de mensageria cetip (ManualdeMensageriaCTPv9.xls)

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Call objMensagem.MontarMensagemCETIP(xmlOperacao)
    Set objMensagem = Nothing

    strOperacao = strProtocolo & xmlOperacao.xml

    EnviarMensagemMQ strOperacao, enumIdentificadorFila.BUS, False, True

    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
        'Transformar o Tipo de Débito/Crédito para Mensageria
        
        'RATS 399 - Ocor 01
        For Each xmlNode In xmlOperacao.selectNodes("//IN_OPER_DEBT_CRED")
            xmlNode.Text = IIf(xmlNode.Text = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
        Next
    End If

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set objMensagem = Nothing
    Set xmlOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarReenvio Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de execução por operação quando esta executado via Tela
'-------------------------------------------------------------------------------------------------

Public Function GerenciarChamadaViaTela(ByRef pvntSeqOperacaoControleIF As Variant, _
                                        ByVal penumFuncionalidadeFuncaoSist As Long, _
                               Optional ByVal pblnIgnoraGradeHorario As Boolean = False, _
                               Optional ByVal pblnOrigemConciliacao As Boolean = False, _
                               Optional ByVal pblnMensagemParaMensagemSPB As Boolean = False, _
                               Optional ByRef xmlRepeticao As MSXML2.DOMDocument40, _
                               Optional ByRef plngRepeticao As Long, _
                               Optional ByVal pstrDtHrUltimaAtlz As String, _
                               Optional ByVal pblnFluxoAutomatico As Boolean = False, _
                               Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString, _
                               Optional ByVal plngNuOperAtiv As Long, _
                               Optional ByVal pblnLiquidacaoPorLote As Boolean = False, _
                               Optional ByVal plngTipoBackOfficeLote As Long) As Boolean

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim xmlTipoLiquidacao                       As MSXML2.DOMDocument40
Dim xmlMesgOrig                             As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim objTipoLiquidacao                       As A6A7A8.clsTipoLiquidacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim strTipoLiquidacao                       As String

On Error GoTo ErrorHandler

    'Liberação de mensagem para geração de nova mensagem
    If pblnMensagemParaMensagemSPB Then
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        objWorkflow.LiberarMensagem pvntSeqOperacaoControleIF, _
                                    penumFuncionalidadeFuncaoSist, _
                                    xmlRepeticao, _
                                    plngRepeticao, _
                                    pblnIgnoraGradeHorario, _
                                    pstrDataRegistroMensagemSPB, _
                                    plngNuOperAtiv
        Set objWorkflow = Nothing

        'Finaliza o MQ para a mensagem ser liberada
        If Not objMQAX200 Is Nothing Then
            objMQAX200.MQEnd
        End If
        Set objMQAX200 = Nothing
        Exit Function
    End If

    'Liberar Mensagens STR
    If penumFuncionalidadeFuncaoSist = enumFuncionalidade.LiberacaoMensagemSTR Then
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        If pblnLiquidacaoPorLote Then
            objWorkflow.GravarMensagemSTR xmlRepeticao, pblnLiquidacaoPorLote, plngTipoBackOfficeLote
        Else
            objWorkflow.GravarMensagemSTR xmlRepeticao
        End If
        Set objWorkflow = Nothing
        Set objMQAX200 = Nothing
        Exit Function
    End If

    'Tratamento especial para Eventos - Carlos 16/06
    If penumFuncionalidadeFuncaoSist = enumFuncionalidade.LiquidarEventos Then
        LiquidarEventos xmlRepeticao
        
        'Finaliza o MQ para a mensagem ser liberada
        If Not objMQAX200 Is Nothing Then
            objMQAX200.MQEnd
        End If
        Set objMQAX200 = Nothing
        Exit Function
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objWorkflow = CreateObject("A8LQS.clsWorkflow")

    'Obter o xml para processamento do envio da mensagem
    Set xmlOperacao = objOperacao.LerXMLOperacao(pvntSeqOperacaoControleIF)
    
    'Pikachu - 19/07/2004
    'Para Solicitacao de cancelamento , obter o XML do cancelamento (TP_SOLI = 4) para CETIP
    If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        If xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = enumTipoAcao.CancelamentoSolicitado Then
            Set xmlOperacao = objOperacao.LerXMLOperacao(pvntSeqOperacaoControleIF, , , , , , True)
        End If
    End If
    '-----------------------------------------------------------------------------------------

    'Obter o codigo da mensagem SPB
    xmlTipoOperacao.loadXML objTipoOperacao.Ler(CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text))
    If (xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text) <> vbNullString Then
        
        'Pikachu - 17/11/2004
        'Evidência 84
        'A BMA 006 DEVERÁ SER EMITIDA SOMENTE PARA A 1º OPERAÇÃO
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecCompromissadaCobertura Or _
           xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecCompromissadaIntermediacao Or _
           xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecDefinitivaCobertura Or _
           xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecDefinitivaIntermediacao Or _
           xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecTermoCobertura Or _
           xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.EspecTermoIntermediacao Then
                                   
            If Not xmlOperacao.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
                If xmlOperacao.documentElement.selectSingleNode("CO_MESG").Text <> "BMA0004" Then
                    fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                End If
            Else
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DespesasBMC Then
                    fgAppendNode xmlOperacao, "MESG", "CO_MESG", "STR0007"
                ElseIf xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
                    If Val(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaReativacao Then
                        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                
                        Set xmlMesgOrig = objMensagemEnviadaRecebida.LerXMLMensagem(, , , "BMC0002R1" _
                                                                                    , , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                
                        If xmlMesgOrig.xml <> vbNullString Then
                            fgAppendNode xmlOperacao, "MESG", "CO_MESG", "BMC0003"
                            fgAppendNode xmlOperacao, "MESG", "NU_CTRL_BMC_ORIG", xmlMesgOrig.selectSingleNode("//NU_CTRL_CAMR").Text
                        Else
                            '4306 - Mensagem original nao localizada para envio de mensagem de reativacao
                            lngCodigoErroNegocio = 4306
                            fgGerarAlerta xmlOperacao, , , enumFatorGeradorAlerta.MensagemNaoLocalizada
                            GoTo ErrorHandler
                        End If
                        Set objMensagemEnviadaRecebida = Nothing
                        Set xmlMesgOrig = Nothing
                    Else
                        fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                    End If
                Else
                    fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
                End If
            End If
        '------------------------------------------------------------------------------------------------------------------------------------------
        Else
            fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
        End If
        
    Else
        'se não tem msg spb verificar se esta na tabela de tipo de liquidação
        If xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "CO_MESG", vbNullString
        Else
            Set objTipoLiquidacao = CreateObject("A6A7A8.clsTipoLiquidacao")
            strTipoLiquidacao = objTipoLiquidacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
            If strTipoLiquidacao <> vbNullString Then
                Set xmlTipoLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
                xmlTipoLiquidacao.loadXML strTipoLiquidacao
                fgAppendNode xmlOperacao, "MESG", "CO_MESG", xmlTipoLiquidacao.documentElement.selectSingleNode("CO_MESG_SPB_LIQU").Text
                Set xmlTipoLiquidacao = Nothing
            Else
                fgAppendNode xmlOperacao, "MESG", "CO_MESG", vbNullString
            End If
            Set objTipoLiquidacao = Nothing
        End If
    End If
    fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text

    'Carlos 01/07/2004
    'Alterado para não deixar confirmar/liberar operações com datas erradas
    If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU") Is Nothing Then
        If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing And _
            Not xmlOperacao.documentElement.selectSingleNode("DT_MESG") Is Nothing Then
            If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text < fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                If xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text <> fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                    '3100 - Data da Operação Inválida.
                    lngCodigoErroNegocio = 3100
                    GoTo ErrorHandler
                End If
            End If
        End If
    Else
        If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.SELIC Or _
            xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
            If Not xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing And _
                Not xmlOperacao.documentElement.selectSingleNode("DT_MESG") Is Nothing Then
                If xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text < fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                    If xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text <> fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) Then
                        '3100 - Data da Operação Inválida.
                        lngCodigoErroNegocio = 3100
                        GoTo ErrorHandler
                    End If
                End If
            End If
        Else
        End If
    End If


    Select Case penumFuncionalidadeFuncaoSist

        Case enumFuncionalidade.Confirmacao
            flGerenciarChamadas xmlOperacao, "Confirmacao"

        Case enumFuncionalidade.Discordancia
            objWorkflow.GerenciarChamadasEstorno xmlOperacao, "Discordancia"

        Case enumFuncionalidade.Liberacao
            If xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = "0" Then
                If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.SELIC Then
                    If pblnOrigemConciliacao = False Then
                        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                      enumStatusOperacao.Liberada, _
                                                      pstrDtHrUltimaAtlz
                    End If
                End If
                
                If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
                    If Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasBoleto Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasTributos Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta Or _
                    Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib Then
                        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                      enumStatusOperacao.Liquidada, _
                                                      pstrDtHrUltimaAtlz
                          
                        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
                    End If
                 End If
                
                If Not xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ") Is Nothing And _
                   pstrDtHrUltimaAtlz <> vbNullString Then
                    xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text = pstrDtHrUltimaAtlz
                End If

                flLiberarOperacao xmlOperacao, True, pblnIgnoraGradeHorario, pblnOrigemConciliacao, xmlRepeticao, pblnFluxoAutomatico
            Else
                If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.SELIC Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
                    flLiberarCancelamentoEstorno xmlOperacao, pblnIgnoraGradeHorario
                Else
                    flLiberarOperacao xmlOperacao, True, pblnIgnoraGradeHorario, pblnOrigemConciliacao, xmlRepeticao
                End If
            End If

        Case enumFuncionalidade.Conciliacao
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.SELIC Then
                objWorkflow.GerenciarChamadas xmlOperacao, "Conciliacao", xmlRepeticao
            Else
                If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DespesaSelic Then
                    flLiquidarDespesa xmlOperacao
                ElseIf CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
                    LiquidarComprVoltaConciliacao xmlOperacao
                Else
                    LiquidarTermoVoltaEventos xmlOperacao, xmlRepeticao
                End If
            End If

        Case enumFuncionalidade.ConcordanciaPrevia
            objWorkflow.GerenciarChamadas xmlOperacao, "ConcordanciaPrevia"

        Case enumFuncionalidade.Concordancia
            objWorkflow.GerenciarChamadas xmlOperacao, "Concordancia"
            
            If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD Then
                If Not xmlRepeticao Is Nothing Then
                    Call xmlRepeticao.loadXML(xmlOperacao.xml)
                End If
            End If
        Case enumFuncionalidade.ConcordanciaBMC0101
            objWorkflow.GerenciarChamadas xmlOperacao, "ConcordanciaBMC0101"

        Case enumFuncionalidade.ConcordanciaBMC0103
            objWorkflow.GerenciarChamadas xmlOperacao, "ConcordanciaBMC0103"

        Case enumFuncionalidade.ConcordanciaLDL0005R2
            objWorkflow.GerenciarChamadas xmlOperacao, "ConcordanciaLDL0005R2"

        Case enumFuncionalidade.RejeicaoConcordanciaPrevia
            objWorkflow.GerenciarChamadasEstorno xmlOperacao, "RejeicaoConcordanciaPrevia"

        Case enumFuncionalidade.RejeicaoConcordancia
            objWorkflow.GerenciarChamadasEstorno xmlOperacao, "RejeicaoConcordancia"

        Case enumFuncionalidade.LiberacaoPrevia
            objWorkflow.GerenciarChamadas xmlOperacao, "LiberacaoPrevia"

        Case enumFuncionalidade.LiberacaoRecebimento
            objWorkflow.GerenciarChamadas xmlOperacao, "LiberacaoRecebimento"

        Case enumFuncionalidade.RegularizarContingencia
            objWorkflow.GerenciarChamadas xmlOperacao, "RegularizarContingencia"

        Case enumFuncionalidade.LiberacaoPagamento
            objWorkflow.GerenciarChamadas xmlOperacao, "LiberacaoPagamento"

        Case enumFuncionalidade.PagamentoContingencia
            objWorkflow.GerenciarChamadas xmlOperacao, "PagamentoContingencia"

    End Select

    Set objWorkflow = Nothing
    Set objOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlTipoOperacao = Nothing

    If Not objMQAX200 Is Nothing Then
        objMQAX200.MQEnd
    End If
    Set objMQAX200 = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set xmlOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set xmlMesgOrig = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamadaViaTela Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar os fluxos alternativo de operações
'-------------------------------------------------------------------------------------------------

Private Function flFluxosAlternativos(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

    On Error GoTo ErrorHandler

    flFluxosAlternativos = True

    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
        Case enumTipoOperacaoLQS.CompromissadaIda
            flFluxosAlternativos = flAlternativoCompromissadaIda(xmlRemessa)
        Case enumTipoOperacaoLQS.CompromissadaVolta
            flFluxosAlternativos = flAlternativoRetornoCompromissada(xmlRemessa)
    End Select

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    flFluxosAlternativos = False

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flFluxosAlternativos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de exeção da operação de Termo na data da liquidação, Eventos financeiros e despesas selic
'-------------------------------------------------------------------------------------------------

Private Function flTermoEventosDespesas(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                        ByRef plngStatusOperacao As Long, _
                                        ByRef pvntSequenciaOperacao As Variant, _
                                        ByRef plngTipoOperacao As Long) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    flTermoEventosDespesas = True

    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
        Case enumTipoMensagemBUS.DespesasSelic, enumTipoMensagemBUS.EventosSelic
            flTermoEventosDespesas = flExcecaoDespesasSelic(xmlRemessa, _
                                                            plngStatusOperacao, _
                                                            pvntSequenciaOperacao, _
                                                            plngTipoOperacao)
        Case enumTipoMensagemBUS.TermoLiquidacao
            flTermoEventosDespesas = flAlternativoTermoLiquidacao(xmlRemessa, _
                                                                  plngStatusOperacao, _
                                                                  pvntSequenciaOperacao, _
                                                                  plngTipoOperacao)
        Case enumTipoMensagemBUS.DespesasCETIP
            flTermoEventosDespesas = flExcecaoDespesasCETIP(xmlRemessa, _
                                                            plngStatusOperacao, _
                                                            pvntSequenciaOperacao, _
                                                            plngTipoOperacao)
    End Select

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTermoEventosDespesas Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo alternativo da operação Compromissada de Ida
'-------------------------------------------------------------------------------------------------

Private Function flAlternativoCompromissadaIda(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Regra para complementação manual
    'Rever se teste será com "0,00000000" ou com ""
    If fgVlrToDBServer(xmlRemessa.documentElement.selectSingleNode("VA_PU_RETN").Text) = 0 Then
         If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
            If CLng("0" & xmlRemessa.documentElement.selectSingleNode("TP_CPRO_OPER_ATIV").Text) = 1 Then
                ' 3036 - Quando tipo do compromisso for 1 - Compra com PU de Retorno definido, o de PU de Retorno deve ser preenchido.
                lngCodigoErroNegocio = 3036
                GoTo ErrorHandler
            End If
         End If
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AComplementar
        flAlternativoCompromissadaIda = False
        Exit Function
    End If

    flAlternativoCompromissadaIda = True

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoCompromissadaIda Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo alternativo da operação Compromissada de Volta
'-------------------------------------------------------------------------------------------------

Private Function flAlternativoRetornoCompromissada(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String
Dim objOperacao                             As A8LQS.clsOperacao
Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao
Dim xmlOperacaoIda                          As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlOperacaoIda = CreateObject("MSXML2.DOMDocument.4.0")
    Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text Or _
        xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text = vbNullString Or _
        xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text = "0" Then
        strXML = objOperacao.LerTodos(, , , CLng("0" & xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text), _
                                      fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_ORIG").Text), _
                                      enumTipoOperacaoLQS.CompromissadaIda, _
                                      , , , _
                                      xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
    Else
        strXML = objOperacao.LerTodos(, , CLng("0" & xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER_RETN").Text), , _
                                      fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_ORIG").Text), _
                                      enumTipoOperacaoLQS.CompromissadaIda, _
                                      , , , _
                                      xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
    End If

    If strXML = vbNullString Then
        
        If Not xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
            If Val(xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ManualEmSer Then
                flAlternativoRetornoCompromissada = False
                Exit Function
            End If
        End If
            
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.CompromissadaIda
        flAlternativoRetornoCompromissada = False
        Exit Function
        
    End If

    xmlOperacaoIda.loadXML strXML

    'Regra para complementação manual
    strXML = objHistoricoOperacao.LerTodos(xmlOperacaoIda.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.AComplementar)

    If strXML <> vbNullString Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AComplementar
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    'Carlos 01/06/2004
    'se a operação de ida não está liquidada para por inconsistencia
    If xmlOperacaoIda.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then 'And _
        'xmlOperacaoIda.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.LiquidadaCliente1 Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.OperacaoNaoLiquidada
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    'Consistir Operação de Ida com a Volta
    If xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.VeicluloLegal
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    If Val(xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CNPT").Text) <> Val(xmlOperacaoIda.documentElement.selectSingleNode("//CO_CNPJ_CNPT").Text) Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.CNPJContraparte
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//NU_ATIV_MERC").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.IdentificadorTitulo
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//DT_VENC_ATIV").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.DataVencimentoTitulo
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If

    flAlternativoRetornoCompromissada = flLeilaoInformalCompromissada(xmlRemessa)

    Set objOperacao = Nothing
    Set objHistoricoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objHistoricoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoRetornoCompromissada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo alternativo da operação Leilão Compromissada
'-------------------------------------------------------------------------------------------------

Private Function flLeilaoInformalCompromissada(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacaoIda                          As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set xmlOperacaoIda = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    strXML = objOperacao.LerTodos(, , CLng("0" & xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER_RETN").Text), , _
                                  fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_ORIG").Text), _
                                  enumTipoOperacaoLQS.LeilaoInformalCompromissada, _
                                  , , , _
                                  xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)

    If strXML = vbNullString Then
        flLeilaoInformalCompromissada = True
        Exit Function
    End If

    xmlOperacaoIda.loadXML strXML

    'Carlos 01/06/2004
    'se a operação de ida não está liquidada para por inconsistencia
    If xmlOperacaoIda.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then 'And _
        'xmlOperacaoIda.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.LiquidadaCliente1 Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.OperacaoNaoLiquidada
        flLeilaoInformalCompromissada = False
        Exit Function
    End If

    'Consistir Operação de Ida com a Volta
    If xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.VeicluloLegal
        flLeilaoInformalCompromissada = False
        Exit Function
    End If

    If Val(xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CNPT").Text) <> Val(xmlOperacaoIda.documentElement.selectSingleNode("//CO_CNPJ_CNPT").Text) Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.CNPJContraparte
        flLeilaoInformalCompromissada = False
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//NU_ATIV_MERC").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.IdentificadorTitulo
        flLeilaoInformalCompromissada = False
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text <> xmlOperacaoIda.documentElement.selectSingleNode("//DT_VENC_ATIV").Text Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Inconsistencia, _
                                            enumJustificativa.DataVencimentoTitulo
        flLeilaoInformalCompromissada = False
        Exit Function
    End If

    flLeilaoInformalCompromissada = True

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLeilaoInformalCompromissada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo na liberação de uma operação
'-------------------------------------------------------------------------------------------------

Private Function flLiberarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                   ByVal pblnLiberarViaTela As Boolean, _
                          Optional ByVal pblnIgnoraGradeHorario As Boolean = False, _
                          Optional ByVal pblnOrigemConciliacao As Boolean = False, _
                          Optional ByRef xmlRepeticao As MSXML2.DOMDocument40, _
                          Optional ByRef pblnFluxoAutomatico As Boolean = False) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objMensagem                             As A8LQS.clsMensagem
Dim objHistSituacaoOperacao                 As A8LQS.clsHistSituacaoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objAlerta                               As A8LQS.clsAlerta
Dim objWorkflow                             As A8LQS.clsWorkflow

Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim blnEnviarMensagem                       As Boolean
Dim blnCliente1                             As Boolean
Dim lngTipoLiquidacao                       As Long
Dim lngCodigoRetornoVerificacao             As Long

Dim blnEnviarInformacaoPJ                   As Boolean

    On Error GoTo ErrorHandler

    'Tratar câmaras diferentes de SELIC
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) <> enumLocalLiquidacao.SELIC Then
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        objWorkflow.LiberarOperacao xmlRemessa, _
                                    pblnLiberarViaTela, _
                                    pblnIgnoraGradeHorario, _
                                    pblnOrigemConciliacao, _
                                    xmlRepeticao, _
                                    pblnFluxoAutomatico
        Set objWorkflow = Nothing
        Exit Function
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    'Esta rotina atende a liberação automática e a liberação via Tela
    If Not pblnLiberarViaTela Then
        'Libera o Registro para Consulta. Está neste ponto pois a transação pode se encerrar em vários pontos.
        objOperacao.OperacaoDisponivelConsulta xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               enumIndicadorSimNao.Sim

        'Verificar a Grade de Horário ..... Carlos - 30/09/2003
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        Set objMensagem = CreateObject("A8LQS.clsMensagem")
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                                                  objMensagem.ObterSequencialISPBDestino(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text)) Then
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         enumJustificativa.GradeHorario
            Set objHistSituacaoOperacao = Nothing
            Set objAlerta = CreateObject("A8LQS.clsAlerta")
            objAlerta.GerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.MensagensBloqueadasGradeHorario
            Set objAlerta = Nothing
            Exit Function
        End If
        Set objMensagem = Nothing
        Set objGradeHorario = Nothing

        Set objRegra = CreateObject("A8LQS.clsRegra")
        'Verifica regra para liberação automática
        If Not objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Liberar, lngCodigoRetornoVerificacao) Then
            Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
            objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                         xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                         IIf(lngCodigoRetornoVerificacao <> 0, _
                                                                lngCodigoRetornoVerificacao, _
                                                                enumJustificativa.RegraWorkflow)
            Set objHistSituacaoOperacao = Nothing
            Exit Function
        End If
        Set objRegra = Nothing
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.LiberadaAutomatica
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica
    Else
        If Not pblnIgnoraGradeHorario Then
            'Verificar a Grade de Horário ..... Carlos - 30/09/2003
            Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
            Set objMensagem = CreateObject("A8LQS.clsMensagem")
            If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                                                      objMensagem.ObterSequencialISPBDestino(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text)) Then
                If Not pblnOrigemConciliacao Then
                    ' 3023 - Horário limite para envio da Operação excedido.
                    lngCodigoErroNegocio = 3023
                    GoTo ErrorHandler
                Else
                    Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                    objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                 xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                                 enumJustificativa.GradeHorario
                    Set objHistSituacaoOperacao = Nothing
                    
                    Set objGradeHorario = Nothing
                    Set objMensagem = Nothing
                    
                    Exit Function
                End If
            End If
            Set objGradeHorario = Nothing
            Set objMensagem = Nothing
        End If

        If pblnOrigemConciliacao Then
            Set objRegra = CreateObject("A8LQS.clsRegra")
            'Verifica regra para liberação automática
            If Not objRegra.VerificarRegraAutomatica(xmlRemessa, enumFuncaoSistema.Liberar, lngCodigoRetornoVerificacao) Then
                Set objHistSituacaoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
                objHistSituacaoOperacao.AlterarJustificativa xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                             xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                                             IIf(lngCodigoRetornoVerificacao <> 0, _
                                                                    lngCodigoRetornoVerificacao, _
                                                                    enumJustificativa.RegraWorkflow)
                Set objHistSituacaoOperacao = Nothing
                Exit Function
            End If
            Set objRegra = Nothing
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.LiberadaAutomatica
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaAutomatica
        End If
    End If

    'Carlos - Gerenciamento de chamada das rotinas
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    'If Not xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '    lngTipoLiquidacao = CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    'End If
    strXML = LerTodos("Liberacao", _
                      CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                      enumIndicadorSimNao.Nao, , _
                      lngTipoLiquidacao)
    xmlControleProcessoOperacao.loadXML strXML

    If strXML = "" Then
        Exit Function
    End If

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlRemessa.getElementsByTagName("IN_ENTR_MANU").Item(0) Is Nothing Then
        If Not xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
        If xmlRemessa.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
            If xmlRemessa.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
            End If
        End If
        End If
    End If

    If fgVerificarCliente1(xmlRemessa) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        If blnCliente1 Then
            fgDebitoCreditoCliente1 xmlRemessa
        End If
        
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        
        blnEnviarInformacaoPJ = True
        If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val(xmlRemessa.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.VinculoDesvinculoTransferencia Then
                If Not xmlRemessa.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                    If Trim$(xmlRemessa.selectSingleNode("//SG_SIST_ORIG").Text) = "YS" Then
                        blnEnviarInformacaoPJ = False
                    End If
                End If
            End If
        End If
        
        If blnEnviarInformacaoPJ Then
            strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlRemessa)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        End If
        
        If blnCliente1 Then
            'irá retornar o debito/credito original
            fgDebitoCreditoCliente1 xmlRemessa
            fgConverterProduto1 xmlRemessa
            strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlRemessa)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgConverterProduto1 xmlRemessa
        End If
        Set objSubReserva = Nothing
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_SPB").Text = enumIndicadorSimNao.Sim Then
            'Verificar compromissada retroativa de volta, não enviar mensagem - Carlos 04/06/2004
        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVolta Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
            If VerificarCompromissadaRetroativa(xmlRemessa) Then
                flLiquidarCompromissada xmlRemessa
                Exit Function
            End If
        End If
        'comentado por Bruno Oliveira 19/07/2010 devido reestruturação das contas Selic
'        If fgVerificarCliente1(xmlRemessa, True) Then
'            flLiquidarCliente1 xmlRemessa
'        Else
            Set objMensagem = CreateObject("A8LQS.clsMensagem")
            strMensagem = objMensagem.MontarMensagem(xmlRemessa, blnEnviarMensagem)
            If blnEnviarMensagem Then
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, True
            End If
            Set objMensagem = Nothing
'        End If
    End If

    Set xmlControleProcessoOperacao = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set xmlControleProcessoOperacao = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo na liberação de um cancelamento ou estorno de uma operação
'-------------------------------------------------------------------------------------------------

Private Function flLiberarCancelamentoEstorno(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                     Optional ByVal pblnIgnoraGradeHorario As Boolean = False) As Boolean

Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagem
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim lngTipoAcao                             As Long
Dim blnEnviarMensagem                       As Boolean

    On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) = enumTipoAcao.CancelamentoSolicitado Then
        If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.SELIC Then
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.Pendencia And _
                CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.EmLancamento Then
                ' 3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
            End If
            lngTipoAcao = enumTipoAcao.CancelamentoEnviado
        ElseIf xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
            
            lngTipoAcao = enumTipoAcao.CancelamentoEnviado
            
            'Pikachu - 20/07/2004
            'O Status das Solicitações de cancelamento CETIP não estavam sendo atualizadas pata Liberada
            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada
            
        End If
    ElseIf CLng(xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) = enumTipoAcao.EstornoSolicitado Then
        lngTipoAcao = enumTipoAcao.EstornoEnviado
    Else
        ' 3027 - Tipo da ação da operação inválido.
        lngCodigoErroNegocio = 3027
        GoTo ErrorHandler
    End If

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    strMensagem = objTipoOperacao.LerTipoOperacaoAcao(CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text), _
                                                      CLng(xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text))

    If strMensagem = vbNullString Then
        ' 3026 - Não é permitido estornar ou cancelar este tipo de operação.
        lngCodigoErroNegocio = 3026
        GoTo ErrorHandler
    End If
    xmlTipoOperacao.loadXML strMensagem

    xmlOperacao.documentElement.selectSingleNode("CO_MESG").Text = xmlTipoOperacao.documentElement.selectSingleNode("CO_MESG_SPB_REGT_ACAO").Text

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    If Not pblnIgnoraGradeHorario Then
        'Verificar a Grade de Horário ..... Carlos - 30/09/2003
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlOperacao.documentElement.selectSingleNode("CO_MESG").Text, _
                                                                  objMensagem.ObterSequencialISPBDestino(xmlOperacao.documentElement.selectSingleNode("CO_MESG").Text)) Then
            ' 3023 - Horário limite para envio da Operação excedido.
            lngCodigoErroNegocio = 3023
            GoTo ErrorHandler
        End If
        Set objGradeHorario = Nothing
    End If
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                  xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                  xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                  , lngTipoAcao

    strMensagem = objMensagem.MontarMensagem(xmlOperacao, blnEnviarMensagem)
    If blnEnviarMensagem Then
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, True
    End If

    Set objOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set objMensagem = Nothing
    Set xmlTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objTipoOperacao = Nothing
    Set objMensagem = Nothing
    Set objGradeHorario = Nothing
    Set xmlTipoOperacao = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberarCancelamentoEstorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo na liquidação de operações de Termo na data da liquidação

Public Function LiquidarTermoVoltaEventos(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                 Optional ByRef xmlConciliacao As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLegado                               As A8LQS.clsLegado
Dim xmlOperacaoIda                          As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean
Dim blnLiquidadaFisicamente                 As Boolean

    On Error GoTo ErrorHandler

    If fgVerificarCliente1(xmlRemessa) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Then
        If blnCliente1 Then
            fgDebitoCreditoCliente1 xmlRemessa
        End If
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        If blnCliente1 Then
            'irá retornar o debito/credito original
            fgDebitoCreditoCliente1 xmlRemessa
            fgConverterProduto1 xmlRemessa
            strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            fgConverterProduto1 xmlRemessa
        End If
        Set objGestaoCaixa = Nothing
    Else
        If Not xmlConciliacao.selectSingleNode("//MESG") Is Nothing Then
            fgAppendNode xmlConciliacao, "MESG", "CO_VEIC_LEGA", xmlRemessa.selectSingleNode("//CO_VEIC_LEGA").Text
        End If
        LiquidarEventos xmlConciliacao
        If Not xmlConciliacao.selectSingleNode("//AlteraStatusOperacao") Is Nothing Then
            If CLng(xmlConciliacao.selectSingleNode("//AlteraStatusOperacao").Text) <> enumIndicadorSimNao.Sim Then
                Exit Function
            End If
        End If
    End If

    If blnCliente1 Then
        fgDebitoCreditoCliente1 xmlRemessa
    End If
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
    If blnCliente1 Then
        'irá retornar o debito/credito original
        fgDebitoCreditoCliente1 xmlRemessa
        fgConverterProduto1 xmlRemessa
        strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        fgConverterProduto1 xmlRemessa
    End If
    Set objSubReserva = Nothing

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Then
        Set xmlOperacaoIda = CreateObject("MSXML2.DOMDocument.4.0")
        'Ler o Termo Ida para alterar o status LIQUIDADA
        strXML = objOperacao.LerTodos(, , xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                      , , enumTipoOperacaoLQS.Termo, , _
                                      xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text, _
                                      xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
        If strXML <> vbNullString Then
            xmlOperacaoIda.loadXML (strXML)
            objOperacao.AlterarStatusOperIndisp xmlOperacaoIda.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.Liquidada
        Set xmlOperacaoIda = Nothing
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Liquidada
    Else
        If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.Liquidada
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", enumStatusOperacao.Liquidada
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
    End If

    Set objOperacao = Nothing

    '*******************************************************
    'Disponibilizar lançamento conta corrente - Carlos 12/09/2003
    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.Disponibilizar xmlRemessa
    Set objContaCorrente = Nothing
    '*******************************************************

    'Enviar Retorno para o sistema legado.
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    Set objLegado = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarTermoVoltaEventos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo na liquidação de eventos financeiros
'-------------------------------------------------------------------------------------------------

Public Function LiquidarEventos(ByRef xmlConciliacao As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLegado                               As A8LQS.clsLegado
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strXML                                  As String
Dim strMensagem                             As String
Dim strVeiculoLegal                         As String

    On Error GoTo ErrorHandler

    'Obter o Produto, qdo tiver operação - Carlos 13/07/04
    If Not xmlConciliacao.documentElement.selectSingleNode("//QuantidadeItemSelOperacao") Is Nothing Then
        If Val("0" & xmlConciliacao.documentElement.selectSingleNode("//QuantidadeItemSelOperacao").Text) > 0 Then
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlConciliacao.selectSingleNode("//Operacao").Text)
            Set objOperacao = Nothing
            If xmlConciliacao.documentElement.selectSingleNode("//CO_PROD") Is Nothing Then
                fgAppendNode xmlConciliacao, "Repeat_Conciliacao", "CO_PROD", xmlOperacao.selectSingleNode("//CO_PROD").Text
            Else
                xmlConciliacao.documentElement.selectSingleNode("//CO_PROD").Text = xmlOperacao.selectSingleNode("//CO_PROD").Text
            End If
            If xmlConciliacao.documentElement.selectSingleNode("//SG_SIST") Is Nothing Then
                fgAppendNode xmlConciliacao, "Repeat_Conciliacao", "SG_SIST", xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
            Else
                xmlConciliacao.documentElement.selectSingleNode("//SG_SIST").Text = xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text
            End If
        Else
            If xmlConciliacao.documentElement.selectSingleNode("//SG_SIST") Is Nothing Then
                fgAppendNode xmlConciliacao, "Repeat_Conciliacao", "SG_SIST", "A8"
            Else
                xmlConciliacao.documentElement.selectSingleNode("//SG_SIST").Text = "A8"
            End If
        End If
    Else
        If Not xmlConciliacao.selectSingleNode("//MESG") Is Nothing Then
            If xmlConciliacao.documentElement.selectSingleNode("//SG_SIST") Is Nothing Then
                fgAppendNode xmlConciliacao, "MESG", "SG_SIST", xmlConciliacao.selectSingleNode("//SG_SIST_ORIG").Text
            Else
                xmlConciliacao.documentElement.selectSingleNode("//SG_SIST").Text = xmlConciliacao.selectSingleNode("//SG_SIST_ORIG").Text
            End If
        Else
            If xmlConciliacao.documentElement.selectSingleNode("//SG_SIST") Is Nothing Then
                fgAppendNode xmlConciliacao, "Repeat_Conciliacao", "SG_SIST", "A8"
            Else
                xmlConciliacao.documentElement.selectSingleNode("//SG_SIST").Text = "A8"
            End If
        End If
    End If

    If Not xmlConciliacao.documentElement.selectSingleNode("//PrimeiroEnvio") Is Nothing Then
        If xmlConciliacao.documentElement.selectSingleNode("//PrimeiroEnvio").Text = enumIndicadorSimNao.Nao Then
            Exit Function
        Else
            xmlConciliacao.documentElement.selectSingleNode("//PrimeiroEnvio").Text = enumIndicadorSimNao.Nao
        End If
    End If

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    If Not xmlConciliacao.selectSingleNode("//MESG") Is Nothing Then
        strVeiculoLegal = xmlConciliacao.selectSingleNode("//CO_VEIC_LEGA").Text
    Else
        strVeiculoLegal = vbNullString
    End If

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    If Not xmlConciliacao.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Mensagem") Is Nothing Then
        For Each objDomNode In xmlConciliacao.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Mensagem")

            Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(objDomNode.selectSingleNode("NumeroControleIF").Text)

            If CLng(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.Conciliada Then
                If Not xmlConciliacao.documentElement.selectSingleNode("//CO_PROD") Is Nothing Then
                    fgAppendNode xmlMensagem, "SISMSG", "CO_PROD", xmlConciliacao.selectSingleNode("//CO_PROD").Text
                End If
    
                xmlMensagem.documentElement.selectSingleNode("//SG_SIST").Text = xmlConciliacao.selectSingleNode("//SG_SIST").Text
    
                strMensagem = objGestaoCaixa.EnviarRealizadoEventos(xmlMensagem)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            End If

        Next
    Else

        If Not xmlConciliacao.selectSingleNode("//NumeroControleIF") Is Nothing Then
            Set xmlMensagem = objMensagemEnviadaRecebida.LerXMLMensagem(xmlConciliacao.selectSingleNode("//NumeroControleIF").Text)

            If CLng(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusMensagem.Conciliada Then
                If strVeiculoLegal <> vbNullString Then
                    xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text = strVeiculoLegal
                End If
    
                If Not xmlConciliacao.documentElement.selectSingleNode("//CO_PROD") Is Nothing Then
                    fgAppendNode xmlMensagem, "SISMSG", "CO_PROD", xmlConciliacao.selectSingleNode("//CO_PROD").Text
                End If
    
                xmlMensagem.documentElement.selectSingleNode("//SG_SIST").Text = xmlConciliacao.selectSingleNode("//SG_SIST").Text
    
                strMensagem = objGestaoCaixa.EnviarRealizadoEventos(xmlMensagem)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            End If
        End If
    End If
    Set objGestaoCaixa = Nothing

    If xmlConciliacao.documentElement.selectSingleNode("//QuantidadeItemSelOperacao") Is Nothing Then
        Exit Function
    End If

    If Val(xmlConciliacao.documentElement.selectSingleNode("//QuantidadeItemSelOperacao").Text) = 0 Then
        Exit Function
    End If

    If Not xmlConciliacao.documentElement.selectSingleNode("//RetornoLegado") Is Nothing Then
        If xmlConciliacao.documentElement.selectSingleNode("//RetornoLegado").Text = enumIndicadorSimNao.Nao Then
            Exit Function
        End If
    Else
        Exit Function
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objLegado = CreateObject("A8LQS.clsLegado")

    If Not xmlConciliacao.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Operacao") Is Nothing Then
        For Each objDomNode In xmlConciliacao.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Operacao")

            Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNode.selectSingleNode("Operacao").Text)

            strMensagem = objLegado.EnviarRetorno(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS

        Next
    Else
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlConciliacao.selectSingleNode("//Operacao").Text)

        strMensagem = objLegado.EnviarRetorno(xmlOperacao)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    End If

    Set objLegado = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarEventos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo na liquidação de operações de compromissada de volta com conciliacao
'-------------------------------------------------------------------------------------------------

Public Function LiquidarComprVoltaConciliacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                     Optional ByVal pblnConciliacaoTela As Boolean = False) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLegado                               As A8LQS.clsLegado
Dim xmlOperacaoIda                          As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    If fgVerificarCliente1(xmlRemessa) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    If blnCliente1 Then
        fgDebitoCreditoCliente1 xmlRemessa
    End If
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    If blnCliente1 Then
        'irá retornar o debito/credito original
        fgDebitoCreditoCliente1 xmlRemessa
        fgConverterProduto1 xmlRemessa
        strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        fgConverterProduto1 xmlRemessa
    End If
    Set objGestaoCaixa = Nothing

    If blnCliente1 Then
        fgDebitoCreditoCliente1 xmlRemessa
    End If
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
    If blnCliente1 Then
        'irá retornar o debito/credito original
        fgDebitoCreditoCliente1 xmlRemessa
        fgConverterProduto1 xmlRemessa
        strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        fgConverterProduto1 xmlRemessa
    End If
    Set objSubReserva = Nothing

    '*******************************************************
    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.Disponibilizar xmlRemessa
    Set objContaCorrente = Nothing
    '*******************************************************

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", enumStatusOperacao.Liquidada
        'fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", enumStatusOperacao.LiquidadaFisicamente
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
        'xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiquidadaFisicamente
    End If

    Set objOperacao = Nothing

    'Enviar Retorno para o sistema legado.
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    Set objLegado = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarComprVoltaConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o envio de mensage para o MQSeries
'-------------------------------------------------------------------------------------------------

Public Function EnviarMensagemMQ(ByRef xmlRemessa As String, _
                                 ByVal plngIdentificadorFila As enumIdentificadorFila, _
                        Optional ByVal pblnEncerrar As Boolean = False, _
                        Optional ByVal pblnMensagemSPB As Boolean = False, _
                        Optional ByVal psHeaderNZ As String) As Long

Dim lngBufferLen                            As Long
Dim strNomeFila                             As String
Dim lngPrioridade                           As Long
Dim dblLog                                  As Double

    On Error GoTo ErrorHandler
    
    fgLogText "EnviMesnMQ start, " & plngIdentificadorFila & " " & xmlRemessa, dblLog

    Select Case plngIdentificadorFila
        Case enumIdentificadorFila.Liberacao
            strNomeFila = "A8Q.E.LIBERACAO"
        Case enumIdentificadorFila.SubReserva
            strNomeFila = "A6Q.E.REMESSASUBRESERVA"
        Case enumIdentificadorFila.BUS
            
            If pblnMensagemSPB Then
                lngPrioridade = 9
            Else
                lngPrioridade = 0
            End If
            
            'Fazer put para A7NET nas mensagens novas da CAM
            If (Left(Trim$(UCase$(psHeaderNZ)), 3) = "CAM" _
                And (Left(Trim$(UCase$(psHeaderNZ)), 7) <> "CAM0001" _
                And Left(Trim$(UCase$(psHeaderNZ)), 7) <> "CAM0002" _
                And Left(Trim$(UCase$(psHeaderNZ)), 7) <> "CAM0003" _
                And Left(Trim$(UCase$(psHeaderNZ)), 7) <> "CAM0004")) _
            Or _
               (Left(Trim$(UCase$(xmlRemessa)), 3) = "CAM" _
                And (Left(Trim$(UCase$(xmlRemessa)), 7) <> "CAM0001" _
                And Left(Trim$(UCase$(xmlRemessa)), 7) <> "CAM0002" _
                And Left(Trim$(UCase$(xmlRemessa)), 7) <> "CAM0003" _
                And Left(Trim$(UCase$(xmlRemessa)), 7) <> "CAM0004")) Then
                strNomeFila = "A7Q.E.ENTRADA_NET"
            Else
                strNomeFila = "A7Q.E.ENTRADA"
            End If
    End Select

    lngBufferLen = Len(xmlRemessa)

    If lngBufferLen = 0 Then
        Exit Function
    End If

    If objMQAX200 Is Nothing Then
        Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    End If

    objMQAX200.MQInit
    objMQAX200.MQOpen strNomeFila, MQOO_OUTPUT
    objMQAX200.MQPut xmlRemessa, lngBufferLen, enumPutOptions.Binario, , , , , , lngPrioridade
    objMQAX200.MQClose

    'If pblnEncerrar Then
        objMQAX200.MQEnd
    'End If
    
    fgLogText "EnviMesnMQ end", dblLog

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarMensagemMQ", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de exeção das despesas selic
'-------------------------------------------------------------------------------------------------

Private Function flExcecaoDespesasSelic(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                        ByRef plngStatusOperacao As Long, _
                                        ByRef pvntSequenciaOperacao As Variant, _
                                        ByRef plngTipoOperacao As Long) As Boolean

Dim strXML                                  As String
Dim xmlDespesas                             As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim objRemessaRejeitada                     As A8LQS.clsRemessaRejeitada
Dim objLegado                               As A8LQS.clsLegado

    On Error GoTo ErrorHandler

    Set xmlDespesas = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    strXML = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                  xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If strXML = vbNullString Then
        
        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
            plngStatusOperacao = enumStatusOperacao.ManualEmSer
        Else
            plngStatusOperacao = enumStatusOperacao.EmSer
        End If
        
        flExcecaoDespesasSelic = True
        Exit Function
    End If

    xmlDespesas.loadXML strXML

    If CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Or _
        CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Concordancia Or _
        CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaAutomatica Or _
        CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.AConciliar Then
        plngStatusOperacao = CLng("0" & xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text)
        pvntSequenciaOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text
        plngTipoOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/TP_OPER").Text
        flExcecaoDespesasSelic = True
    Else
        '3007 - Solicitação de Complementação Inválida para Eventos/Despesas Selic.
        lngCodigoErroNegocio = 3007
        GoTo ErrorHandler
    End If

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcecaoDespesasSelic Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo alternativo das operações de Termo na data da liquidação

Private Function flAlternativoTermoLiquidacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                              ByRef plngStatusOperacao As Long, _
                                              ByRef pvntSequenciaOperacao As Variant, _
                                              ByRef plngTipoOperacao As Long) As Boolean

Dim strXML                                  As String
Dim strMensagem                             As String
Dim objOperacao                             As A8LQS.clsOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objRemessaRejeitada                     As A8LQS.clsRemessaRejeitada
Dim objLegado                               As A8LQS.clsLegado
Dim xmlDespesas                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    Set xmlDespesas = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    strXML = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                  xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If strXML = vbNullString Or _
       Trim$(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text) = vbNullString Then
        
        plngStatusOperacao = enumStatusOperacao.AConciliar
        flAlternativoTermoLiquidacao = True
        Exit Function
    
    End If

    xmlDespesas.loadXML strXML

    If CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.AConciliar Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text)
        plngStatusOperacao = enumStatusOperacao.AConciliar
        pvntSequenciaOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text
        plngTipoOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/TP_OPER").Text
        flAlternativoTermoLiquidacao = True
    Else
        '3008 - Status Inválido para Complementação do Termo na Data da Liquidação.
        lngCodigoErroNegocio = 3008
        GoTo ErrorHandler
    End If

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoTermoLiquidacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo alternativo das operações de Antecipação de Resgate de Contrato de Derivativo

Private Function flAntecipacaoResgateContratoDerivativo(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                                        ByVal pSiglaSistema As String, _
                                                        ByVal pblnEstorno As Boolean) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim strDataAux                              As String
Dim vntValorAux                             As Variant

    On Error GoTo ErrorHandler

    'Converte a data e valor para o envio da operação original
    strDataAux = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text
    xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text
    xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text = strDataAux
    'vntValorAux = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text
    'xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_ORIG").Text
    'xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_ORIG").Text = vntValorAux

    If pSiglaSistema = "PJ" Then
        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
        If pblnEstorno Then
            strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
            Call flEstornarRealPrevistoCETIP(xmlOperacao)
        Else
            strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
        End If
        Set objGestaoCaixa = Nothing
    End If

    If pSiglaSistema = "A6" Then
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        If pblnEstorno Then
            strMensagem = objSubReserva.EstornarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
        Else
            strMensagem = objSubReserva.EnviarPrevisao(xmlOperacao)
            EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
        End If
        Set objSubReserva = Nothing
    End If

    strDataAux = xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text
    xmlOperacao.documentElement.selectSingleNode("DT_VENC_OPER_ATIV_ORIG").Text = xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text
    xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataAux
    'vntValorAux = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_ORIG").Text
    'xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV_ORIG").Text = xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text
    'xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = vntValorAux

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAntecipacaoResgateContratoDerivativo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de operações
'-------------------------------------------------------------------------------------------------

Public Function GerenciarCancelamento(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                      ByVal pblnAlterarStatus As Boolean) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objRegra                                As A8LQS.clsRegra
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strMensagemME                           As String
Dim strXML                                  As String
Dim blnCliente1                             As Boolean
Dim lngTipoLiquidacao                       As Long

    On Error GoTo ErrorHandler

    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text <> enumLocalLiquidacao.SELIC Then
        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
        objWorkflow.GerenciarCancelamento xmlOperacao, pblnAlterarStatus
        Set objWorkflow = Nothing
        Exit Function
    End If

    'Altera Status quando vem o cancelamento via Legado
    If pblnAlterarStatus Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.CanceladaOrigem
        Set objOperacao = Nothing
    End If

    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
    'é enviada por remessa
    'If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
    '    lngTipoLiquidacao = CLng(xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)
    'End If
    strXML = LerTodos("RecebimentoCancelamento", _
                      CLng(xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                      enumIndicadorSimNao.Sim, , _
                      lngTipoLiquidacao)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao
    
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    
    If Not xmlOperacao.documentElement.selectSingleNode("//TP_OPER") Is Nothing Then
        If Val(xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesBMC _
        Or Val(xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao _
        Or Val(xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica Then
            
            strMensagemME = objGestaoCaixa.EstornarPrevisaoME(xmlOperacao)
            EnviarMensagemMQ strMensagemME, enumIdentificadorFila.BUS, False, False
            
        End If
    End If

    'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
    'é enviada por remessa
    If (xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliar Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Conciliada Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica) Then
'        xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaCliente1 Or
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
            If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
                If fgVerificarCliente1(xmlOperacao) Then
                    blnCliente1 = True
                    fgDebitoCreditoCliente1 xmlOperacao
                Else
                    blnCliente1 = False
                End If
                
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                
                If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras Then
                    fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Sim
                    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Nao
                End If
                If blnCliente1 Then
                    'irá retornar o debito/credito original
                    fgDebitoCreditoCliente1 xmlOperacao
                    fgConverterProduto1 xmlOperacao
                    strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                    fgConverterProduto1 xmlOperacao
                End If
            End If
        End If
    End If

    Set objGestaoCaixa = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarCancelamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Ler Todos os registro da tabela de controle conforme filtro
'-------------------------------------------------------------------------------------------------

Public Function LerTodos(Optional ByVal pstrFuncionalidade As String = "", _
                         Optional ByVal plngTipoOperacao As Long = 0, _
                         Optional ByVal plngEstorno As Long = 2, _
                         Optional ByVal plngStatusOperacao As Long = 0, _
                         Optional ByVal plngTipoLiquidacao As Long = 0) As String

Dim strSQL                                  As String
Dim blnWhere                                As Boolean

On Error GoTo ErrorHandler

    strSQL = " Select   TP_OPER, " & vbCrLf & _
             "          NO_PROC_OPER_ATIV, " & vbCrLf & _
             "          IN_ENVI_PREV_PJ, " & vbCrLf & _
             "          IN_ENVI_PREV_A6, " & vbCrLf & _
             "          IN_ENVI_RELZ_PJ, " & vbCrLf & _
             "          IN_ENVI_RELZ_SOLI_A6, " & vbCrLf & _
             "          IN_ENVI_RELZ_CONF_A6, " & vbCrLf & _
             "          IN_VERI_REGR_CONF, " & vbCrLf & _
             "          IN_VERI_REGR_CNCL, " & vbCrLf & _
             "          IN_VERI_REGR_LIBE, " & vbCrLf & _
             "          IN_ENVI_MESG_RETN, " & vbCrLf & _
             "          IN_ENVI_MESG_SPB, " & vbCrLf & _
             "          IN_DISP_LANC_CNTA_CRRT, " & vbCrLf & _
             "          IN_ESTO_PJ_A6, " & vbCrLf & _
             "          IN_ENVI_ALER, " & _
             "          CO_SITU_PROC, " & _
             "          TP_LIQU_OPER_ATIV " & _
             " From     A8.TB_CTRL_PROC_OPER_ATIV "

    If pstrFuncionalidade <> "" Then
        strSQL = strSQL & " Where    NO_PROC_OPER_ATIV  =   '" & pstrFuncionalidade & "'"
        blnWhere = True
    End If

    If plngTipoOperacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     TP_OPER    =   " & plngTipoOperacao
        Else
            strSQL = strSQL & " Where   TP_OPER    =   " & plngTipoOperacao
            blnWhere = True
        End If
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If blnWhere Then
            strSQL = strSQL & " And     IN_ESTO_PJ_A6    =   " & plngEstorno
        Else
            strSQL = strSQL & " Where   IN_ESTO_PJ_A6    =   " & plngEstorno
            blnWhere = True
        End If
    End If

    If plngStatusOperacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_SITU_PROC    =   " & plngStatusOperacao
        Else
            strSQL = strSQL & " Where   CO_SITU_PROC    =   " & plngStatusOperacao
            blnWhere = True
        End If
    End If

    If plngTipoLiquidacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     TP_LIQU_OPER_ATIV   =   " & plngTipoLiquidacao
        Else
            strSQL = strSQL & " Where   TP_LIQU_OPER_ATIV   =   " & plngTipoLiquidacao
            blnWhere = True
        End If
    End If

'--------
    strSQL = ""
    
    If pstrFuncionalidade <> "" Then
        strSQL = strSQL & " NO_PROC_OPER_ATIV = '" & pstrFuncionalidade & "'"
        blnWhere = True
    End If

    If plngTipoOperacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     TP_OPER    =   " & plngTipoOperacao
        Else
            strSQL = strSQL & " TP_OPER = " & plngTipoOperacao
            blnWhere = True
        End If
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If blnWhere Then
            strSQL = strSQL & " And     IN_ESTO_PJ_A6    =   " & plngEstorno
        Else
            strSQL = strSQL & "  IN_ESTO_PJ_A6 = " & plngEstorno
            blnWhere = True
        End If
    End If

    If plngStatusOperacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     CO_SITU_PROC    =   " & plngStatusOperacao
        Else
            strSQL = strSQL & " CO_SITU_PROC = " & plngStatusOperacao
            blnWhere = True
        End If
    End If

    If plngTipoLiquidacao > 0 Then
        If blnWhere Then
            strSQL = strSQL & " And     TP_LIQU_OPER_ATIV   =   " & plngTipoLiquidacao
        Else
            strSQL = strSQL & " TP_LIQU_OPER_ATIV = " & plngTipoLiquidacao
            blnWhere = True
        End If
    End If

'--------
'    If fgEstacaoDesenvolvimento Then
    strSQL = "select * from a8.tb_ctrl_proc_oper_ativ where " & strSQL
    LerTodos = fgQueryXMLLerTodos("ProcessoOperacao", strSQL, App.EXEName & "." & TypeName(Me))
'    Else
'        LerTodos = fgQueryXMLLerTodosCACHE("ProcessoOperacao", strSQL, App.EXEName & "." & TypeName(Me), "CONTROLE_PROCESSAMENTO")
'    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'-------------------------------------------------------------------------------------------------
'Verificar a situação do lançamento em conta corrente da operação
'-------------------------------------------------------------------------------------------------

Private Function flVerificarContaCorrente(ByVal pvntNumeroSequencial As Variant, _
                                 Optional ByVal penumStatusIntegracao As enumStatusIntegracao = enumStatusIntegracao.CanceladoOrigem) As Boolean

Dim objLancamentoContabil                   As A8LQS.clsLancamentoCCContabil

Dim strLancamentoContabil                   As String
Dim xmlLancamentoContabil                   As MSXML2.DOMDocument40

Dim strLancamentoEstorno                    As String
Dim xmlLancamentoEstorno                    As MSXML2.DOMDocument40

Dim intCanalVenda                           As Integer
Dim lngStatusLancamento                     As Long

    On Error GoTo ErrorHandler

    Set objLancamentoContabil = CreateObject("A8LQS.clsLancamentoCCContabil")

    strLancamentoContabil = objLancamentoContabil.Ler(pvntNumeroSequencial, _
                            enumTipoLancamentoIntegracao.Normal)
                            
    If strLancamentoContabil = vbNullString Then
        Exit Function
    End If

    Set xmlLancamentoContabil = CreateObject("MSXML2.DOMDocument.4.0")

    xmlLancamentoContabil.loadXML strLancamentoContabil
    
    lngStatusLancamento = 0
    If Not xmlLancamentoContabil.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        lngStatusLancamento = Val(xmlLancamentoContabil.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
    End If

    intCanalVenda = 0
    If Not xmlLancamentoContabil.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val(xmlLancamentoContabil.selectSingleNode("//TP_CNAL_VEND").Text)
    End If

    If lngStatusLancamento = enumStatusIntegracao.IntegradoCC Or _
       lngStatusLancamento = enumStatusIntegracao.ErroEstornoCC Or _
       lngStatusLancamento = enumStatusIntegracao.EnviadoCC Or _
       lngStatusLancamento = enumStatusIntegracao.EnviadoCCAutomatico Or _
       lngStatusLancamento = enumStatusIntegracao.EnviadoEstornoCC Or _
       lngStatusLancamento = enumStatusIntegracao.CanceladoOrigem Then
       
        If lngStatusLancamento <> enumStatusIntegracao.IntegradoCC Or intCanalVenda <> enumCanalDeVenda.SGC Then
            'Situação do lançamento em conta corrente não permite cancelamento
            lngCodigoErroNegocio = 3099
            GoTo ErrorHandler
        End If
    End If

    If lngStatusLancamento = enumStatusIntegracao.Integrado Then
        strLancamentoEstorno = objLancamentoContabil.Ler(pvntNumeroSequencial, _
                                                         enumTipoLancamentoIntegracao.Estorno)
                                                         
        If intCanalVenda <> enumCanalDeVenda.SGC Then
            If strLancamentoEstorno = vbNullString Then
                'Situação do lançamento em conta corrente não permite cancelamento
                lngCodigoErroNegocio = 3099
                GoTo ErrorHandler
                Exit Function
            End If
            
            Set xmlLancamentoEstorno = CreateObject("MSXML2.DOMDocument.4.0")
            xmlLancamentoEstorno.loadXML strLancamentoEstorno
    
            If Val(xmlLancamentoEstorno.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) <> enumStatusIntegracao.Integrado Then
                'Situação do lançamento em conta corrente não permite cancelamento
                lngCodigoErroNegocio = 3099
                GoTo ErrorHandler
            End If
        End If
                                                         
    End If
    
    If lngStatusLancamento = enumStatusIntegracao.Antecipado Or _
       lngStatusLancamento = enumStatusIntegracao.Disponível Or _
       lngStatusLancamento = enumStatusIntegracao.Suspenso Or _
       lngStatusLancamento = enumStatusIntegracao.ErroCC Or _
       lngStatusLancamento = enumStatusIntegracao.ErroSaldoCC Then
       
       objLancamentoContabil.AlterarStatus pvntNumeroSequencial, _
                                           enumTipoLancamentoIntegracao.Normal, _
                                           penumStatusIntegracao, _
                                           , , False
       
    ElseIf (lngStatusLancamento = enumStatusIntegracao.Integrado Or _
            lngStatusLancamento = enumStatusIntegracao.IntegradoCC) And _
            intCanalVenda = enumCanalDeVenda.SGC Then
       
       objLancamentoContabil.AlterarStatus pvntNumeroSequencial, _
                                           enumTipoLancamentoIntegracao.Normal, _
                                           penumStatusIntegracao, _
                                           , , False
       
    End If

    Set xmlLancamentoContabil = Nothing
    Set objLancamentoContabil = Nothing
    Set xmlLancamentoEstorno = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objLancamentoContabil = Nothing
    Set xmlLancamentoContabil = Nothing
    Set xmlLancamentoEstorno = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "flVerificarContaCorrente", lngCodigoErroNegocio, intNumeroSequencialErro
End Function

'--------------------------------------------------------------------------------------------------
'Comparar Campos da Remessa de cancelamento com a operação original
'--------------------------------------------------------------------------------------------------

Private Function flVerificarCamposOperacaoOriginal(ByVal strXMLOriginal As String, _
                                                   ByVal strXMLOperacao As String) As String

Dim xmlOriginal                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlErros                                As MSXML2.DOMDocument40

    Set xmlOriginal = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlOriginal.loadXML("<Grupo_Original>" & strXMLOriginal & "</Grupo_Original>")
    Call xmlOperacao.loadXML("<Grupo_Operacao>" & strXMLOperacao & "</Grupo_Operacao>")
    
    If fgVlrXml_To_Decimal(xmlOriginal.selectSingleNode("//VA_OPER_ATIV").Text) <> _
       fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text) Then
        fgAdicionaErro xmlErros, 3072, enumJustificativa.ValorFinanceiro
    End If

    If fgVlrXml_To_Decimal(xmlOriginal.selectSingleNode("//VA_MOED_ESTR_BMC").Text) <> _
       fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//VA_MOED_ESTR_BMC").Text) Then
        fgAdicionaErro xmlErros, 3135, enumJustificativa.ValorMoedaEstrangeira
    End If

    If xmlOriginal.selectSingleNode("//IN_OPER_DEBT_CRED").Text <> _
       xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text Then
        fgAdicionaErro xmlErros, 3067, enumJustificativa.DebitoCredito
    End If
              
    If xmlErros.selectNodes("//Grupo_ErrorInfo").length > 0 Then
        flVerificarCamposOperacaoOriginal = xmlErros.xml
    Else
        flVerificarCamposOperacaoOriginal = vbNullString
    End If
    
    Set xmlOperacao = Nothing
    Set xmlOriginal = Nothing
    Set xmlErros = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlOperacao = Nothing
    Set xmlOriginal = Nothing
    Set xmlErros = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarCamposOperacaoOriginal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma operação
'-------------------------------------------------------------------------------------------------

Private Function flCancelarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                           Optional ByRef strError As String = vbNullString) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim xmlOperacao                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
        Case enumTipoMensagemBUS.CompromissadaIda, _
             enumTipoMensagemBUS.Redesconto, _
             enumTipoMensagemBUS.PagamentoRedesconto, _
             enumTipoMensagemBUS.OperacaoCompromissadaCETIP, _
             enumTipoMensagemBUS.OperacaoRetornoAntecipacaoCETIP
            
            flCancelarCompromissadaIdaRSDC xmlRemessa
            Exit Function
        
        Case enumTipoMensagemBUS.LiquidacaoFisicaOperacaoBMA
            
            flCancelarLiquidacaoFisicaOper xmlRemessa
            Exit Function
        
        Case enumTipoMensagemBUS.OperacoesCorretoras
            
            flCancelarOperacaoCorretoras xmlRemessa
            Exit Function
        
        Case enumTipoMensagemBUS.RegistroOperacoesRodaDolar
            
            Call CancelarRodaDolarBMC(xmlRemessa)
            Exit Function
        
        'KIDA - CBLC
        Case enumTipoMensagemBUS.RegistroLiquidacaoBrutaCBLC
        
            Call flCancelarBrutaCBLC(xmlRemessa)
            Exit Function
    
        'KIDA - CBLC - EVENTOS
        Case enumTipoMensagemBUS.RegistroLiquidacaoEventoCBLC
        
            Call flCancelarEventoCBLC(xmlRemessa)
            Exit Function
    
    End Select

    'Cancelar operação CETIP
    If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        flCancelarOperacaoCETIP xmlRemessa
        Exit Function
    End If

    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoLocalizada
        lngCodigoErroNegocio = 3009
        GoTo ErrorHandler
    End If
    
    flVerificarContaCorrente xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)
        Case enumTipoMensagemBUS.RegistroOperacoesBMC
            
            strError = flVerificarCamposOperacaoOriginal(xmlRemessa.xml, xmlOperacao.xml)
            
            If strError <> vbNullString Then
                Exit Function
            End If
    
        Case enumTipoMensagemLQS.TermoD0
    
            If Not xmlOperacao.selectSingleNode("//TP_LIQU_DEFN") Is Nothing And _
               Not xmlRemessa.selectSingleNode("//TP_LIQU_DEFN") Is Nothing Then
                If Trim$(xmlOperacao.selectSingleNode("//TP_LIQU_DEFN").Text) <> _
                   Trim$(xmlRemessa.selectSingleNode("//TP_LIQU_DEFN").Text) Then
                    'Tipo de Liquidação Definitiva informado, é diferente da operação original. Solicitação não permitida.
                    lngCodigoErroNegocio = 4273
                    GoTo ErrorHandler
                End If
            End If
        
            If Not xmlOperacao.selectSingleNode("//TP_LIQU_COMP") Is Nothing And _
               Not xmlRemessa.selectSingleNode("//TP_LIQU_COMP") Is Nothing Then
                If Trim$(xmlOperacao.selectSingleNode("//TP_LIQU_COMP").Text) <> _
                   Trim$(xmlRemessa.selectSingleNode("//TP_LIQU_COMP").Text) Then
                    'Tipo de Liquidação Compromissada informado, é diferente da operação original. Solicitação não permitida.
                    lngCodigoErroNegocio = 4274
                    GoTo ErrorHandler
                End If
            End If
        
            If Not xmlOperacao.selectSingleNode("//PU_ATIV_MERC_RETN") Is Nothing And _
               Not xmlRemessa.selectSingleNode("//PU_ATIV_MERC_RETN") Is Nothing Then
                If Trim$(xmlOperacao.selectSingleNode("//PU_ATIV_MERC_RETN").Text) <> _
                   Trim$(xmlRemessa.selectSingleNode("//PU_ATIV_MERC_RETN").Text) Then
                    'Preço Unitário Retorno informado, é diferente da operação original. Solicitação não permitida.
                    lngCodigoErroNegocio = 4275
                    GoTo ErrorHandler
                End If
            End If
        
            If Not xmlOperacao.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing And _
               Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing Then
                If Trim$(xmlOperacao.selectSingleNode("//VA_OPER_ATIV_RETN").Text) <> _
                   Trim$(xmlRemessa.selectSingleNode("//VA_OPER_ATIV_RETN").Text) Then
                    'Valor Financeiro Retorno informado, é diferente da operação original. Solicitação não permitida.
                    lngCodigoErroNegocio = 4276
                    GoTo ErrorHandler
                End If
            End If
        
    End Select
    
    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        
        Case enumStatusOperacao.Inicial, _
             enumStatusOperacao.EmSer, _
             enumStatusOperacao.AConciliar, _
             enumStatusOperacao.ConcordanciaBackoffice, _
             enumStatusOperacao.Concordancia, _
             enumStatusOperacao.ConcordanciaAutomatica, _
             enumStatusOperacao.ConcordanciaBalcao, _
             enumStatusOperacao.ConcordanciaBalcaoAutomatica, _
             enumStatusOperacao.Conciliada, _
             enumStatusOperacao.ConciliadaAutomatica
             'enumStatusOperacao.LiberadaCliente1

            GerenciarCancelamento xmlOperacao, True
            
        Case enumStatusOperacao.AConciliarRegistro
            
             If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
                CancelarRegistroOperacaoBMCBalcao xmlRemessa
             Else
                '3025 - Situação da Operação não permite cancelamento.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
             End If
            
        Case enumStatusOperacao.AConciliarAceite
             
             If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacoesBMC Then
                GerenciarCancelamento xmlOperacao, True
             Else
                '3025 - Situação da Operação não permite cancelamento.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
             End If

        Case enumStatusOperacao.Rejeitada, _
             enumStatusOperacao.Expirada, _
             enumStatusOperacao.Registrada, _
             enumStatusOperacao.RegistradaAutomatica, _
             enumStatusOperacao.RejeitadaPiloto, _
             enumStatusOperacao.AConciliarBMA0013
            
             If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.RegistroOperacoesBMC Then
                '3025 - Situação da Operação não permite cancelamento.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
             Else
                GerenciarCancelamento xmlOperacao, True
            End If

        Case enumStatusOperacao.ConcordanciaAdmArea, _
             enumStatusOperacao.DiscordanciaLib, _
             enumStatusOperacao.ConcordanciaBackofficePrevia, _
             enumStatusOperacao.ConcordanciaLib, _
             enumStatusOperacao.ConcordanciaLibAuto

             If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralCBLC Or _
                xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoBrutaCBLC Or _
                xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC Or _
                xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF Then
                '3025 - Situação da Operação não permite cancelamento.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
             Else
                GerenciarCancelamento xmlOperacao, True
             End If

        Case enumStatusOperacao.Liquidada, _
             enumStatusOperacao.LiquidadaConvertida, _
             enumStatusOperacao.LiquidadaFisicamente, _
             enumStatusOperacao.LiquidadaViaContingencia, _
             enumStatusOperacao.LiquidadaFisicamenteAutomatica

             If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.OperacaoDefinitivaInternaBMA Then
                '3011 - Operação já Liquidada, não podendo ser cancelada.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3011
                GoTo ErrorHandler
             Else
                GerenciarCancelamento xmlOperacao, True
             End If

        Case enumStatusOperacao.Liberada

             If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais Then
                '3025 - Situação da Operação não permite cancelamento.
                'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
             Else
                GerenciarCancelamento xmlOperacao, True
             End If

        Case Else

            '3025 - Situação da Operação não permite cancelamento.
            'fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            fgAppendNode xmlRemessa, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    
    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarOperacao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma operação Cetip
'-------------------------------------------------------------------------------------------------

Private Function flCancelarOperacaoCETIP(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim xmlOperacao                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                 xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If
    
    flVerificarContaCorrente xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.Inicial, _
             enumStatusOperacao.EmSer, _
             enumStatusOperacao.AConciliar, _
             enumStatusOperacao.Concordancia, _
             enumStatusOperacao.ConcordanciaAutomatica, _
             enumStatusOperacao.Expirada, _
             enumStatusOperacao.Rejeitada, _
             enumStatusOperacao.RejeitadaPiloto, _
             enumStatusOperacao.Registrada, _
             enumStatusOperacao.RegistradaAutomatica, _
             enumStatusOperacao.Liquidada
             
            GerenciarCancelamento xmlOperacao, True

        'Pikachu - 08/04/05
        'RATS 228- ocor 3
        'Case enumStatusOperacao.Registrada, enumStatusOperacao.RegistradaAutomatica
        '    GerenciarCancelamento xmlOperacao, True

        'Pikachu - 08/04/05
        'RATS 228- ocor 3
        'Case enumStatusOperacao.Registrada, enumStatusOperacao.RegistradaAutomatica
        '    GerenciarCancelamento xmlOperacao, True

        Case Else
            '3025 - Situação da Operação não permite cancelamento.
            fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarOperacaoCETIP Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma liquidação com corretoras
'-------------------------------------------------------------------------------------------------

Private Function flCancelarOperacaoCorretoras(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim xmlOperacao                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If
    
    flVerificarContaCorrente xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.Inicial, enumStatusOperacao.EmSer
            GerenciarCancelamento xmlOperacao, True

        Case Else
            '3025 - Situação da Operação não permite cancelamento.
            fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarOperacaoCorretoras Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de registro de roda de dólar
'-------------------------------------------------------------------------------------------------

Public Function CancelarRodaDolarBMC(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If
    
    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.Conciliada, enumStatusOperacao.ConciliadaAutomatica
             
            Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text))
            
            If xmlConciliacao.xml <> vbNullString Then
                Call objMensagemEnviadaRecebida.AlterarStatus(xmlConciliacao.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              xmlConciliacao.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              enumStatusMensagem.AConciliar, _
                                                              vbNullString, _
                                                              , _
                                                              False, _
                                                              , _
                                                              "BMC0011")
            End If
            
        Case enumStatusOperacao.EmSer, enumStatusOperacao.ManualEmSer
            
        Case Else
            '3025 - Situação da Operação não permite cancelamento.
            fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    
    End Select

    Call objOperacao.AlterarStatusAcao( _
            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
            enumStatusOperacao.CanceladaOrigem, _
            vbNullString, , , _
            False)

    Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
    Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
    
    Call objOperacaoMensagemInterna.Incluir(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                            xmlRemessa)
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set objConciliacao = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set objConciliacao = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlConciliacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarRodaDolarBMC Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento do Registro de Operacao BMC Balcao
'-------------------------------------------------------------------------------------------------

Public Function CancelarRegistroOperacaoBMCBalcao(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgOrig                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMesgOrig = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If
    
    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.AConciliarRegistro
            
            Set xmlMesgOrig = objMensagemEnviadaRecebida.LerXMLMensagem(, , , "BMC0001R1" _
                                                                        , , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            
            If xmlMesgOrig.xml <> vbNullString Then
                
                fgAppendNode xmlOperacao, "MESG", "CO_MESG", "BMC0002"
                fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", enumTipoMensagemLQS.RegistroOperacoesBMCRetorno
                fgAppendNode xmlOperacao, "MESG", "CO_CANC_BMC", fgObterCodCancelBMC0002
                fgAppendNode xmlOperacao, "MESG", "NU_CTRL_BMC_ORIG", xmlMesgOrig.selectSingleNode("//NU_CTRL_CAMR").Text
                
                Call flLiberarOperacao(xmlOperacao, True)
            Else
                '4302 - Mensagem original nao localizada para envio de mensagem de cancelamento
                lngCodigoErroNegocio = 4302
                fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.MensagemNaoLocalizada
                GoTo ErrorHandler
            End If

        Case Else
            '3025 - Situação da Operação não permite cancelamento.
            fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    
    End Select
    
    Call objOperacaoMensagemInterna.Incluir(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                            xmlRemessa)
    
    Set objOperacao = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgOrig = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMesgOrig = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarRegistroOperacaoBMCBalcao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma operação compromissada de ida e/ou redesconto
'-------------------------------------------------------------------------------------------------
Private Function flCancelarCompromissadaIdaRSDC(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objLegado                               As A8LQS.clsLegado
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoEstorno                      As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strXML                                  As String
Dim strMensagem                             As String
Dim strOperacao                             As String
Dim blnCliente1                             As Boolean
Dim vntCodigoXML                            As Variant
Dim blnCancelouOperacao                     As Boolean

Dim lngStatusOriginal                       As Long
Dim lngTipoMensagem                         As Long
Dim lngTipoOperacao                         As Long

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")

    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Then
        objOperacao.ComplOperacaoComRDCOriginal xmlRemessa
    End If

    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    strOperacao = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                       xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                       , , , _
                                       xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text, _
                                       , , _
                                       xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)

    If Trim(strOperacao) = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        GoTo ErrorHandler
    End If

    xmlLerTodos.loadXML strOperacao

    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    strXML = LerTodos("RecebimentoCancelamento", _
                      xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text, _
                      enumIndicadorSimNao.Sim)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    For Each objDomNode In xmlLerTodos.selectNodes("Repeat_Operacao/Grupo_Operacao")
    
        blnCancelouOperacao = False

        Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

        If xmlOperacao.Text = vbNullString Then
            '3009 - Operação não localizada no SLCC para cancelamento.
            lngCodigoErroNegocio = 3009
            GoTo ErrorHandler
        End If

        If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
            flVerificarContaCorrente objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
        Else
            'Para operação compromissada tem uma situação que a operação pode estar liquidada e ser cancelada
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemBUS.CompromissadaIda Then
                '3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler
            End If
        End If

        Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
            Case enumStatusOperacao.Pendencia, enumStatusOperacao.EmLancamento, _
                 enumStatusOperacao.Liberada, enumStatusOperacao.LiberadaAutomatica, _
                 enumStatusOperacao.Estornada ', enumStatusOperacao.LiquidadaCliente1
                '3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler

            Case enumStatusOperacao.CanceladaOrigem
                'Verificar se ja foi solicitado Cancelamento pela Origem
                If objOperacaoMensagemInterna.ExisteTipoSolicitacao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                    enumTipoSolicitacao.Cancelamento) Then
                    '3025 - Situação da Operação não permite cancelamento.
                    lngCodigoErroNegocio = 3025
                    GoTo ErrorHandler
                End If

            Case enumStatusOperacao.Inicial, enumStatusOperacao.EmSer, enumStatusOperacao.AConciliar, _
                 enumStatusOperacao.Concordancia, enumStatusOperacao.ConcordanciaAutomatica, _
                 enumStatusOperacao.Conciliada, enumStatusOperacao.ConciliadaAutomatica, _
                 enumStatusOperacao.Rejeitada, enumStatusOperacao.Expirada, enumStatusOperacao.Liquidada, _
                 enumStatusOperacao.LiquidadaConvertida, _
                 enumStatusOperacao.RejeitadaPiloto, enumStatusOperacao.AComplementar, enumStatusOperacao.Cancelada
                 'enumStatusOperacao.LiberadaCliente1

                If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
                    lngStatusOriginal = Val(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text)
                
                    objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                        enumStatusOperacao.CanceladaOrigem
                    
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.CanceladaOrigem
                    blnCancelouOperacao = True
                End If

                If objOperacaoMensagemInterna.ExisteTipoSolicitacao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                    enumTipoSolicitacao.Inclusao, _
                                                                    vntCodigoXML) Then
                    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim Then
                        'Obter o xml da operação de inclusão para estorno do previsto
                        Set xmlOperacaoEstorno = CreateObject("MSXML2.DOMDocument.4.0")
                        xmlOperacaoEstorno.loadXML fgSelectVarchar4000(vntCodigoXML)
                        
                        fgAppendNode xmlOperacaoEstorno, "MESG", "TP_OPER", objDomNode.selectSingleNode("TP_OPER").Text
                        fgAppendNode xmlOperacaoEstorno, "MESG", "NU_SEQU_OPER_ATIV", objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                        
                        If Not objDomNode.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
                            fgAppendNode xmlOperacaoEstorno, "MESG", "CO_ULTI_SITU_PROC", objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text
                        End If
                        
                        If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
                            lngTipoMensagem = Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
                        End If
                        
                        If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
                            lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
                        End If
                        
                        If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
                            
                            If lngStatusOriginal = enumStatusOperacao.Concordancia Or _
                               lngStatusOriginal = enumStatusOperacao.ConcordanciaAutomatica Then
                                
'                                If Not xmlOperacao.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
'                                    fgAppendNode xmlOperacao, "MESG", "VA_OPER_ATIV_COMPL", xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text
'                                End If
                            
                                Call EnviarPrevisao(xmlOperacao, xmlControleProcessoOperacao)
                            
                            Else
                                fgAppendNode xmlOperacaoEstorno, "MESG", "NAO_ENVIAR_VENCIMENTO", vbNullString
                                Call EstornarPrevisao(xmlOperacaoEstorno, xmlControleProcessoOperacao)
                            
                            End If
                        
                        Else
                            EstornarPrevisao xmlOperacaoEstorno, xmlControleProcessoOperacao
                        
                        End If
                        
                        Set xmlOperacaoEstorno = Nothing
                    End If
                End If

                'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
                'é enviada por remessa
                If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Or _
                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica Or _
                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliar Or _
                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Conciliada Or _
                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica Or _
                  (lngStatusOriginal = enumStatusOperacao.Concordancia Or _
                   lngStatusOriginal = enumStatusOperacao.ConcordanciaAutomatica And _
                  (lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                   lngTipoMensagem = enumTipoMensagemBUS.OperacaoRetornoAntecipacaoCETIP)) Then
                   'Or xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaCliente1
                    
                    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                        If fgVerificarCliente1(xmlOperacao) Then
                            blnCliente1 = True
                            fgDebitoCreditoCliente1 xmlOperacao
                        Else
                            blnCliente1 = False
                        End If
                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        If blnCliente1 Then
                            'irá retornar o debito/credito original
                            fgDebitoCreditoCliente1 xmlOperacao
                            fgConverterProduto1 xmlOperacao
                            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                            fgConverterProduto1 xmlOperacao
                        End If
                        Set objGestaoCaixa = Nothing
                    End If
                End If
        
                Set objLegado = CreateObject("A8LQS.clsLegado")
                strMensagem = objLegado.EnviarRetorno(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                Set objLegado = Nothing

        End Select

        If blnCancelouOperacao Then
            'Armazenar a mensagem original para todas as operações canceladas
            objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                               CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                               xmlRemessa
        End If

    Next

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarCompromissadaIdaRSDC Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma operação compromissada de ida e/ou redesconto
'-------------------------------------------------------------------------------------------------
Private Function flCancelarPorLastro(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objLegado                               As A8LQS.clsLegado

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim rsOperacao                              As ADODB.Recordset
Dim strXML                                  As String
Dim strMensagem                             As String
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.ConsolidacaoRedescontosIntradia Then
        '4016 - Tipo de solicitação inválido.
        lngCodigoErroNegocio = 4016
        GoTo ErrorHandler
    End If

    If xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemBUS.PagamentoRedesconto Then
        objOperacao.ComplOperacaoComRDCOriginal xmlRemessa
    End If

    Set rsOperacao = objOperacao.LerPorSolicitacao(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                   enumTipoSolicitacao.Complementacao, _
                                                   , , _
                                                   xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("CO_IDEF_LAST").Text)

    If rsOperacao.EOF Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        GoTo ErrorHandler
    End If

    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    strXML = LerTodos("RecebimentoCancelamento", _
                      rsOperacao!TP_OPER, _
                      enumIndicadorSimNao.Sim)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    Do While Not rsOperacao.EOF

        Set xmlOperacao = objOperacao.LerXMLOperacao(rsOperacao!NU_SEQU_OPER_ATIV)

        If xmlOperacao.Text = vbNullString Then
            '3009 - Operação não localizada no SLCC para cancelamento.
            lngCodigoErroNegocio = 3009
            GoTo ErrorHandler
        End If

        Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
            Case enumStatusOperacao.Pendencia, enumStatusOperacao.EmLancamento, _
                 enumStatusOperacao.Liberada, enumStatusOperacao.LiberadaAutomatica, _
                 enumStatusOperacao.Liquidada, _
                 enumStatusOperacao.Cancelada, enumStatusOperacao.CanceladaOrigem, _
                 enumStatusOperacao.Estornada
                 'enumStatusOperacao.LiquidadaCliente1

                '3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                GoTo ErrorHandler

            Case enumStatusOperacao.Inicial, enumStatusOperacao.EmSer, enumStatusOperacao.AConciliar, _
                 enumStatusOperacao.Concordancia, enumStatusOperacao.ConcordanciaAutomatica, _
                 enumStatusOperacao.Conciliada, enumStatusOperacao.ConciliadaAutomatica, _
                 enumStatusOperacao.Rejeitada, enumStatusOperacao.Expirada, _
                 enumStatusOperacao.LiquidadaConvertida, enumStatusOperacao.RejeitadaPiloto ', enumStatusOperacao.LiberadaCliente1

                objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                    enumStatusOperacao.CanceladaOrigem

                xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.CanceladaOrigem
                
                If (xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliar Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Conciliada Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica) Then
                    'xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiberadaCliente1
                    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                        If fgVerificarCliente1(xmlOperacao) Then
                            blnCliente1 = True
                            fgDebitoCreditoCliente1 xmlOperacao
                        Else
                            blnCliente1 = False
                        End If
                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        If blnCliente1 Then
                            'irá retornar o debito/credito original
                            fgDebitoCreditoCliente1 xmlOperacao
                            fgConverterProduto1 xmlOperacao
                            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                            fgConverterProduto1 xmlOperacao
                        End If
                        Set objGestaoCaixa = Nothing
                    End If
                End If
                
                Set objLegado = CreateObject("A8LQS.clsLegado")
                strMensagem = objLegado.EnviarRetorno(xmlOperacao)
                EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                Set objLegado = Nothing
                
        End Select

        rsOperacao.MoveNext

    Loop

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarPorLastro Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de cancelamento de uma liquidação física de operação
'-------------------------------------------------------------------------------------------------

Private Function flCancelarLiquidacaoFisicaOper(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim rsOperacao                              As ADODB.Recordset
Dim strXML                                  As String
Dim strMensagem                             As String
Dim blnEstornar                             As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Set rsOperacao = objOperacao.LerPorSolicitacao(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                   enumTipoSolicitacao.Complementacao, _
                                                   "")

    If rsOperacao.EOF Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If

    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    'Criar no enumTipoOperacao
    strXML = LerTodos("RecebimentoCancelamento", _
                      enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA, _
                      enumIndicadorSimNao.Sim)

    If strXML = "" Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    blnEstornar = True
    Do While Not rsOperacao.EOF

        Set xmlOperacao = objOperacao.LerXMLOperacao(rsOperacao!NU_SEQU_OPER_ATIV)

        If xmlOperacao.Text = vbNullString Then
            '3009 - Operação não localizada no SLCC para cancelamento.
            lngCodigoErroNegocio = 3009
            fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
            GoTo ErrorHandler
        End If
        
        flVerificarContaCorrente rsOperacao!NU_SEQU_OPER_ATIV

        Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
            Case enumStatusOperacao.Inicial, enumStatusOperacao.EmSer, enumStatusOperacao.AConciliar, _
                 enumStatusOperacao.Concordancia, enumStatusOperacao.ConcordanciaAutomatica, _
                 enumStatusOperacao.Conciliada, enumStatusOperacao.ConciliadaAutomatica, _
                 enumStatusOperacao.Rejeitada, enumStatusOperacao.Expirada, enumStatusOperacao.Liquidada, enumStatusOperacao.RejeitadaPiloto

                If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
                    objOperacao.AlterarStatusOperIndisp xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                        enumStatusOperacao.CanceladaOrigem
                End If

                EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao

                'Esta seleção de nós está assim pois a estrutura do XML é diferente quando é chamada pela tela e quando
                'é enviada por remessa
                If (xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Concordancia Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaAutomatica Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliar Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Conciliada Or _
                    xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica) Then
                    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                        EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                        If fgVerificarCliente1(xmlOperacao) Then
                            strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                            EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
                            fgDebitoCreditoCliente1 xmlOperacao
                        End If
                        Set objGestaoCaixa = Nothing
                    End If
                End If

            Case enumStatusOperacao.Cancelada, enumStatusOperacao.CanceladaOrigem, _
                 enumStatusOperacao.Pendencia, enumStatusOperacao.EmLancamento, _
                 enumStatusOperacao.Liberada, enumStatusOperacao.LiberadaAutomatica, _
                 enumStatusOperacao.Estornada, enumStatusOperacao.Registrada, _
                 enumStatusOperacao.RegistradaAutomatica, _
                 enumStatusOperacao.ConcordanciaBackoffice, enumStatusOperacao.ConcordanciaAdmArea, _
                 enumStatusOperacao.DiscordanciaBackoffice, enumStatusOperacao.DiscordanciaAdmArea, _
                 enumStatusOperacao.PagamentoLib
                '3025 - Situação da Operação não permite cancelamento.
                lngCodigoErroNegocio = 3025
                fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoPermiteCancelamento
                GoTo ErrorHandler
        End Select
        rsOperacao.MoveNext
    Loop

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarLiquidacaoFisicaOper Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo liquidação de cliente 1
'-------------------------------------------------------------------------------------------------

Public Function flLiquidarCliente1(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                          Optional ByVal pblnLiquidar As Boolean = False) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLegado                               As A8LQS.clsLegado
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim strDataOperacao                         As String
Dim lngStatusOperacao                       As Long
Dim vntSequenciaOperacao                    As Variant

Dim lngIndDebtCredOrig                      As Long
Dim lngCodProdOrig                          As Long

On Error GoTo ErrorHandler

    If pblnLiquidar Then
        If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
            Exit Function
        End If
    End If
    
    lngIndDebtCredOrig = xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    lngCodProdOrig = xmlRemessa.selectSingleNode("//CO_PROD").Text

    If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        strDataOperacao = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    Else
        strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
    End If

    'Ler o status da operação da tesouraria
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    lngStatusOperacao = objOperacao.ObterStatusOperInvertida(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                             strDataOperacao, _
                                                             xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                             xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text, _
                                                             xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text, _
                                                             xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                             vntSequenciaOperacao)

    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", 0
    End If

    If pblnLiquidar Then
        'Neste ponto é no R1
        If lngStatusOperacao <> enumStatusOperacao.LiberadaCliente1 Then
            Exit Function
        Else
            objOperacao.AlterarStatusOperIndisp vntSequenciaOperacao, _
                                                enumStatusOperacao.LiquidadaCliente1
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiquidadaCliente1
        End If
    Else
        If lngStatusOperacao <> enumStatusOperacao.Liquidada Then
            'Alterar somente o status pois será liquidado no retorno da ponta da tesouraria
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.LiberadaCliente1
            Exit Function
        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.LiquidadaCliente1
        End If
    End If

    Set objOperacao = Nothing

    If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", enumStatusOperacao.Liquidada
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("CO_FORM_LIQU") Is Nothing Then
        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        objContaCorrente.Disponibilizar xmlRemessa
        Set objContaCorrente = Nothing
    End If

    'Incluido 12/05/2004 *************** Inicio
    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    fgDebitoCreditoCliente1 xmlRemessa
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
    'debito/Credito e Produto Original
    xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    xmlRemessa.selectSingleNode("//CO_PROD").Text = lngCodProdOrig
    
    fgConverterProduto1 xmlRemessa
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
    'debito/Credito e Produto Original
    xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = lngIndDebtCredOrig
    xmlRemessa.selectSingleNode("//CO_PROD").Text = lngCodProdOrig

    Set objSubReserva = Nothing
    'Incluido 12/05/2004 *************** Fim

    'Enviar Retorno para o sistema legado.
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
    Set objLegado = Nothing

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidarCliente1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de liquidação de uma operação compromissada de ida
'-------------------------------------------------------------------------------------------------

Private Function flLiquidarCompromissada(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objLegado                               As A8LQS.clsLegado
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim strMensagem                             As String
Dim strDataOperacao                         As String
Dim lngStatusOperacao                       As Long
Dim vntSequenciaOperacao                    As Variant
Dim blnCliente1                             As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                        enumStatusOperacao.Liquidada
    xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada

    If Not xmlRemessa.documentElement.selectSingleNode("CO_FORM_LIQU") Is Nothing Then
        Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
        objContaCorrente.Disponibilizar xmlRemessa
        Set objContaCorrente = Nothing
    End If

    If fgVerificarCliente1(xmlRemessa) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    If blnCliente1 Then
        fgDebitoCreditoCliente1 xmlRemessa
    End If
    strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
    If blnCliente1 Then
        'irá retornar o debito/credito original
        fgDebitoCreditoCliente1 xmlRemessa
        fgConverterProduto1 xmlRemessa
        strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlRemessa)
        EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
        fgConverterProduto1 xmlRemessa
    End If
    Set objSubReserva = Nothing

    'Enviar Retorno para o sistema legado.
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
    Set objLegado = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidarCompromissada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de envio do realizado de despesas selic para o sistema PJ
'-------------------------------------------------------------------------------------------------

Public Function EnviarDespesasPJ(ByRef xmlConciliacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim dblValorMensagem                        As Double
Dim dblValorOperacao                        As Double
Dim dblValorDiferenca                       As Double

    On Error GoTo ErrorHandler

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    dblValorMensagem = fgVlrToDBServer(xmlConciliacao.documentElement.selectSingleNode("//ValorTotalMensagem").Text)
    dblValorOperacao = fgVlrToDBServer(xmlConciliacao.documentElement.selectSingleNode("//ValorTotalOperacao").Text)

    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlConciliacao.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

    If dblValorMensagem > dblValorOperacao Then
        xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoEntradaSaida.Saida
    ElseIf dblValorMensagem < dblValorOperacao Then
        xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoEntradaSaida.ENTRADA
    Else
        Exit Function
    End If

    'Produto definido para o Net de Despesas após a conciliação
    xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "530"
    xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text = xmlConciliacao.documentElement.selectSingleNode("//ValorDiferenca").Text

    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False

    Set objGestaoCaixa = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objGestaoCaixa = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarDespesasPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de liquidação de despesas selic
'-------------------------------------------------------------------------------------------------

Private Function flLiquidarDespesa(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim strMensagem                             As String

    On Error GoTo ErrorHandler

    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarRealizado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
    Set objGestaoCaixa = Nothing

    Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
    strMensagem = objSubReserva.EnviarRealizadoSolicitado(xmlRemessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva, False, False
    Set objSubReserva = Nothing

    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    objContaCorrente.Disponibilizar xmlRemessa
    Set objContaCorrente = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiquidarDespesa Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de cancelamento de uma operação com mensagem SPB

Private Function flCancelarOperacaoComMensagem(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strConciliarRegistroOperacao            As String
Dim intCanalVenda                           As Integer
Dim intTipoContraparte                      As Integer
Dim lngLayoutEntrada                        As Long

    On Error GoTo ErrorHandler

    lngLayoutEntrada = 0
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        lngLayoutEntrada = Val(xmlRemessa.selectSingleNode("//TP_MESG").Text)
    End If
    
    If lngLayoutEntrada = enumTipoMensagemBUS.RegistroContratoSWAP Then
        flValidarComplementoContratoSWAP xmlRemessa
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        GoTo ErrorHandler
    End If

    intCanalVenda = 0
    If Not xmlRemessa.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val(xmlRemessa.selectSingleNode("//TP_CNAL_VEND").Text)
    End If
    
    intTipoContraparte = 0
    If Not xmlRemessa.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(xmlRemessa.selectSingleNode("//TP_CNPT").Text)
    End If
    
    fgAppendNode xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
    fgAppendNode xmlRemessa, "MESG", "TP_OPER", xmlOperacao.selectSingleNode("//TP_OPER").Text

    Select Case xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text
        Case enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao, _
             enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP, _
             enumTipoOperacaoLQS.RetiradaCustodiaConciliacao
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            
            fgAppendNode xmlRemessa, "MESG", "DH_ULTI_ATLZ", xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text
            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CUTD_SELIC_CNPT", xmlOperacao.selectSingleNode("//CO_CNTA_CUTD_SELIC_CNPT").Text
            fgAppendNode xmlRemessa, "MESG", "DE_ATIV_MERC", xmlOperacao.selectSingleNode("//CO_SUB_TIPO_ATIV").Text
            fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC", xmlOperacao.selectSingleNode("//NU_ATIV_MERC_CETIP").Text
            fgAppendNode xmlRemessa, "MESG", "QT_ATIV_MERC", fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//QT_TITU_CETIP").Text)
            
            If xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "Crédito"
            Else
                xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "Débito"
            End If
            
            If Not xmlOperacao.selectSingleNode("//CO_AREA") Is Nothing And _
               Not xmlRemessa.selectSingleNode("//CO_AREA") Is Nothing Then
               
                If xmlOperacao.selectSingleNode("//CO_AREA").Text <> xmlRemessa.selectSingleNode("//CO_AREA").Text Then
                    '4235 - Código de Área divergente da operação original
                    lngCodigoErroNegocio = 4235
                    GoTo ErrorHandler
                End If
                
            End If
            
            strConciliarRegistroOperacao = objConciliacao.ConciliarRegistroOperacao(enumAcaoConciliacao.BOConcordar, _
                                                                                    xmlRemessa, _
                                                                                    , _
                                                                                    True, , , True)
            If strConciliarRegistroOperacao <> vbNullString Then
                Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
                xmlErro.loadXML strConciliarRegistroOperacao
                'Erros de consiliação
                lngCodigoErroNegocio = xmlErro.selectSingleNode("//Number").Text
                GoTo ErrorHandler
            End If
            If xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "Crédito" Then
                xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            Else
                xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            End If
            Set objConciliacao = Nothing
    End Select

    flVerificarContaCorrente xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                             enumStatusIntegracao.Cancelado

    fgAppendNode xmlRemessa, "MESG", "CO_ULTI_SITU_PROC", xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.Registrada, _
             enumStatusOperacao.RegistradaAutomatica, _
             enumStatusOperacao.Pendencia

            'Armazenar o novo comando da operação que devera ser enviado para o SPB
            objOperacao.AlterarComandoOperacao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                               CLng("0" & xmlRemessa.documentElement.selectSingleNode("CO_OPER_CETIP").Text)

            Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
            objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
            Set objWorkflow = Nothing

        Case enumStatusOperacao.Liquidada

            If intCanalVenda = enumCanalDeVenda.SGC Then
                
                'RATS 942
                If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroContratoSWAPSemOpcaoBarreira Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo_CETIP21 Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP_CETIP21 Then
                
                    
                    If Not xmlRemessa.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                        fgAppendNode xmlRemessa, "MESG", "NU_CTRL_MESG_SPB_ORIG", xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                    End If
                    
                    If Not xmlRemessa.selectSingleNode("//NU_COMD_OPER_ORIG") Is Nothing Then
                        xmlRemessa.selectSingleNode("//NU_COMD_OPER_ORIG").Text = xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                    Else
                        fgAppendNode xmlRemessa, "MESG", "NU_COMD_OPER_ORIG", xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                    End If
                    
                    If Not xmlRemessa.selectSingleNode("//NU_ATIV_MERC_CETIP") Is Nothing Then
                        xmlRemessa.selectSingleNode("//NU_ATIV_MERC_CETIP").Text = xmlOperacao.selectSingleNode("//NU_ATIV_MERC").Text
                    Else
                        fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC_CETIP", xmlOperacao.selectSingleNode("//NU_ATIV_MERC").Text
                    End If
                
                    If Not xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                        xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "2"
                    Else
                        fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", "2"
                    End If
                
                End If
                '---------------------------------
                
                Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                Set objWorkflow = Nothing
            
            Else
                If lngLayoutEntrada = enumTipoMensagemBUS.RegistroContratoSWAP Then
                    If xmlRemessa.documentElement.selectSingleNode("IN_TIPO_CNTR").Text = "N" Then
    
                        'Armazenar o novo comando da operação que devera ser enviado para o SPB
                        objOperacao.AlterarComandoOperacao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                           CLng("0" & xmlRemessa.documentElement.selectSingleNode("CO_OPER_CETIP").Text)
    
                        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                        objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                        Set objWorkflow = Nothing
                    Else
                        '3025 - Situação da Operação não permite cancelamento.
                        lngCodigoErroNegocio = 3025
                        GoTo ErrorHandler
                    End If
    
                ElseIf lngLayoutEntrada = enumTipoMensagemBUS.RegistroContratoTermoCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.LancamentoPU_CETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.EspecifQtdeCotasFundosInvestimento Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.RegistroOperacaoesCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.OperacaoRetornoAntecipacaoCETIP Then
    
                    If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                        objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                        Set objWorkflow = Nothing
                    Else
                        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade Then
                            Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                            objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                            Set objWorkflow = Nothing
                        Else
                            '3025 - Situação da Operação não permite cancelamento.
                            lngCodigoErroNegocio = 3025
                            GoTo ErrorHandler
                        End If
                    End If
    
                ElseIf lngLayoutEntrada = enumTipoMensagemBUS.RegistroOperacaoesCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.AntecipacaoResgateContratoDerivativoCETIP Or _
                       lngLayoutEntrada = enumTipoMensagemBUS.EventosJurosCETIP Then
    
                    If intTipoContraparte = enumTipoContraparte.Cliente1 Then
                        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                        objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                        Set objWorkflow = Nothing
                    Else
                        '3025 - Situação da Operação não permite cancelamento.
                        lngCodigoErroNegocio = 3025
                        GoTo ErrorHandler
                    End If
    
                ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RetiradaCustodia Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RetiradaCustodiaConciliacao Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DepositoFundoInvestimentoCETIP Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP Then
    
                    If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
                        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                        objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                        Set objWorkflow = Nothing
                    Else
                        If xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = enumTipoLiquidacao.SemModalidade Then
                            Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                            objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                            Set objWorkflow = Nothing
                        Else
                            '3025 - Situação da Operação não permite cancelamento.
                            lngCodigoErroNegocio = 3025
                            GoTo ErrorHandler
                        End If
                    End If
                
                'KIDA - RATS 932
                ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroContratoSWAPSemOpcaoBarreira Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroContratoSWAPComOpcaoBarreira Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo_CETIP21 Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Or _
                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP_CETIP21 Then
                        
                        'RATS 942
                        'Armazenar o novo comando da operação que devera ser enviado para o SPB
                        objOperacao.AlterarComandoOperacao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                           CLng("0" & xmlRemessa.documentElement.selectSingleNode("CO_OPER_CETIP").Text)
                                            
                        If Not xmlRemessa.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "NU_CTRL_MESG_SPB_ORIG", xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                        End If
                        
                        If Not xmlRemessa.selectSingleNode("//NU_COMD_OPER_ORIG") Is Nothing Then
                            xmlRemessa.selectSingleNode("//NU_COMD_OPER_ORIG").Text = xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                        Else
                            fgAppendNode xmlRemessa, "MESG", "NU_COMD_OPER_ORIG", xmlOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                        End If
                        
                        If Not xmlRemessa.selectSingleNode("//NU_ATIV_MERC_CETIP") Is Nothing Then
                            xmlRemessa.selectSingleNode("//NU_ATIV_MERC_CETIP").Text = xmlOperacao.selectSingleNode("//NU_ATIV_MERC").Text
                        Else
                            fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC_CETIP", xmlOperacao.selectSingleNode("//NU_ATIV_MERC").Text
                        End If
                        
                        If Not xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                            xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "2"
                        Else
                            fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", "2"
                        End If
                        
                        
                        
                        '------------------------------------
                        
                        Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                        objWorkflow.GerenciarCancelamentoMensagem xmlRemessa
                        Set objWorkflow = Nothing
    
                Else
                    '3025 - Situação da Operação não permite cancelamento.
                    lngCodigoErroNegocio = 3025
                    GoTo ErrorHandler
                
                End If
            
            End If

        Case Else
            '3025 - Situação da Operação não permite cancelamento.
            lngCodigoErroNegocio = 3025
            GoTo ErrorHandler
    
    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       lngLayoutEntrada, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    Exit Function

ErrorHandler:
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlErro = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarOperacaoComMensagem Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o fluxo de alteração de uma operação

Private Function flAlterarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objWorkflow                             As A8LQS.clsWorkflow
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim lngTipoLiquidacao                       As enumTipoLiquidacao
Dim strCampoValor                           As String

Dim lngTipoOperacao                         As Long

    On Error GoTo ErrorHandler

    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)

        Case enumTipoMensagemBUS.EventosJurosCETIP

            flValidarEventosJurosCETIP xmlRemessa
            lngTipoLiquidacao = Val("0" & xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text)

    End Select

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.CompromissadaIda Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.Redesconto Or _
        CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.PagamentoRedesconto Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                     xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                     , , , , _
                                                     xmlRemessa.documentElement.selectSingleNode("CO_IDEF_LAST").Text)
    Else
        Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                     xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    End If

    If xmlOperacao.Text = vbNullString Then
        '3114 - Operação não localizada no SLCC para alteração.
        lngCodigoErroNegocio = 3114
        GoTo ErrorHandler
    End If

    lngTipoOperacao = 0
    If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlOperacao.selectSingleNode("//TP_OPER").Text)
    End If
    
    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.EmSer, enumStatusOperacao.Concordancia, _
             enumStatusOperacao.ConcordanciaAutomatica, enumStatusOperacao.Registrada, _
             enumStatusOperacao.RegistradaAutomatica, enumStatusOperacao.Rejeitada, _
             enumStatusOperacao.Expirada, enumStatusOperacao.AComplementar, _
             enumStatusOperacao.AConciliar

            If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
                If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.LiquidacaoMultilateralBMC Then
                    '3115 - Situação da Operação não permite alteração.
                    lngCodigoErroNegocio = 3115
                    GoTo ErrorHandler
                End If
            
                If Val(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.TermoD0 Then
                    
                    If Not xmlOperacao.selectSingleNode("//TP_LIQU_DEFN") Is Nothing And _
                       Not xmlRemessa.selectSingleNode("//TP_LIQU_DEFN") Is Nothing Then
                        If Trim$(xmlOperacao.selectSingleNode("//TP_LIQU_DEFN").Text) <> _
                           Trim$(xmlRemessa.selectSingleNode("//TP_LIQU_DEFN").Text) Then
                            'Tipo de Liquidação Definitiva informado, é diferente da operação original. Solicitação não permitida.
                            lngCodigoErroNegocio = 4273
                            GoTo ErrorHandler
                        End If
                    End If
                
                    If Not xmlOperacao.selectSingleNode("//TP_LIQU_COMP") Is Nothing And _
                       Not xmlRemessa.selectSingleNode("//TP_LIQU_COMP") Is Nothing Then
                        If Trim$(xmlOperacao.selectSingleNode("//TP_LIQU_COMP").Text) <> _
                           Trim$(xmlRemessa.selectSingleNode("//TP_LIQU_COMP").Text) Then
                            'Tipo de Liquidação Compromissada informado, é diferente da operação original. Solicitação não permitida.
                            lngCodigoErroNegocio = 4274
                            GoTo ErrorHandler
                        End If
                    End If
                
                    If Not xmlOperacao.selectSingleNode("//PU_ATIV_MERC_RETN") Is Nothing And _
                       Not xmlRemessa.selectSingleNode("//PU_ATIV_MERC_RETN") Is Nothing Then
                        If Trim$(xmlOperacao.selectSingleNode("//PU_ATIV_MERC_RETN").Text) <> _
                           Trim$(xmlRemessa.selectSingleNode("//PU_ATIV_MERC_RETN").Text) Then
                            'Preço Unitário Retorno informado, é diferente da operação original. Solicitação não permitida.
                            lngCodigoErroNegocio = 4275
                            GoTo ErrorHandler
                        End If
                    End If
                
                    If Not xmlOperacao.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing And _
                       Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing Then
                        If Trim$(xmlOperacao.selectSingleNode("//VA_OPER_ATIV_RETN").Text) <> _
                           Trim$(xmlRemessa.selectSingleNode("//VA_OPER_ATIV_RETN").Text) Then
                            'Valor Financeiro Retorno informado, é diferente da operação original. Solicitação não permitida.
                            lngCodigoErroNegocio = 4276
                            GoTo ErrorHandler
                        End If
                    End If
                
                End If
            End If
            
            'Alterar a operação
            objOperacao.AlterarComandoOperacao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                               , _
                                               lngTipoLiquidacao

        Case enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado, enumStatusOperacao.DebitoMoedaNacionalLiquidado

            If Not xmlRemessa.documentElement.selectSingleNode("CO_NATU_ALTE") Is Nothing Then
                
                If xmlRemessa.documentElement.selectSingleNode("CO_NATU_ALTE").Text = "ME" And _
                   xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.DebitoMoedaNacionalLiquidado Then
                    
                    'Alteração de Valor de Moeda Estrangeira só é permitida para Operações em situação de Débito Moeda Nacional Liquidado
                    lngCodigoErroNegocio = 4233
                    GoTo ErrorHandler
                    
                ElseIf xmlRemessa.documentElement.selectSingleNode("CO_NATU_ALTE").Text = "MN" And _
                       xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado Then
                    
                    'Alteração de Valor de Moeda Nacional só é permitida para Operações em situação de Débito Moeda Estrangeira Liquidado
                    lngCodigoErroNegocio = 4234
                    GoTo ErrorHandler
                    
                Else
                    
                    strCampoValor = IIf(xmlRemessa.documentElement.selectSingleNode("CO_NATU_ALTE").Text = "MN", "VA_OPER_ATIV", "VA_MOED_ESTR_BMC")
    
                    'Alterar a operação
                    objOperacao.AlterarValorMoedasBMC xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                      xmlRemessa.documentElement.selectSingleNode("CO_NATU_ALTE").Text, _
                                                      xmlRemessa.documentElement.selectSingleNode(strCampoValor).Text
                                                      
                End If

            Else
                
                '4232 - Natureza Alteração deve ser igual a ME (Moeda Estrageira) ou MN (Moeda Nacional)
                lngCodigoErroNegocio = 4232
                GoTo ErrorHandler
            
            End If
        
        Case enumStatusOperacao.Liquidada
            If fgVerificaOperacaoCC(lngTipoOperacao) = True Then
               
                Call flAlterarValorLancamentoCC(xmlOperacao, xmlRemessa)
                
            Else
                
                '3115 - Situação da Operação não permite alteração.
                lngCodigoErroNegocio = 3115
                GoTo ErrorHandler
            
            End If

        Case Else
            '3115 - Situação da Operação não permite alteração.
            lngCodigoErroNegocio = 3115
            GoTo ErrorHandler
    
    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarOperacao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verifica Operacoes Autorizadas a alterar valor cc
Public Function fgVerificaOperacaoCC(ByVal plngTipoOperacao As Long) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT DISTINCT(TP_OPER) " & _
             " FROM   A8.TB_FCAO_SIST_TIPO_OPER " & _
             " WHERE  CO_FCAO_SIST = 16 " & _
             " AND    TP_OPER = " & plngTipoOperacao
             
    Set objRS = fgQuerySQL(strSQL)
    
    If Not objRS.EOF Then
        fgVerificaOperacaoCC = True
    Else
        fgVerificaOperacaoCC = False
    End If
    
    objRS.Close

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "fgVerificaOperacaoCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'=====================================================
'Autor........: Nick
'Data criação.: 16/09/2008
'Finalidade...: Alterar Valor de Lançamentos em CC
'Demanda......: GSB-08-03375-0A - Alteração de Lançamentos em CC
'=====================================================
Private Sub flAlterarValorLancamentoCC(ByRef pxmlOperacaoOriginal As MSXML2.DOMDocument40, _
                                       ByRef pxmlRemessaAlteracao As MSXML2.DOMDocument40)

Dim intModalidadeLiquidacao                 As Integer
Dim intTipoContraparte                      As Integer
Dim intFormaLiquidacao                      As Integer
Dim intMargemAtualizacao                    As Integer
Dim strParametroGeral                       As String
Dim strLancamentoCC                         As String
Dim datDataOperacao                         As Date
Dim dblValorLimite                          As Double
Dim dblValorLancamento                      As Double
Dim lngSequencialLancamento                 As Long
Dim lngSequencialLancamentoAtual            As Long
Dim lngStatusLancamento                     As Long
    
Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim xmlLancamentoCCContabil                 As MSXML2.DOMDocument40

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objRegra                                As A8LQS.clsRegra

    On Error GoTo ErrorHandler
    
    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlLancamentoCCContabil = CreateObject("MSXML2.DOMDocument.4.0")
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    Set objRegra = CreateObject("A8LQS.clsRegra")

    'Validação da Modalidade de Liquidação
    intModalidadeLiquidacao = 0
    If Not pxmlOperacaoOriginal.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intModalidadeLiquidacao = Val(pxmlOperacaoOriginal.selectSingleNode("//TP_LIQU_OPER_ATIV").Text)
    End If
    
    If intModalidadeLiquidacao <> enumTipoLiquidacao.SemModalidade Then
        lngCodigoErroNegocio = 4010
        GoTo ErrorHandler
    End If
    
    'Ivan - 02/08/2010
    'Alteracao efetuada de acordo com Ulysses (projeto Book3.02/Teimosinha) / A data da operacao nao pode ser menor que D-1
    'Validação da Data da Operação
    If Not pxmlOperacaoOriginal.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
        datDataOperacao = fgDtXML_To_Date(pxmlOperacaoOriginal.selectSingleNode("//DT_OPER_ATIV ").Text)
    End If

    If fgAdicionarDiasUteis(datDataOperacao, 1, Proximo) < fgDataHoraServidor(DataAux) Then
        lngCodigoErroNegocio = 4012
        GoTo ErrorHandler
    End If
    
    'Validação do Tipo de Contraparte
    intTipoContraparte = 0
    If Not pxmlOperacaoOriginal.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = Val(pxmlOperacaoOriginal.selectSingleNode("//TP_CNPT").Text)
    End If
    
    If intTipoContraparte <> enumTipoContraparte.Cliente1 Then
        lngCodigoErroNegocio = 4182
        GoTo ErrorHandler
    End If
        
    'Validação da Forma de Liquidação
    intFormaLiquidacao = 0
    If Not pxmlOperacaoOriginal.selectSingleNode("//CO_FORM_LIQU") Is Nothing Then
        intFormaLiquidacao = Val(pxmlOperacaoOriginal.selectSingleNode("//CO_FORM_LIQU").Text)
    End If
    
    If intFormaLiquidacao <> enumFormaLiquidacao.ContaCorrente Then
        lngCodigoErroNegocio = 4026
        GoTo ErrorHandler
    End If
        
    'Leitura do Lançamento em Conta Corrente original
    strLancamentoCC = objLancamentoCCContabil.LerTodos(pxmlOperacaoOriginal.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                       enumTipoLancamentoIntegracao.Normal, , , , , , , , , , , _
                                                       "NR_SEQU_LANC DESC")
    
    If strLancamentoCC = vbNullString Then
        '4341 - Lançamento em Conta Corrente não encontrado para alteração.
        lngCodigoErroNegocio = 4341
        GoTo ErrorHandler
    End If
    
    Call xmlLancamentoCCContabil.loadXML(strLancamentoCC)
    
    lngSequencialLancamento = 0
    If Not xmlLancamentoCCContabil.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
        lngSequencialLancamento = Val(xmlLancamentoCCContabil.selectSingleNode("//NR_SEQU_LANC").Text)
    End If
    
    dblValorLancamento = 0
    If Not xmlLancamentoCCContabil.selectSingleNode("//VA_LANC_CC") Is Nothing Then
        dblValorLancamento = fgVlrXml_To_Decimal(xmlLancamentoCCContabil.selectSingleNode("//VA_LANC_CC").Text)
    End If
    
    'Validação do Status do Lançamento Original
    lngStatusLancamento = 0
    If Not xmlLancamentoCCContabil.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        lngStatusLancamento = Val(xmlLancamentoCCContabil.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
    End If
    
    Select Case lngStatusLancamento
        Case enumStatusIntegracao.Disponível, _
             enumStatusIntegracao.ErroCC, _
             enumStatusIntegracao.ErroSaldoCC, _
             enumStatusIntegracao.Suspenso, _
             enumStatusIntegracao.Integrado, _
             enumStatusIntegracao.IntegradoCC
             
        Case Else
            '4194 - Status de lançamento conta corrente não permite alteração.
            lngCodigoErroNegocio = 4194
            GoTo ErrorHandler
            
    End Select
    
    'Validação da Margem do Valor a ser Alterado
    'Leitura ao Cadastro de Parâmetros Gerais no A7
    strParametroGeral = fgSelectVarchar4000(0, False)
    xmlParametroGeral.preserveWhiteSpace = True
    Call xmlParametroGeral.loadXML(strParametroGeral)
    
    'Obtendo a margem de atualização
    intMargemAtualizacao = 0
    If Not xmlParametroGeral.selectSingleNode("//MARGEM_PERC_ALTER_VALOR_CC/PERCENTUAL") Is Nothing Then
        intMargemAtualizacao = Val(xmlParametroGeral.selectSingleNode("//MARGEM_PERC_ALTER_VALOR_CC/PERCENTUAL ").Text)
    End If
    
    If intMargemAtualizacao = 0 Then
        '4342 - Margem percentual para alteração de lançamento em CC não cadastrada, ou inválida no Cadastro de Parâmetros Gerais do Sistema A7. O valor deve ser diferente de 0 (Zero).
        lngCodigoErroNegocio = 4342
        GoTo ErrorHandler
    End If
    
    'Calculando limite para alteração
    dblValorLimite = dblValorLancamento * intMargemAtualizacao / 100
    
    If fgVlrXml_To_Decimal(pxmlRemessaAlteracao.selectSingleNode("//VA_LANC_CC").Text) > dblValorLancamento + dblValorLimite Or _
       fgVlrXml_To_Decimal(pxmlRemessaAlteracao.selectSingleNode("//VA_LANC_CC").Text) < dblValorLancamento - dblValorLimite Then
        '4343 - Valor informado na remessa está fora da margem permitida para alteração do lançamento em CC.
        lngCodigoErroNegocio = 4343
        GoTo ErrorHandler
    End If
    
    'Incluindo TAGs necessárias para os próximos passos
    If pxmlOperacaoOriginal.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
        Call fgAppendNode(pxmlOperacaoOriginal, "MESG", "NR_SEQU_LANC", vbNullString)
    End If
    pxmlOperacaoOriginal.selectSingleNode("//NR_SEQU_LANC").Text = lngSequencialLancamento
    
    If pxmlOperacaoOriginal.selectSingleNode("//VA_LANC_CC") Is Nothing Then
        Call fgAppendNode(pxmlOperacaoOriginal, "MESG", "VA_LANC_CC", vbNullString)
    End If
    
    If pxmlOperacaoOriginal.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
        Call fgAppendNode(pxmlOperacaoOriginal, "MESG", "TP_LANC_ITGR", vbNullString)
    End If
    
    If pxmlOperacaoOriginal.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
        Call fgAppendNode(pxmlOperacaoOriginal, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
    End If
    
    Select Case lngStatusLancamento
        Case enumStatusIntegracao.Integrado, enumStatusIntegracao.IntegradoCC
            
            'Estorno do Lançamento Original
            pxmlOperacaoOriginal.selectSingleNode("//TP_LANC_ITGR").Text = enumTipoLancamentoIntegracao.Estorno
            
            Call objContaCorrente.Disponibilizar(pxmlOperacaoOriginal)
        
        Case enumStatusIntegracao.ErroSaldoCC
        
            pxmlOperacaoOriginal.selectSingleNode("//TP_LANC_ITGR").Text = enumTipoLancamentoIntegracao.Normal
             
        Case Else
            
            'Cancelamento do Lançamento CC Original
            Call objLancamentoCCContabil.AlterarStatus(pxmlOperacaoOriginal.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                       enumTipoLancamentoIntegracao.Normal, _
                                                       enumStatusIntegracao.Cancelado, _
                                                       vbNullString, , _
                                                       False, , , , _
                                                       lngSequencialLancamento)
        
            'Envio de mensagem de retorno ao sistema origem
            pxmlOperacaoOriginal.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.Cancelado
            Call objFluxoProcessamento.EnviarMensagemRetorno(pxmlOperacaoOriginal, False)
            
            pxmlOperacaoOriginal.selectSingleNode("//TP_LANC_ITGR").Text = enumTipoLancamentoIntegracao.Normal
    
    End Select
    
    'Disponibilização do lançamento em CC com o novo valor
    lngSequencialLancamentoAtual = lngSequencialLancamento
    lngSequencialLancamento = lngSequencialLancamento + 1
    pxmlOperacaoOriginal.selectSingleNode("//NR_SEQU_LANC").Text = lngSequencialLancamento
    pxmlOperacaoOriginal.selectSingleNode("//VA_LANC_CC").Text = pxmlRemessaAlteracao.selectSingleNode("//VA_LANC_CC").Text
    
    Call objContaCorrente.Disponibilizar(pxmlOperacaoOriginal)
             
    'Ivan - 02/08/2010
    'Retirado a integracao, pois agora e feita na funcao Incluir apos o Disponibilizar
    'Verificação de parâmetro para Integração Automática
    If objRegra.VerificarRegraAutomatica(pxmlOperacaoOriginal, enumFuncaoSistema.IntegracaoCC) Then
        Call objContaCorrente.Integrar(pxmlOperacaoOriginal.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                       enumIntegrarCCContabilidade.ContaCorrente, _
                                       False, _
                                       enumTipoLancamentoIntegracao.Normal, , _
                                       pxmlOperacaoOriginal.selectSingleNode("//VA_LANC_CC").Text, , , _
                                       lngSequencialLancamentoAtual, True)
    
        'Envio de mensagem de retorno ao sistema origem
    '    pxmlOperacaoOriginal.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.EnviadoCC
    '    Call objFluxoProcessamento.EnviarMensagemRetorno(pxmlOperacaoOriginal, False)
    End If
    
    Set xmlParametroGeral = Nothing
    Set xmlLancamentoCCContabil = Nothing
    Set objLancamentoCCContabil = Nothing
    Set objContaCorrente = Nothing
    Set objFluxoProcessamento = Nothing
    Set objRegra = Nothing

    Exit Sub

ErrorHandler:
    Set xmlParametroGeral = Nothing
    Set xmlLancamentoCCContabil = Nothing
    Set objLancamentoCCContabil = Nothing
    Set objContaCorrente = Nothing
    Set objFluxoProcessamento = Nothing
    Set objRegra = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarValorLancamentoCC Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Controlar o fluxo de reativacao de uma operação

Private Function flProcessaReativacao(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objRegra                                As A8LQS.clsRegra
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMesgOrig                             As MSXML2.DOMDocument40
Dim lngTipoLiquidacao                       As enumTipoLiquidacao
Dim strCampoValor                           As String
Dim lngCodigoRetornoVerificacao             As Long

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
                                                   
    If xmlOperacao.Text = vbNullString Then
        '3114 - Operação não localizada no SLCC para alteração.
        lngCodigoErroNegocio = 3114
        GoTo ErrorHandler
    End If
    
    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        Case enumStatusOperacao.Cancelada
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao Then
                
                Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
                
                Set xmlMesgOrig = objMensagemEnviadaRecebida.LerXMLMensagem(, , , "BMC0002R1" _
                                                                            , , , xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                
                If xmlMesgOrig.xml <> vbNullString Then
                    'Altera o status da operacao para reativacao solicitada
                    Call objOperacao.AlterarStatusOperIndisp(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                             enumStatusOperacao.ReativacaoSolicitada)
                    'Verifica regra para Confirmacao
                    Set objRegra = CreateObject("A8LQS.clsRegra")
                    If objRegra.VerificarRegraAutomatica(xmlOperacao, enumFuncaoSistema.ConfirmarReativacao, lngCodigoRetornoVerificacao) Then
                        Call objOperacao.AlterarStatusOperIndisp(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                 enumStatusOperacao.ConcordanciaReativacaoAutomatica)
                        'Verifica regra para Liberacao
                        If objRegra.VerificarRegraAutomatica(xmlOperacao, enumFuncaoSistema.LiberarReativacao, lngCodigoRetornoVerificacao) Then
                            'Atualiza a Status da Operacao
                            xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaReativacaoAutomatica
                            
                            fgAppendNode xmlOperacao, "MESG", "CO_MESG", "BMC0003"
                            fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", enumTipoMensagemLQS.RegistroOperacoesBMCRetorno
                            fgAppendNode xmlOperacao, "MESG", "NU_CTRL_BMC_ORIG", xmlMesgOrig.selectSingleNode("//NU_CTRL_CAMR").Text
                    
                            Call flLiberarOperacao(xmlOperacao, True)
                        
                        End If
                    
                    End If
                Else
                    '4306 - Mensagem original nao localizada para envio de mensagem de reativacao
                    lngCodigoErroNegocio = 4306
                    fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.MensagemNaoLocalizada
                    GoTo ErrorHandler
                End If
            
            End If

        Case Else
            'Somente operacoes com situacao CANCELADA podem ser reativadas
            lngCodigoErroNegocio = 4305
            GoTo ErrorHandler

    End Select

    'Armazenar a mensagem original
    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    objOperacaoMensagemInterna.Incluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                       CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                       xmlRemessa
    
    Set objOperacao = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objRegra = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgOrig = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set objOperacao = Nothing
    Set objOperacaoMensagemInterna = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objRegra = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgOrig = Nothing
    
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaReativacao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Validar o complemento de um Contrato de SWAP

Private Function flValidarComplementoContratoSWAP(ByRef pxmlremessa As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim rsQuery                                 As ADODB.Recordset
Dim lngErro                                 As Long
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strNumComando                           As String
Dim strCodigoVeiculoLegal                   As String
Dim vntIdentificadorParticipante            As Variant
Dim vntIdentificadorContraparte             As Variant
Dim strDHInicio                             As String
Dim strDHFim                                As String

On Error GoTo ErrorHandler
    
    With pxmlremessa
        
        Set xmlNode = .selectSingleNode("//NU_COMD_OPER")
        strNumComando = Trim$(xmlNode.Text)
        
        Set xmlNode = .selectSingleNode("//CO_PARP_CAMR")
        vntIdentificadorParticipante = Val(xmlNode.Text)
        
        Set xmlNode = .selectSingleNode("//CO_CNPT_CAMR")
        vntIdentificadorContraparte = Val(xmlNode.Text)
        
    End With
    
    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
    
    strSQL = "SELECT    A.CO_ULTI_SITU_PROC " & _
             "  FROM    A8.TB_OPER_ATIV A,  " & _
             "          A8.TB_VEIC_LEGA B   " & _
             " Where    A.CO_VEIC_LEGA          = B.CO_VEIC_LEGA " & _
             "   AND    A.CO_CNPT_CAMR          =  " & vntIdentificadorContraparte & _
             "   AND    B.ID_PART_CAMR_CETIP    =  " & vntIdentificadorParticipante & _
             "   AND    A.NU_COMD_OPER          = '" & strNumComando & "'" & _
             "   AND    A.TP_OPER               =  " & enumTipoOperacaoLQS.RegDadosComplemContratoSWAP & _
             "   AND DT_OPER_ATIV BETWEEN " & strDHInicio & _
             "   AND " & strDHFim
       
    Set rsQuery = fgQuerySQL(strSQL)
    
    If Not rsQuery.EOF Then
        
        Select Case rsQuery!CO_ULTI_SITU_PROC
            
            Case enumStatusOperacao.Cancelada, _
                 enumStatusOperacao.CanceladaOrigem
                '3111 - Situação do complemento de contrato SWAP inválido.
                lngCodigoErroNegocio = 3111
                GoTo ErrorHandler
                
        End Select
        
    End If
    
    rsQuery.Close
    
    Set rsQuery = Nothing

    Exit Function
    
ErrorHandler:

    Set rsQuery = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidarComplementoContratoSWAP Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Validar dados de alteração para Eventos de Juros CETIP

Private Function flValidarEventosJurosCETIP(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    strOperacao = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                       xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If strOperacao = vbNullString Then
        Exit Function
    End If

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    xmlOperacao.loadXML strOperacao

    If Val("0" & xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bruta Then
        If xmlOperacao.documentElement.selectSingleNode("//TP_LIQU_OPER_ATIV").Text <> enumTipoLiquidacao.Bilateral Then
            '3113 - Modalidade de Liquidação da operação inválida para alteração.
            lngCodigoErroNegocio = 3113
            GoTo ErrorHandler
        End If
    ElseIf Val("0" & xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) <> enumTipoLiquidacao.Bilateral And _
            Val("0" & xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text) <> enumTipoLiquidacao.SemModalidade Then
        '3112 - Modalidade de Liquidação para alteração invalida
        lngCodigoErroNegocio = 3112
        GoTo ErrorHandler
    End If

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidarEventosJurosCETIP Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o recebimento de uma operação compromissada com data retroativa

Public Function VerificarCompromissadaRetroativa(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoIda                          As MSXML2.DOMDocument40
Dim strXML                                  As String

    On Error GoTo ErrorHandler

    VerificarCompromissadaRetroativa = False

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacaoIda = CreateObject("MSXML2.DOMDocument.4.0")

    strXML = objOperacao.LerTodos(, , , CLng("0" & xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text), _
                                  fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_ORIG").Text), _
                                  enumTipoOperacaoLQS.CompromissadaIda, _
                                  , , , _
                                  xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlOperacaoIda.loadXML strXML

    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacaoIda.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

    If xmlOperacao.Text = vbNullString Then
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text And _
        xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text Then
        VerificarCompromissadaRetroativa = True
    End If

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoIda = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarCompromissadaRetroativa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Operações geradas via entrada manual e tenham conciliação parem com o status de A CONCILIAR
Private Function flAlternativoEntradaManual(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlControleProcessoOperacaoAux          As MSXML2.DOMDocument40
Dim strXML                                  As String

On Error GoTo ErrorHandler

    Set xmlControleProcessoOperacaoAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    strXML = objProcessoOperacao.LerTodos("RecebimentoOperacao", _
                                          CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text), _
                                          enumIndicadorSimNao.Nao)

    Set objProcessoOperacao = Nothing
    If strXML = "" Then
        Exit Function
    End If
            
    xmlControleProcessoOperacaoAux.loadXML strXML

    'Alterado para as operações geradas via entrada manual e tenham conciliação parem com o status de A CONCILIAR
    If xmlControleProcessoOperacaoAux.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CNCL").Text = enumIndicadorSimNao.Sim Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AConciliar
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.AConciliar
        Set objOperacao = Nothing
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Chamada para a rotina de alteração de dados de conta corrente

Private Function flAlterarDadosContaCorrente(ByRef pxmlremessa As MSXML2.DOMDocument40)

'Pikachu - 21/10/2004
'Release 2 - A8
'Demanda 967 - Alteração dos controles de operações recebidas pelo SLCC,
'Alteração de dados de conta corrente

Dim objLancamentoCC                         As A8LQS.clsLancamentoCCContabil
Dim objOperacaoInterna                      As A8LQS.clsOperacaoMensagemInterna
Dim objLegado                               As A8LQS.clsLegado
Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim vntNuSequOperAtiv                       As Variant
Dim lngTipoOperacao                         As Long
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    'Verifica se existe operação original para alteração dados Conta Corrente
    strSQL = "  SELECT  NU_SEQU_OPER_ATIV ,  " & vbCrLf & _
             "          A.TP_OPER             " & vbCrLf & _
             "  FROM    A8.TB_OPER_ATIV A, " & vbCrLf & _
             "          A8.TB_TIPO_OPER B  " & vbCrLf & _
             "  WHERE   A.TP_OPER           =   B.TP_OPER " & vbCrLf & _
             "    AND   DT_OPER_ATIV        =  " & fgDtXML_To_Oracle(pxmlremessa.selectSingleNode("//DT_OPER_ATIV").Text) & vbCrLf & _
             "    AND   CO_VEIC_LEGA        = '" & pxmlremessa.selectSingleNode("//CO_VEIC_LEGA").Text & "'" & vbCrLf & _
             "    AND   IN_OPER_DEBT_CRED   =  " & pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text & vbCrLf & _
             "    AND   VA_OPER_ATIV        =  " & Replace(pxmlremessa.selectSingleNode("//VA_OPER_ATIV").Text, ",", ".") & vbCrLf & _
             "    AND   TP_MESG_RECB_INTE   = '" & Val(pxmlremessa.selectSingleNode("//TP_MESG_ORIG").Text) & "'" & vbCrLf & _
             "    AND   CO_OPER_ATIV        = '" & pxmlremessa.selectSingleNode("//CO_OPER_ATIV").Text & "'"
    
    Set objRS = fgQuerySQL(strSQL)
    
    vntNuSequOperAtiv = objRS!NU_SEQU_OPER_ATIV
    lngTipoOperacao = objRS!TP_OPER
    
    objRS.Close
    
    Set objOperacaoInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    
    Call objOperacaoInterna.Incluir(vntNuSequOperAtiv, _
                                    enumTipoMensagemLQS.AlteracaoDadosContaCorrente, _
                                    enumTipoSolicitacao.Alteracao, _
                                    pxmlremessa)
                                    
    Set objOperacaoInterna = Nothing
    
    
    Set objLancamentoCC = CreateObject("A8LQS.clsLancamentoCCContabil")
    Call objLancamentoCC.AlterarContaCorrente(vntNuSequOperAtiv, _
                                              enumTipoLancamentoIntegracao.Normal, _
                                              pxmlremessa.selectSingleNode("//CO_AGEN").Text, _
                                              pxmlremessa.selectSingleNode("//NU_CC").Text)
    Set objLancamentoCC = Nothing
    
    If pxmlremessa.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
        Call fgAppendNode(pxmlremessa, "MESG", "NU_SEQU_OPER_ATIV", vntNuSequOperAtiv)
    End If
    
    Set objLegado = CreateObject("A8LQS.clsLegado")
    strMensagem = objLegado.EnviarRetorno(pxmlremessa)
    EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, False, False
    Set objLegado = Nothing
    
    
    Exit Function
ErrorHandler:
    
    Set objLegado = Nothing
    Set objLancamentoCC = Nothing
    Set objOperacaoInterna = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarDadosContaCorrente Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Sub flAlteracaoRemessaCETIP(ByRef pxmlremessa As MSXML2.DOMDocument40)

Dim xmlParametrosGerais                     As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
        
    If Not pxmlremessa.selectSingleNode("//CO_PARP_CAMR") Is Nothing Then
        If pxmlremessa.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
            Call fgAppendNode(pxmlremessa, pxmlremessa.documentElement.nodeName, "CO_CNTA_CUTD_SELIC_VEIC_LEGA", pxmlremessa.selectSingleNode("//CO_PARP_CAMR").Text)
        Else
            pxmlremessa.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text = pxmlremessa.selectSingleNode("//CO_PARP_CAMR").Text
        End If
    End If
        
    '-----------------------------------------------------------------------------
    'Pikachu - 09/03/2005
    'Alteração temporária para o SAC e SIGOM, pois estão com valor FIXO - QF
    'Aguardar SAC corrigir
    '-----------------------------------------------------------------------------
    
    If Not pxmlremessa.selectSingleNode("//CO_SUB_TIPO_ATIV") Is Nothing Then
        If Trim$(pxmlremessa.selectSingleNode("//CO_SUB_TIPO_ATIV").Text) = "QF" Then
            pxmlremessa.selectSingleNode("//CO_SUB_TIPO_ATIV").Text = "CFA"
        End If
    End If
        
    If Not pxmlremessa.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val(pxmlremessa.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.DespesasCETIP Then
            
            If pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
                Call fgAppendNode(pxmlremessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito)
            Else
                pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            End If
        
        ElseIf Val(pxmlremessa.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
        
            Set xmlParametrosGerais = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlParametrosGerais.loadXML(fgSelectVarchar4000(0, False))
    
            If Not xmlParametrosGerais.selectSingleNode("//ALTERAR") Is Nothing Then
                If UCase(xmlParametrosGerais.selectSingleNode("//ALTERAR").Text) = "S" Then
                
                    If Not pxmlremessa.selectSingleNode("//TP_FLUX_CAIX") Is Nothing Then
                        pxmlremessa.selectSingleNode("//TP_FLUX_CAIX").Text = "0"
                    End If
                
                End If
            End If
            
            Set xmlParametrosGerais = Nothing
        
        End If
    End If
        
    If Not pxmlremessa.selectSingleNode("//PE_TAXA_NEGO_CETIP") Is Nothing Then
        If pxmlremessa.selectSingleNode("//PE_TAXA_NEGO") Is Nothing Then
            Call fgAppendNode(pxmlremessa, "MESG", "PE_TAXA_NEGO", pxmlremessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text)
        End If
    End If
        
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlteracaoRemessaCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flAlteracaoRemessaCCR(ByRef xmlRemessa As MSXML2.DOMDocument40)

Dim xmlParametrosGerais                     As MSXML2.DOMDocument40
Dim intLayouEntrada                         As Integer
Dim strRemessaAux                           As String

On Error GoTo ErrorHandler
        
    intLayouEntrada = 0
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        intLayouEntrada = Val(xmlRemessa.selectSingleNode("//TP_MESG").Text)
    End If
            
    fgAppendNode xmlRemessa, "MESG", "CO_ISPB_IF_CNPT", xmlRemessa.selectSingleNode("//CNPJBaseEntRespons").Text
    
    If Not xmlRemessa.selectSingleNode("//TpComerc") Is Nothing Then
        
        If xmlRemessa.selectSingleNode("//TpInstntoCCR") Is Nothing Then
            strRemessaAux = xmlRemessa.selectSingleNode("//TpComerc").Text & _
                            "   "
        
        Else
            strRemessaAux = xmlRemessa.selectSingleNode("//TpComerc").Text & _
                            xmlRemessa.selectSingleNode("//TpInstntoCCR").Text
        End If
        
        If UCase(Trim(xmlRemessa.selectSingleNode("//TpComerc").Text)) = "EX" Then
            fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Credito
        Else
            fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoDebitoCredito.Debito
        End If
                        
    End If
    
    fgAppendNode xmlRemessa, "MESG", "NU_ATIV_MERC", strRemessaAux
                    
    If Not xmlRemessa.selectSingleNode("//CodReemb") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_CTRL_MESG_SPB_ORIG", xmlRemessa.selectSingleNode("//CodReemb").Text
    End If
            
     If intLayouEntrada = enumTipoMensagemBUS.EmissaoOperacaoCCR Then
         
         If Not xmlRemessa.selectSingleNode("//DtEms") Is Nothing Then
             
'             If Trim(xmlRemessa.selectSingleNode("//DtEms").Text) = "" Then
'                 fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", fgDt_To_Xml(fgDataHoraServidor(DataAux))
'             Else
'                 If xmlRemessa.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
'                     fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.selectSingleNode("//DtEms").Text
'                 Else
'                     xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//DtEms").Text
'                 End If
'             End If
         Else
             fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", fgDt_To_Xml(fgDataHoraServidor(DataAux))
         End If
                             
                             
         If Not xmlRemessa.selectSingleNode("//DtExprc") Is Nothing Then
             If xmlRemessa.selectSingleNode("//DT_VENC_ATIV") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "DT_VENC_ATIV", xmlRemessa.selectSingleNode("//DtExprc").Text
             Else
                 xmlRemessa.selectSingleNode("//DT_VENC_ATIV").Text = xmlRemessa.selectSingleNode("//DtExprc").Text
             End If
         End If
         
         If xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
             fgAppendNode xmlRemessa, "MESG", "VA_OPER_ATIV", xmlRemessa.selectSingleNode("//VlrEms").Text
         Else
             xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//VlrEms").Text
         End If

     
     ElseIf intLayouEntrada = enumTipoMensagemBUS.NegociacaoOperacaoCCR Then
         
'         If xmlRemessa.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
'             fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.selectSingleNode("//DtNegc").Text
'         Else
'             xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//DtNegc").Text
'         End If
         
         If xmlRemessa.selectSingleNode("//DT_VENC_ATIV") Is Nothing Then
             fgAppendNode xmlRemessa, "MESG", "DT_VENC_ATIV", xmlRemessa.selectSingleNode("//DtVenc").Text
         Else
             xmlRemessa.selectSingleNode("//DT_VENC_ATIV").Text = xmlRemessa.selectSingleNode("//DtVenc").Text
         End If
         
         If xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
             fgAppendNode xmlRemessa, "MESG", "VA_OPER_ATIV", xmlRemessa.selectSingleNode("//VlrNegc").Text
         Else
             xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//VlrNegc").Text
         End If
     
     ElseIf intLayouEntrada = enumTipoMensagemBUS.DevolucaoRecolhimentoEstornoReembolsoCCR Then
         
         If xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
             fgAppendNode xmlRemessa, "MESG", "VA_OPER_ATIV", xmlRemessa.selectSingleNode("//VlrDevRecolht_EstReemb").Text
         Else
             xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text = xmlRemessa.selectSingleNode("//VlrDevRecolht_EstReemb").Text
         End If
    End If

        
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlteracaoRemessaCCR Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub flAlteracaoRemessaCorretoras(ByRef pxmlremessa As MSXML2.DOMDocument40)

    On Error GoTo ErrorHandler

    If Val(pxmlremessa.selectSingleNode("//CO_PROD").Text) = 254 Or _
       Val(pxmlremessa.selectSingleNode("//CO_PROD").Text) = 467 Then
        pxmlremessa.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMD
    Else
        pxmlremessa.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CLBCAcoes
    End If
        
    Exit Sub

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlteracaoRemessaCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Controlar o fluxo de exeção das despesas CETIP
Private Function flExcecaoDespesasCETIP(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                        ByRef plngStatusOperacao As Long, _
                                        ByRef pvntSequenciaOperacao As Variant, _
                                        ByRef plngTipoOperacao As Long) As Boolean

Dim strXML                                  As String
Dim xmlDespesas                             As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim objRemessaRejeitada                     As A8LQS.clsRemessaRejeitada
Dim objLegado                               As A8LQS.clsLegado

    On Error GoTo ErrorHandler

    Set xmlDespesas = CreateObject("MSXML2.DOMDocument.4.0")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    strXML = objOperacao.LerTodos(xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                  xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If strXML = vbNullString Then
        
        If Not xmlRemessa.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
            plngStatusOperacao = enumStatusOperacao.ManualEmSer
        Else
            plngStatusOperacao = enumStatusOperacao.EmSer
        End If

        flExcecaoDespesasCETIP = True
        Exit Function
    End If

    xmlDespesas.loadXML strXML

    If CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Or _
        CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Concordancia Or _
        CLng(xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaAutomatica Then
        plngStatusOperacao = CLng("0" & xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/CO_ULTI_SITU_PROC").Text)
        pvntSequenciaOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text
        plngTipoOperacao = xmlDespesas.documentElement.selectSingleNode("Grupo_Operacao/TP_OPER").Text
        flExcecaoDespesasCETIP = True
    Else
        '3132 - Solicitação de Complementação Inválida para Despesas CETIP.
        lngCodigoErroNegocio = 3132
        GoTo ErrorHandler
    End If

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set xmlDespesas = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcecaoDespesasCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


Private Sub flAlterarStatusBMC0112(ByVal pstrChaveBMC0112 As String, _
                                   ByVal pvntSEQ_OPER_ATIV As Variant)

Dim objMensagem                             As A8LQS.clsProcessoMensagemBMA

On Error GoTo ErrorHandler
        
    Set objMensagem = CreateObject("A8LQS.clsProcessoMensagemBMA")
    Call objMensagem.AtualizaStatusBMA0112(pstrChaveBMC0112, pvntSEQ_OPER_ATIV)
    Set objMensagem = Nothing

    Exit Sub
ErrorHandler:
    
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarStatusBMC0112 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub


'KIDA - CBLC - 03/10/2008
Private Function flCancelarBrutaCBLC(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objOperacaoMensagemInterna              As A8LQS.clsOperacaoMensagemInterna
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.Text = vbNullString Then
        '3009 - Operação não localizada no SLCC para cancelamento.
        lngCodigoErroNegocio = 3009
        fgGerarAlerta xmlRemessa, , , enumFatorGeradorAlerta.OperacaoNaoLocalizada
        GoTo ErrorHandler
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.EmSer And _
       xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaBackoffice Then
        '3025 - Situação da Operação não permite cancelamento.
        lngCodigoErroNegocio = 3025
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        
        Case enumStatusOperacao.ConcordanciaBackoffice
             
            If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoBrutaCBLC Then
                If fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                End If
                
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
                
                'Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoRealizadoSolicitado)
                
            Else
                If fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                End If
                
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            End If
                         
            Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text))
            
            If xmlConciliacao.xml <> vbNullString Then
                Call objMensagemEnviadaRecebida.AlterarStatus(xmlConciliacao.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                              xmlConciliacao.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                              enumStatusMensagem.AConciliar, _
                                                              vbNullString, _
                                                              , _
                                                              False)
            End If
            
            objConciliacao.ExcluirConciliacaoOperacao "", xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
            
            
            
        Case enumStatusOperacao.EmSer, _
             enumStatusOperacao.ManualEmSer
            
            'KIDA - CBCL - 03/10/2008
            If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoBrutaCBLC Then
                If fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                    Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.EstornoPrevisto)
                End If
            Else
                If fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                End If
            End If
            
    End Select
    
    

    Call objOperacao.AlterarStatusAcao( _
            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
            enumStatusOperacao.CanceladaOrigem, _
            vbNullString, , , _
            False)

    Set objOperacaoMensagemInterna = CreateObject("A8LQS.clsOperacaoMensagemInterna")
    Call objOperacaoMensagemInterna.Incluir(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text), _
                                            CLng(xmlRemessa.documentElement.selectSingleNode("TP_SOLI").Text), _
                                            xmlRemessa)
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objOperacaoMensagemInterna = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarRodaDolarBMC Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'KIDA - CBLC - EVENTOS
Private Function flCancelarEventoCBLC(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(, xmlRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text, _
                                                   xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

    If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.EmSer And _
       xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaBackoffice And _
       xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.ConcordanciaBackofficeAutomatico Then
        '3025 - Situação da Operação não permite cancelamento.
        lngCodigoErroNegocio = 3025
        GoTo ErrorHandler
    End If
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Select Case xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text
        
        Case enumStatusOperacao.ConcordanciaBackoffice, _
             enumStatusOperacao.ConcordanciaBackofficeAutomatico
             
            If xmlOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC Then
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            Else
                If fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_NEGO").Text) = fgDataHoraServidor(DataAux) Then
                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoPrevisto)
                End If
                
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
            End If
            
            objConciliacao.ExcluirConciliacaoOperacao "", xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
            
        Case enumStatusOperacao.EmSer, _
             enumStatusOperacao.ManualEmSer
            
    End Select
    

    Call objOperacao.AlterarStatusAcao( _
            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
            enumStatusOperacao.CanceladaOrigem, _
            vbNullString, , , _
            False)

    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objConciliacao = Nothing
    
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    'Set objOperacaoMensagemInterna = Nothing
    Set objConciliacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarEventoCBLC", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




Private Function flEnviarPrivistoMEPJ(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim strMensagem                             As String

On Error GoTo ErrorHandler

    flEnviarPrivistoMEPJ = False

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    strMensagem = objGestaoCaixa.EnviarPrevisaoME(xmlRemessa)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing
    Set objGestaoCaixa = Nothing

    flEnviarPrivistoMEPJ = True

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarRealizadoPJ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
