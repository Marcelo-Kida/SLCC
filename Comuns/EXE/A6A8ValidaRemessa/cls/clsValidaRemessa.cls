VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValidaRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Validar as remessas dos sistemas A6 e A8

Option Explicit

'Validar as remessas do sistema A8

Public Function ValidarMensagemA8(ByRef pstrRemessa As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMensagem.loadXML pstrRemessa

    Select Case Val("0" & xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text)

        Case enumTipoMensagemLQS.LiquidacaoEventosBMA

            ValidarMensagemA8 = fgConsisteLiquidacaoEventos_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.TransferenciaLDL_BMA

            ValidarMensagemA8 = fgConsisteOperacaoTransfCamaraLDL_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.IntermediacaoOperacoesInternasBMA

            ValidarMensagemA8 = fgConsisteIntermediacaoOperInterna_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.LiquidacaoOperacoesBMA

            ValidarMensagemA8 = fgConsisteLiquidacaoOperacao_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.LiquidacaoFisicaOperacaoBMA

            ValidarMensagemA8 = fgConsisteLiquidacaoFisicaOperacao(xmlMensagem)

        Case enumTipoMensagemLQS.EspecificacaoOperacoesBMA

            ValidarMensagemA8 = fgConsisteEspecificacaoOperacoes_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.RegistroOperacaoBMA

            ValidarMensagemA8 = fgConsisteRegistroOperacao_BMA(xmlMensagem)

        Case enumTipoMensagemLQS.OperacoesComCorretorasCETIP, enumTipoMensagemLQS.MovimentacoesInstFinancCETIP, _
             enumTipoMensagemLQS.ResgateFundoInvestimentoCETIP, enumTipoMensagemLQS.MovimentacoesCustodiaCETIP, _
             enumTipoMensagemLQS.ExercicioDesistenciaCETIP, enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP, _
             enumTipoMensagemLQS.ConversaoPermutaValorImobCETIP, enumTipoMensagemLQS.EspecificacaoQuantidadesCotasCETIP, _
             enumTipoMensagemLQS.OperacaoDefinitivaCETIP, enumTipoMensagemLQS.OperacaoCompromissadaCETIP, _
             enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP, enumTipoMensagemLQS.OperacaoRetornoAntecipacaoCETIP, _
             enumTipoMensagemLQS.OperacaoRetencaoIRF_CETIP, enumTipoMensagemLQS.RegistroContratoSWAP, _
             enumTipoMensagemLQS.RegistroOperacaoesCETIP, enumTipoMensagemLQS.RegistroContratoTermoCETIP, _
             enumTipoMensagemLQS.ExercicioOpcaoContratoSwapCETIP, enumTipoMensagemLQS.AntecipacaoResgateContratoDerivativoCETIP, _
             enumTipoMensagemLQS.EventoJurosCETIP, enumTipoMensagemLQS.DespesasCETIP, _
             enumTipoMensagemLQS.LancamentoPU_CETIP, enumTipoMensagemLQS.MovimentacoesContratoDerivativo

             ValidarMensagemA8 = fgConsisteMensagemCETIP(xmlMensagem)

        'Pikachu - 21/10/2004
        'Release 2 - A8
        'Demanda 967 - Alteração dos controles de operações recebidas pelo SLCC,
        'Alteração de dados de conta corrente

        Case enumTipoMensagemLQS.AlteracaoDadosContaCorrente

            ValidarMensagemA8 = fgConsisteMensagemAlteracaoCC(xmlMensagem)
        '--------------------------------------------------------------------------------

        Case enumTipoMensagemLQS.LivreMovimentacao

            ValidarMensagemA8 = fgConsisteMensagemLivreMovimentacao(xmlMensagem)

        Case enumTipoMensagemLQS.RegistroLiquidacaoMultilateralCBLC, _
             enumTipoMensagemLQS.RegistroLiquidacaoBrutaCBLC, _
             enumTipoMensagemLQS.RegistroLiquidacaoEventoCBLC

            ValidarMensagemA8 = fgConsisteMensagemCBLC(xmlMensagem)

        Case enumTipoMensagemLQS.RegistroLiquidacaoMultilateralBMF

            ValidarMensagemA8 = fgConsisteMensagemBMF(xmlMensagem)

        Case enumTipoMensagemLQS.RegistroOperacoesBMC, _
             enumTipoMensagemLQS.LiquidacaoMultilateralBMC, _
             enumTipoMensagemLQS.TransferenciasBMC, _
             enumTipoMensagemLQS.DespesasBMC, _
             enumTipoMensagemLQS.RegistroOperacoesRodaDolar

            ValidarMensagemA8 = fgConsisteMensagemBMC(xmlMensagem)
        
                'Inclusão de str 150 || 15/03/07
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Case enumTipoMensagemLQS.EnvioTEDClientes
            ValidarMensagemA8 = fgConsisteMensagemTED(xmlMensagem)

        'Inclusao de str 154 08/02/07
        Case enumTipoMensagemLQS.EnvioPagDespesas
            ValidarMensagemA8 = fgConsisteMensagemPagDespesas(xmlMensagem)

        Case enumTipoMensagemLQS.MensagemNZ, _
             enumTipoMensagemLQS.MensagemErroNZ, _
             enumTipoMensagemLQS.MensagemCCDV, _
             enumTipoMensagemLQS.MensagemCCBG, _
             enumTipoMensagemLQS.MensagemErroPZ, _
             enumTipoMensagemLQS.MensagemR1PZ, _
             enumTipoMensagemLQS.MensagemR2PZ, _
             enumTipoMensagemLQS.MensagemSTR0010R2PZA8
            
            ValidarMensagemA8 = vbNullString
        
        Case Else

            ValidarMensagemA8 = fgConsisteMensagemA8(xmlMensagem)

    End Select

    ValidarMensagemA8 = fgConsisteDatasRemessa(xmlMensagem, ValidarMensagemA8)
    
    
    Set xmlMensagem = Nothing
    
    Exit Function

ErrorHandler:
    
    Set xmlMensagem = Nothing
    
    ValidarMensagemA8 = Err.Description

End Function


'Validar as remessas do sistema A6

Public Function ValidarMensagemA6(ByRef pstrRemessa As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlErroNegocio                          As MSXML2.DOMDocument40
Dim objMovimentoCaixa                       As A6A8ValidaRemessa.clsMovimentoCaixa

Dim strErro                                 As String

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMensagem.loadXML pstrRemessa
    
'    fgGravaArquivo "LOGA6_ValidarMensagemA6_TP_MESG_" & xmlMensagem.selectSingleNode("/MESG/TP_MESG").Text & "_", vbNullString
    
    Select Case Val(xmlMensagem.selectSingleNode("/MESG/TP_MESG").Text)
        Case 100
            ValidarMensagemA6 = fgConsisteRemessaEstatística(xmlMensagem)
        Case 101
            ValidarMensagemA6 = fgConsisteRemessaOperacao(xmlMensagem)
            
            If ValidarMensagemA6 = vbNullString Then
                Set objMovimentoCaixa = CreateObject("A6A8ValidaRemessa.clsMovimentoCaixa")
                
                With objMovimentoCaixa
                    pstrRemessa = xmlMensagem.xml
                    Call .ObterItemCaixa(pstrRemessa, CaixaFuturo)
                    If .VerificarMovimentoSubReserva(pstrRemessa) Then
                        Call .ObterItemCaixa(pstrRemessa, CaixaSubReserva)
                    End If
                End With
                
                Set objMovimentoCaixa = Nothing
            End If
        Case 102
            ValidarMensagemA6 = fgConsisteRemessaSaldoFechamento(xmlMensagem)
        Case 103
            ValidarMensagemA6 = fgConsisteRemessaViaSistemaA8(xmlMensagem)
            
            If ValidarMensagemA6 = vbNullString Then
                Set objMovimentoCaixa = CreateObject("A6A8ValidaRemessa.clsMovimentoCaixa")
                
                With objMovimentoCaixa
                    
                    pstrRemessa = xmlMensagem.xml
                    Call .ObterItemCaixaOnLine(pstrRemessa, _
                                               enumTipoCaixa.CaixaSubReserva, _
                                               gstrGrupoVeiculoLegal, _
                                               glngTipoBackOffice)
                    
                    If xmlMensagem.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.Previsto Or _
                       xmlMensagem.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevisto Or _
                       xmlMensagem.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.PrevistoCompromIda Or _
                       xmlMensagem.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevistoCompromIda Then
                        
                        Call .ObterItemCaixaOnLine(pstrRemessa, _
                                                   enumTipoCaixa.CaixaFuturo, _
                                                   gstrGrupoVeiculoLegal, _
                                                   glngTipoBackOffice)
                    
                    End If
                    
                End With
                
                Set objMovimentoCaixa = Nothing
            End If
        Case 110
            '*******************************************************
            'Verificar se ira validar remessa Veiculo Legal
            '*******************************************************
        Case Else
            Set xmlErroNegocio = CreateObject("MSXML2.DOMDocument.4.0")
            'Tipo de mensagem inválida
            fgAdicionaErro xmlErroNegocio, 4003
            ValidarMensagemA6 = xmlErroNegocio.xml
    End Select

    Set xmlMensagem = Nothing
    Set xmlErroNegocio = Nothing
    Set objMovimentoCaixa = Nothing


    Exit Function
ErrorHandler:

    strErro = Err.Description
    fgGravaArquivo "LOGA6_ValidarMensagemA6_Error", strErro

    Set xmlMensagem = Nothing
    Set xmlErroNegocio = Nothing
    Set objMovimentoCaixa = Nothing

    ValidarMensagemA6 = strErro

End Function

Public Function AdicionaUsuario(ByVal pstrUsuario As String, _
                                ByVal pstrXMLSegresagao As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
'    With rsControleAcesso
'        .AddNew
'        .Fields!Usuario = Left(pstrUsuario, 8)
'        .Fields!XMLSegregacao = pstrXMLSegresagao
'        .Update
'    End With
    
    AdicionaUsuario = vbNullString
    
    Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "AdicionaUsuario", 0
End Function

Public Function VerificaUsuario(ByVal pstrUsuario As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    VerificaUsuario = vbNullString
    
'    If rsControleAcesso Is Nothing Then
'        Set rsControleAcesso = CreateObject("ADODB.Recordset")
'
'        'Inicializa o recordset que controla o usuário
'        With rsControleAcesso.Fields
'            .Append "Usuario", adChar, 8
'            .Append "XMLSegregacao", adChar, 65000
'            rsControleAcesso.Open
'        End With
'        VerificaUsuario = vbNullString
'        Exit Function
'    End If
'
'    If Not rsControleAcesso Is Nothing Then
'        rsControleAcesso.Filter = "Usuario = '" & pstrUsuario & "'"
'        If rsControleAcesso.RecordCount > 0 Then
'            VerificaUsuario = Trim$(rsControleAcesso!XMLSegregacao)
'        Else
'            VerificaUsuario = vbNullString
'        End If
'
'        rsControleAcesso.Filter = ""
'        rsControleAcesso.MoveFirst
'    Else
'        VerificaUsuario = vbNullString
'    End If
    
    Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "VerificaUsuario", 0
End Function

Public Function RemoveUsuariosCache() As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    'Set rsControleAcesso = Nothing

    Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "RemoveUsuariosCache", 0
End Function

Private Sub flAtualizarParametroCache()

Dim strXMLParametro                         As String
Dim xmlParametro                            As MSXML2.DOMDocument40
Dim objParametro                            As Object
Dim strXML                                  As String

On Error GoTo ErrorHandler
    
    Set xmlParametro = CreateObject("MSXML2.DOMDocument.4.0")
    Set objParametro = CreateObject("A7Server.clsParametrosGerais")
    strXML = objParametro.flLer(0)
    xmlParametro.loadXML strXML
    xmlParametro.selectSingleNode("//PARM_GERL/CACHE_CONTROLE_ACESSO/PERIODICIDADE_EM_MINUTOS").Text = "N"
    Call objParametro.Alterar(xmlParametro.xml)
    Set objParametro = Nothing
    Set xmlParametro = Nothing
    Exit Sub
ErrorHandler:
    
    Set objParametro = Nothing
    Set xmlParametro = Nothing
    
    Err.Clear
End Sub

Public Function CarregarControleProcessamento(ByVal pstrFiltro As String) As ADODB.Recordset

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    If rsCtrlProcOperAtiv Is Nothing Then
        strSQL = "SELECT * FROM A8.TB_CTRL_PROC_OPER_ATIV ORDER BY TP_OPER "
        Set rsCtrlProcOperAtiv = QuerySQL(strSQL)
    End If
    rsCtrlProcOperAtiv.Filter = ""
    rsCtrlProcOperAtiv.Filter = pstrFiltro

    If rsCtrlProcOperAtiv.RecordCount > 0 Then
        Set CarregarControleProcessamento = rsCtrlProcOperAtiv
        CarregarControleProcessamento.Filter = pstrFiltro
    End If

    Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "CarregarControleProcessamento", 0
End Function

Public Function RecarregarCache()

On Error GoTo ErrorHandler
    
    'Pikachu - Implementação de recarregar o cache sem parar o serviço - Ativa Valida Remessa
    '29/07/2005
    
'    If Not rsControleAcesso Is Nothing Then rsControleAcesso.Close: Set rsControleAcesso = Nothing
'    If Not rsCtrlProcOperAtiv Is Nothing Then rsCtrlProcOperAtiv.Close: Set rsCtrlProcOperAtiv = Nothing
'    If Not rsEmpresaFusionada Is Nothing Then rsEmpresaFusionada.Close: Set rsEmpresaFusionada = Nothing
'    If Not rsEmpresa Is Nothing Then rsEmpresa.Close: Set rsEmpresa = Nothing
'    If Not rsBanco Is Nothing Then rsBanco.Close: Set rsBanco = Nothing
'    If Not rsLocalLiquidacao Is Nothing Then rsLocalLiquidacao.Close: Set rsLocalLiquidacao = Nothing
'    If Not rsSistema Is Nothing Then rsSistema.Close: Set rsSistema = Nothing
'    If Not rsProduto Is Nothing Then rsProduto.Close: Set rsProduto = Nothing
'    If Not rsTipoLiquidacao Is Nothing Then rsTipoLiquidacao.Close: Set rsTipoLiquidacao = Nothing
'    If Not rsVeiculoLegal Is Nothing Then rsVeiculoLegal.Close: Set rsVeiculoLegal = Nothing
'    If Not rsPartNegociacao Is Nothing Then rsPartNegociacao.Close: Set rsPartNegociacao = Nothing
'    If Not rsDataFeriado Is Nothing Then rsDataFeriado.Close: Set rsDataFeriado = Nothing
'    If Not rsTipoMensagem Is Nothing Then rsTipoMensagem.Close: Set rsTipoMensagem = Nothing
'    If Not rsTipoConta Is Nothing Then rsTipoConta.Close: Set rsTipoConta = Nothing
'    If Not rsSegmento Is Nothing Then rsSegmento.Close: Set rsSegmento = Nothing
'    If Not rsEventoFinanceiro Is Nothing Then rsEventoFinanceiro.Close: Set rsEventoFinanceiro = Nothing
'    If Not rsIndexador Is Nothing Then rsIndexador.Close: Set rsIndexador = Nothing
'    If Not rsControleDominio Is Nothing Then rsControleDominio.Close: Set rsControleDominio = Nothing
'    If Not rsTipoTransferencia Is Nothing Then rsTipoTransferencia.Close: Set rsTipoTransferencia = Nothing
'    If Not rsTipoTransferenciaLDL Is Nothing Then rsTipoTransferenciaLDL.Close: Set rsTipoTransferenciaLDL = Nothing
'    If Not rsTipoCompromisso Is Nothing Then rsTipoCompromisso.Close: Set rsTipoCompromisso = Nothing
'    If Not rsTipoTitular Is Nothing Then rsTipoTitular.Close: Set rsTipoTitular = Nothing
'    If Not rsTipoCompromissoRetn Is Nothing Then rsTipoCompromissoRetn.Close: Set rsTipoCompromissoRetn = Nothing
'    If Not rsTipoLiquidacaoMens Is Nothing Then rsTipoLiquidacaoMens.Close: Set rsTipoLiquidacaoMens = Nothing
'    If Not rsTipoLeilao Is Nothing Then rsTipoLeilao.Close: Set rsTipoLeilao = Nothing
'    If Not rsTipoPgtoLDL Is Nothing Then rsTipoPgtoLDL.Close: Set rsTipoPgtoLDL = Nothing
'    If Not rsTipoPagtoRedesc Is Nothing Then rsTipoPagtoRedesc.Close: Set rsTipoPagtoRedesc = Nothing
'    If Not rsTipoBackOffice Is Nothing Then rsTipoBackOffice.Close: Set rsTipoBackOffice = Nothing
'    If Not rsTipoNegociacaoBMA Is Nothing Then rsTipoNegociacaoBMA.Close: Set rsTipoNegociacaoBMA = Nothing
'    If Not rsSubTipoNegociacaoBMA Is Nothing Then rsSubTipoNegociacaoBMA.Close: Set rsSubTipoNegociacaoBMA = Nothing
'    If Not rsSubTipoAtivo Is Nothing Then rsSubTipoAtivo.Close: Set rsSubTipoAtivo = Nothing
'    If Not rsTipoAtivo Is Nothing Then rsTipoAtivo.Close: Set rsTipoAtivo = Nothing
'    If Not rsModalidadeLiquidacao Is Nothing Then rsModalidadeLiquidacao.Close: Set rsModalidadeLiquidacao = Nothing
'    If Not rsTipoContratoSwap Is Nothing Then rsTipoContratoSwap.Close: Set rsTipoContratoSwap = Nothing
'    If Not rsCodigoMoeda Is Nothing Then rsCodigoMoeda.Close: Set rsCodigoMoeda = Nothing
'    If Not rsTipoMovimentoFinanceiro Is Nothing Then rsTipoMovimentoFinanceiro.Close: Set rsTipoMovimentoFinanceiro = Nothing
'    If Not rsFinalidadeIF Is Nothing Then rsFinalidadeIF.Close: Set rsFinalidadeIF = Nothing
'    If Not rsIndicadorBoletim Is Nothing Then rsIndicadorBoletim.Close: Set rsIndicadorBoletim = Nothing
'    If Not rsTipoTrigger Is Nothing Then rsTipoTrigger.Close: Set rsTipoTrigger = Nothing
'    If Not rsTipoFonte Is Nothing Then rsTipoFonte.Close: Set rsTipoFonte = Nothing
'    If Not rsIndicadorFormaLiqu Is Nothing Then rsIndicadorFormaLiqu.Close: Set rsIndicadorFormaLiqu = Nothing
'    If Not rsIndicadorPeriodicidade Is Nothing Then rsIndicadorPeriodicidade.Close: Set rsIndicadorPeriodicidade = Nothing
'    If Not rsIndexadorCetip Is Nothing Then rsIndexadorCetip.Close: Set rsIndexadorCetip = Nothing
'    If Not rsTipoCliente Is Nothing Then rsTipoCliente.Close: Set rsTipoCliente = Nothing
'    If Not rsIndicadorExecrcOpcao Is Nothing Then rsIndicadorExecrcOpcao.Close: Set rsIndicadorExecrcOpcao = Nothing
'    If Not rsGrupoLancFinc Is Nothing Then rsGrupoLancFinc.Close: Set rsGrupoLancFinc = Nothing
'    If Not rsFinalidadeAberturaConta Is Nothing Then rsFinalidadeAberturaConta.Close: Set rsFinalidadeAberturaConta = Nothing
'    If Not rsItemCaixa Is Nothing Then rsItemCaixa.Close: Set rsItemCaixa = Nothing
'    If Not rsItemCaixaGenerico Is Nothing Then rsItemCaixaGenerico.Close: Set rsItemCaixaGenerico = Nothing

    Exit Function
ErrorHandler:

fgRaiseError App.EXEName, "basA6A8ValidaRemessa", "CarregarControleProcessamento", 0
End Function

