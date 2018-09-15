VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 2  'RequiresTransaction
END
Attribute VB_Name = "clsControleProcessoOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : clsOperacaoDefinitriva\r\nData Criação   : \r\nObjetivo       : Gerenciar e Controlar os fluxos operacionais das operações definitivas.\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F2946630040"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Componente     :
'Classe         : clsOperacaoDefinitriva
'Data Criação   :
'Objetivo       : Gerenciar e Controlar os fluxos operacionais das operações definitivas.
'Analista       :
'
'Programador    :
'Data           :
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

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

'Este método tem como função gerenciar o cancelamento da operação definitiva.
'
'Teste se a operação existe.
'
'Se existir e o status da operação for INICIAL ou EM SER:
' Chama o método EstornarPrevisão, da classe   clsGestaoCaixa, para montar e  enviar o Estorno   da  Previsão para o Sistema PJ-GestãoCaixa.
' Chama o método EstornarPrevisao, da classe  clsSubReserva, para montar e enviar o Estorno da  Previsão  para o Sistema de Sub-Reserva.
' Alterar status da operação  para CANCELADA ORIGEM.
'
'Se existir e o status da operação for CONCORDÂNCIA ou CONCORDÂNCIA AUTOMÁTICA:
' Chamar o método EstornarRealizado, da classe  clsGestaoCaixa, para montar e enviar os Estorno de  Realizado para o sistema PJ-Gestão Caixa.
' Proceder como Status INICIAL ou EM SER.
'
'Se existir e o status da operação for PENDENTE ou EM LANCAMENTO, chamar o método GerarSolicitacaoCancelamento, da classe clsOperacao, para gerar Solicitação de Cancelamento, e terminar o fluxo.
'
'Se existir e o status da operação for REJEITADA ou EXPIRADA:
' Altera o status da operação para CANCELADA ORIGEM e  termina o fluxo.
'
'Se existir e o status da operação for LIBERADA ou LIQUIDADA:
' Chamar o método Salvar, da classe clsremessaRejeitada,  para incluir remessa rejeitada
' Chamar o método MontarMensagemRejeicao, da classe cls Legado, para montar e enviar para o legado a rejeição da operação.
Private Function flCancelarOperacao(objDomRemessa As MSXML2.DOMDocument40) As Boolean
Attribute flCancelarOperacao.VB_Description = "Este método tem como função gerenciar o cancelamento da operação definitiva.\r\n\r\nTeste se a operação existe.\r\n\r\nSe existir e o status da operação for INICIAL ou EM SER:\r\n Chama o método EstornarPrevisão, da classe   clsGestaoCaixa, para montar e  enviar o Estorno   da  Previsão para o Sistema PJ-GestãoCaixa.\r\n Chama o método EstornarPrevisao, da classe  clsSubReserva, para montar e enviar o Estorno da  Previsão  para o Sistema de Sub-Reserva.\r\n Alterar status da operação  para CANCELADA ORIGEM.\r\n\r\nSe existir e o status da operação for CONCORDÂNCIA ou CONCORDÂNCIA AUTOMÁTICA:\r\n Chamar o método EstornarRealizado, da classe  clsGestaoCaixa, para montar e enviar os Estorno de  Realizado para o sistema PJ-Gestão Caixa.\r\n Proceder como Status INICIAL ou EM SER.\r\n\r\nSe existir e o status da operação for PENDENTE ou EM LANCAMENTO, chamar o método GerarSolicitacaoCancelamento, da classe clsOperacao, para gerar Solicitação de Cancelamento, e termin"
    On Error GoTo ErrorHandler


    '## your code goes here...

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCancelarOperacao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flComplementarOperacao(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean
Attribute flComplementarOperacao.VB_Description = "Este método tem como função gerenciar a complementação da operação definitiva.\r\n\r\nRecebe o remessa com a Ação 2 (Complementar)\r\n\r\nConsulta a operação para verificar sua existência.\r\n\r\nCaso exista, chamar o método ComplementarOperacao da classe clsOperacao, que irá atualizar os dados definitivos da operação. Senão, executar o método flIncluirOperacao.\r\n\r\nVerificar se a operação não é retroativa com mercado.\r\n\r\nSe for, encerrar o fluxo.\r\nSenão, verificar se existe Regra de Confirmação\r\nAutomática.\r\n\r\nSe não puder Confirmar, encerrar o fluxo.\r\n\r\nSenão, atualizar o status para CONCORDÂNCIA AUTOMÁTICA, chamando o método AlterarStatus da classe  clsOperacao.\r\nChamar o método EnviarRealizado, da classe clsGestaoCaixa, que porovidenciará o envio do Realizado para o Sistema PJ-Gestão Caixa.\r\n\r\nChamar o método LigacaoAssincrona, que irá incluir a operação na fila MQ, para continuar o processo.\r\n\r\nEncerrar o processo,\r\n\r\n"

Dim strOperacao                             As String
Dim objOperacao                             As A8Liquidacao.clsOperacao
Dim objDomOperacao                          As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")
    strOperacao = objOperacao.LerTodos(objDomRemessa.documentElement.selectSingleNode("CO_OPER_ATIV").Text)

    'Verificar no caso das compromissadas regras de complementacao para gravar somente um registro de status
    ' ex EM SER / A COMPLEMENTAR / EM SER


    'Verifica se Inclui ou Complementa a Operação
    If Trim(strOperacao) <> vbNullString Then
        Set objDomOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        objDomOperacao.loadXML strOperacao
        objOperacao.ComplementarOperacao objDomRemessa, _
                                         CLng(objDomOperacao.documentElement.selectSingleNode("Grupo_Operacao/NU_SEQU_OPER_ATIV").Text), _
                                         CLng(objDomOperacao.documentElement.selectSingleNode("Grupo_Operacao/TP_OPER").Text)
    Else
        flProcessarBoleto objDomRemessa, enumStatusOperacao.EmSer
    End If

    If Not flFluxosAlternativos(objDomRemessa) Then
        Exit Function
    End If

    flGerenciarChamadaAposCompl objDomRemessa

    Set objOperacao = Nothing
    Set objDomOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objDomOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flComplementarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flProcessarBoleto(ByRef objDomBoleto As MSXML2.DOMDocument40, _
                                   ByRef penumStatusOperacao As enumStatusOperacao) As Boolean

Dim objOperacao                             As A8Liquidacao.clsOperacao
Dim objGestaoCaixa                          As A8Liquidacao.clsGestaoCaixa
Dim objSubReserva                           As A8Liquidacao.clsSubReserva
Dim objDomGerenciamentoEnvio                As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")
    objOperacao.IncluirOperacao objDomBoleto, _
                                penumStatusOperacao
    Set objOperacao = Nothing

    'Carlos - Gerenciamento de Envio de Mensagens
    Set objDomGerenciamentoEnvio = CreateObject("MSXML2.DOMDocument.4.0")
    Set objDomGerenciamentoEnvio = fgObterPermissaoEnvio(enumFuncionalidade.RecebimentoBoleto, _
                                                         CLng(objDomBoleto.documentElement.selectSingleNode("TP_OPER").Text))

    If objDomGerenciamentoEnvio.documentElement.selectSingleNode("PrevisaoPJ").Text = enumIndicadorSimNao.Sim Then
        Set objGestaoCaixa = CreateObject("A8Liquidacao.clsGestaoCaixa")
        If Not objGestaoCaixa.EnviarPrevisao(objDomBoleto) Then
            '
        End If
        Set objGestaoCaixa = Nothing
    End If

    If objDomGerenciamentoEnvio.documentElement.selectSingleNode("PrevisaoSBR").Text = enumIndicadorSimNao.Sim Then
        Set objSubReserva = CreateObject("A8Liquidacao.clsSubReserva")
        If Not objSubReserva.EnviarPrevisao(objDomBoleto) Then
            '
        End If
        Set objGestaoCaixa = Nothing
    End If

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessarBoleto Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function GerenciarRecebimentoOperacao(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean
Attribute GerenciarRecebimentoOperacao.VB_Description = "Receber a Operação Defintiva.\r\n\r\nChamar o método flIdentificarAcao e testar o retorno.\r\nSe houver erro, terminar o processo transacional.\r\n\r\nChamar o método flValidarOperacao e testar o retorno.\r\nSe houver erro, terminar o processo transacional.\r\n\r\nSe a ação for = 1 (Inclusão), chamar p método                      flProcessarBoleto, passando o xml da remessa.\r\n\r\nSe a ação for = 2 (Complementação), chamar o método flComplementarOperacao, passando o xml da remessa.\r\n\r\nSe a ação for = 3 (Cancelamento), chamar o método flCancelarOperacao.\r\n\r\nSe houver erro, terminar o processo transacional.\r\n\r\nSenão, efetivar a transação."

    On Error GoTo ErrorHandler

    'flValidarOperacao objDomRemessa

    Select Case CLng(objDomRemessa.documentElement.selectSingleNode("TP_ACAO").Text)
        Case enumTipoAcao.Incluir
            'Verificar se é uma operação sem o Boleto (Numero de Operação Selic preenchido)
            'If objDomRemessa.documentElement.selectSingleNode("").Text = vbNullString Then
            flProcessarBoleto objDomRemessa, enumStatusOperacao.Inicial
        Case enumTipoAcao.Complementar
            flComplementarOperacao objDomRemessa
        Case enumTipoAcao.Cancelar
        Case enumTipoAcao.Liberar
            flLiberarOperacao objDomRemessa
    End Select

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flGerenciarChamadaAposCompl(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8Liquidacao.clsOperacao
Dim objGestaoCaixa                          As A8Liquidacao.clsGestaoCaixa
Dim objSubReserva                           As A8Liquidacao.clsSubReserva
Dim objRegra                                As A8Liquidacao.clsRegra
Dim objDomGerenciarChamadas                 As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    'Carlos - Gerenciamento de chamada das rotinas
    Set objDomGerenciarChamadas = CreateObject("MSXML2.DOMDocument.4.0")
    Set objDomGerenciarChamadas = fgObterPermissaoEnvio(enumFuncionalidade.RecebimentoOperacao, _
                                                        CLng(objDomRemessa.documentElement.selectSingleNode("TP_OPER").Text))

    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")

    If objDomGerenciarChamadas.documentElement.selectSingleNode("Confirmacao").Text = enumIndicadorSimNao.Sim Then
        Set objRegra = CreateObject("A8Liquidacao.clsRegra")
        'Verifica regra para confirmação automática
        If objRegra.VerificarRegraAutomatica(objDomRemessa, enumFuncaoSistema.Confirmar) Then
            objOperacao.AlterarStatus objDomRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      enumStatusOperacao.ConcordanciaAutomatica
        Else
            Exit Function
        End If
        Set objRegra = Nothing
    End If

    If objDomGerenciarChamadas.documentElement.selectSingleNode("RealizadoSolicitadoPJ").Text = enumIndicadorSimNao.Sim Then
        Set objGestaoCaixa = CreateObject("A8Liquidacao.clsGestaoCaixa")
        objGestaoCaixa.EnviarRealizado objDomRemessa
        Set objGestaoCaixa = Nothing
    End If

    If objDomGerenciarChamadas.documentElement.selectSingleNode("Conciliacao").Text = enumIndicadorSimNao.Sim Then
        Set objRegra = CreateObject("A8Liquidacao.clsRegra")
        'Verifica regra para conciliação automática
        If objRegra.VerificarRegraAutomatica(objDomRemessa, enumFuncaoSistema.Conciliar) Then
            objOperacao.ConciliarOperacao objDomRemessa
        Else
            Exit Function
        End If
        Set objRegra = Nothing
    End If

    'Termo Liquidacao e Eventos Selic envia neste momento e não após a Liberação.
    If objDomRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TermoLiquidacao Then
        If objDomGerenciarChamadas.documentElement.selectSingleNode("RealizadoConfirmadoSBR").Text = enumIndicadorSimNao.Sim Then
            Set objSubReserva = CreateObject("A8Liquidacao.clsSubReserva")
            objSubReserva.EnviarRealizadoConfirmado objDomRemessa
            Set objSubReserva = Nothing
        End If
    End If

    If objDomGerenciarChamadas.documentElement.selectSingleNode("Liberacao").Text = enumIndicadorSimNao.Sim Then
        'Alterado tipo de ação para liberação
        objDomRemessa.documentElement.selectSingleNode("TP_ACAO").Text = enumti
        flLigacaoAssincrona objDomRemessa
    End If

    'Vai para a Liberação
    'If objDomGerenciarChamadas.documentElement.selectSingleNode("RealizadoConfirmadoSBR").Text = enumIndicadorSimNao.Sim Then
    '    Set objSubReserva = CreateObject("A8Liquidacao.clsSubReserva")
    '    objSubReserva.EnviarRealizadoConfirmado objDomRemessa
    '    Set objSubReserva = Nothing
    'End If


    Set objOperacao = Nothing
    Set objDomGerenciarChamadas = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objDomGerenciarChamadas = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamadaAposCompl Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flFluxosAlternativos(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8Liquidacao.clsOperacao

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")

    Select Case CLng(objDomRemessa.documentElement.selectSingleNode("TP_OPER").Text)
        Case enumTipoOperacaoLQS.CompromissadaIda
            If Not flAlternativoCompromissadaIda(objDomRemessa) Then
                flFluxosAlternativos = False
            End If
        Case enumTipoOperacaoLQS.CompromissadaVolta
            If Not flAlternativoRetornoCompromissada(objDomRemessa) Then
                flFluxosAlternativos = False
            End If
        Case enumTipoOperacaoLQS.TermoLiquidacao
            If Not flAlternativoRetornoTermo(objDomRemessa) Then
                flFluxosAlternativos = False
            End If
    End Select

    Set objOperacao = Nothing

    flFluxosAlternativos = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flFluxosAlternativos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAlternativoCompromissadaIda(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8Liquidacao.clsOperacao

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")

    'Regra para complementação manual
    If objDomRemessa.documentElement.selectSingleNode("PURetorno").Text = vbNullString And _
        objDomRemessa.documentElement.selectSingleNode("TP_CPRO_OPER_ATIV").Text = "01" Then
        objOperacao.AlterarStatus objDomRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                  enumStatusOperacao.AComplementar
        flAlternativoCompromissadaIda = False
        Exit Function
    End If

    flAlternativoCompromissadaIda = False

    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoCompromissadaIda Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flAlternativoRetornoCompromissada(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String
Dim objOperacao                             As A8Liquidacao.clsOperacao
Dim objHistoricoOperacao                    As A8Liquidacao.clsHistoricoSitOper
Dim objDomOperacaoIda                       As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler

    Set objDomOperacaoIda = CreateObject("MSXML2.DOMDocument.4.0")
    Set objHistoricoOperacao = CreateObject("A8Liquidacao.clsHistoricoSitOper")
    Set objOperacao = CreateObject("A8Liquidacao.clsOperacao")

    strXML = objOperacao.LerTodos(, , objDomRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text)

    If strXML = vbNullString Then
        'Verificar se deve dar erro caso não ache a operacao de ida !!!! Carlos 10/09/2003
    End If

    objDomOperacaoIda.loadXML strXML

    'Regra para complementação manual
    strXML = objHistoricoOperacao.LerTodos(objDomOperacaoIda.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                           enumStatusOperacao.AComplementar)

    If strXML <> vbNullString Then
        objOperacao.AlterarStatus objDomRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                  enumStatusOperacao.AComplementar
        flAlternativoRetornoCompromissada = False
        Exit Function
    End If


    'Falta conciliar a Volta com a Ida, a operação de ida já está no objOperacao - Carlos 10/09/2003
    ' Coincide Venc. SEM SEL1611
    ' Coincide Venc. COM SEL1611
    ' Coincide Venc. COM SEL1611 mas com valor diferente

    flAlternativoRetornoCompromissada = False

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
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoRetornoCompromissada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flAlternativoRetornoTermo(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String

    On Error GoTo ErrorHandler


    flAlternativoRetornoTermo = False

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlternativoRetornoTermo Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flLigacaoAssincrona(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String

    On Error GoTo ErrorHandler

    'Fazer o PUT para a liberação

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLigacaoAssincrona Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Private Function flLiberarOperacao(ByRef objDomRemessa As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String

    On Error GoTo ErrorHandler






    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
