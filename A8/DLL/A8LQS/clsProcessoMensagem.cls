VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo recebimento das mensagens enviadas pela Mensageria e
'' controlar a liquidação das operações com as mensagens R1.

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

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo do recebimento de uma mensagem SPB
'-------------------------------------------------------------------------------------------------

Public Function GerenciarRecebimentoMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pstrHeaderNZ As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux
Dim objMensagem                             As A8LQS.clsMensagem
Dim objMensagemBMA                          As A8LQS.clsProcessoMensagemBMA
Dim objProcessoMensagemSTR                  As A8LQS.clsProcessoMensagemSTR
Dim objProcessoMensagemCCR                  As A8LQS.clsProcessoMensagemCCR
Dim objWorkflow                             As A8LQS.clsWorkflow

Dim blnMensagemTratada                      As Boolean
Dim lngLocalLiquidacao                      As Long
Dim strLeitura                              As String

    On Error GoTo ErrorHandler
    
    Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
    Set objMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    
    lngLocalLiquidacao = objMensagemBMA.VerificaCamaraAtiva(xmlMensagem, _
                                                            xmlMensagem.selectSingleNode("//CodMsg").Text)
                                                            
    strLeitura = objWorkflow.ObterInatividadeFluxoRecebimento(, _
                                                              xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                              lngLocalLiquidacao)
        
    If strLeitura <> vbNullString Then
        Call objMensagemBMA.ProcessaMensagem(xmlMensagem, pstrHeaderNZ, MensagemInativa)
        Set objWorkflow = Nothing
        Set objMensagemBMA = Nothing
        Exit Function
    End If

    If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "CTP" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "BMA" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "LDL" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "LTR" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "CAM" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "BMC" Then

TrataMensagemErroBMA:
        Set objMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
        Call objMensagemBMA.GerenciarRecebimentoMensagemBMA(xmlMensagem, pstrHeaderNZ)
        Set objMensagemBMA = Nothing
        
    'Tratamento mensagens STR
    ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "STR" Or _
           Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "PAG" Then

        Set objProcessoMensagemSTR = CreateObject("A8LQS.clsProcessoMensagemSTR")
        objProcessoMensagemSTR.GerenciarRecebimentoMensagem xmlMensagem, pstrHeaderNZ
        Set objProcessoMensagemSTR = Nothing
    
    'KIDA - CCR
    'Tratamento mensagens CCR
    ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "CCR" Then

        Set objProcessoMensagemCCR = CreateObject("A8LQS.clsProcessoMensagemCCR")
        objProcessoMensagemCCR.GerenciarRecebimentoMensagem xmlMensagem, pstrHeaderNZ
        Set objProcessoMensagemCCR = Nothing

    'MENSAGENS SELIC
    Else
            
        Set objMensagem = CreateObject("A8LQS.clsMensagem")
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "GEN" Then
            If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 2) <> "R1" Then
                udtProtocoloErroNZAux.String = pstrHeaderNZ
                LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
                If Left$(udtProtocoloErroNZ.CodigoMensagem, 3) = "LDL" Then
                    GoTo TrataMensagemErroBMA
                Else
                    blnMensagemTratada = objMensagem.VerificaMensagemTratadaSLCC(Mid(udtProtocoloErroNZ.CodigoMensagem, 1, 7))
                End If
            End If
        Else
            blnMensagemTratada = objMensagem.VerificaMensagemTratadaSLCC(Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7))
        End If
        
        Set objMensagem = Nothing
    
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 2) = "R1" And _
            blnMensagemTratada And _
            xmlMensagem.selectSingleNode("//CodMsg").Text <> "RDC0005R1" Then
            
            flMensagemTratada xmlMensagem, enumStatusMensagem.R1, pstrHeaderNZ
        
        ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 1) = "E" And _
            blnMensagemTratada Then
            
            flMensagemTratada xmlMensagem, enumStatusMensagem.ErroNegocioSPB, pstrHeaderNZ, True
        
        ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "GEN" And _
            blnMensagemTratada Then
            
            flMensagemTratada xmlMensagem, _
                              IIf(Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "GEN0004", enumStatusMensagem.GEN0004, enumStatusMensagem.GEN0009), _
                              pstrHeaderNZ, _
                              True, udtProtocoloErroNZ.ControleRemessaNZ, udtProtocoloErroNZ.CodigoMensagem
        
        ElseIf xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0011" Or _
                xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0015" Or _
                xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0016" Or _
                xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0005R1" Then
            
            flMensagemTratadaRDC xmlMensagem, pstrHeaderNZ
        
        ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1400" Or _
                Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1100" Or _
                Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "RDC0014" Then
            
            flMensagemCancelamentoEstorno xmlMensagem, pstrHeaderNZ
        
        ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1069" Then
            
            flTratarSEL1069 xmlMensagem, pstrHeaderNZ
        
        'Mensagens com Erro do NZ
        ElseIf blnMensagemTratada Then
            
            flMensagemTratada xmlMensagem, enumStatusMensagem.ErroNegocioSPB, pstrHeaderNZ, True
        
        Else
            
            flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
        
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimentoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de gravação de uma mensagem SPB não identificada como enviada pelo Sistema SLCC
'-------------------------------------------------------------------------------------------------

Private Function flMensagemNaoTratada(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                      ByRef pstrHeaderNZ As String) As Boolean

      Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
      Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
      Dim objConciliacao                          As A8LQS.clsConciliacao
      Dim objAlerta                               As A8LQS.clsAlerta
      Dim objTipoConta                            As A8LQS.clsTipoConta
      Dim udtHeaderMensagem                       As udtHeaderMensagem
      Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
      Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
      Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux
      Dim lngStatusMensagem                       As Long
      Dim lngNumeroComando                        As Long
      Dim blnAtualizaOriginal                     As Boolean
      Dim lngStatusMensagemOriginal               As Long
      Dim strControleRemessaNZ                    As String
      Dim lngCodigoEmpresa                        As Long
      Dim strSiglaSistema                         As String
      Dim strDataRegistroMensagem                 As String
      Dim strCodigoVeiculoLegal                   As String
      Dim lngTipoBackOffice                       As Long

      Dim strContaContraparte                     As String
      Dim strTipoDebitoCredito                    As String
      Dim blnCliente1                             As Boolean

10    On Error GoTo ErrorHandler

20        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

          'Gravar a mensagem
30        If (Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 8, 1) = "E" _
          Or Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "GEN" _
          Or xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1082" _
          Or xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1095") _
          And Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 2) <> "R1" Then
40            udtProtocoloErroNZAux.String = pstrHeaderNZ
50            LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
60            strControleRemessaNZ = udtProtocoloErroNZ.ControleRemessaNZ
70            lngCodigoEmpresa = udtProtocoloErroNZ.CodigoEmpresa
80            strSiglaSistema = ""
90        Else
100           udtHeaderMensagemAux.String = pstrHeaderNZ
110           LSet udtHeaderMensagem = udtHeaderMensagemAux
120           strControleRemessaNZ = udtHeaderMensagem.ControleRemessaNZ
130           lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
140           strSiglaSistema = udtHeaderMensagem.SiglaSistemaEnviouNZ
150       End If

160       blnAtualizaOriginal = False
170       strDataRegistroMensagem = ""
          
180       If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "SEL1611" Then

190           lngStatusMensagem = enumStatusMensagem.Informação
200           Select Case Val("0" & xmlMensagem.documentElement.selectSingleNode("//TpInfEvtSEL").Text)
                  Case enumTipoEventoSEL1611.OperacaoProcessadaAutomaticamente
                  
210                   Select Case Val("0" & xmlMensagem.documentElement.selectSingleNode("//CodOpSEL").Text)
                          
                          Case enumCodigoOperacaoSelic.LiquidacaoTermosDataLiquidacao, _
                               enumCodigoOperacaoSelic.Juros, _
                               enumCodigoOperacaoSelic.Resgate, _
                               enumCodigoOperacaoSelic.Amortizacao, _
                               enumCodigoOperacaoSelic.RetornoCompromissosTitulosVencemDia
                              
220                           lngStatusMensagem = flObterStatusSEL1611(xmlMensagem.documentElement.selectSingleNode("//SitOpSEL").Text)

230                       Case Else
                              
240                           lngStatusMensagem = enumStatusMensagem.Informação
                              
250                   End Select
                  
260               Case enumTipoEventoSEL1611.LancamentoRegistradoContraparte
                      
270                   Select Case Val("0" & xmlMensagem.documentElement.selectSingleNode("//CodOpSEL").Text)
                          
                          Case enumCodigoOperacaoSelic.LiquidacaoLeilaoVendaPrimario, _
                               enumCodigoOperacaoSelic.LiquidacaoLeilaoVendaSecundario, _
                               enumCodigoOperacaoSelic.LiquidacaoLeilaoCompraSecundario, _
                               enumCodigoOperacaoSelic.Despesas
                              
280                           lngStatusMensagem = flObterStatusSEL1611(xmlMensagem.documentElement.selectSingleNode("//SitOpSEL").Text)

290                       Case Else
                              
300                           lngStatusMensagem = enumStatusMensagem.Informação
                              
310                   End Select
                  
320               Case enumTipoEventoSEL1611.PreRegistroIntermediacaoLote
330                   lngStatusMensagem = enumStatusMensagem.Informação
                      
340           End Select
              
350           lngNumeroComando = Val(xmlMensagem.documentElement.selectSingleNode("//NumOpSEL").Text)
360           strDataRegistroMensagem = xmlMensagem.documentElement.selectSingleNode("//DtHrSit").Text

              'Para as contas de cliente-1 será necessário realizar a pesquisa utilizando as quatro primeiras posições
              'da conta de custódia concatenado com 4 zeros para identificarmos o Tipo de Back Office correto.
              
370           Set objTipoConta = CreateObject("A8LQS.clsTipoConta")
              
380           strTipoDebitoCredito = "D"
390           If Not xmlMensagem.selectSingleNode("//TpDeb_Cred") Is Nothing Then
400               strTipoDebitoCredito = xmlMensagem.selectSingleNode("//TpDeb_Cred").Text
410           End If
              
420           blnCliente1 = True
430           If strTipoDebitoCredito = "D" Then
440               If Not xmlMensagem.documentElement.selectSingleNode("//CtCed") Is Nothing Then
450                   strContaContraparte = Mid(fgCompletaString(xmlMensagem.documentElement.selectSingleNode("//CtCed").Text, "0", 9, True), 1, 4) & "0000"
460               End If
470           Else
480               If Not xmlMensagem.documentElement.selectSingleNode("//CtCes") Is Nothing Then
490                   strContaContraparte = Mid(fgCompletaString(xmlMensagem.documentElement.selectSingleNode("//CtCes").Text, "0", 9, True), 1, 4) & "0000"
500               End If
510           End If
              
520           Set objTipoConta = Nothing

530           fgObterCodigoVeiculoLegal "SEL1611", _
                                        strContaContraparte, _
                                        lngCodigoEmpresa, _
                                        strCodigoVeiculoLegal, _
                                        strSiglaSistema, _
                                        lngTipoBackOffice, _
                                        , , _
                                        blnCliente1

540           If lngTipoBackOffice = 0 Then
550               lngTipoBackOffice = enumTipoBackOffice.Todos
560           End If

570       Else

580           lngTipoBackOffice = enumTipoBackOffice.Todos
590           strCodigoVeiculoLegal = vbNullString

600           If Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 8, 2) = "R1" Then
610               lngStatusMensagem = enumStatusMensagem.R1
620               lngStatusMensagemOriginal = enumStatusMensagem.Respondida
630               lngNumeroComando = 0
640               blnAtualizaOriginal = True
650               If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "SEL1069R1" Then
660                   If xmlMensagem.documentElement.selectSingleNode("//SitOpSEL").Text = "ATU" Then
670                       lngStatusMensagemOriginal = enumStatusMensagem.Encerrada
680                   End If
690               End If
700           ElseIf Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 8, 1) = "E" _
                  Or xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1082" _
                  Or xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1095" Then
710               lngStatusMensagem = enumStatusMensagem.ErroNegocioSPB
720               lngStatusMensagemOriginal = enumStatusMensagem.Respondida
730               lngNumeroComando = 0
740               blnAtualizaOriginal = True
750           ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "GEN0004" Then
760               lngStatusMensagem = enumStatusMensagem.GEN0004
770               lngStatusMensagemOriginal = enumStatusMensagem.Respondida
780               lngNumeroComando = 0
790               blnAtualizaOriginal = True
800           ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "GEN0009" Then
810               lngStatusMensagem = enumStatusMensagem.GEN0009
820               lngStatusMensagemOriginal = enumStatusMensagem.Respondida
830               lngNumeroComando = 0
840               blnAtualizaOriginal = True
850           Else
860               lngStatusMensagem = enumStatusMensagem.Informação
870               lngNumeroComando = 0
880           End If
890       End If

900       objMensagemEnviadaRecebida.Incluir xmlMensagem.xml, _
                                             Trim(strControleRemessaNZ), _
                                             0, _
                                             0, _
                                             xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, _
                                             lngTipoBackOffice, _
                                             lngCodigoEmpresa, _
                                             lngNumeroComando, _
                                             "", _
                                             lngStatusMensagem, _
                                             enumIndicadorSimNao.Nao, _
                                             "", _
                                             enumLocalLiquidacao.SELIC, _
                                             strCodigoVeiculoLegal, _
                                             strSiglaSistema, _
                                             , _
                                             strDataRegistroMensagem

          'Alterar o Status da Mensagem Original
910       If blnAtualizaOriginal Then
920           Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
930           Set xmlMensagemEnviadaRecebida = CreateObject("MSXML2.DOMDocument.4.0")
940           xmlMensagemEnviadaRecebida.loadXML objMensagemEnviadaRecebida.LerMensagemIda(Trim(strControleRemessaNZ))
950           objMensagemEnviadaRecebida.AlterarStatus xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                       xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                       lngStatusMensagemOriginal, _
                                                       xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
960       End If
970       Set objMensagemEnviadaRecebida = Nothing

980       Select Case xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text
              Case "SEL1611"
990               If lngStatusMensagem = enumStatusMensagem.AConciliar Then
                      'Verificar a conciliação automática
1000                  Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
1010                  objConciliacao.ConciliarMensagem xmlMensagem, lngCodigoEmpresa
1020                  Set objConciliacao = Nothing
1030              End If
1040          Case "SEL1099"
1050              Set objAlerta = CreateObject("A8LQS.clsAlerta")
1060              objAlerta.GerarAlertaSemOperacao xmlMensagem, enumFatorGeradorAlerta.AvisoDebitoCreditoRB
1070              Set objAlerta = Nothing
1080      End Select

1090      If Not ObjectContext Is Nothing Then
1100          ObjectContext.SetComplete
1110      End If

1120      Exit Function
ErrorHandler:

1130      Set objMensagemEnviadaRecebida = Nothing

1160      If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
1170      Call fgRaiseError(App.EXEName, TypeName(Me), "flMensagemNaoTratada Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Linha:" + Erl)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de gravação e distribuição de uma mensagem SPB enviada pelo Sistema SLCC
'-------------------------------------------------------------------------------------------------

Private Function flMensagemTratada(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                   ByVal penumStatusMensagem As enumStatusMensagem, _
                                   ByRef pstrHeaderNZ As String, _
                          Optional ByRef pblnMensagemErro As Boolean = False, _
                          Optional ByVal pstrNumeroControleIF As String = "", _
                          Optional ByVal pstrCodigoMensagem As String = "") As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoMensagemBMA                  As A8LQS.clsProcessoMensagemBMA

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlGrupoMensagem                        As MSXML2.DOMDocument40
Dim xmlTagDominioStatus                     As MSXML2.DOMDocument40
Dim domNode                                 As MSXML2.IXMLDOMNode
Dim strGrupoMensagem                        As String
Dim strTagDominioStatus                     As String
Dim strNomeTag                              As String
Dim strMensagem                             As String
Dim strSituacaoMensagem                     As String
Dim strSiglaSistema                         As String

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux
Dim lngStatusOperacaoErro                   As Long

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Set xmlTagDominioStatus = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlGrupoMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a Operação
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    If pstrNumeroControleIF = vbNullString Then
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1069" Then
            Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.selectSingleNode("//NumCtrlIF").Text, False)
            If xmlOperacao.xml = vbNullString Then
                flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
                Exit Function
            End If
        ElseIf Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1023" Then
            Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.selectSingleNode("//NumCtrlIF").Text, False)
            If xmlOperacao.xml = vbNullString Then
                flTratarSEL1023 xmlMensagem, pstrHeaderNZ
                Exit Function
            End If
        Else
            Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.selectSingleNode("//NumCtrlIF").Text, , _
                                                                     Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7))
        End If
        pstrNumeroControleIF = xmlMensagem.selectSingleNode("//NumCtrlIF").Text
    Else
        Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(pstrNumeroControleIF)
    End If

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
        '3038 - Operação já liquidada.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
    
    'Pikachu - 16/07/2004
    'Processo mensagem para Leilao BMA
    If xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA Then
                    
        Set objProcessoMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
        Call objProcessoMensagemBMA.ProcessaMensagemSELIC_BMA(xmlMensagem, pstrHeaderNZ)
        Set objProcessoMensagemBMA = Nothing
        
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
        Exit Function
    End If

    'Eduardo Quaresma 05/01/2004
    'Verifica o código da mensagem
    If xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0007R1" Or _
        xmlMensagem.selectSingleNode("//CodMsg").Text = "RDC0008R1" Then
        'Valida o redesconto
        Call flValidarRedesconto(xmlOperacao.documentElement.selectSingleNode("//NU_CTRL_RDSC_ORIG").Text, _
                                 CLng(xmlOperacao.documentElement.selectSingleNode("//CO_EMPR").Text), _
                                 fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text))
    End If
    
    If pstrCodigoMensagem = vbNullString Then
        strGrupoMensagem = objMensagem.LerGrupoMensagem(Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7))
        pstrCodigoMensagem = xmlMensagem.selectSingleNode("//CodMsg").Text
    Else
        strGrupoMensagem = objMensagem.LerGrupoMensagem(Mid(pstrCodigoMensagem, 1, 7))
    End If
    If strGrupoMensagem = vbNullString Then
        '3038 - Mensagem sem grupo cadastrado.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
    xmlGrupoMensagem.loadXML strGrupoMensagem

    strTagDominioStatus = LerTodosTagDominioStatus(xmlGrupoMensagem.documentElement.selectSingleNode("CO_GRUP").Text)
    If strTagDominioStatus = vbNullString Then
        '3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
        lngCodigoErroNegocio = 3017
        GoTo ErrorHandler
    End If
    xmlTagDominioStatus.loadXML strTagDominioStatus

    strNomeTag = xmlTagDominioStatus.documentElement.selectSingleNode("//NO_TAG").Text

    'Incluir a Mensagem
    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag) Is Nothing Then
        strSituacaoMensagem = ""
    Else
        strSituacaoMensagem = xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text
    End If
    
    If Not xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG") Is Nothing Then
        strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
    ElseIf Not xmlOperacao.documentElement.selectSingleNode("SG_SIST") Is Nothing Then
        strSiglaSistema = xmlOperacao.documentElement.selectSingleNode("SG_SIST").Text
    End If
    
    objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                       pstrNumeroControleIF, _
                                       xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       0, pstrCodigoMensagem, _
                                       xmlOperacao.documentElement.selectSingleNode("TP_BKOF").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                       strSituacaoMensagem, _
                                       penumStatusMensagem, _
                                       enumIndicadorSimNao.Nao, _
                                       "", _
                                       xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       strSiglaSistema
    
    Set objMensagemRecebidaEnviada = Nothing
    
    'Verificar qual o status da mensagem e qual deverá ser o próximo status da Operação
    If pblnMensagemErro = False Then
        
        For Each domNode In xmlTagDominioStatus.documentElement.selectNodes("//Grupo_TagDominioStatus")
            
            If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text = domNode.selectSingleNode("DE_DOMI").Text Then
                
                Select Case domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                    Case enumStatusOperacao.Liquidada, _
                         enumStatusOperacao.EmLancamento, _
                         enumStatusOperacao.Pendencia
                         
                        Call xmlOperacao.loadXML(objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text).xml)
                         
                        flGerenciarChamada xmlOperacao, _
                                           xmlMensagem, _
                                           CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text)), _
                                           CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                           xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                           enumIndicadorSimNao.Nao, _
                                           pstrNumeroControleIF
                                            
                        If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "SEL1054R1" And _
                           domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = enumStatusOperacao.Liquidada Then
                            If Not xmlMensagem.documentElement.selectSingleNode("//NumOpSELRet") Is Nothing Then
                               objOperacao.AlterarComandoOperacaoRetn xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                      xmlMensagem.documentElement.selectSingleNode("//NumOpSELRet").Text
                            End If
                        End If
                        Exit For
                    
                    Case enumStatusOperacao.Expirada, _
                         enumStatusOperacao.Rejeitada, _
                         enumStatusOperacao.Cancelada, _
                         enumStatusOperacao.Estornada, _
                         enumStatusOperacao.RejeitadaPiloto
                         
                         flGerenciarChamada xmlOperacao, _
                                            xmlMensagem, _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text)), _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                            xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                            enumIndicadorSimNao.Sim, _
                                            pstrNumeroControleIF
                        Exit For
                End Select
            End If
        Next
    Else
        
         strSituacaoRecebida = vbNullString
         lngStatusOperacaoErro = enumStatusOperacao.Rejeitada

        'Pikachu - 14/06/2004
        'Alteração para tratamento de rejeicao das mensagens pelo Piloto - WorkFlow (PK)
        If penumStatusMensagem = enumStatusMensagem.ErroNegocioSPB Then
            
            udtProtocoloErroNZAux.String = pstrHeaderNZ
            LSet udtProtocoloErroNZ = udtProtocoloErroNZAux

            If Trim(udtProtocoloErroNZ.CodigoErro1) = "5156" Then
                lngStatusOperacaoErro = enumStatusOperacao.RejeitadaPiloto
            Else
                strSituacaoRecebida = Trim(udtProtocoloErroNZ.CodigoErro1)
            End If
            
        End If
        
        'Mensagens com Erros de Negócios ou NZ
        flGerenciarChamada xmlOperacao, _
                           xmlMensagem, _
                           lngStatusOperacaoErro, _
                           enumStatusMensagem.Respondida, _
                           "ERR", _
                           enumIndicadorSimNao.Sim, _
                           pstrNumeroControleIF, _
                           strSituacaoRecebida
    End If

    If Not xmlMensagem.selectSingleNode("//NumCtrlRDC") Is Nothing Then
        objOperacao.GravarControleSPBOriginal xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                              xmlMensagem.selectSingleNode("//NumCtrlRDC").Text
    End If

    If pblnMensagemErro = False Then
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoInformalCompromissada Then
            If xmlMensagem.documentElement.selectSingleNode("//SitOpSEL").Text = "ATU" And _
                Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1007" Then
                objOperacao.GravarControleSPBOriginal xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                      xmlMensagem.selectSingleNode("//NumOpSELRet").Text
            End If
        End If
    End If

    'Tratar os Repasses
    If Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1066" Or _
        Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 1, 7) = "STR0004" Then
        flTratarRepasse xmlOperacao
    End If

    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set objProcessoMensagemBMA = Nothing
    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMensagemTratada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de gravação e distribuição de mensagens SPB de cancelamento e/ou estorno
'-------------------------------------------------------------------------------------------------

Private Function flMensagemCancelamentoEstorno(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                               ByRef pstrHeaderNZ As String) As Boolean

Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim strSituacaoMensagem                     As String

    On Error GoTo ErrorHandler

    'Obter a Operação
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.selectSingleNode("//NumCtrlIF").Text, False)

    If xmlOperacao.xml = vbNullString Then
        'Mensagem sem Operação
        flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
        Exit Function
    End If

    If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "SEL1400" Or _
        Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "RDC0014" Then
        If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada Then
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If
        lngStatusOperacao = enumStatusOperacao.Cancelada
    Else
        lngStatusOperacao = enumStatusOperacao.Estornada
    End If

    If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 1) = "E" Then
        lngStatusMensagem = enumStatusMensagem.ErroNegocioSPB
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "RDC0014" Then
            strSituacaoMensagem = "ERR"
        Else
            If xmlMensagem.selectSingleNode("//SitOpSEL") Is Nothing Then
                strSituacaoMensagem = "ERR"
            Else
                strSituacaoMensagem = xmlMensagem.selectSingleNode("//SitOpSEL").Text
            End If
        End If
    Else
        lngStatusMensagem = enumStatusMensagem.R1
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "RDC0014" Then
            strSituacaoMensagem = xmlMensagem.selectSingleNode("//SitLancRDC").Text
        Else
            If xmlMensagem.selectSingleNode("//SitOpSEL") Is Nothing Then
                strSituacaoMensagem = vbNullString
            Else
                strSituacaoMensagem = xmlMensagem.selectSingleNode("//SitOpSEL").Text
            End If
        End If
    End If

    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                       xmlMensagem.selectSingleNode("//NumCtrlIF").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       0, xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("TP_BKOF").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                       strSituacaoMensagem, _
                                       lngStatusMensagem, _
                                       enumIndicadorSimNao.Nao, _
                                       "", _
                                       xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                       xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
    Set objMensagemRecebidaEnviada = Nothing

    If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 8, 1) = "E" Then
        'Mensagens com Erros de Negócios ou NZ
        flGerenciarChamada xmlOperacao, _
                           xmlMensagem, _
                           xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                           enumStatusMensagem.ErroNegocioSPB, _
                           "ERR", _
                           enumIndicadorSimNao.Nao, _
                           xmlMensagem.selectSingleNode("//NumCtrlIF").Text
    Else
        If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7) = "RDC0014" Then
            If xmlMensagem.selectSingleNode("//SitLancRDC").Text <> "1" Then
                Exit Function
            Else
                strSituacaoMensagem = xmlMensagem.selectSingleNode("//SitLancRDC").Text
            End If
        Else
            If xmlMensagem.selectSingleNode("//SitOpSEL").Text <> "ATU" Then
                Exit Function
            Else
                strSituacaoMensagem = xmlMensagem.selectSingleNode("//SitOpSEL").Text
            End If
        End If
        flGerenciarChamada xmlOperacao, _
                           xmlMensagem, _
                           lngStatusOperacao, _
                           enumStatusMensagem.Respondida, _
                           strSituacaoMensagem, _
                           enumIndicadorSimNao.Sim, _
                           xmlMensagem.selectSingleNode("//NumCtrlIF").Text
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMensagemCancelamentoEstorno Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de gravação e distribuição de mensagens SPB do grupo RDC
'-------------------------------------------------------------------------------------------------

Private Function flMensagemTratadaRDC(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                      ByRef pstrHeaderNZ As String) As Boolean

Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objLegado                               As A8LQS.clsLegado
Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoAux                          As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim strMensagem                             As String
Dim lngStatusOperacao                       As Long
Dim lngStatusMensagem                       As Long
Dim blnEstorno                              As Boolean
Dim lngCodigoEmpresa                        As Long
Dim strSiglaSistema                         As String
Dim strOperacao                             As String
Dim strControleRDC                          As String
Dim lngTipoBackOffice                       As Long

On Error GoTo ErrorHandler

    Select Case xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text
        Case "RDC0011"
            lngStatusOperacao = 0
            lngStatusMensagem = enumStatusMensagem.Informação
            blnEstorno = False
        Case "RDC0015"
            lngStatusOperacao = enumStatusOperacao.LiquidadaConvertida
            lngStatusMensagem = enumStatusMensagem.Informação
            blnEstorno = False
        Case "RDC0005R1"
            lngStatusOperacao = enumStatusOperacao.LiquidadaConvertida
            lngStatusMensagem = enumStatusMensagem.R1
            blnEstorno = False
        Case "RDC0016"
            lngStatusOperacao = enumStatusOperacao.Cancelada
            lngStatusMensagem = enumStatusMensagem.Informação
            blnEstorno = True
    End Select

    If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "RDC0005R1" Then
        If Not flGravarRDC0005R1(xmlMensagem, pstrHeaderNZ, strControleRDC) Then
            Exit Function
        End If
    Else
        strControleRDC = xmlMensagem.documentElement.selectSingleNode("//NumCtrlRDCOr").Text
    End If

    'Obter a Operação - Se achou é a conciliação
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    udtHeaderMensagemAux.String = pstrHeaderNZ
    LSet udtHeaderMensagem = udtHeaderMensagemAux
    strOperacao = objOperacao.LerTodos(, , , , , , _
                                       strControleRDC, , _
                                       CLng(udtHeaderMensagem.CodigoEmpresa))

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    If strOperacao <> vbNullString Then
        xmlOperacao.loadXML strOperacao
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        Set xmlOperacaoAux = xmlOperacao
    End If

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    If xmlOperacao.xml <> vbNullString Then
        If Not xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
            If xmlOperacao.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada Then
                If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "RDC0015" Or _
                    xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "RDC0005R1" Then
                    If Not xmlMensagem.documentElement.selectSingleNode("//NumCtrlIF") Is Nothing Then
                        Set xmlOperacaoAux = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.documentElement.selectSingleNode("//NumCtrlIF").Text)
                    Else
                        Set xmlOperacaoAux = xmlOperacao
                    End If
                    flEnviarConversaoRDC xmlOperacaoAux
                Else
                    If blnEstorno = False Then
                        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                        strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
                        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                        Set objSubReserva = Nothing

                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
                        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                        Set objGestaoCaixa = Nothing
                    Else
                        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
                        strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
                        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                        Set objSubReserva = Nothing

                        Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
                        strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                        Set objGestaoCaixa = Nothing
                    End If
                    Set xmlOperacaoAux = xmlOperacao
                End If

                Set objLegado = CreateObject("A8LQS.clsLegado")
                strMensagem = objLegado.EnviarRetorno(xmlOperacaoAux, xmlMensagem)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                Set objLegado = Nothing

                If lngStatusOperacao > 0 Then
                    objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                  lngStatusOperacao, _
                                                  xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                  , , False
                End If
                Set objProcessoOperacao = Nothing

                If Not xmlMensagem.selectSingleNode("//NumCtrlRDC") Is Nothing Then
                    objOperacao.GravarControleSPBOriginal xmlOperacaoAux.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                          xmlMensagem.selectSingleNode("//NumCtrlRDC").Text
                End If

            End If
        End If

        If xmlOperacaoAux.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
            lngTipoBackOffice = enumTipoBackOffice.Todos
        Else
            lngTipoBackOffice = xmlOperacaoAux.documentElement.selectSingleNode("TP_BKOF").Text
        End If

        'Incluir a Mensagem
        Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                           strControleRDC, _
                                           xmlOperacaoAux.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                           0, xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, _
                                           lngTipoBackOffice, _
                                           xmlOperacaoAux.documentElement.selectSingleNode("CO_EMPR").Text, _
                                           xmlOperacaoAux.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                           "", _
                                           lngStatusMensagem, _
                                           enumIndicadorSimNao.Nao, _
                                           "", _
                                           xmlOperacaoAux.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                           xmlOperacaoAux.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                           xmlOperacaoAux.documentElement.selectSingleNode("SG_SIST_ORIG").Text
        Set objMensagemRecebidaEnviada = Nothing
    Else
        'Incluir a Mensagem
        udtHeaderMensagemAux.String = pstrHeaderNZ
        LSet udtHeaderMensagem = udtHeaderMensagemAux
        lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
        strSiglaSistema = udtHeaderMensagem.SiglaSistemaEnviouNZ
        Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                           strControleRDC, _
                                           0, _
                                           0, xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, _
                                           enumTipoBackOffice.Todos, _
                                           lngCodigoEmpresa, _
                                           "", _
                                           "", _
                                           enumStatusMensagem.Informação, _
                                           enumIndicadorSimNao.Nao, _
                                           "", _
                                           enumLocalLiquidacao.SELIC, _
                                           "", _
                                           strSiglaSistema
        Set objMensagemRecebidaEnviada = Nothing
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMensagemTratadaRDC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Controlar o fluxo de distribuição de uma operação identifica pela mensagem SPB recebida
'-------------------------------------------------------------------------------------------------

Private Function flGerenciarChamada(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                    ByVal plngStatusOperacao As Long, _
                                    ByVal plngStatusMensagem As Long, _
                                    ByVal pstrSituacaoMensagem As String, _
                                    ByVal plngEstorno As enumIndicadorSimNao, _
                                    ByVal pstrNumeroControleIF As String, _
                           Optional ByVal pstrSituacaoRecebida As String = vbNullString) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objSubReserva                           As A8LQS.clsSubReserva
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objLegado                               As A8LQS.clsLegado
Dim objContaCorrente                        As A8LQS.clsContaCorrente
Dim objAlerta                               As A8LQS.clsAlerta
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strXML                                  As String
Dim blnCliente1                             As Boolean
Dim blnEnviarInformacaoPJ                   As Boolean
Dim blnIncluirHistorico                     As Boolean

    On Error GoTo ErrorHandler

    blnEnviarInformacaoPJ = True
    
    'Alterar Status Operação
    If plngStatusOperacao > 0 Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      plngStatusOperacao, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , , False, , pstrSituacaoRecebida
        Set objOperacao = Nothing
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = plngStatusOperacao
    End If

    'Alterar o Status da Mensagem
    If plngStatusMensagem > 0 Then
        Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Set xmlMensagemEnviadaRecebida = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagemEnviadaRecebida.loadXML objMensagemEnviadaRecebida.LerMensagemIda(pstrNumeroControleIF)
        
        blnIncluirHistorico = True
        If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
            Select Case Val(xmlOperacao.selectSingleNode("//TP_MESG").Text)
                Case enumTipoMensagemBUS.Redesconto
                    blnIncluirHistorico = False
            End Select
        End If
        
        objMensagemEnviadaRecebida.AlterarStatus xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                 plngStatusMensagem, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                 , , , , , , , , , , _
                                                 blnIncluirHistorico
        
        Set objMensagemEnviadaRecebida = Nothing
    End If
    
    'Carlos - Gerenciamento de chamada das rotinas
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    strXML = objProcessoOperacao.LerTodos("RecebimentoR1", _
                                          CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text), _
                                          plngEstorno, _
                                          plngStatusOperacao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

    'Caso o sistema esteja em contingência na entrada manual, é para respeitar a solicitação de envio PJ e A8 que vier
    'na entrada manual.
    If Not xmlOperacao.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing And _
       Not xmlOperacao.documentElement.selectSingleNode("IN_SIST_SITU_CNTG") Is Nothing Then
       
        If xmlOperacao.documentElement.selectSingleNode("IN_ENTR_MANU").Text = enumIndicadorSimNao.Sim And _
           xmlOperacao.documentElement.selectSingleNode("IN_SIST_SITU_CNTG").Text = enumIndicadorSimNao.Sim Then
           
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_PJ").Text = enumIndicadorSimNao.Nao
            End If
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_PJ").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Nao
            End If

            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_PREV_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_PREV_A6").Text = enumIndicadorSimNao.Nao
            End If
            If xmlOperacao.documentElement.selectSingleNode("IN_ENVI_RELZ_SIST_A6").Text = enumIndicadorSimNao.Sim Then
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim
            Else
                xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Nao
            End If
            
        End If
        
    End If

    If fgVerificarCliente1(xmlOperacao) Then
        blnCliente1 = True
    Else
        blnCliente1 = False
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            If blnCliente1 Then
                fgDebitoCreditoCliente1 xmlOperacao
            End If
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            
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
                strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            End If
            
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgDebitoCreditoCliente1 xmlOperacao
                fgConverterProduto1 xmlOperacao
                strMensagem = objSubReserva.EnviarRealizadoConfirmado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                fgConverterProduto1 xmlOperacao
            End If
            Set objSubReserva = Nothing
        End If
    End If

    If Not xmlOperacao.selectSingleNode("//NumOpSELRet") Is Nothing Then
        If Not xmlMensagem.selectSingleNode("//NumOpSELRet") Is Nothing Then
            xmlOperacao.selectSingleNode("//NumOpSELRet").Text = xmlMensagem.selectSingleNode("//NumOpSELRet").Text
        End If
    End If
    
    If Not xmlOperacao.selectSingleNode("//SitOpSEL") Is Nothing Then
        If Not xmlMensagem.selectSingleNode("//SitOpSEL") Is Nothing Then
            xmlOperacao.selectSingleNode("//SitOpSEL").Text = xmlMensagem.selectSingleNode("//SitOpSEL").Text
        End If
    End If

    If plngEstorno = enumIndicadorSimNao.Nao Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Disponibilizar xmlOperacao, False
            Set objContaCorrente = Nothing
        End If
    Else
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_DISP_LANC_CNTA_CRRT").Text = enumIndicadorSimNao.Sim Then
            Set objContaCorrente = CreateObject("A8LQS.clsContaCorrente")
            objContaCorrente.Cancelar xmlOperacao, False, False
            Set objContaCorrente = Nothing
        End If
    End If
    
    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_MESG_RETN").Text = enumIndicadorSimNao.Sim Then
        Set objLegado = CreateObject("A8LQS.clsLegado")
        strMensagem = objLegado.EnviarRetorno(xmlOperacao, xmlMensagem)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        Set objLegado = Nothing
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_ALER").Text = enumIndicadorSimNao.Sim Then
        Set objAlerta = CreateObject("A8LQS.clsAlerta")
        objAlerta.GerarAlerta xmlOperacao, xmlMensagem, pstrSituacaoMensagem
        Set objAlerta = Nothing
    End If

    'Estornos de previsão
    objProcessoOperacao.EstornarPrevisao xmlOperacao, xmlControleProcessoOperacao

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_SOLI_A6").Text = enumIndicadorSimNao.Sim Then
        If blnCliente1 Then
            fgDebitoCreditoCliente1 xmlOperacao
        End If
        Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
        strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
        objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
        If blnCliente1 Then
            'irá retornar o debito/credito original
            fgDebitoCreditoCliente1 xmlOperacao
            fgConverterProduto1 xmlOperacao
            strMensagem = objSubReserva.EstornarRealizadoSolicitado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            fgConverterProduto1 xmlOperacao
        End If
        Set objSubReserva = Nothing
    End If

    If plngEstorno = enumIndicadorSimNao.Sim Then
        If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_CONF_A6").Text = enumIndicadorSimNao.Sim Then
            If blnCliente1 Then
                fgDebitoCreditoCliente1 xmlOperacao
            End If
            Set objSubReserva = CreateObject("A8LQS.clsSubReserva")
            strMensagem = objSubReserva.EstornarRealizadoConfirmado(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgDebitoCreditoCliente1 xmlOperacao
                fgConverterProduto1 xmlOperacao
                strMensagem = objSubReserva.EstornarRealizadoConfirmado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.SubReserva
                fgConverterProduto1 xmlOperacao
            End If
            Set objSubReserva = Nothing
        End If
    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_ENVI_RELZ_PJ").Text = enumIndicadorSimNao.Sim Then
        If xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text <> vbNullString Then
            If blnCliente1 Then
                fgDebitoCreditoCliente1 xmlOperacao
            End If
            
            If plngEstorno = enumIndicadorSimNao.Nao Then
                If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
                   CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
                   CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
                   CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
                    Call fgAppendNode(xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim)
                End If
            End If
            
            Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
            
            If blnEnviarInformacaoPJ Then
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            End If
            
            If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TransferenciaCamaras And blnEnviarInformacaoPJ Then
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Sim
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                fgCriarTAGTransferencia xmlOperacao, enumIndicadorSimNao.Nao
            End If
            
            If blnCliente1 Then
                'irá retornar o debito/credito original
                fgDebitoCreditoCliente1 xmlOperacao
                fgConverterProduto1 xmlOperacao
                strMensagem = objGestaoCaixa.EstornarRealizado(xmlOperacao)
                objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
                fgConverterProduto1 xmlOperacao
            End If
            Set objGestaoCaixa = Nothing
        End If
    End If

    'incluído por Carlos 12/05/2004
    'comentado por Bruno Oliveira 19/07/2010 devido reestruturação das contas Selic
'    If fgVerificarCliente1(xmlOperacao, , False) Then
'        objProcessoOperacao.flLiquidarCliente1 xmlOperacao, True
'    End If

    'Enviar previsao do pagto do redesconto
    If plngStatusOperacao = enumStatusOperacao.Liquidada Then
        If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
            CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
            objProcessoOperacao.EnviarPrevisaoRDSC xmlOperacao
        End If
    End If

    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlControleProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
    '    ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGerenciarChamada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'------------------------------------------------------------------------------------------
'Obter o nome da tag de situação da operação / mensagem
'------------------------------------------------------------------------------------------

Public Function ObterNomeTagSituacao(ByVal pstrCodigoGrupoMensagem As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   DISTINCT(NO_TAG) as NO_TAG " & vbCrLf & _
             "   FROM   A8.TB_SITU_SPB_SITU_PROC " & vbCrLf & _
             "  WHERE   SG_GRUP_MESG_SPB    =   '" & pstrCodigoGrupoMensagem & "'"

    ObterNomeTagSituacao = fgQueryXMLLerTodos("NomeTagSituacao", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNomeTagSituacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'-------------------------------------------------------------------------------------------------
'Executar a leitura de todos os registros tabela de situação do processamento
'-------------------------------------------------------------------------------------------------

Public Function LerTodosTagDominioStatus(ByVal pstrCodigoGrupoMensagem As String, _
                                Optional ByVal pstrNomeTag As String = vbNullString) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   A.SG_GRUP_MESG_SPB,             " & vbCrLf & _
             "          A.NO_TAG,                       " & vbCrLf & _
             "          A.DE_DOMI,                      " & vbCrLf & _
             "          A.CO_SITU_PROC_OPER_ATIV,       " & vbCrLf & _
             "          A.CO_SITU_PROC_MESG_SPB,        " & vbCrLf & _
             "          A.CO_USUA_ULTI_ATLZ,            " & vbCrLf & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,       " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                 " & vbCrLf & _
             "          B.SQ_TIPO_TAG                   " & vbCrLf & _
             " FROM     A8.TB_SITU_SPB_SITU_PROC A,     " & vbCrLf & _
             "          A8.TB_TAG                B      " & vbCrLf & _
             " WHERE    A.NO_TAG           = B.NO_TAG   " & vbCrLf & _
             " AND      A.SG_GRUP_MESG_SPB = '" & pstrCodigoGrupoMensagem & "'" & vbCrLf
    
    If pstrNomeTag <> vbNullString Then
        strSQL = strSQL & " AND      A.NO_TAG = '" & pstrNomeTag & "'"
    End If
    
    LerTodosTagDominioStatus = fgQueryXMLLerTodos("TagDominioStatus", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosTagDominioStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo do repasse financeiro, caso  se encontre pendente

Private Function flTratarRepasse(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim domNode                                 As MSXML2.IXMLDOMNode
Dim strXML                                  As String

On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    If objOperacao.VerificarRepassePendente(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text) Then
        Exit Function
    End If

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strXML = objConciliacao.LerTodos(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    If strXML = vbNullString Then
        Exit Function
    End If

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlConciliacao.loadXML strXML

    For Each domNode In xmlConciliacao.documentElement.selectNodes("//Grupo_Conciliacao")
        objOperacao.AlterarStatusOperIndisp xmlConciliacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Liquidada
    Next

    Set objOperacao = Nothing
    Set objConciliacao = Nothing
    Set xmlConciliacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objOperacao = Nothing
    Set objConciliacao = Nothing
    Set xmlConciliacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTratarRepasse Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de recebimento e gravação da mensagem SBP SEL1069

Private Function flTratarSEL1069(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                 ByRef pstrHeaderNZ As String) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoConciliada                   As MSXML2.DOMDocument40
Dim xmlGrupoMensagem                        As MSXML2.DOMDocument40
Dim xmlTagDominioStatus                     As MSXML2.DOMDocument40
Dim xmlMensagemIda                          As MSXML2.DOMDocument40
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim domNode                                 As MSXML2.IXMLDOMNode
Dim domNodeOperacao                         As MSXML2.IXMLDOMNode
Dim strGrupoMensagem                        As String
Dim strTagDominioStatus                     As String
Dim strNomeTag                              As String
Dim strMensagem                             As String
Dim strMensagemIda                          As String
Dim strSituacaoMensagem                     As String
Dim strOperacao                             As String
Dim strCodigoMensagem                       As String
Dim strTipoOperacao                         As String
Dim blnGravar                               As Boolean
Dim blnAlterarStatus                        As Boolean

'Pikachu - 14/06/2004
'Alteração de mensagem Rejeitada Piloto
Dim lngStatusMensagem                       As Long

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Set xmlTagDominioStatus = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlGrupoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacaoConciliada = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    'Travar para processar uma mensagem por vez
    objTipoOperacao.BloquearTipoOperacaoParaInclusao

    strMensagemIda = objMensagemRecebidaEnviada.LerMensagemIda(xmlMensagem.selectSingleNode("//NumCtrlIF").Text)

    If strMensagemIda = vbNullString Then
        ' 3022 - Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    Set xmlMensagemIda = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMensagemIda.loadXML strMensagemIda

    'Ler todas mensagem com a mesma sequencia
    If Not flLerSEL1069(xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text) Then
        flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
        Exit Function
    End If

    'Obter a Operações conciliadas
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strOperacao = objConciliacao.LerSequenciaOperacao(xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)

    If strOperacao = vbNullString Then
        ' 3022 - Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    xmlOperacaoConciliada.loadXML strOperacao

    strGrupoMensagem = objMensagem.LerGrupoMensagem(Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7))

    If strGrupoMensagem = vbNullString Then
        '3016 - Mensagem sem grupo cadastrado.
        lngCodigoErroNegocio = 3016
        GoTo ErrorHandler
    End If
    xmlGrupoMensagem.loadXML strGrupoMensagem

    strTagDominioStatus = LerTodosTagDominioStatus(xmlGrupoMensagem.documentElement.selectSingleNode("CO_GRUP").Text)
    If strTagDominioStatus = vbNullString Then
        '3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
        lngCodigoErroNegocio = 3017
        GoTo ErrorHandler
    End If
    xmlTagDominioStatus.loadXML strTagDominioStatus

    strNomeTag = xmlTagDominioStatus.documentElement.selectSingleNode("//NO_TAG").Text

    'Incluir a Mensagem
    If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag) Is Nothing Then
        strSituacaoMensagem = ""
    Else
        strSituacaoMensagem = xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text
    End If

    strSituacaoRecebida = vbNullString
    If xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG") Is Nothing Then
        strSituacaoRecebida = Format$(xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG").Text, "00000") & "|" & _
                              strSituacaoMensagem
    End If
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    blnGravar = True
    blnAlterarStatus = True

    'Para SEL1069 sem operação
    If Val(xmlOperacaoConciliada.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) = 0 Then
        flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
        Exit Function
    End If

    For Each domNodeOperacao In xmlOperacaoConciliada.documentElement.selectNodes("//Grupo_Conciliacao")

        Set xmlOperacao = objOperacao.LerXMLOperacao(domNodeOperacao.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

        If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If

        strTipoOperacao = objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        xmlTipoOperacao.loadXML strTipoOperacao

        If xmlOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text
        Else
            xmlOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text = xmlTipoOperacao.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text
        End If

        If blnGravar Then
            
            'Pikachu - 14/06/2004
            'Alteração de mensagem Rejeitada Piloto
            If xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1069E" Then
                lngStatusMensagem = enumStatusMensagem.ErroNegocioSPB
            Else
                lngStatusMensagem = enumStatusMensagem.R1
            End If
            '---------------------------------------
            
            objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                               xmlMensagem.selectSingleNode("//NumCtrlIF").Text, _
                                               xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                               xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                               enumTipoBackOffice.Todos, _
                                               xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                               "", _
                                               strSituacaoMensagem, _
                                               lngStatusMensagem, _
                                               enumIndicadorSimNao.Nao, _
                                               "", _
                                               xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                               "", _
                                               xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
            blnGravar = False
        End If

        If xmlMensagem.selectSingleNode("//CodMsg").Text <> "SEL1069E" Then

            'Verificar qual o status da mensagem e qual deverá ser o próximo status da Operação
            For Each domNode In xmlTagDominioStatus.documentElement.selectNodes("//Grupo_TagDominioStatus")
                If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text = domNode.selectSingleNode("DE_DOMI").Text Then
                    If blnAlterarStatus = False Then
                        domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = ""
                    End If
                    blnAlterarStatus = False
                    Select Case domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                        Case enumStatusOperacao.Liquidada, _
                             enumStatusOperacao.EmLancamento, _
                             enumStatusOperacao.Pendencia
                             flGerenciarChamada xmlOperacao, _
                                                xmlMensagem, _
                                                CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text)), _
                                                CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                                xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                                enumIndicadorSimNao.Nao, _
                                                xmlMensagem.selectSingleNode("//NumCtrlIF").Text
                             Exit For
                        Case enumStatusOperacao.Expirada, _
                             enumStatusOperacao.Rejeitada, _
                             enumStatusOperacao.Cancelada, _
                             enumStatusOperacao.Estornada
                             flGerenciarChamada xmlOperacao, _
                                                xmlMensagem, _
                                                0, _
                                                CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                                xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                                enumIndicadorSimNao.Sim, _
                                                xmlMensagem.selectSingleNode("//NumCtrlIF").Text, _
                                                strSituacaoRecebida
                            Exit For
                    End Select
                End If
            Next
        Else
            'Mensagens com Erros de Negócios ou NZ
            flGerenciarChamada xmlOperacao, _
                               xmlMensagem, _
                               0, _
                               enumStatusMensagem.ErroNegocioSPB, _
                               "ERR", _
                               enumIndicadorSimNao.Sim, _
                               xmlMensagem.selectSingleNode("//NumCtrlIF").Text
        End If
    Next

    Set objMensagemRecebidaEnviada = Nothing
    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoConciliada = Nothing
    Set objTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoConciliada = Nothing
    Set objTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTratarSEL1069 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Executar a leitura de mensagem SEL1069 com a situação de encerrada para uma conciliação específica

Private Function flLerSEL1069(ByVal pvntSequenciaConciliacao As Variant) As Boolean

Dim rsMesgDespesa                           As ADODB.Recordset
Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_CTRL_IF,                                                                            " & vbNewLine & _
             "          DH_REGT_MESG_SPB,                                                                      " & vbNewLine & _
             "          CO_ULTI_SITU_PROC                                                                      " & vbNewLine & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB                                                               " & vbNewLine & _
             " WHERE    NU_SEQU_CNCL_OPER_ATIV_MESG  =  " & pvntSequenciaConciliacao & "                       " & vbNewLine & _
             " AND      CO_ULTI_SITU_PROC            <> " & enumStatusMensagem.Encerrada & vbNewLine & _
             " AND      CO_MESG_SPB                  = 'SEL1069' "

    Set rsMesgDespesa = fgQuerySQL(strSQL)

    If rsMesgDespesa.RecordCount > 1 Then
        flLerSEL1069 = False
    Else
        flLerSEL1069 = True
    End If

    Set rsMesgDespesa = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set rsMesgDespesa = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLerSEL1069 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de gravação da mensagem SBP RDC0005R1

Private Function flGravarRDC0005R1(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                   ByRef pstrHeaderNZ As String, _
                                   ByRef pstrControleRDC As String) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMensagemEnviadaRecebida              As MSXML2.DOMDocument40
Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim udtHeaderMensagem                       As udtHeaderMensagem
Dim udtHeaderMensagemAux                    As udtHeaderMensagemAux
Dim lngCodigoEmpresa                        As Long
Dim strSiglaSistema                         As String

    On Error GoTo ErrorHandler

    flGravarRDC0005R1 = False

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlMensagem.documentElement.selectSingleNode("//SitLancRDC").Text = "1" Then
        Set xmlOperacao = objOperacao.ObterOperacaoPorControleIF(xmlMensagem.documentElement.selectSingleNode("//NumCtrlIF").Text)
        If xmlOperacao.xml = vbNullString Then
            flMensagemNaoTratada xmlMensagem, pstrHeaderNZ
            Exit Function
        End If
        If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      enumStatusOperacao.Liquidada, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                      , , False
        Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Set xmlMensagemEnviadaRecebida = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagemEnviadaRecebida.loadXML objMensagemRecebidaEnviada.LerMensagemIda(xmlMensagem.documentElement.selectSingleNode("//NumCtrlIF").Text)
        objMensagemRecebidaEnviada.AlterarStatus xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("NU_CTRL_IF").Text, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                 enumStatusMensagem.Encerrada, _
                                                 xmlMensagemEnviadaRecebida.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
        Set objMensagemRecebidaEnviada = Nothing
        pstrControleRDC = xmlOperacao.documentElement.selectSingleNode("NU_CTRL_RDSC_ORIG").Text
    Else
        'Incluir a Mensagem
        udtHeaderMensagemAux.String = pstrHeaderNZ
        LSet udtHeaderMensagem = udtHeaderMensagemAux
        lngCodigoEmpresa = udtHeaderMensagem.CodigoEmpresa
        strSiglaSistema = udtHeaderMensagem.SiglaSistemaEnviouNZ
        Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                           xmlMensagem.documentElement.selectSingleNode("//NumCtrlIF").Text, _
                                           0, _
                                           0, xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, _
                                           enumTipoBackOffice.Todos, _
                                           lngCodigoEmpresa, _
                                           "", _
                                           "", _
                                           enumStatusMensagem.R1, _
                                           enumIndicadorSimNao.Nao, _
                                           "", _
                                           0, _
                                           "", _
                                           strSiglaSistema
        Set objMensagemRecebidaEnviada = Nothing
        Exit Function
    End If

    flGravarRDC0005R1 = True

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGravarRDC0005R1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de envio de Previsao e Realizado para o sistema PJ para operações de conversão de redesconto

Private Function flEnviarConversaoRDC(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim strMensagem                             As String
Dim lngCodigoProduto                        As Long
Dim datDataD1                               As Date
Dim intDebitoCredito                        As Integer

    On Error GoTo ErrorHandler

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")

    lngCodigoProduto = CLng(xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text)

    '316 - Redesconto Intradia - Entrada
    intDebitoCredito = CInt(xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text)
    xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = intDebitoCredito

    '68 - Redesconto Over - Entrada
    xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "68"
    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS

    '68 - Redesconto Over - Entrada
    xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "68"
    strMensagem = objGestaoCaixa.EnviarRealizado(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS

    'Adiciona mais 1 dia
    Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    datDataD1 = objA6A7A8Funcoes.AdicionarDiasUteis(fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text), _
                                                    1, enumPaginacao.Proximo)

    '317 - Redesconto Over - Saida
    intDebitoCredito = CInt(xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text)
    xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
    xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "317"
    xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = fgDt_To_Xml(datDataD1)
    strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
    objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
    xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = intDebitoCredito

    xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = lngCodigoProduto

    Set objGestaoCaixa = Nothing
    Set objProcessoOperacao = Nothing
    Set objA6A7A8Funcoes = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objProcessoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flEnviarConversaoRDC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Controlar o fluxo de gravação e distribuição da mensagem SEL1023

Private Function flTratarSEL1023(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                 ByRef pstrHeaderNZ As String) As Boolean

Dim objMensagem                             As A8LQS.clsMensagem
Dim objMensagemRecebidaEnviada              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objProcessoMensagemBMA                   As A8LQS.clsProcessoMensagemBMA

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoConciliada                   As MSXML2.DOMDocument40
Dim xmlGrupoMensagem                        As MSXML2.DOMDocument40
Dim xmlTagDominioStatus                     As MSXML2.DOMDocument40
Dim xmlMensagemIda                          As MSXML2.DOMDocument40
Dim domNode                                 As MSXML2.IXMLDOMNode
Dim domNodeOperacao                         As MSXML2.IXMLDOMNode
Dim strGrupoMensagem                        As String
Dim strTagDominioStatus                     As String
Dim strNomeTag                              As String
Dim strMensagem                             As String
Dim strMensagemIda                          As String
Dim strSituacaoMensagem                     As String
Dim strOperacao                             As String
Dim strCodigoMensagem                       As String
Dim blnPrimeiro                             As Boolean
    
'Pikachu - 07/10/2004
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim strRet                                  As String
Dim blnAcaoPagtoContingencia                As Boolean

Dim strSituacaoRecebida                     As String

On Error GoTo ErrorHandler

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set objMensagemRecebidaEnviada = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    'Travar para processar uma mensagem por vez
    objTipoOperacao.BloquearTipoOperacaoParaInclusao

    strMensagemIda = objMensagemRecebidaEnviada.LerMensagemIda(xmlMensagem.selectSingleNode("//NumCtrlIF").Text)
    
    If strMensagemIda = vbNullString Then
        ' 3022 - Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
    
    Set xmlMensagemIda = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMensagemIda.loadXML strMensagemIda
    
    'Pikachu - 9/07/2004
    'Tratamento para SEL1023 - Leilao BMA
    If Not xmlMensagemIda.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If xmlMensagemIda.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA Then
            'Tratamento para SEL1023 - Leilao BMA
            'Senão Trata Liquidação Fisica BMA
            If Trim$(xmlMensagemIda.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) <> vbNullString Then
                Set objProcessoMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
                Call objProcessoMensagemBMA.ProcessaMensagemSELIC_BMA(xmlMensagem, pstrHeaderNZ)
                Set objProcessoMensagemBMA = Nothing
                
                Set objTipoOperacao = Nothing
                Set objMensagemRecebidaEnviada = Nothing
                Exit Function
            End If
        End If
    End If
    '---------------------------------------------------
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    Set xmlTagDominioStatus = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlGrupoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacaoConciliada = CreateObject("MSXML2.DOMDocument.4.0")

    'Obter a Operações conciliadas
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strOperacao = objConciliacao.LerTodos(xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)

    If strOperacao = vbNullString Then
        ' 3022 - Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If

    xmlOperacaoConciliada.loadXML strOperacao

    strGrupoMensagem = objMensagem.LerGrupoMensagem(Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 7))

    If strGrupoMensagem = vbNullString Then
        '3038 - Mensagem sem grupo cadastrado.
        lngCodigoErroNegocio = 3038
        GoTo ErrorHandler
    End If
    xmlGrupoMensagem.loadXML strGrupoMensagem

    strTagDominioStatus = LerTodosTagDominioStatus(xmlGrupoMensagem.documentElement.selectSingleNode("CO_GRUP").Text)
    If strTagDominioStatus = vbNullString Then
        '3017 - Grupo de mensagem sem Situação SPB X Situação Processamento cadastrado.
        lngCodigoErroNegocio = 3017
        GoTo ErrorHandler
    End If
    xmlTagDominioStatus.loadXML strTagDominioStatus

    strNomeTag = xmlTagDominioStatus.documentElement.selectSingleNode("//NO_TAG").Text

    'Incluir a Mensagem
    If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag) Is Nothing Then
        strSituacaoMensagem = ""
    Else
        strSituacaoMensagem = xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text
    End If

    strSituacaoRecebida = vbNullString
    If xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG") Is Nothing Then
        strSituacaoRecebida = Format$(xmlTagDominioStatus.selectSingleNode("//SQ_TIPO_TAG").Text, "00000") & "|" & _
                              strSituacaoMensagem
    End If
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    blnPrimeiro = True

    For Each domNodeOperacao In xmlOperacaoConciliada.documentElement.selectNodes("//Grupo_Conciliacao")

        Set xmlOperacao = objOperacao.LerXMLOperacao(domNodeOperacao.selectSingleNode("NU_SEQU_OPER_ATIV").Text)

        'Pikachu - 07/10/2004
        'Verifica se foi pagamento em contingencia
        '-------------------------------------------------------------------------
        Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
        
        strRet = objHistMensagem.LerTodos(domNodeOperacao.selectSingleNode("//NU_CTRL_IF").Text, _
                                          domNodeOperacao.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                          domNodeOperacao.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
         
        xmlHistorico.loadXML strRet
                        
        blnAcaoPagtoContingencia = IIf(xmlHistorico.selectSingleNode("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']") Is Nothing, False, True)
        
        Set objHistMensagem = Nothing
        Set xmlHistorico = Nothing
        '----------------------------------------------------------------------------

        If CLng(xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Then
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If

        If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.SELIC Then
            objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                               domNodeOperacao.selectSingleNode("NU_CTRL_IF").Text, _
                                               xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                               xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                               xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                               enumTipoBackOffice.Todos, _
                                               xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                               "", _
                                               strSituacaoMensagem, _
                                               enumStatusMensagem.R1, _
                                               enumIndicadorSimNao.Nao, _
                                               "", _
                                               xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                               "", _
                                               xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
        Else
            If blnPrimeiro Then
                blnPrimeiro = False
                objMensagemRecebidaEnviada.Incluir xmlMensagem.xml, _
                                                   domNodeOperacao.selectSingleNode("NU_CTRL_IF").Text, _
                                                   0, _
                                                   xmlMensagemIda.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                                   xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                   enumTipoBackOffice.Todos, _
                                                   xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                   "", _
                                                   strSituacaoMensagem, _
                                                   enumStatusMensagem.R1, _
                                                   enumIndicadorSimNao.Nao, _
                                                   "", _
                                                   xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                                   "", _
                                                   xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text
            End If
        End If

        'Verificar qual o status da mensagem e qual deverá ser o próximo status da Operação
        For Each domNode In xmlTagDominioStatus.documentElement.selectNodes("//Grupo_TagDominioStatus")
            If xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text = domNode.selectSingleNode("DE_DOMI").Text Then
                Select Case domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                    Case enumStatusOperacao.Liquidada, _
                         enumStatusOperacao.EmLancamento, _
                         enumStatusOperacao.Pendencia
                         
                        'Pikachu - 07/10/2004
                        'Liquidacao Fisica
                        Select Case domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text
                            Case enumStatusOperacao.Liquidada
                                domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = enumStatusMensagem.MensagemLiquidada
                            Case enumStatusOperacao.EmLancamento
                                domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = enumStatusMensagem.MensagemEmLancamento
                            Case enumStatusOperacao.Pendencia
                                domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = enumStatusMensagem.MensagemPendente
                        End Select
                         
                         If blnAcaoPagtoContingencia Then
                            domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = ""
                         End If
                         
                         flGerenciarChamada xmlOperacao, _
                                            xmlMensagem, _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text)), _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                            xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                            enumIndicadorSimNao.Nao, _
                                            xmlMensagem.selectSingleNode("//NumCtrlIF").Text
                         Exit For
                    Case enumStatusOperacao.Expirada, _
                         enumStatusOperacao.Rejeitada, _
                         enumStatusOperacao.Cancelada, _
                         enumStatusOperacao.Estornada
                         flGerenciarChamada xmlOperacao, _
                                            xmlMensagem, _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_OPER_ATIV").Text)), _
                                            CLng(IIf(domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text = "", 0, domNode.selectSingleNode("CO_SITU_PROC_MESG_SPB").Text)), _
                                            xmlMensagem.documentElement.selectSingleNode("//" & strNomeTag).Text, _
                                            enumIndicadorSimNao.Sim, _
                                            xmlMensagem.selectSingleNode("//NumCtrlIF").Text, _
                                            strSituacaoRecebida
                        Exit For
                End Select
            End If
        Next
    Next

    Set objMensagemRecebidaEnviada = Nothing
    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoConciliada = Nothing
    Set objTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objHistMensagem = Nothing
    Set xmlHistorico = Nothing
    Set objMensagem = Nothing
    Set xmlTagDominioStatus = Nothing
    Set xmlGrupoMensagem = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoConciliada = Nothing
    Set objTipoOperacao = Nothing
    Set objProcessoMensagemBMA = Nothing
        
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTratarSEL1023 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o status interno do sistema SLCC de acordo com o status recebido na SEL1611

Private Function flObterStatusSEL1611(ByVal pstrSituacaoMensagem As String) As Long

Dim strSituacaoMensagem                     As String

On Error GoTo ErrorHandler
        
    Select Case Trim(pstrSituacaoMensagem)
        Case "ATU", "CON", "LAN", "LIB" ' A Concliar
            flObterStatusSEL1611 = enumStatusMensagem.AConciliar
        Case "EXP"                      ' Expirada
            flObterStatusSEL1611 = enumStatusMensagem.MensagemExpirada
        Case "PEN", "PEO"               ' Pendente
            flObterStatusSEL1611 = enumStatusMensagem.MensagemPendente
        Case "RLF", "RST", "STR", "EST" 'Rejeitada
            flObterStatusSEL1611 = enumStatusMensagem.MensagemRejeitada
        Case Else
            flObterStatusSEL1611 = enumStatusMensagem.Informação
    End Select
    
'A CONCILIAR ATU Atualizada
'A CONCILIAR CON Confirmada
'A CONCILIAR LAN Lançada
'A CONCILIAR LIB Liberada
'EXPIRADA    EXP Operação Expirada por Tempo
'PENDENTE    PEN Pendente de Título
'PENDENTE    PEO Pendente de Operação
'REJEITADA   RLF Operação Rejeitada por Falta de Limite Financeiro
'REJEITADA   RST Operação Rejeitada pelo STR
'REJEITADA   STR Operação Enviada ao STR
'REJEITADA   EST Estornada
    
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterStatusSEL1611 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Eduardo Quaresma 05/01/2004
Private Function flValidarRedesconto(ByVal pstrCtrlRDSCOrigem As String, _
                                     ByVal plngCodigoEmpresa As Long, _
                                     ByVal pdatDataOperacao As Date) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacaoAux                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objGestaoCaixa                          As A8LQS.clsGestaoCaixa
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim strRetorno                              As String
Dim intDebitoCredito                        As Integer
Dim datOperacaoAtivd                        As Date
Dim strMensagem                             As String
Dim strCodigoProduto                        As String

On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacaoAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set objGestaoCaixa = CreateObject("A8LQS.clsGestaoCaixa")
    Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    strRetorno = objOperacao.LerTodos(, , , , , enumTipoOperacaoLQS.ConversaoRedesconto, pstrCtrlRDSCOrigem, , plngCodigoEmpresa)

    If Trim(strRetorno) = "" Then
        strRetorno = objOperacao.LerTodos(, , , , , enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil, pstrCtrlRDSCOrigem, , plngCodigoEmpresa)
    End If

    If xmlOperacaoAux.loadXML(strRetorno) Then
    'Se a data de operação for igual a data de operação original
        If fgDtXML_To_Date(xmlOperacaoAux.selectSingleNode("//DT_OPER_ATIV").Text) = pdatDataOperacao Then

            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacaoAux.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

            'Estorna o provisto D+1 do produto 317
            'Adiciona mais 1 dia
            Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
            datOperacaoAtivd = objA6A7A8Funcoes.AdicionarDiasUteis(pdatDataOperacao, _
                                                                   1, enumPaginacao.Proximo)
            '317 - Redesconto Over - Saida
            intDebitoCredito = CInt(xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text)
            xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumIndicadorEntradaSaida.Saida
            xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "317"
            xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = fgDt_To_Xml(datOperacaoAtivd)
            strMensagem = objGestaoCaixa.EstornarPrevisao(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
            'Realiza previsão D0

            'Produto original
            xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = intDebitoCredito
            xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = fgDt_To_Xml(pdatDataOperacao)
            xmlOperacao.documentElement.selectSingleNode("CO_PROD").Text = "317"
            strMensagem = objGestaoCaixa.EnviarPrevisao(xmlOperacao)
            objProcessoOperacao.EnviarMensagemMQ strMensagem, enumIdentificadorFila.BUS
        End If
    End If

    flValidarRedesconto = True

    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set objGestaoCaixa = Nothing
    Set objA6A7A8Funcoes = Nothing
    Set objProcessoOperacao = Nothing

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidarRedesconto Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'para mensagens de Erro do NZ, inclui atributo CodErro nas respectivas tags criticadas no HeaderNZ
Public Function flIncluirAtributoCodErroTagsErroNZ(ByVal pstrMensagem As String, _
                                                   ByVal pstrHeader As String) As String

Dim udtProtocoloErroNZ                      As udtProtocoloErroNZ
Dim udtProtocoloErroNZAux                   As udtProtocoloErroNZAux
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strNomeTag                              As String
Dim strCodigoMensagem                       As String
            
On Error GoTo ErrorHandler
        
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pstrMensagem
    
    udtProtocoloErroNZAux.String = pstrHeader
    LSet udtProtocoloErroNZ = udtProtocoloErroNZAux
    strCodigoMensagem = Trim(udtProtocoloErroNZ.CodigoMensagem)
        
    If Len(strCodigoMensagem) = 7 Then 'mensagens de erro vindas do NZ/PK
        
        If strCodigoMensagem = "GEN0001" Then
            strCodigoMensagem = "GEN0004"
        End If
        
        If Trim(udtProtocoloErroNZ.CodigoErro1) <> "" Then
            strNomeTag = Left(udtProtocoloErroNZ.NomeDoCampo1, InStr(1, Replace(udtProtocoloErroNZ.NomeDoCampo1, ",", " "), " ") - 1)
            If strNomeTag <> vbNullString Then
                If Not xmlMensagem.selectSingleNode("//" & strNomeTag) Is Nothing Then
                    fgAppendAttribute xmlMensagem, strNomeTag, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro1)
                Else
                    fgAppendAttribute xmlMensagem, strCodigoMensagem, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro1)
                End If
            End If
        End If
        If Trim(udtProtocoloErroNZ.CodigoErro2) <> "" Then
            strNomeTag = Left(udtProtocoloErroNZ.NomeDoCampo2, InStr(1, Replace(udtProtocoloErroNZ.NomeDoCampo2, ",", " "), " ") - 1)
            If strNomeTag <> vbNullString Then
                If Not xmlMensagem.selectSingleNode("//" & strNomeTag) Is Nothing Then
                    fgAppendAttribute xmlMensagem, strNomeTag, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro2)
                Else
                    fgAppendAttribute xmlMensagem, strCodigoMensagem, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro2)
                End If
            End If
        End If
        If Trim(udtProtocoloErroNZ.CodigoErro3) <> "" Then
            strNomeTag = Left(udtProtocoloErroNZ.NomeDoCampo3, InStr(1, Replace(udtProtocoloErroNZ.NomeDoCampo3, ",", " "), " ") - 1)
            If strNomeTag <> vbNullString Then
                If Not xmlMensagem.selectSingleNode("//" & strNomeTag) Is Nothing Then
                    fgAppendAttribute xmlMensagem, strNomeTag, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro3)
                Else
                    fgAppendAttribute xmlMensagem, strCodigoMensagem, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro3)
                End If
            End If
        End If
        If Trim(udtProtocoloErroNZ.CodigoErro4) <> "" Then
            strNomeTag = Left(udtProtocoloErroNZ.NomeDoCampo4, InStr(1, Replace(udtProtocoloErroNZ.NomeDoCampo4, ",", " "), " ") - 1)
            If strNomeTag <> vbNullString Then
                If Not xmlMensagem.selectSingleNode("//" & strNomeTag) Is Nothing Then
                    fgAppendAttribute xmlMensagem, strNomeTag, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro4)
                Else
                    fgAppendAttribute xmlMensagem, strCodigoMensagem, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro4)
                End If
            End If
        End If
        If Trim(udtProtocoloErroNZ.CodigoErro5) <> "" Then
            strNomeTag = Left(udtProtocoloErroNZ.NomeDoCampo5, InStr(1, Replace(udtProtocoloErroNZ.NomeDoCampo5, ",", " "), " ") - 1)
            If strNomeTag <> vbNullString Then
                If Not xmlMensagem.selectSingleNode("//" & strNomeTag) Is Nothing Then
                    fgAppendAttribute xmlMensagem, strNomeTag, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro5)
                Else
                    fgAppendAttribute xmlMensagem, strCodigoMensagem, "CodErro", Trim(udtProtocoloErroNZ.CodigoErro5)
                End If
            End If
        End If
   End If

   flIncluirAtributoCodErroTagsErroNZ = xmlMensagem.xml

Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirAtributoCodErroTagsErroNZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

