VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGestaoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF769A40363"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável por montar as mensagens de envio para o Sistema PJ.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
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

'Enviar Precisão PJ
'   - Montar protocolo de integração A7
'   - Montar Remessa Previsão PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EnviarPrevisao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                     Optional ByVal pblnCompromissadaIda As Boolean = False, _
                     Optional ByVal strDataVencimento As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strPrevisao                             As String

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1004" 'Verificar qual será o evento do BUS para esta função.
    LSet udtProtocoloAux = udtProtocolo

    strPrevisao = flMontarPrevisao(xmlRemessa, enumTipoMovimentoPJ.Previsto, pblnCompromissadaIda, strDataVencimento)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.Previsto, _
            udtProtocoloAux.String & strPrevisao

    EnviarPrevisao = udtProtocoloAux.String & strPrevisao

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisao", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Previsao PJ - Moeda Estrangeira
'   - Montar protocolo de integração A7
'   - Montar Remessa Previsão PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EnviarPrevisaoME(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strPrevisao                             As String

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1014"
    LSet udtProtocoloAux = udtProtocolo

    strPrevisao = flMontarMoviMoedaEstrangeria(xmlRemessa, enumTipoMovimentoPJ.Previsto)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.Previsto, _
            udtProtocoloAux.String & strPrevisao

    EnviarPrevisaoME = udtProtocoloAux.String & strPrevisao

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisaoME", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Realizado PJ
'   - Montar protocolo de integração A7
'   - Montar Remessa realizado PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EnviarRealizado(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                       Optional ByVal pblnCompromissadaIda As Boolean = False, _
                       Optional ByVal strDataVencimento As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strMaioresValores                       As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1001" 'Verificar qual será o evento do BUS para esta função.
    LSet udtProtocoloAux = udtProtocolo

    If xmlRemessa.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.Previsto, pblnCompromissadaIda)
        Else
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.Realizado, pblnCompromissadaIda)
        End If
    Else
        If CLng(xmlRemessa.documentElement.selectSingleNode("ENVIAR_PREV").Text) = enumIndicadorSimNao.Sim Then
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.Previsto, pblnCompromissadaIda, strDataVencimento)
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
                strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.Previsto, pblnCompromissadaIda)
            Else
                strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.Realizado, pblnCompromissadaIda)
            End If
        End If
    End If

    If strMaioresValores = vbNullString Then
        EnviarRealizado = vbNullString
        Exit Function
    End If

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.RealizadoConfirmado, _
            udtProtocoloAux.String & strMaioresValores

    EnviarRealizado = udtProtocoloAux.String & strMaioresValores

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealizado Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Realizado Eventos SELIC PJ
'   - Montar protocolo de integração A7
'   - Montar Remessa realizado PJ (String)

Public Function EnviarRealizadoEventos(ByRef xmlMensagem As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strMaioresValores                       As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1001"
    LSet udtProtocoloAux = udtProtocolo

    strMaioresValores = flMaioresValoresEventos(xmlMensagem, enumTipoMovimentoPJ.Realizado)

    If strMaioresValores = vbNullString Then
        EnviarRealizadoEventos = vbNullString
        Exit Function
    End If

    'Grava Historico
    'Incluir xmlMensagem.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.RealizadoConfirmado, _
            udtProtocoloAux.String & strMaioresValores

    EnviarRealizadoEventos = udtProtocoloAux.String & strMaioresValores

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealizadoEventos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Estorno Previsão PJ
'   - Montar protocolo de integração A7
'   - Montar Remessa estorno Previsão PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EstornarPrevisao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                        Optional ByVal pblnCompromissadaIda As Boolean = False, _
                        Optional ByVal strDataVencimento As String = vbNullString, _
                        Optional ByVal strValorComplementacao As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strPrevisao                             As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1004" 'Verificar qual será o evento do BUS para esta função.
    LSet udtProtocoloAux = udtProtocolo

    strPrevisao = flMontarPrevisao(xmlRemessa, _
                                   enumTipoMovimentoPJ.EstornoPrevisto, _
                                   pblnCompromissadaIda, _
                                   strDataVencimento, _
                                   strValorComplementacao)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.EstornoPrevisto, _
            udtProtocoloAux.String & strPrevisao

    EstornarPrevisao = udtProtocoloAux.String & strPrevisao

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarPrevisao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Estorno Previsao PJ - Moeda Estrangeira
'   - Montar protocolo de integração A7
'   - Montar Remessa estorno Previsão PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EstornarPrevisaoME(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strPrevisao                             As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1014"
    LSet udtProtocoloAux = udtProtocolo

    strPrevisao = flMontarMoviMoedaEstrangeria(xmlRemessa, enumTipoMovimentoPJ.EstornoPrevisto)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.EstornoPrevisto, _
            udtProtocoloAux.String & strPrevisao

    EstornarPrevisaoME = udtProtocoloAux.String & strPrevisao

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarPrevisaoME Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar Estorno Realizado PJ
'   - Montar protocolo de integração A7
'   - Montar Remessa estorno realizado PJ (XML)
'   - Incluir registro no histórico de envio PJ

Public Function EstornarRealizado(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                         Optional ByVal pblnCompromissadaIda As Boolean = False, _
                         Optional ByVal strDataVencimento As String = vbNullString, _
                         Optional ByVal strValorComplementacao As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strMaioresValores                       As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1001" 'Verificar qual será o evento do BUS para esta função.
    LSet udtProtocoloAux = udtProtocolo

    If xmlRemessa.documentElement.selectSingleNode("ENVIAR_PREV") Is Nothing Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.EstornoPrevisto, pblnCompromissadaIda)
        Else
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.EstornoRealizado, pblnCompromissadaIda)
        End If
    Else
        If CLng(xmlRemessa.documentElement.selectSingleNode("ENVIAR_PREV").Text) = enumIndicadorSimNao.Sim Then
            strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.EstornoPrevisto, pblnCompromissadaIda, strDataVencimento, strValorComplementacao)
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
                strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.EstornoPrevisto, pblnCompromissadaIda)
            Else
                strMaioresValores = flMontarMaioresValores(xmlRemessa, enumTipoMovimentoPJ.EstornoRealizado, pblnCompromissadaIda)
            End If
        End If
    End If

    If strMaioresValores = vbNullString Then
        EstornarRealizado = vbNullString
        Exit Function
    End If

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.EstornoRealizadoSolicitado, _
            udtProtocoloAux.String & strMaioresValores

    EstornarRealizado = udtProtocoloAux.String & strMaioresValores

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarRealizado Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Incluir registro na tabela A8.TB_HIST_ENVI_INFO_GEST_CAIX

Public Function Incluir(ByRef pvntSequenciaOperacao As Variant, _
                        ByRef plngTipoMovimento As enumTipoMovimento, _
                        ByRef pstrMensagem As String) As Boolean

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long

    On Error GoTo ErrorHandler

    lngCodigoTxtXML = fgInsertVarchar4000(pstrMensagem)

    strSQL = "Insert Into A8.TB_HIST_ENVI_INFO_GEST_CAIX (" & vbCrLf & _
             "  NU_SEQU_OPER_ATIV, " & vbCrLf & _
             "  DH_ENVI_GEST_CAIX, " & vbCrLf & _
             "  CO_SITU_MOVI_GEST_CAIX, " & vbCrLf & _
             "  CO_TEXT_XML) " & vbCrLf & _
             " Values ( " & vbCrLf & _
             pvntSequenciaOperacao & ", " & vbCrLf & _
             "  SYSDATE, " & vbCrLf & _
             plngTipoMovimento & ", " & vbCrLf & _
             lngCodigoTxtXML & ")"

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar remessa Maiores Valores PJ

Private Function flMontarMaioresValores(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                        ByRef plngTipoMovimentoPJ As enumTipoMovimentoPJ, _
                               Optional ByVal pblnCompromissadaIda As Boolean = False, _
                               Optional ByVal strDataVencimento As String = vbNullString, _
                               Optional ByVal strValorComplementacao As String = vbNullString) As String

Dim udtMaioresValores                       As udtMaioresValores
Dim udtMaioresValoresAux                    As udtMaioresValoresAux
Dim strCodigoRemessaIni                     As String
Dim strCodigoRemessaFim                     As String
Dim strDataOperacao                         As String
Dim vntValor                                As Variant
Dim lngIndicadorDebitoCredito               As Long
Dim strNomeCliente                          As String
Dim lngCodigoProduto                        As Long
Dim lngTipoMensagem                         As Long

Dim blnZerarLocalLiquidacao                 As Boolean
Dim blnAlterarCodigoProduto                 As Boolean

Dim strBanco                                As String
Dim strAgencia                              As String
Dim strConta                                As String
Dim strCNPJCPF                              As String

    On Error GoTo ErrorHandler

    blnZerarLocalLiquidacao = False
    blnAlterarCodigoProduto = True
    lngTipoMensagem = 0
    
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        lngTipoMensagem = Val(xmlRemessa.selectSingleNode("//TP_MESG").Text)
        
        If lngTipoMensagem = enumTipoMensagemBUS.OperacaoCompromissadaCETIP Or _
           lngTipoMensagem = enumTipoMensagemBUS.OperacaoRetornoAntecipacaoCETIP Then
            If Not xmlRemessa.selectSingleNode("//TP_CNPT") Is Nothing Then
                If Val(xmlRemessa.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Cliente1 Then
                    blnZerarLocalLiquidacao = True
                End If
            End If
        End If
    
        If lngTipoMensagem = enumTipoMensagemBUS.OperacaoRetornoAntecipacaoCETIP Then
            blnAlterarCodigoProduto = False
        End If
    End If
    
    With udtMaioresValores

        If strDataVencimento = vbNullString Then
            If xmlRemessa.getElementsByTagName("DT_OPER_ATIV").Item(0) Is Nothing Then
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
        Else
            strDataOperacao = strDataVencimento
        End If

        'Tem mensagem que não tem no lay-out a data da operação
        If pblnCompromissadaIda Then
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text
            If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing Then
                If Val("0" & xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text) > 0 Then
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text, ",", , vbBinaryCompare)
                Else
                    If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                        vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                    Else
                        flMontarMaioresValores = vbNullString
                        Exit Function
                    End If
                End If
            Else
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                   vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                Else
                   flMontarMaioresValores = vbNullString
                   Exit Function
                End If
            End If
            lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
            lngCodigoProduto = flObterProdutoPJCompromissadaVolta(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text)
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesRodaDolar Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesBMC Then
                
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Vinculo Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Desvinculo Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TransferenciaSemFinanceiro Then
                    
                    lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                    lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                    
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TransferenciaCamaras Then
                    If xmlRemessa.documentElement.selectSingleNode("PROD_CAMR") Is Nothing Then
                        'Utilizar o produto normal
                        lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                        lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                    Else
                        If xmlRemessa.documentElement.selectSingleNode("PROD_CAMR").Text = enumIndicadorSimNao.Sim Then
                            lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida)
                            lngCodigoProduto = CLng("0" & xmlRemessa.documentElement.selectSingleNode("CO_PROD_CAMR").Text)
                        Else
                            'Utilizar o produto normal
                            lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                            lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                        End If
                    End If
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DespesasBMC Then
                
                lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
            
                If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
                    strDataOperacao = xmlRemessa.documentElement.selectSingleNode("//DT_OPER_ATIV").Text
                Else
                    strDataOperacao = xmlRemessa.documentElement.selectSingleNode("//DT_LIQU_OPER_MOED_NACI").Text
                End If
                
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("//CO_PROD").Text
                If Not xmlRemessa.documentElement.selectSingleNode("//TP_OPER_CAMB") Is Nothing Then
                    If xmlRemessa.documentElement.selectSingleNode("//TP_OPER_CAMB").Text = "V" Then
                        lngIndicadorDebitoCredito = enumTipoEntradaSaida.ENTRADA
                    Else ' "C"
                        lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                    End If
                Else
                    lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                End If
                
            Else
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
            End If
            
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_MOED_NACIO").Text, ",", , vbBinaryCompare)
            Else
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                Else
                    flMontarMaioresValores = vbNullString
                    Exit Function
                End If
            End If
        
            'RATS 1033 | 06/dez/10 | Bruno Oliveira | para layout 94 (RegistroContratoSWAPCetip21) deve ser enviado o valor do prêmio
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_PREM") Is Nothing Then
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_PREM").Text, ",", , vbBinaryCompare)
                Else
                    If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
                        vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                    Else
                        flMontarMaioresValores = vbNullString
                        Exit Function
                    End If
                End If
            End If
        
            If strValorComplementacao <> vbNullString Then
                vntValor = Split(strValorComplementacao, ",", , vbBinaryCompare)
            End If
        End If

        'Se for Cliente1 foi tratado anteriormente
        If Not xmlRemessa.documentElement.selectSingleNode("CLIENTE1") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("CLIENTE1").Text = enumIndicadorSimNao.Sim Then
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
            End If
        End If

        fgObterIdentificadorRemessaPJ CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), _
                                      fgDataHoraServidor(DataAux), _
                                      1, _
                                      strCodigoRemessaIni, _
                                      strCodigoRemessaFim
        .TipoRemessa = "200"
        .CodigoRemessa = fgCompletaString(strCodigoRemessaIni, " ", 23, False)
        .DataRemessa = Format(Date, "YYYYMMDD")
        .HoraRemessa = Format(Time, "HHMM")
        .CodigoEmpresa = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, "0", 5, True)
        .SiglaSistema = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False)
        .CodigoMoeda = gstrCodigoMoeda
        .CodigoBanqueiro = Space(12)
        .TipoCaixa = fgCompletaString(enumTipoCaixaPJ.Futuro, "0", 3, True)
        .CodigoItemCaixa = fgCompletaString("", "0", 9, True)
        
        .CodigoProduto = fgCompletaString(lngCodigoProduto, "0", 4, True)
        
        If strDataVencimento <> vbNullString And blnAlterarCodigoProduto Then
            .CodigoProduto = fgCompletaString(fgAlterarCodigoProduto(lngCodigoProduto), "0", 4, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_CONT") Is Nothing Then
            .TipoConta = fgCompletaString("", "0", 3, True)
        Else
            .TipoConta = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("TP_CONT").Text, "0", 3, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("CO_SEGM") Is Nothing Then
            .CodigoSegmento = fgCompletaString("", "0", 3, True)
        Else
            .CodigoSegmento = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_SEGM").Text, "0", 3, True)
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC") Is Nothing Then
            .CodigoEventoFinanceiro = fgCompletaString("", "0", 3, True)
        Else
            .CodigoEventoFinanceiro = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC").Text, "0", 3, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("CO_INDX") Is Nothing Then
            .CodigoIndexador = fgCompletaString("", "0", 3, True)
        Else
            .CodigoIndexador = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_INDX").Text, "0", 3, True)
        End If

        .CodigoLocalLiquidacao = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, "0", 4, True)
        
        If Val(.CodigoLocalLiquidacao) = enumLocalLiquidacao.PAG Then
            .CodigoLocalLiquidacao = fgCompletaString(enumLocalLiquidacao.CIP, "0", 4, True)
        ElseIf Val(.CodigoLocalLiquidacao) = enumLocalLiquidacao.CONTA_CORRENTE Then
            blnZerarLocalLiquidacao = True
        End If
        
        If blnZerarLocalLiquidacao Then
            .CodigoLocalLiquidacao = "0000"
        End If
        
        .TipoMovimento = fgCompletaString(plngTipoMovimentoPJ, "0", 3, True)
        .DataMovimento = strDataOperacao
        .HoraMovimento = Format(Now, "HHmm")
        
        If strDataVencimento <> vbNullString Then
            lngIndicadorDebitoCredito = IIf(lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.ENTRADA, _
                                                enumIndicadorEntradaSaida.Saida, _
                                                enumIndicadorEntradaSaida.ENTRADA)
        End If
        
        .TipoEntradaSaida = fgCompletaString(lngIndicadorDebitoCredito, "0", 1, False)

        If UBound(vntValor) = 0 Then
            .ValorMovimento = Format(vntValor(0), String(15, "0")) & "00"
        ElseIf UBound(vntValor) = -1 Then
            .ValorMovimento = Format("0", String(15, "0")) & "00"
        Else
            .ValorMovimento = Format(vntValor(0), String(15, "0")) & fgCompletaString(vntValor(1), "0", 2, False)
        End If
        
        'solicitado pela Isban enviar Banco, Agencia e Conta p/ o sistema PJ nas operações de Compromissada CETIP - Bruno Oliveira - 25/nov/2011
        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.OperacaoCompromissadaCETIP _
        Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP _
        Or xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.OperacaoRetornoCompromissadaCETIP Then
            If Not xmlRemessa.documentElement.selectSingleNode("CO_BANC") Is Nothing Then
                strBanco = xmlRemessa.documentElement.selectSingleNode("CO_BANC").Text
            End If
            If Not xmlRemessa.documentElement.selectSingleNode("CO_AGEN") Is Nothing Then
                strAgencia = xmlRemessa.documentElement.selectSingleNode("CO_AGEN").Text
            End If
            If Not xmlRemessa.documentElement.selectSingleNode("NU_CC") Is Nothing Then
                strConta = xmlRemessa.documentElement.selectSingleNode("NU_CC").Text
            End If
            If Not xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CNPT") Is Nothing Then
                strCNPJCPF = xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CNPT").Text
            End If
        End If
        .CodigoBanco = fgCompletaString(strBanco, "0", 3, True)
        .CodigoAgencia = fgCompletaString(strAgencia, "0", 5, True)
        .NumeroContaCorrente = fgCompletaString(strConta, "0", 13, True)
        
        .TipoPessoa = 2
        .CodigoCNPJ_CPF = fgCompletaString(strCNPJCPF, "0", 15, True)

        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.DespesaSelic Then
            strNomeCliente = "SELIC"
        ElseIf xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RedescontoIntradia Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.ConversaoRedesconto Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.PagamentoRedesconto Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.PagamentoRedescontoAssociadoVenda Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Or _
            xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.ConsolidacaoRedescontosIntradia Then
            strNomeCliente = "BANCO CENTRAL DO BRASIL"
        ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
            strNomeCliente = xmlRemessa.documentElement.selectSingleNode("//NM_CLIE_MOED_ESTR").Text
        Else
            If xmlRemessa.documentElement.selectSingleNode("NO_CNPT") Is Nothing Then
                strNomeCliente = vbNullString
            Else
                strNomeCliente = xmlRemessa.documentElement.selectSingleNode("NO_CNPT").Text
            End If
        End If

        If lngTipoMensagem = enumTipoMensagemBUS.EnvioTEDClientes Then
            strNomeCliente = xmlRemessa.selectSingleNode("//NO_TITU_DEBT_1").Text
        End If
        
        .NomeCliente = fgCompletaString(strNomeCliente, " ", 64, False)
        .TipoProcessamento = enumTipoProcessamentoPJ.OnLine
        .TipoEnvio = enumTipoEnvioPJ.Parcial
        
        If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.Definitiva Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaIda Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.CompromissadaVolta Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.Termo Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoPrimario Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.VincDesvInstrumentoFinanceiro Or _
           xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.EventosJuros Then
            .NumeroOperacaoSelic = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, "0", 6, True)
            .ContaCedenteCessionaria = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text, "0", 9, True)
        Else
            .NumeroOperacaoSelic = String$(6, "0")
            .ContaCedenteCessionaria = String$(9, "0")
        End If

        .FILLER = Space(5)
    End With

    LSet udtMaioresValoresAux = udtMaioresValores

    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaLiquidacaoInterbancaria Then
        vntValor = xmlRemessa.documentElement.selectSingleNode("//VA_MOED_NACIO").Text
    Else
        If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            vntValor = xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text
        Else
            flMontarMaioresValores = vbNullString
            Exit Function
        End If
    End If

    If flVerificarEnvioMaiorValor(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                  lngCodigoProduto, _
                                  vntValor) Then
        flMontarMaioresValores = udtMaioresValoresAux.String
    Else
        flMontarMaioresValores = vbNullString
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMaioresValores Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Montar remessa Maiores Valores PJ para eventos SELIC

Private Function flMaioresValoresEventos(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                         ByRef plngTipoMovimentoPJ As enumTipoMovimentoPJ) As String

Dim udtMaioresValores                       As udtMaioresValores
Dim udtMaioresValoresAux                    As udtMaioresValoresAux
Dim strCodigoRemessaIni                     As String
Dim strCodigoRemessaFim                     As String
Dim strDataOperacao                         As String
Dim vntValor                                As Variant
Dim lngIndicadorDebitoCredito               As Long
Dim strNomeCliente                          As String
Dim lngCodigoProduto                        As Long
Dim lngTipoBackOffice                       As Long

    On Error GoTo ErrorHandler

    With udtMaioresValores

        strDataOperacao = xmlMensagem.documentElement.selectSingleNode("//DtMovto").Text
        vntValor = Split(xmlMensagem.documentElement.selectSingleNode("//VlrFinanc").Text, ",", , vbBinaryCompare)
        lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.ENTRADA

        If xmlMensagem.documentElement.selectSingleNode("//CO_PROD") Is Nothing Then
            If xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then
                lngTipoBackOffice = fgObterTipoBackOfficeUsuario
                If lngTipoBackOffice = enumTipoBackOffice.FundosProprios Or _
                    lngTipoBackOffice = enumTipoBackOffice.FundosTerceiros Then
                    lngCodigoProduto = 523
                Else
                    lngCodigoProduto = 522
                End If
            Else
                If Trim(xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text) = vbNullString Then
                    lngTipoBackOffice = fgObterTipoBackOfficeUsuario
                    If lngTipoBackOffice = enumTipoBackOffice.FundosProprios Or _
                        lngTipoBackOffice = enumTipoBackOffice.FundosTerceiros Then
                        lngCodigoProduto = 523
                    Else
                        lngCodigoProduto = 522
                    End If
                Else
                    lngTipoBackOffice = fgObterTipoBackOffice(xmlMensagem.selectSingleNode("//CO_VEIC_LEGA").Text)
                    If lngTipoBackOffice = enumTipoBackOffice.FundosProprios Or _
                        lngTipoBackOffice = enumTipoBackOffice.FundosTerceiros Then
                        lngCodigoProduto = 523
                    Else
                        lngCodigoProduto = 522
                    End If
                End If
            End If
        Else
            lngCodigoProduto = xmlMensagem.documentElement.selectSingleNode("//CO_PROD").Text
        End If

        fgObterIdentificadorRemessaPJ CLng(xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text), _
                                      fgDataHoraServidor(DataAux), _
                                      1, _
                                      strCodigoRemessaIni, _
                                      strCodigoRemessaFim
        .TipoRemessa = "200"
        .CodigoRemessa = fgCompletaString(strCodigoRemessaIni, " ", 23, False)
        .DataRemessa = Format(Date, "YYYYMMDD")
        .HoraRemessa = Format(Time, "HHMM")
        .CodigoEmpresa = fgCompletaString(xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text, "0", 5, True)
        '.SiglaSistema = "A8 "
        .SiglaSistema = fgCompletaString(xmlMensagem.documentElement.selectSingleNode("SG_SIST").Text, " ", 3, False)
        .CodigoMoeda = gstrCodigoMoeda
        .CodigoBanqueiro = Space(12)
        .TipoCaixa = fgCompletaString(enumTipoCaixaPJ.Futuro, "0", 3, True)
        .CodigoItemCaixa = fgCompletaString("", "0", 9, True)
        .CodigoProduto = fgCompletaString(lngCodigoProduto, "0", 4, True)
        .TipoConta = fgCompletaString("", "0", 3, True)
        .CodigoSegmento = fgCompletaString("", "0", 3, True)
        .CodigoEventoFinanceiro = fgCompletaString("", "0", 3, True)
        .CodigoIndexador = fgCompletaString("", "0", 3, True)
        .CodigoLocalLiquidacao = fgCompletaString(enumLocalLiquidacao.SELIC, "0", 4, True)
        .TipoMovimento = fgCompletaString(plngTipoMovimentoPJ, "0", 3, True)
        .DataMovimento = strDataOperacao
        .HoraMovimento = Format(Now, "HHmm")
        .TipoEntradaSaida = fgCompletaString(lngIndicadorDebitoCredito, "0", 1, False)

        If UBound(vntValor) = 0 Then
            .ValorMovimento = Format(vntValor(0), String(15, "0")) & "00"
        ElseIf UBound(vntValor) = -1 Then
            .ValorMovimento = Format("0", String(15, "0")) & "00"
        Else
            .ValorMovimento = Format(vntValor(0), String(15, "0")) & fgCompletaString(vntValor(1), "0", 2, False)
        End If
        .CodigoBanco = fgCompletaString("", "0", 3, False)
        .CodigoAgencia = fgCompletaString("", "0", 5, False)
        .NumeroContaCorrente = fgCompletaString("", "0", 13, False)
        .TipoPessoa = 2
        .CodigoCNPJ_CPF = fgCompletaString("", "0", 15, False)

        .NomeCliente = fgCompletaString(strNomeCliente, " ", 64, False)
        .TipoProcessamento = enumTipoProcessamentoPJ.OnLine
        .TipoEnvio = enumTipoEnvioPJ.Parcial
        .FILLER = Space(20)
    End With

    LSet udtMaioresValoresAux = udtMaioresValores

    If flVerificarEnvioMaiorValor(xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text, _
                                  lngCodigoProduto, _
                                  xmlMensagem.documentElement.selectSingleNode("//VlrFinanc").Text) Then
        flMaioresValoresEventos = udtMaioresValoresAux.String
    Else
        flMaioresValoresEventos = vbNullString
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMaioresValoresEventos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Montar remessa de previsão PJ

Private Function flMontarPrevisao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                  ByRef plngTipoMovimentoPJ As enumTipoMovimentoPJ, _
                         Optional ByVal pblnCompromissadaIda As Boolean = False, _
                         Optional ByVal strDataVencimento As String = vbNullString, _
                         Optional ByVal strValorComplementacao As String = vbNullString) As String

Dim udtRemessaMovimento                     As udtRemessaMovimento
Dim udtRemessaMovimentoAux                  As udtRemessaMovimentoAux
Dim strCodigoRemessaIni                     As String
Dim strCodigoRemessaFim                     As String
Dim strDataOperacao                         As String
Dim vntValor                                As Variant
Dim lngIndicadorDebitoCredito               As Long
Dim lngCodigoProduto                        As Long
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes

Dim blnZerarLocalLiquidacao                 As Boolean
Dim blnAlterarCodigoProduto                 As Boolean
Dim lngTipoOperacao                         As Long

    On Error GoTo ErrorHandler

    blnZerarLocalLiquidacao = False
    blnAlterarCodigoProduto = True
    
    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
        lngTipoOperacao = Val(xmlRemessa.selectSingleNode("//TP_OPER").Text)
        
        If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoCompromissadaCETIP Or _
           lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
            If Not xmlRemessa.selectSingleNode("//TP_CNPT") Is Nothing Then
                If Val(xmlRemessa.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Cliente1 Then
                    blnZerarLocalLiquidacao = True
                End If
            End If
        End If
        
        If lngTipoOperacao = enumTipoOperacaoLQS.OperacaoAntecipacaoCompromissadaCETIP Then
            blnAlterarCodigoProduto = False
        End If
    End If
    
    With udtRemessaMovimento

        If xmlRemessa.getElementsByTagName("DT_OPER_ATIV").Item(0) Is Nothing Then
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
        Else
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
        End If

        'Tem mensagem que não tem no lay-out a data da operação
        If pblnCompromissadaIda Then
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
                If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Then
                    Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
                    strDataOperacao = fgDt_To_Xml(objA6A7A8Funcoes.AdicionarDiasUteis(fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text), 1, enumPaginacao.Proximo))
                    Set objA6A7A8Funcoes = Nothing
                Else
                    strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
                End If
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                lngCodigoProduto = flObterProdutoPJPagtoRedesconto(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text)
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
                
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                lngCodigoProduto = flObterProdutoPJCompromissadaVolta(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text)
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DespesasBMC Then
                
                lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN") Is Nothing Then
                    If Val("0" & xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text) > 0 Then
                        vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV_RETN").Text, ",", , vbBinaryCompare)
                    Else
                        vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                    End If
                Else
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                End If
                lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                lngCodigoProduto = flObterProdutoPJCompromissadaVolta(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text)
            End If
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.OperacaoTermoCobertaBMA Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Or _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesRodaDolar Then
                
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Vinculo Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Desvinculo Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TransferenciaSemFinanceiro Then
                    
                lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TransferenciaCamaras Then
                
                If xmlRemessa.documentElement.selectSingleNode("PROD_CAMR") Is Nothing Then
                    'Utilizar o produto normal
                    lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                    lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                Else
                    If xmlRemessa.documentElement.selectSingleNode("PROD_CAMR").Text = enumIndicadorSimNao.Sim Then
                        lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida)
                        lngCodigoProduto = CLng("0" & xmlRemessa.documentElement.selectSingleNode("CO_PROD_CAMR").Text)
                    Else
                        'Utilizar o produto normal
                        lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Debito, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
                        lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                    End If
                End If
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
                   CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
                    
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesBMC Then
                    
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                    
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.DespesasBMC Then
            
                lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
            
            ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
            
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_MOED_NACI").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
                If Not xmlRemessa.documentElement.selectSingleNode("TP_OPER_CAMB") Is Nothing Then
                    If xmlRemessa.documentElement.selectSingleNode("TP_OPER_CAMB").Text = "V" Then
                        lngIndicadorDebitoCredito = enumTipoEntradaSaida.ENTRADA
                    Else ' "C"
                        lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                    End If
                Else
                    lngIndicadorDebitoCredito = enumTipoEntradaSaida.Saida
                End If
                
            Else
                
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
                lngCodigoProduto = xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text
            
            End If
            
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
            Or CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_MOED_NACIO").Text, ",", , vbBinaryCompare)
            Else
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
            End If
        
            'RATS 1033 | 06/dez/10 | Bruno Oliveira | para layout 94 (RegistroContratoSWAPCetip21) deve ser enviado o valor do prêmio
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                If Not xmlRemessa.documentElement.selectSingleNode("//VA_PREM") Is Nothing Then
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_PREM").Text, ",", , vbBinaryCompare)
                Else
                    vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
                End If
            End If
        
            If strValorComplementacao <> vbNullString Then
                vntValor = Split(strValorComplementacao, ",", , vbBinaryCompare)
            End If
            
            
        End If

        'Se for Cliente1 foi tratado anteriormente
        If Not xmlRemessa.documentElement.selectSingleNode("CLIENTE1") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("CLIENTE1").Text = enumIndicadorSimNao.Sim Then
                lngIndicadorDebitoCredito = xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text
            End If
        End If

        fgObterIdentificadorRemessaPJ CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), _
                                      fgDataHoraServidor(DataAux), _
                                      1, _
                                      strCodigoRemessaIni, _
                                      strCodigoRemessaFim

        .TipoRemessa = "100"
        .CodigoRemessa = fgCompletaString(strCodigoRemessaIni, " ", 23, False)
        .DataRemessa = Format(Date, "YYYYMMDD")
        .HoraRemessa = Format(Time, "HHMM")
        .CodigoEmpresa = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, "0", 5, True)
        '.SiglaSistema = "A8 "
        .SiglaSistema = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False)
        .CodigoMoeda = gstrCodigoMoeda
        .CodigoBanqueiro = fgCompletaString("", " ", 12, False)
        .TipoCaixa = fgCompletaString(enumTipoCaixaPJ.Futuro, "0", 3, True)
        .CodigoItemCaixa = fgCompletaString("", "0", 9, True)
        .TipoAtivoPassivo = 0
        
        .CodigoProduto = fgCompletaString(lngCodigoProduto, "0", 4, True)
        
        If strDataVencimento <> vbNullString And blnAlterarCodigoProduto Then
            .CodigoProduto = fgCompletaString(fgAlterarCodigoProduto(lngCodigoProduto), "0", 4, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("TP_CONT") Is Nothing Then
            .TipoConta = fgCompletaString("", "0", 3, True)
        Else
            .TipoConta = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("TP_CONT").Text, "0", 3, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("CO_SEGM") Is Nothing Then
            .CodigoSegmento = fgCompletaString("", "0", 3, True)
        Else
            .CodigoSegmento = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_SEGM").Text, "0", 3, True)
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC") Is Nothing Then
            .EventoFinanceiro = fgCompletaString("", "0", 3, True)
        Else
            .EventoFinanceiro = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_EVEN_FINC").Text, "0", 3, True)
        End If

        If xmlRemessa.documentElement.selectSingleNode("CO_INDX") Is Nothing Then
            .CodigoIndexador = fgCompletaString("", "0", 3, True)
        Else
            .CodigoIndexador = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_INDX").Text, "0", 3, True)
        End If

        .CodigoLocalLiquidacao = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, "0", 4, True)
        
        If Val(.CodigoLocalLiquidacao) = enumLocalLiquidacao.PAG Then
            .CodigoLocalLiquidacao = fgCompletaString(enumLocalLiquidacao.CIP, "0", 4, True)
        ElseIf Val(.CodigoLocalLiquidacao) = enumLocalLiquidacao.CONTA_CORRENTE Then
            blnZerarLocalLiquidacao = True
        End If
        
        If blnZerarLocalLiquidacao Then
            .CodigoLocalLiquidacao = "0000"
        End If
        
        .CodigoFaixaValor = fgCompletaString("", "0", 3, True)
        .TipoMovimento = fgCompletaString(plngTipoMovimentoPJ, "0", 3, True)
        .DataMovimento = IIf(strDataVencimento = vbNullString, strDataOperacao, strDataVencimento)
        .HoraMovimento = Format(Now, "HHmm")
        
        If strDataVencimento <> vbNullString Then
            lngIndicadorDebitoCredito = IIf(lngIndicadorDebitoCredito = enumTipoDebitoCredito.Debito, _
                                                enumTipoDebitoCredito.Credito, _
                                                enumTipoDebitoCredito.Debito)
        End If

        .TipoEntradaSaida = fgCompletaString(lngIndicadorDebitoCredito, "0", 1, False)
        
        If UBound(vntValor) = 0 Then
            .ValorMovimento = fgCompletaString(vntValor(0), "0", 17, True) & "00"
            .ValorContabil = fgCompletaString(vntValor(0), "0", 17, True) & "00"
        ElseIf UBound(vntValor) = -1 Then
            .ValorMovimento = fgCompletaString("0", "0", 17, True) & "00"
            .ValorContabil = fgCompletaString("0", "0", 17, True) & "00"
        Else
            .ValorMovimento = fgCompletaString(vntValor(0), "0", 17, True) & fgCompletaString(vntValor(1), "0", 2, False)
            .ValorContabil = fgCompletaString(vntValor(0), "0", 17, True) & fgCompletaString(vntValor(1), "0", 2, False)
        End If

        .TipoProcessamento = enumTipoProcessamentoPJ.OnLine
        .TipoEnvio = enumTipoEnvioPJ.Parcial
        .FILLER = fgCompletaString("", " ", 46, False)

    End With

    LSet udtRemessaMovimentoAux = udtRemessaMovimento

    flMontarPrevisao = udtRemessaMovimentoAux.String

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarPrevisao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o Produto PJ para compromissada de volta

Private Function flObterProdutoPJCompromissadaVolta(ByVal plngProdutoPJ As Long) As Long

    On Error GoTo ErrorHandler

    Select Case plngProdutoPJ
        'Eventos Mercado
        Case 79
            flObterProdutoPJCompromissadaVolta = 228

        Case 230
            flObterProdutoPJCompromissadaVolta = 78

        Case 280
            flObterProdutoPJCompromissadaVolta = 284

        Case 285
            flObterProdutoPJCompromissadaVolta = 279

        'Interno
        Case 502
            flObterProdutoPJCompromissadaVolta = 499

        Case 501
            flObterProdutoPJCompromissadaVolta = 500

        Case 423
            flObterProdutoPJCompromissadaVolta = 424

        Case 421
            flObterProdutoPJCompromissadaVolta = 422

        Case Else
            flObterProdutoPJCompromissadaVolta = plngProdutoPJ

    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterProdutoPJCompromissadaVolta Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o Produto PJ para pagamento de Redesconto

Private Function flObterProdutoPJPagtoRedesconto(ByVal plngProdutoPJ As Long) As Long

    On Error GoTo ErrorHandler

    Select Case plngProdutoPJ

        Case 67
            flObterProdutoPJPagtoRedesconto = 316

        Case 68
            flObterProdutoPJPagtoRedesconto = 317

        Case 318
            flObterProdutoPJPagtoRedesconto = 320

        Case 319
            flObterProdutoPJPagtoRedesconto = 321

        Case Else
            flObterProdutoPJPagtoRedesconto = plngProdutoPJ

    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterProdutoPJPagtoRedesconto Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Selecionar na tabela de Produto o valor minimo para o envio de registro para PJ;
Private Function flVerificarEnvioMaiorValor(ByVal plngCodigoEmpresa As Long, _
                                            ByVal plngProdutoPJ As Long, _
                                            ByVal pvntValorOperacao As Variant) As Boolean

Dim strSQL                                  As String
Dim rsProduto                               As ADODB.Recordset

    On Error GoTo ErrorHandler

    Set rsProduto = CreateObject("ADODB.Recordset")

    strSQL = " SELECT   A.VA_MINI_MAIR_VALO     " & vbCrLf & _
             "   FROM   A8.TB_PRODUTO       A,  " & vbCrLf & _
             "          A8.TB_EMPRESA_HO    B   " & vbCrLf & _
             "  WHERE   A.CO_EMPR_FUSI      = B.CO_EMPR_FUSI " & vbCrLf & _
             "    AND   B.CO_EMPR           = " & plngCodigoEmpresa & vbCrLf & _
             "    AND   A.CO_PROD           =   " & plngProdutoPJ

    Set rsProduto = fgQuerySQL(strSQL)

    If rsProduto.EOF Then
        flVerificarEnvioMaiorValor = True
    Else
    
        If IsNull(rsProduto!VA_MINI_MAIR_VALO) Then
            flVerificarEnvioMaiorValor = True
        Else
            If Val(rsProduto!VA_MINI_MAIR_VALO) > Val(pvntValorOperacao) Then
                flVerificarEnvioMaiorValor = False
            Else
                flVerificarEnvioMaiorValor = True
            End If
        End If
    
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarEnvioMaiorValor Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter a descrição do tipo de movimento PJ
Private Function flObterDescricaoTipoMovimento(ByVal plngTipoMovimento As Long) As String

    If Trim(plngTipoMovimento) = enumTipoMovimentoPJ.Previsto Then
        flObterDescricaoTipoMovimento = "Previsto"
    ElseIf Trim(plngTipoMovimento) = enumTipoMovimentoPJ.Realizado Then
        flObterDescricaoTipoMovimento = "Realizado"
    ElseIf Trim(plngTipoMovimento) = enumTipoMovimentoPJ.EstornoPrevisto Then
        flObterDescricaoTipoMovimento = "Estorno Previsto"
    ElseIf Trim(plngTipoMovimento) = enumTipoMovimentoPJ.EstornoRealizado Then
        flObterDescricaoTipoMovimento = "Estorno Realizado"
    End If

End Function

Private Function flMontarMoviMoedaEstrangeria(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                              ByVal plngTipoMovimentoPJ As enumTipoMovimentoPJ) As String

Dim udtMovi_PJ_MoedaEstrangeira             As udtMovi_PJ_MoedaEstrangeira
Dim udtMovi_PJ_MoedaEstrangeiraAux          As udtMovi_PJ_MoedaEstrangeiraAux

Dim strCodigoRemessaIni                     As String
Dim strCodigoRemessaFim                     As String
Dim strCodigoBanqueiroSwift                As String
Dim strNomeCliente                          As String
Dim strContaBanqueiro                       As String
Dim strCodigoMoedaISO                       As String

Dim strDataOperacao                         As String
Dim vntValor                                As Variant
Dim lngIndicadorDebitoCredito               As Long
Dim lngCodigoProduto                        As Long
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If Not xmlRemessa.selectSingleNode("//DT_LIQU_OPER_ATIV_MOED_ESTR") Is Nothing Then
        strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV_MOED_ESTR").Text
    ElseIf Not xmlRemessa.selectSingleNode("//DT_LIQU_OPER") Is Nothing Then
        strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER").Text
    Else
        strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
    End If
    
    If Not xmlRemessa.selectSingleNode("//VA_MOED_ESTR_BMC") Is Nothing Then
        vntValor = Split(xmlRemessa.documentElement.selectSingleNode("VA_MOED_ESTR_BMC").Text, ",", , vbBinaryCompare)
    Else
        If Not xmlRemessa.selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            vntValor = Split(xmlRemessa.documentElement.selectSingleNode("VA_OPER_ATIV").Text, ",", , vbBinaryCompare)
        Else
            vntValor = Split("0,0", ",", , vbBinaryCompare)
        End If
    End If
    
    If Not xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing Then
        lngIndicadorDebitoCredito = IIf(CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumIndicadorEntradaSaida.ENTRADA, enumIndicadorEntradaSaida.Saida, enumIndicadorEntradaSaida.ENTRADA)
    Else
        lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
    End If
    
    If Not xmlRemessa.selectSingleNode("//CO_PROD") Is Nothing Then
        lngCodigoProduto = Val("0" & xmlRemessa.documentElement.selectSingleNode("//CO_PROD").Text)
    End If

    fgObterIdentificadorRemessaPJ CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), _
                                  fgDataHoraServidor(DataAux), _
                                  1, _
                                  strCodigoRemessaIni, _
                                  strCodigoRemessaFim
    
    If Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroOperacaoInterbancaria _
    Or Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroOperacaoArbitragem _
    Or Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.IFInformaLiquidacaoInterbancaria Then
        
        lngCodigoProduto = Val("0" & xmlRemessa.documentElement.selectSingleNode("//CO_PROD_MOED_ESTR").Text)
        
        If Not xmlRemessa.documentElement.selectSingleNode("//CO_BANQ_SWIFT") Is Nothing Then
            strCodigoBanqueiroSwift = xmlRemessa.documentElement.selectSingleNode("//CO_BANQ_SWIFT").Text
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//NM_CLIE_MOED_ESTR") Is Nothing Then
            strNomeCliente = xmlRemessa.documentElement.selectSingleNode("//NM_CLIE_MOED_ESTR").Text
        End If
        
        If Not xmlRemessa.documentElement.selectSingleNode("//NR_CNTA_BANQ") Is Nothing Then
            strContaBanqueiro = xmlRemessa.documentElement.selectSingleNode("//NR_CNTA_BANQ").Text
        End If
        
        'para operações de arbitragem precisar obter as informações do Grupo de Contratação da moeda USD
        If Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroOperacaoArbitragem Then
        
            For Each objDomNode In xmlRemessa.selectNodes("//GR_CONTR")
            
                If Trim(objDomNode.selectSingleNode("CO_MOED_ISO").Text) = "USD" Then
            
                    If Not objDomNode.selectSingleNode("VA_MOED_ESTRG") Is Nothing Then
                        vntValor = Split(objDomNode.selectSingleNode("VA_MOED_ESTRG").Text, ",", , vbBinaryCompare)
                    Else
                        vntValor = Split("0,0", ",", , vbBinaryCompare)
                    End If
                    
                    If Not objDomNode.selectSingleNode("TP_OPER_CAMB") Is Nothing Then
                        If objDomNode.selectSingleNode("TP_OPER_CAMB").Text = "C" Then
                            lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.ENTRADA
                        Else ' = "V"
                            lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
                        End If
                    Else
                        lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
                    End If
                    
                    If Not objDomNode.selectSingleNode("CO_MOED_ISO") Is Nothing Then
                        strCodigoMoedaISO = objDomNode.selectSingleNode("CO_MOED_ISO").Text
                    End If
                    
                End If
            
            Next

        Else
        
            If Not xmlRemessa.selectSingleNode("//VA_MOED_ESTRG") Is Nothing Then
                vntValor = Split(xmlRemessa.documentElement.selectSingleNode("//VA_MOED_ESTRG").Text, ",", , vbBinaryCompare)
            Else
                vntValor = Split("0,0", ",", , vbBinaryCompare)
            End If
            
            If Not xmlRemessa.selectSingleNode("//TP_OPER_CAMB") Is Nothing Then
                If xmlRemessa.selectSingleNode("//TP_OPER_CAMB").Text = "C" Then
                    lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.ENTRADA
                Else ' = "V"
                    lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
                End If
            Else
                lngIndicadorDebitoCredito = enumIndicadorEntradaSaida.Saida
            End If
            
            If Not xmlRemessa.documentElement.selectSingleNode("//CO_MOED_ISO") Is Nothing Then
                strCodigoMoedaISO = xmlRemessa.documentElement.selectSingleNode("//CO_MOED_ISO").Text
            End If
        
        End If
        
    End If
    
    With udtMovi_PJ_MoedaEstrangeira
        .TipoRemessa = "250"
        .CodigoEmpresa = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, "0", 5, True)
        .SiglaSistema = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False)
        .IdentificadorMovimento = fgCompletaString(strCodigoRemessaIni, " ", 25, False)
        
        If strCodigoMoedaISO <> vbNullString Then
            .CodigoMoeda = fgCompletaString(strCodigoMoedaISO, "0", 3, True)
        Else
            If Not xmlRemessa.documentElement.selectSingleNode("CO_MOED_ESTR") Is Nothing Then
                .CodigoMoeda = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_MOED_ESTR").Text, "0", 3, True)
            Else
                .CodigoMoeda = fgCompletaString("", "0", 3, True)
            End If
        End If
        
        If strCodigoBanqueiroSwift <> vbNullString Then
            .CodigoBanqueiroSwift = fgCompletaString(strCodigoBanqueiroSwift, " ", 30, False)
        Else
            .CodigoBanqueiroSwift = fgCompletaString(fgObterCodigoBanqueiroSwiftPJME, " ", 30, False)
        End If
        
        .CodigoProduto = fgCompletaString(lngCodigoProduto, "0", 4, True)
        .DataMovimento = strDataOperacao
        
        If Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemLQS.RegistroOperacoesBMC Then
            If Not xmlRemessa.documentElement.selectSingleNode("TP_NEGO") Is Nothing Then
                If xmlRemessa.documentElement.selectSingleNode("TP_NEGO").Text = 1 _
                Or xmlRemessa.documentElement.selectSingleNode("TP_NEGO").Text = 2 Then
                    .CodigoReferenciaSwift = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("NR_IDEF_NEGO_BMC").Text, " ", 16, False)
                ElseIf xmlRemessa.documentElement.selectSingleNode("TP_NEGO").Text = 3 Then
                    .CodigoReferenciaSwift = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_CNTR_SISB").Text, " ", 16, False)
                Else
                    .CodigoReferenciaSwift = fgCompletaString("", " ", 16, False)
                End If
            Else
                .CodigoReferenciaSwift = fgCompletaString("", " ", 16, False)
            End If
        Else
            .CodigoReferenciaSwift = fgCompletaString("", " ", 16, False)
        End If
        
        .TipoEntradaSaida = fgCompletaString(lngIndicadorDebitoCredito, "0", 1, False)
        
        If UBound(vntValor) = 0 Then
            .ValorMovimento = fgCompletaString(vntValor(0), "0", 17, True) & "00"
        Else
            .ValorMovimento = fgCompletaString(vntValor(0), "0", 17, True) & fgCompletaString(vntValor(1), "0", 2, False)
        End If
        
        If strNomeCliente <> vbNullString Then
            .NomeCliente = fgCompletaString(strNomeCliente, " ", 50, False)
        Else
            .NomeCliente = fgCompletaString(fgObterNomeClientePJME, " ", 50, False)
        End If
        
        .TipoMovimento = fgCompletaString(plngTipoMovimentoPJ, "0", 3, True)
        .TipoProcessamento = enumTipoProcessamentoPJ.OnLine
        
        If strContaBanqueiro <> vbNullString Then
            .ContaBanqueiro = fgCompletaString(strContaBanqueiro, " ", 35, False)
        Else
            .ContaBanqueiro = fgCompletaString(fgObterContaBanqueiroPJME(Val(xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)), "0", 35, False)
        End If
        
        .FILLER = fgCompletaString("", " ", 93, False)
    End With

    LSet udtMovi_PJ_MoedaEstrangeiraAux = udtMovi_PJ_MoedaEstrangeira

    flMontarMoviMoedaEstrangeria = udtMovi_PJ_MoedaEstrangeiraAux.String

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMoviMoedaEstrangeria Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Enviar Estorno Realizado PJ - Moeda Estrangeira
'   - Montar protocolo de integração A7
'   - Montar Remessa estorno Realizado PJ (String)
'   - Incluir registro no histórico de envio PJ

Public Function EstornarRealizadoME(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strRealizado                            As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1014"
    LSet udtProtocoloAux = udtProtocolo

    strRealizado = flMontarMoviMoedaEstrangeria(xmlRemessa, enumTipoMovimentoPJ.EstornoRealizado)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.EstornoRealizadoConfirmado, _
            udtProtocoloAux.String & strRealizado

    EstornarRealizadoME = udtProtocoloAux.String & strRealizado

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarRealizadoME Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function EnviarRealizadoME(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strRealizado                            As String

    On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PJ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1014"
    LSet udtProtocoloAux = udtProtocolo

    strRealizado = flMontarMoviMoedaEstrangeria(xmlRemessa, enumTipoMovimentoPJ.Realizado)

    'Grava Historico
    Incluir xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
            enumTipoMovimento.RealizadoConfirmado, _
            udtProtocoloAux.String & strRealizado

    EnviarRealizadoME = udtProtocoloAux.String & strRealizado

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealizadoME Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

