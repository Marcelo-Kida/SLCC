VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF769B00382"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'' Objeto responsável pelo envio e estorno de informações para o módulo de Sub-
'' Reserva. As informações são classificadas em:
''      Previsão
''      Realizado Solicitado
''      Realizado Confirmado
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

'Montar remessa de previsão para A6 - Sub Reserva

Public Function EnviarPrevisao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                      Optional ByVal pblnCompromissadaIda As Boolean = False, _
                      Optional ByVal strDataVencimento As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim strDataOperacao                         As String
Dim strDataOperacaoAux                      As String
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        If pblnCompromissadaIda Then
            fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.PrevistoCompromIda
        Else
            fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.Previsto
        End If
    Else
        If pblnCompromissadaIda Then
            xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.PrevistoCompromIda
        Else
            xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.Previsto
        End If
    End If
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMC Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(Val(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If pblnCompromissadaIda Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
            strDataOperacaoAux = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Then
                Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
                strDataOperacao = fgDt_To_Xml(objA6A7A8Funcoes.AdicionarDiasUteis(fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text), 1, enumPaginacao.Proximo))
                Set objA6A7A8Funcoes = Nothing
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        
        ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            
            If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
            
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA And _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA And _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA And _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaTermo And _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaGenericaAVista And _
               CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaGenericaATermo Then

                'RATS 491 / 003 - PM 1951271 - Adrian
                If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoInformalCompromissada Then
                    strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
                    xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text
                Else
                    strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
                    xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text
                End If
                
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        End If
    Else
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
        End If
    End If

    If Val(strTipoMensagem) = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
       Val(strTipoMensagem) = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
        
        If strDataVencimento <> vbNullString Then
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataVencimento
        
            If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        
            If Not xmlRemessa.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
                xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text = fgAlterarCodigoProduto(CLng(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text))
            End If
        End If
        
    End If
    
    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    flAlterarDataMovimento xmlRemessa

    EnviarPrevisao = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    If pblnCompromissadaIda Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacaoAux
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        Else
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaEspecificaTermo And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaGenericaAVista And _
                CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) <> enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
                
                'RATS 491 / 003 - PM 1951271 - Adrian
                If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text = enumTipoOperacaoLQS.LeilaoInformalCompromissada Then
                    xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text = strDataOperacao
                Else
                    xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
                End If
                
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        End If
    End If

    If Val(strTipoMensagem) = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
       Val(strTipoMensagem) = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
        
        If strDataVencimento <> vbNullString Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            
            If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        
            If Not xmlRemessa.documentElement.selectSingleNode("CO_PROD") Is Nothing Then
                xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text = fgAlterarCodigoProduto(CLng(xmlRemessa.documentElement.selectSingleNode("CO_PROD").Text))
            End If
        End If
    
    End If
    
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarPrevisao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar remessa de realizado Solicitado para A6 - Sub Reserva
Public Function EnviarRealizadoSolicitado(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.RealizadoSolicitado
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.RealizadoSolicitado
    End If
    'Fixo por ser SELIC
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(Val(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    flAlterarDataMovimento xmlRemessa

    EnviarRealizadoSolicitado = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealizadoSolicitado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar remessa de estorno de realizado Solictado para A6 - Sub Reserva
Public Function EstornarRealizadoSolicitado(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.EstornoRealizadoSolicitado
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoRealizadoSolicitado
    End If
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    flAlterarDataMovimento xmlRemessa

    EstornarRealizadoSolicitado = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarRealizadoSolicitado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar remessa de estorno de previsão para A6 - Sub Reserva
Public Function EstornarPrevisao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                        Optional ByVal pblnCompromissadaIda As Boolean = False, _
                        Optional ByVal strDataVencimento As String = vbNullString, _
                        Optional ByVal strValorComplementacao As String = vbNullString) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim strDataOperacao                         As String
Dim strDataOperacaoAux                      As String
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    If strValorComplementacao <> vbNullString Then
        xmlRemessa.selectSingleNode("//VA_OPER_ATIV").Text = strValorComplementacao
    End If
    
    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        If pblnCompromissadaIda Then
            fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.EstornoPrevistoCompromIda
        Else
            fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.EstornoPrevisto
        End If
    Else
        If pblnCompromissadaIda Then
            xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevistoCompromIda
        Else
            xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoPrevisto
        End If
    End If
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If pblnCompromissadaIda Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
            strDataOperacaoAux = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Then
                Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
                strDataOperacao = fgDt_To_Xml(objA6A7A8Funcoes.AdicionarDiasUteis(fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text), 1, enumPaginacao.Proximo))
                Set objA6A7A8Funcoes = Nothing
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        Else
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        End If
    Else
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Termo Then
            If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
            Else
                strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            End If
        End If
    End If

    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    If Val(strTipoMensagem) = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
       Val(strTipoMensagem) = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
        
        If strDataVencimento <> vbNullString Then
            strDataOperacao = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataVencimento
        
            If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        End If

    End If
    
    flAlterarDataMovimento xmlRemessa

    EstornarPrevisao = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    If pblnCompromissadaIda Then
        If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradia Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoPrazoUmDiaUtil Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssociadoAquisicao Or _
            CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.RedescontoIntradiaAssocLiberacaoGarantia Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacaoAux
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        Else
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
            Else
                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
            End If
        End If
    End If

    If Val(strTipoMensagem) = enumTipoMensagemBUS.OperacaoDefinitivaCETIP Or _
       Val(strTipoMensagem) = enumTipoMensagemBUS.MovimentacaoInstrumentoFinanceiro Then
        
        If strDataVencimento <> vbNullString Then
            xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strDataOperacao
            
            If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                If CLng(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito
                Else
                    xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito
                End If
            End If
        End If
    
    End If
    
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarPrevisao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Montar remessa de estorno de realizado confirmado para A6 - Sub Reserva
Public Function EstornarRealizadoConfirmado(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.EstornoRealizadoConfirmado
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.EstornoRealizadoConfirmado
    End If
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    flAlterarDataMovimento xmlRemessa

    EstornarRealizadoConfirmado = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EstornarRealizadoConfirmado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montar remessa de realizado confirmado para A6 - Sub Reserva
Public Function EnviarRealizadoConfirmado(ByRef xmlRemessa As MSXML2.DOMDocument40) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strTipoMensagem                         As String

    On Error GoTo ErrorHandler

    'Mensagem da Sub-Reserva
    strTipoMensagem = xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text
    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = 103
    'Criação de TAGs para Sub-Reserva
    If xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.RealizadoConfirmado
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = enumTipoMovimento.RealizadoConfirmado
    End If
    If CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMA Then
        'Fixo por ser BMA
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "Multilateral"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "Multilateral"
        End If
    ElseIf CLng(xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = flTipoLiquidacao(CLng(xmlRemessa.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text))
        End If
    Else
        'Fixo por ser SELIC
        If xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "DE_TIPO_LIQU", "LBTR"
        Else
            xmlRemessa.documentElement.selectSingleNode("DE_TIPO_LIQU").Text = "LBTR"
        End If
    End If

    If xmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED") Is Nothing And _
       xmlRemessa.selectSingleNode("//IN_MOVI_ENTR_SAID") Is Nothing Then
        
        Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
            Case enumTipoOperacaoLQS.DespesasBMC
                fgAppendNode xmlRemessa, "MESG", "IN_OPER_DEBT_CRED", enumTipoEntradaSaida.Saida
        End Select
        
    End If
    
    flAlterarDataMovimento xmlRemessa

    EnviarRealizadoConfirmado = xmlRemessa.xml

    flAlterarDataMovimento xmlRemessa

    xmlRemessa.documentElement.selectSingleNode("TP_MESG").Text = strTipoMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarRealizadoConfirmado Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar a Data de Movimento da remessa
Private Function flAlterarDataMovimento(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim strData                                 As String

    On Error GoTo ErrorHandler

'    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.TermoDescobertaBMA Then
'        Exit Function
'    End If

    If CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaEspecificaTermo Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaGenericaAVista Or _
       CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaGenericaATermo Then
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV") Is Nothing Then
        Exit Function
    End If

    If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        Exit Function
    End If
    
    'Para Operações a termo a data da previsão sera DT_LIQU_OPER_ATIV
    strData = xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text
    xmlRemessa.documentElement.selectSingleNode("DT_LIQU_OPER_ATIV").Text = xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text

    xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text = strData

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterarDataMovimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a descrição do tipo de liquidação
Private Function flTipoLiquidacao(ByRef plngTipoLiquidacao As Long) As String

Dim objTipoLiquidacao                       As A6A7A8.clsTipoLiquidacao
Dim xmlTipoLiquidacao                       As MSXML2.DOMDocument40
Dim strTipoLiquidacao                       As String

    On Error GoTo ErrorHandler

    Set objTipoLiquidacao = CreateObject("A6A7A8.clsTipoLiquidacao")
    strTipoLiquidacao = objTipoLiquidacao.Ler(plngTipoLiquidacao)
    Set objTipoLiquidacao = Nothing

    If strTipoLiquidacao = vbNullString Then
        Exit Function
    End If

    Set xmlTipoLiquidacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoLiquidacao.loadXML strTipoLiquidacao

    flTipoLiquidacao = xmlTipoLiquidacao.documentElement.selectSingleNode("//NO_TIPO_LIQU_OPER_ATIV").Text

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flTipoLiquidacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

