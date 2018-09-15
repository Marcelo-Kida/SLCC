VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConciliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável por promover todas as verificações para a conciliação de
'' operações e/ou  mensagens

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const MESG_DESPESA                  As String = "SEL1069"

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

'' Consultar todos os atributos dos registros conciliados (operações e/ou
'' mensagens) para a exibição de seu detalhe em listas e GRID s

Public Function ConsultaConciliacao(Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                                    Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                                    Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString, _
                                    Optional ByVal pstrDataJustificativa As String = vbNullString, _
                                    Optional ByVal plngTipoOperacao As Long = 0, _
                                    Optional ByVal plngEmpresa As Long = 0, _
                                    Optional ByVal plngCodigoSituProc As Long = 0, _
                                    Optional ByVal pvntSequenciaConciliacao As Variant = 0) As String

Dim strSQL                                  As String
Dim datDataCorte                            As Date

On Error GoTo ErrorHandler

    strSQL = " SELECT   A.NU_SEQU_CNCL_OPER_ATIV_MESG,                                              " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                                        " & vbNewLine & _
             "          A.NU_CTRL_IF,                                                               " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                                                         " & vbNewLine & _
             "          A.NU_SEQU_CNTR_REPE,                                                        " & vbNewLine & _
             "          B.TP_JUST_CNCL,                                                             " & vbNewLine & _
             "          C.NO_TIPO_JUST_CNCL,                                                        " & vbNewLine & _
             "          B.TX_JUST,                                                                  " & vbNewLine & _
             "          B.CO_USUA_ATLZ,                                                             " & vbNewLine & _
             "          B.CO_ETCA_TRAB_ATLZ,                                                        " & vbNewLine & _
             "          B.DH_JUST_CNCL,                                                             " & vbNewLine & _
             "          E.NU_COMD_OPER,                                                             " & vbNewLine & _
             "          NVL(E.CO_MESG_SPB, 'Inexistente')   CO_MESG_SPB,                            " & vbNewLine & _
             "          E.CO_TEXT_XML,                                                              " & vbNewLine

    strSQL = strSQL & _
             "          D.NU_ATIV_MERC as NU_ATIV_MERC_MESG,                                        " & vbNewLine & _
             "          D.PU_ATIV_MERC AS PU_ATIV_MERC_MESG,                                        " & vbNewLine & _
             "          D.QT_ATIV_MERC AS QT_ATIV_MERC_MESG,                                        " & vbNewLine & _
             "          D.IN_OPER_DEBT_CRED,                                                        " & vbNewLine & _
             "          D.NU_COMD_OPER_ORIG,                                                        " & vbNewLine & _
             "          F.NU_ATIV_MERC AS NU_ATIV_MERC_OPER,                                        " & vbNewLine & _
             "          F.PU_ATIV_MERC AS PU_ATIV_MERC_OPER,                                        " & vbNewLine & _
             "          F.QT_ATIV_MERC AS QT_ATIV_MERC_OPER,                                        " & vbNewLine & _
             "          F.VA_OPER_ATIV,                                                             " & vbNewLine & _
             "          F.NU_COMD_OPER as NU_COMD_OPER_OPER,                                        " & vbNewLine & _
             "          F.TP_OPER,                                                                  " & vbNewLine & _
             "          E.IN_CONF_MESG_LTR                                                          " & vbNewLine
             
    strSQL = strSQL & _
             " FROM     A8.TB_CNCL_OPER_ATIV                A,                                      " & vbNewLine & _
             "          A8.TB_JUST_CNCL_OPER_ATIV_MESG      B,                                      " & vbNewLine & _
             "          A8.TB_TIPO_JUST_CNCL                C,                                      " & vbNewLine & _
             "          A8.TB_MESG_RECB_SPB_CNCL            D,                                      " & vbNewLine & _
             "          A8.TB_MESG_RECB_ENVI_SPB            E,                                      " & vbNewLine & _
             "          A8.TB_OPER_ATIV                     F                                       " & vbNewLine
    
    strSQL = strSQL & _
             " WHERE    A.NU_SEQU_CNCL_OPER_ATIV_MESG       = B.NU_SEQU_CNCL_OPER_ATIV_MESG         " & vbNewLine & _
             " AND      B.TP_JUST_CNCL                      = C.TP_JUST_CNCL                    (+) " & vbNewLine & _
             " AND      A.NU_CTRL_IF                        = D.NU_CTRL_IF                      (+) " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB                  = D.DH_REGT_MESG_SPB                (+) " & vbNewLine & _
             " AND      A.NU_SEQU_CNTR_REPE                 = D.NU_SEQU_CNTR_REPE               (+) " & vbNewLine & _
             " AND      D.NU_CTRL_IF                        = E.NU_CTRL_IF                      (+) " & vbNewLine & _
             " AND      D.DH_REGT_MESG_SPB                  = E.DH_REGT_MESG_SPB                (+) " & vbNewLine & _
             " AND      D.NU_SEQU_CNTR_REPE                 = E.NU_SEQU_CNTR_REPE               (+) " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV                 = F.NU_SEQU_OPER_ATIV               (+) " & vbNewLine
    
    If pvntSequenciaOperacao <> Empty Then
        
        If pvntSequenciaOperacao < 0 Then
            strSQL = Replace$(strSQL, "A8.TB_CNCL_OPER_ATIV", "A8HIST.TB_CNCL_OPER_ATIV")
            strSQL = Replace$(strSQL, "A8.TB_JUST_CNCL_OPER_ATIV_MESG", "A8HIST.TB_JUST_CNCL_OPER_ATIV_MESG")
            strSQL = Replace$(strSQL, "A8.TB_MESG_RECB_SPB_CNCL", "A8HIST.TB_MESG_RECB_SPB_CNCL")
            strSQL = Replace$(strSQL, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
            strSQL = Replace$(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
            pvntSequenciaOperacao = pvntSequenciaOperacao * -1
        End If
    
        If IsNumeric(pvntSequenciaOperacao) Then
            strSQL = strSQL & vbNewLine & _
                " AND       A.NU_SEQU_OPER_ATIV                 = " & pvntSequenciaOperacao & vbCrLf
        Else
            strSQL = strSQL & vbNewLine & _
                " AND       A.NU_SEQU_OPER_ATIV                   " & pvntSequenciaOperacao & vbCrLf
        End If
    
    End If
    
    If pvntSequenciaConciliacao <> 0 Then
        
        strSQL = strSQL & vbNewLine & _
            " AND       A.NU_SEQU_CNCL_OPER_ATIV_MESG                 = " & pvntSequenciaConciliacao & vbCrLf
    
    End If
    
    If Trim(pstrNumeroControleIF) <> "" Then
        strSQL = strSQL & vbNewLine & _
            " AND       A.NU_CTRL_IF                        = '" & pstrNumeroControleIF & "'" & vbCrLf
    End If
    
    If Trim(pstrDataRegistroMensagemSPB) <> "" Then
        
        datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
        
        If fgDtHrStr_To_DateTime(pstrDataRegistroMensagemSPB) < datDataCorte Then
            strSQL = Replace$(strSQL, "A8.TB_CNCL_OPER_ATIV", "A8HIST.TB_CNCL_OPER_ATIV")
            strSQL = Replace$(strSQL, "A8.TB_JUST_CNCL_OPER_ATIV_MESG", "A8HIST.TB_JUST_CNCL_OPER_ATIV_MESG")
            strSQL = Replace$(strSQL, "A8.TB_MESG_RECB_SPB_CNCL", "A8HIST.TB_MESG_RECB_SPB_CNCL")
            strSQL = Replace$(strSQL, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
            strSQL = Replace$(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
        End If
    
        strSQL = strSQL & vbNewLine & _
            " AND       A.DH_REGT_MESG_SPB                  = " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB) & vbCrLf
    End If
    
    If Trim(pstrDataJustificativa) <> vbNullString Then
        strSQL = strSQL & _
            " AND       B.DH_JUST_CNCL >= " & fgDtXML_To_Oracle(pstrDataJustificativa) & vbCrLf
    End If
    
    If plngTipoOperacao <> 0 Then
        strSQL = strSQL & _
            " AND       F.TP_OPER =  " & plngTipoOperacao & vbCrLf
    End If
    
    If plngEmpresa <> 0 Then
        strSQL = strSQL & _
            " AND       F.CO_EMPR = " & plngEmpresa & vbCrLf
    End If
    
    If plngCodigoSituProc <> 0 Then
        strSQL = strSQL & _
            " AND       F.CO_ULTI_SITU_PROC = " & plngCodigoSituProc & vbCrLf
    End If
    
    strSQL = strSQL & _
             "ORDER BY A.DH_REGT_MESG_SPB DESC "
    
    ConsultaConciliacao = fgQueryXMLLerTodos("ConsultaConciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not ObjectContext Is Nothing Then
'       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConsultaConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Consultar as quantidades, relativas às operações parcialmente conciliadas, para promover o controle de visualização em tela
'--------------------------------------------------------------------------------------------------

Public Function ObterQtdConciliadaOperacao(Optional ByVal pvntSequenciaOperacao As Variant = 0) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT    /*+ RULE */                                                 " & vbNewLine & _
             "           a.NU_SEQU_OPER_ATIV,                                        " & vbNewLine & _
             "           SUM(b.QT_ATIV_MERC_CNCL)  QT_ATIV_MERC_CNCL                 " & vbNewLine & _
             " FROM      A8.TB_OPER_ATIV            a,                               " & vbNewLine & _
             "           A8.TB_CNCL_OPER_ATIV       b                                " & vbNewLine & _
             " WHERE     a.NU_SEQU_OPER_ATIV  = B.NU_SEQU_OPER_ATIV                  " & vbNewLine & _
             " AND       a.DT_OPER_ATIV >= " & fgDtXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD"))
             
    If pvntSequenciaOperacao <> 0 Then
        strSQL = strSQL & vbNewLine & _
                 " AND      b.NU_SEQU_CNCL_OPER_ATIV_MESG         = " & pvntSequenciaOperacao & "       " & vbNewLine
    End If
    
    strSQL = strSQL & vbNewLine & _
             " GROUP BY a.NU_SEQU_OPER_ATIV"
    
    ObterQtdConciliadaOperacao = fgQueryXMLLerTodos("QtdConciliadaOperacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QtdConciliadaOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Consultar as quantidades, relativas às mensagens parcialmente conciliadas, para promover o controle de visualização em tela
'--------------------------------------------------------------------------------------------------

Public Function ObterQtdConciliadaMensagem(Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                                           Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   NU_CTRL_IF,                                                               " & vbNewLine & _
             "          DH_REGT_MESG_SPB,                                                         " & vbNewLine & _
             "          SUM(QT_ATIV_MERC_CNCL)  QT_ATIV_MERC_CNCL                                 " & vbNewLine & _
             " FROM     A8.TB_CNCL_OPER_ATIV                                                      " & vbNewLine & _
             " WHERE    NU_CTRL_IF          IS NOT NULL                                           " & vbNewLine & _
             " AND      DH_REGT_MESG_SPB    IS NOT NULL                                           " & vbNewLine
             
    
    If Trim(pstrNumeroControleIF) <> "" Then
        strSQL = strSQL & vbNewLine & _
            " AND         NU_CTRL_IF                        = '" & pstrNumeroControleIF & "'" & vbNewLine
    End If
    
    If Trim(pstrDataRegistroMensagemSPB) <> "" Then
        strSQL = strSQL & vbNewLine & _
            " AND       TRUNC(DH_REGT_MESG_SPB)             = " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB)
    End If
    
    strSQL = strSQL & " AND       DH_REGT_MESG_SPB >= " & fgDtXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD"))
    
    strSQL = strSQL & vbNewLine & _
             " GROUP BY NU_CTRL_IF,                                                               " & vbNewLine & _
             "          DH_REGT_MESG_SPB"
    
    ObterQtdConciliadaMensagem = fgQueryXMLLerTodos("QtdConciliadaMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQtdConciliadaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Promover a conciliação entre operações e mensagens, por quantidades e, especificamente para o caso de despesas, por valores, de acordo com os critérios abaixo:
'     Por Quantidades:
'    1) 1 mensagem para N operações
'        sem divergência
'    2) 1 operação para N mensagens
'        sem divergência
'    3) 1 operação para 1 mensagem
'        Com ou sem divergência
'    4) 1 operação para nenhuma mensagem
'    5) 1 mensagem para nenhuma operação'
'
'     Por Valores:
'    1) N operações para N mensagens
'        Com ou sem divergência'
'
'    Para as conciliações de despesa, verificar a necessidade de envio da mensagem SEL1069 (IF requisita Transferência para pagamento dos custos de utilização), a partir da mensagem SEL1611 (SEL informa Operações processadas na rotina de abertura), uma vez que a SEL1069 ainda não tenha sido enviada.

'Em seguida, verificar a necessidade de envio de movimento REALIZADO para o PJ e Sub-Reserva, de acordo com os seguintes tipos de operações e condições:
'    Leilão Primário
'    Leilão Secundário   e
'    Eventos de Resgate
'    Eventos de Amortização
'    Eventos de Juros,
'        sem envio de Repasse Financeiro
'--------------------------------------------------------------------------------------------------

Public Function Conciliar(ByVal xmlLoteOperacaoMensagem As MSXML2.DOMDocument40) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessaOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlDespesa                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim intTipoOperacao                         As enumTipoOperacaoLQS
Dim intStatusOperacao                       As enumStatusOperacao
Dim intStatusMensagem                       As enumStatusMensagem
Dim lngSequencialConciliacaoOperacao        As Long
Dim lngEmpresa                              As Long
Dim intTipoJustificativaConciliacao         As Integer
Dim strComentario                           As String
Dim strxmlRetornoValidacao                  As String
Dim blnExisteOperacaoDespesa                As Boolean
Dim blnEncerraStatusMesgOpDespesa           As Boolean
Dim blnExisteMensagemDespesa                As Boolean
Dim blnDespesaEncerrada                     As Boolean
Dim strxmlRetornoInfo                       As String
Dim blnGeraRealizado                        As Boolean
Dim blnGeraPJSel1611                        As Boolean
Dim intFuncionalidade                       As enumFuncionalidade
Dim intCont                                 As Integer
Dim strNumeroControleDespesa                As String
Dim datDataMensagemDespesa                  As Date
Dim blnPrimeiro                             As Boolean
Dim blnExisteMensagemDespesaAux             As Boolean
Dim xmlCamposExtras                         As MSXML2.DOMDocument40
Dim xmlAux                                  As MSXML2.DOMDocument40


On Error GoTo ErrorHandler

    'Captura a Empresa do Lote
    lngEmpresa = Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Empresa").Text)

    'Status da mensagem - Carlos 17/06/2004
    If xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/AlteraStatusMensagem") Is Nothing Then
        intStatusMensagem = enumStatusMensagem.Conciliada
    Else
        If Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/AlteraStatusMensagem").Text) = enumIndicadorSimNao.Sim Then
            intStatusMensagem = enumStatusMensagem.Conciliada
        Else
            intStatusMensagem = enumStatusMensagem.AConciliar
        End If
    End If

    'Captura o Tipo de Operação e configura seu próximo STATUS
    intTipoOperacao = Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/TipoOperacao").Text)
    Select Case intTipoOperacao
        Case enumTipoOperacaoLQS.CompromissadaVoltaConciliacao
            'intStatusOperacao = enumStatusOperacao.LiquidadaFisicamente
            intStatusOperacao = enumStatusOperacao.Liquidada
            blnGeraRealizado = True
            intFuncionalidade = enumFuncionalidade.Conciliacao

        Case enumTipoOperacaoLQS.TermoDataLiquidacaoCerta, enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta
            intStatusOperacao = enumStatusOperacao.Liquidada
            blnGeraRealizado = True
            intFuncionalidade = enumFuncionalidade.Conciliacao

        Case enumTipoOperacaoLQS.LeilaoPrimario, _
             enumTipoOperacaoLQS.LeilaoSecundario, _
             enumTipoOperacaoLQS.LeilaoSecundarioCompra
            intStatusOperacao = enumStatusOperacao.Conciliada

            blnGeraRealizado = True
            intFuncionalidade = enumFuncionalidade.Liberacao

        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
        
            intStatusOperacao = enumStatusOperacao.Concordancia
            blnGeraRealizado = True
            intFuncionalidade = enumFuncionalidade.Conciliacao

        Case enumTipoOperacaoLQS.EventosResgate, _
             enumTipoOperacaoLQS.EventosAmortização, _
             enumTipoOperacaoLQS.EventosJuros

            'Verifica se é necessário disponibilizar o Repasse Financeiro (1 - Sim, 0 - Não)
            'If Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DispRepasseFinanceiro").Text) = enumIndicadorSimNao.sim Then
            '    intStatusOperacao = enumStatusOperacao.LiquidadaFisicamente
            '    intStatusMensagem = enumStatusMensagem.LiquidadaFisicamente

            '    If Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeItemSelMensagem").Text) > 0 Then
            '        blnGeraPJSel1611 = True
            '        intFuncionalidade = enumFuncionalidade.LiquidarEventos
            '    End If
            'Else
                intStatusOperacao = enumStatusOperacao.Liquidada

                If Not xmlLoteOperacaoMensagem.selectSingleNode("//Operacao") Is Nothing Then
                    blnGeraRealizado = True
                    intFuncionalidade = enumFuncionalidade.Conciliacao
                Else
                    blnGeraPJSel1611 = True
                    intFuncionalidade = enumFuncionalidade.LiquidarEventos
                End If
            'End If

        Case enumTipoOperacaoLQS.DespesaSelic
            intStatusOperacao = enumStatusOperacao.LiquidadaFisicamente
            blnGeraRealizado = True
            intFuncionalidade = enumFuncionalidade.Conciliacao
            blnExisteMensagemDespesaAux = False
    End Select

    'Valida o lote de conciliação recebido
    strxmlRetornoValidacao = flValidaConciliacao(intTipoOperacao, lngEmpresa, xmlLoteOperacaoMensagem)
    If Trim(strxmlRetornoValidacao) <> vbNullString Then
        Conciliar = strxmlRetornoValidacao                                      '<-- Retorno Inconsistências de Negócio

        Exit Function
    End If

    'Verifica se existe tipo de justificativa...
    If xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/TipoJustificativa").length > 0 Then
        intTipoJustificativaConciliacao = _
            Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/TipoJustificativa").Text)

        'Verifica se existe comentário para a justificativa...
        If xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Comentario").length > 0 Then
            strComentario = _
                xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Comentario").Text
        End If
    End If

    'Inclui o Tipo de Justificativa / Comentário para a conciliação e captura o SEQUENCIAL gerado
    lngSequencialConciliacaoOperacao = IncluirJustificativaConciliacaoOperacao(intTipoJustificativaConciliacao, _
                                                                               strComentario)
    
    'Verifica se é uma conciliação 1 MSG -> N OP...                 (Neste caso, sem divergência de QUANTIDADES)
    If UCase(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/PontaN").Text) = "O" Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        
        For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Operacao")
            Call objOperacao.AlterarStatusAcao(objDomNode.selectSingleNode("Operacao").Text, _
                                               intStatusOperacao, _
                                               objDomNode.selectSingleNode("DataUltimaAtualizacaoOperacao").Text)
            
            Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                            objDomNode.selectSingleNode("Operacao").Text, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                            Val(objDomNode.selectSingleNode("QuantidadeConciliada").Text))
        Next
        
        Set objOperacao = Nothing
        
        Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Call objMensagem.AlterarStatus(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                       xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                       intStatusMensagem, _
                                       xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoMensagem").Text)
        Set objMensagem = Nothing
        
    'Verifica se é uma conciliação 1 OP -> N MSG...                 (Neste caso, sem divergência de QUANTIDADES)
    ElseIf UCase(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/PontaN").Text) = "M" Then
        Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

        For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Mensagem")
            Call objMensagem.AlterarStatus(objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                           objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text, _
                                           intStatusMensagem, _
                                           objDomNode.selectSingleNode("DataUltimaAtualizacaoMensagem").Text)

            Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                            objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                            objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text, _
                                            Val(objDomNode.selectSingleNode("QuantidadeConciliada").Text))
        Next
        
        Set objMensagem = Nothing
        
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Call objOperacao.AlterarStatusAcao(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                           intStatusOperacao, _
                                           xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoOperacao").Text)
        Set objOperacao = Nothing
        
    'Verifica se é uma conciliação N MSG -> N OP...                 (Neste caso, com ou sem divergência de VALORES)
    ElseIf UCase(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/PontaN").Text) = "MO" Then
        
        'Verifica se as mensagens foram vinculadas às operações...
        If xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao").length > 0 Then
            '...se sim, atualiza o STATUS das operações
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            
            For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao/Grupo_CodigoOperacao")
                Call objOperacao.AlterarStatusAcao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                   intStatusOperacao, _
                                                   objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text)
            Next
            
            Set objOperacao = Nothing
            
            blnExisteOperacaoDespesa = True
            blnExisteMensagemDespesaAux = True
        End If
        
        'Processa mensagens recebidas
        Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
        Set xmlDespesa = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlDespesa, "", "Repeat_Info", "")
        
        For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Mensagem")
            'Atualiza o STATUS das mensagens
            Call objMensagem.AlterarStatus(objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                           objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text, _
                                           intStatusMensagem, _
                                           objDomNode.selectSingleNode("DataUltimaAtualizacaoMensagem").Text)
            
            'Verifica se as mensagens foram vinculadas às operações...
            If blnExisteOperacaoDespesa Then
                '...se sim, inclui a conciliação (MENSAGEM COM OPERAÇÕES)
                For Each objDomNodeAux In _
                    xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao/Grupo_CodigoOperacao")
                    
                    Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                                    objDomNodeAux.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                    objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                                    objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text)
                Next
            
            '...se não, inclui a conciliação (MENSAGEM SEM OPERAÇÕES)
            Else
                Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                                , _
                                                objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                                objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text)
            End If
            
            'Verifica se a mensagem SEL1069 para o Número de Comando e Empresa correntes,
            'já chegou e está com STATUS ENCERRADA
            blnExisteMensagemDespesa = False
            If objDomNode.selectSingleNode("NumeroComando").Text <> vbNullString Then
                intCont = intCont + 1
                
                blnDespesaEncerrada = flVerificaMensagemDespesa(lngEmpresa, _
                                                                objDomNode.selectSingleNode("NumeroComando").Text, _
                                                                objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text, _
                                                                blnExisteMensagemDespesa, _
                                                                strNumeroControleDespesa, _
                                                                datDataMensagemDespesa)
                If intCont = 1 Then
                    blnEncerraStatusMesgOpDespesa = blnDespesaEncerrada
                Else
                    blnEncerraStatusMesgOpDespesa = blnEncerraStatusMesgOpDespesa And blnDespesaEncerrada
                End If
            Else
                blnEncerraStatusMesgOpDespesa = False
            End If

            Call fgAppendNode(xmlDespesa, "Repeat_Info", "Grupo_Info", "")

            'Se não existir a MENSAGEM DE PAGAMENTO DE DESPESA,
            'então gerar a SEL1069 a partir da SEL1611, com o STATUS CONCORDÂNCIA
            If Not blnExisteMensagemDespesa Then
                'Captura o XML da mensagem SEL1611...
                Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
                Call xmlMensagem.loadXML(fgSelectVarchar4000(objDomNode.selectSingleNode("TextoXML").Text))

                '...altera o código da mensagem para SEL1069...
                xmlMensagem.selectSingleNode("//CodMsg").Text = MESG_DESPESA

                'Alteração de Debito/Credito da SEL1611 para gerar a SEL1069
                If xmlMensagem.selectSingleNode("//TpDeb_Cred").Text = "C" Then
                    xmlMensagem.selectSingleNode("//TpDeb_Cred").Text = "D"
                Else
                    xmlMensagem.selectSingleNode("//TpDeb_Cred").Text = "C"
                End If

                '...grava a mensagem SEL1069, com o Tipo BackOffice TESOURARIA e Local de Liquidação SELIC
                Call objMensagem.Salvar(xmlMensagem, _
                                        lngEmpresa, _
                                        enumTipoBackOffice.Tesouraria, _
                                        enumLocalLiquidacao.SELIC, _
                                        "", _
                                        "A8", _
                                        objDomNode.selectSingleNode("NumeroComando").Text, _
                                        enumStatusMensagem.Concordancia, _
                                        enumIndicadorSimNao.Nao, _
                                        lngSequencialConciliacaoOperacao)

                Set xmlMensagem = Nothing

                Call fgAppendNode(xmlDespesa, "Grupo_Info", "Mensagem", MESG_DESPESA, "Repeat_Info")
            Else

                'Atualiza a Sequência de Conciliação na tabela de Mensagens Recebidas e Enviadas
                Call objMensagem.AlterarSequenciaConciliacao(strNumeroControleDespesa, _
                                                             datDataMensagemDespesa, _
                                                             lngSequencialConciliacaoOperacao, _
                                                             1)

            End If
            
            Call fgAppendNode(xmlDespesa, "Grupo_Info", "NumeroComando", _
                        objDomNode.selectSingleNode("NumeroComando").Text, "Repeat_Info")
        Next
        
        strxmlRetornoInfo = xmlDespesa.xml
        Set xmlDespesa = Nothing
        
        'Verifica se o encerramento de mensagens e operações deve ser processado...
        If blnEncerraStatusMesgOpDespesa Then
            For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Mensagem")
                'Atualiza o STATUS das mensagens para ENCERRADA
                Call objMensagem.AlterarStatus(objDomNode.selectSingleNode("NumeroControleIF").Text, _
                                               objDomNode.selectSingleNode("DataRegistroMensagemSPB").Text, _
                                               enumStatusMensagem.Encerrada, _
                                               objDomNode.selectSingleNode("DataUltimaAtualizacaoMensagem").Text, _
                                               , _
                                               False)
            Next
            
            If blnExisteOperacaoDespesa Then
                Set objOperacao = CreateObject("A8LQS.clsOperacao")
                'Atualiza o STATUS das operações para LIQUIDADA
                For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao/Grupo_CodigoOperacao")
                    Call objOperacao.AlterarStatusAcao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                       enumStatusOperacao.Liquidada, _
                                                       objDomNode.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                                       , _
                                                       , _
                                                       False)
                Next
                
                Set objOperacao = Nothing
            End If
        End If
        
        Set objMensagem = Nothing
    Else
        'Verifica se é uma conciliação 1 OP -> 1 MSG...
        If xmlLoteOperacaoMensagem.selectNodes("//Operacao").length > 0 And _
           xmlLoteOperacaoMensagem.selectNodes("//NumeroControleIF").length > 0 Then
    
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Call objOperacao.AlterarStatusAcao(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                               IIf(Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/AlteraStatusOperacao").Text) = 1, _
                                                        intStatusOperacao, _
                                                        enumStatusOperacao.AConciliar), _
                                               xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoOperacao").Text)
            Set objOperacao = Nothing
            
            Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
            Call objMensagem.AlterarStatus(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                           xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                           intStatusMensagem, _
                                           xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoMensagem").Text)
            Set objMensagem = Nothing

            Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                            Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeConciliada").Text))
        
        'Verifica se é uma conciliação 1 OP -> 0 MSG...
        ElseIf xmlLoteOperacaoMensagem.selectNodes("//Operacao").length > 0 Then
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Call objOperacao.AlterarStatusAcao(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                               IIf(Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/AlteraStatusOperacao").Text) = 1, _
                                                        intStatusOperacao, _
                                                        enumStatusOperacao.AConciliar), _
                                               xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoOperacao").Text)
            Set objOperacao = Nothing
            
            Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Operacao").Text, _
                                            vbNullString, _
                                            vbNullString, _
                                            Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeConciliada").Text))

        'Verifica se é uma conciliação 1 MSG -> 0 OP...
        ElseIf xmlLoteOperacaoMensagem.selectNodes("//NumeroControleIF").length > 0 Then
            Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
            Call objMensagem.AlterarStatus(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                           xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                           intStatusMensagem, _
                                           xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataUltimaAtualizacaoMensagem").Text)
            Set objMensagem = Nothing

            Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                            vbNullString, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/NumeroControleIF").Text, _
                                            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/DataRegistroMensagemSPB").Text, _
                                            Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeConciliada").Text))

        End If

    End If

    If intTipoOperacao = enumTipoOperacaoLQS.DespesaSelic And _
        blnEncerraStatusMesgOpDespesa Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao/Grupo_CodigoOperacao")
            Call objOperacao.AlterarStatusOperIndisp(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                     enumStatusOperacao.Liquidada)
        Next
        Set objOperacao = Nothing
    End If
    
    '>>> Verifica se deve gerar o REALIZADO para o PJ (Gestão do Caixa) e SBR (Sub Reserva) -----------------------
    If blnGeraRealizado Then

        Set objProcessaOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        If intTipoOperacao = enumTipoOperacaoLQS.DespesaSelic And _
           blnExisteMensagemDespesaAux And _
           blnExisteOperacaoDespesa Then

            'Enviar a diferença da conciliação de Despesas
            objProcessaOperacao.EnviarDespesasPJ xmlLoteOperacaoMensagem

            For Each objDomNode In xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/ConciliacaoOperacao/Repeat_CodigoOperacao/Grupo_CodigoOperacao")
                Call objProcessaOperacao.GerenciarChamadaViaTela(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                 intFuncionalidade, _
                                                                 False, _
                                                                 True, , _
                                                                 xmlLoteOperacaoMensagem)
            Next

        Else
            Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
            Set xmlCamposExtras = CreateObject("MSXML2.DOMDocument.4.0")

            If Not xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Operacao") Is Nothing Then
                
                'Criado para conciliar evento com quantidade de mensagens SPB maior que operacao
                'envia Realizado PJ só na primeira passagem
                Call fgAppendNode(xmlLoteOperacaoMensagem, "Grupo_Conciliacao", "PrimeiroEnvio", enumIndicadorSimNao.Sim)
                
                'Apenas para conciliações com mais de uma operação, ou seja, 1 Mensagem para N Operações,
                'neste caso as quantidades devem estar batidas e o comando é sempre para finalizar a conciliação
                For Each objDomNode In _
                    xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Repeat_Operacao")

                    xmlAux.loadXML xmlLoteOperacaoMensagem.xml
                    If intTipoOperacao = enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA Then
                        'Campos extras para geração da mensagem
                        Set objOperacao = CreateObject("A8LQS.clsOperacao")
                        objOperacao.CamposExtrasLeilaoBMA objDomNode.selectSingleNode("Operacao").Text, _
                                                          xmlCamposExtras
                        Set objOperacao = Nothing
                        fgAppendXML xmlAux, "Repeat_Conciliacao", xmlCamposExtras.xml
                    End If

                    Call objProcessaOperacao.GerenciarChamadaViaTela(objDomNode.selectSingleNode("Operacao").Text, _
                                                                     intFuncionalidade, _
                                                                     False, _
                                                                     True, , _
                                                                     xmlAux)

                    xmlLoteOperacaoMensagem.selectSingleNode("//PrimeiroEnvio").Text = xmlAux.selectSingleNode("//PrimeiroEnvio").Text

                Next

            ElseIf intTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
                    intTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
                    intTipoOperacao = enumTipoOperacaoLQS.EventosResgate Then

                    If Not xmlLoteOperacaoMensagem.selectSingleNode("//Operacao") Is Nothing Then
                        '...apenas entrar no gerenciar chamadas se o comando for de finalização de conciliação
                        If Not xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusMensagem") Is Nothing Then
                            If Val(xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusMensagem").Text) = 1 Then

                                Call objProcessaOperacao.GerenciarChamadaViaTela(xmlLoteOperacaoMensagem.selectSingleNode("//Operacao").Text, _
                                                                                 intFuncionalidade, _
                                                                                 False, _
                                                                                 True, , _
                                                                                 xmlLoteOperacaoMensagem)
                            ElseIf Not xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusOperacao") Is Nothing Then
                                If Val(xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusOperacao").Text) = 1 Then

                                    Call objProcessaOperacao.GerenciarChamadaViaTela(xmlLoteOperacaoMensagem.selectSingleNode("//Operacao").Text, _
                                                                                     intFuncionalidade, _
                                                                                     False, _
                                                                                     True, , _
                                                                                     xmlLoteOperacaoMensagem)

                                End If
                            End If
                        ElseIf Not xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusOperacao") Is Nothing Then
                            If Val(xmlLoteOperacaoMensagem.selectSingleNode("//AlteraStatusOperacao").Text) = 1 Then
    
                                Call objProcessaOperacao.GerenciarChamadaViaTela(xmlLoteOperacaoMensagem.selectSingleNode("//Operacao").Text, _
                                                                                 intFuncionalidade, _
                                                                                 False, _
                                                                                 True, , _
                                                                                 xmlLoteOperacaoMensagem)
    
                            End If
                        End If
                    End If
            Else

                'Apenas para conciliações com uma operação, ou seja, 1 Operação para 0 mensagem ou 1 mensagem,
                'neste caso as quantidades podem ou não estar batidas...
                If Not xmlLoteOperacaoMensagem.selectSingleNode("//Operacao") Is Nothing Then
                    '...apenas entrar no gerenciar chamadas se o comando for de finalização de conciliação
                    If Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/AlteraStatusOperacao").Text) = 1 Then
                        
                        xmlAux.loadXML xmlLoteOperacaoMensagem.xml
                        If intTipoOperacao = enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA Then
                            'Campos extras para geração da mensagem
                            Set objOperacao = CreateObject("A8LQS.clsOperacao")
                            objOperacao.CamposExtrasLeilaoBMA xmlLoteOperacaoMensagem.selectSingleNode("//Operacao").Text, _
                                                              xmlCamposExtras
                            Set objOperacao = Nothing
                            fgAppendXML xmlAux, "Repeat_Conciliacao", xmlCamposExtras.xml
                        End If
                        
                        Call objProcessaOperacao.GerenciarChamadaViaTela(xmlLoteOperacaoMensagem.selectSingleNode("//Operacao").Text, _
                                                                         intFuncionalidade, _
                                                                         False, _
                                                                         True, , _
                                                                         xmlAux)

                    End If
                End If
            End If
        End If

        Set objProcessaOperacao = Nothing
    End If
    '--------------------------------------------------------------------------------------------------------------

    'Gerar PJ através das SEL1611 - Carlos 17/06/04
    If blnGeraPJSel1611 Then
        'Criado para enviar o retorno para o Legado quando tiver operacao
        Call fgAppendNode(xmlLoteOperacaoMensagem, "Grupo_Conciliacao", "RetornoLegado", enumIndicadorSimNao.Sim)

        Set objProcessaOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        Call objProcessaOperacao.GerenciarChamadaViaTela(0, _
                                                         intFuncionalidade, _
                                                         False, _
                                                         , , _
                                                         xmlLoteOperacaoMensagem)
        Set objProcessaOperacao = Nothing
    End If
    '--------------------------------------------------------------------------------------------------------------

    Conciliar = strxmlRetornoInfo                                               '<-- Retorno OK
                                                                                '    Para a DESPESA retorna XML INFO

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    Set xmlDespesa = Nothing
    Set objProcessaOperacao = Nothing

    If Not ObjectContext Is Nothing Then
    '   'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0        '<-- Retorno Erros
    Call fgRaiseError(App.EXEName, TypeName(Me), "Conciliar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

    Exit Function
    Resume

End Function

'--------------------------------------------------------------------------------------------------
'Promover a inclusão de registros de Justificativa de Conciliação na tabela:
'    TB_JUST_CNCL_OPER_ATIV_MESG
'--------------------------------------------------------------------------------------------------

Public Function IncluirJustificativaConciliacaoOperacao(Optional ByVal pintTipoJustificativaConciliacao As Integer = 0, _
                                                        Optional ByVal pstrTextoConciliacao As String = vbNullString) As Long

Dim vntSequencialConciliacaoOperacao        As Variant
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    vntSequencialConciliacaoOperacao = fgExecuteSequence("A8.SQ_A8_NU_SEQU_CNCL_OPER_MESG")
    
    strSQL = " INSERT INTO A8.TB_JUST_CNCL_OPER_ATIV_MESG (                 " & vbNewLine & _
             "         NU_SEQU_CNCL_OPER_ATIV_MESG,                         " & vbNewLine & _
             "         TP_JUST_CNCL,                                        " & vbNewLine & _
             "         TX_JUST,                                             " & vbNewLine & _
             "         CO_USUA_ATLZ,                                        " & vbNewLine & _
             "         CO_ETCA_TRAB_ATLZ,                                   " & vbNewLine & _
             "         DH_JUST_CNCL)                                        " & vbNewLine & _
             " VALUES (" & vntSequencialConciliacaoOperacao & ",            " & vbNewLine & _
             "         " & IIf(pintTipoJustificativaConciliacao = 0, _
                                "NULL", _
                                pintTipoJustificativaConciliacao) & ",      " & vbNewLine & _
             "         " & IIf(pstrTextoConciliacao = vbNullString, _
                                "NULL", _
                                "'" & pstrTextoConciliacao & "'") & ",      " & vbNewLine & _
             "        '" & fgUsuarioRede & "',                              " & vbNewLine & _
             "        '" & fgEstacaoTrabalhoUsuario & "',                   " & vbNewLine & _
             "         SYSDATE)"
    
    fgExecuteSQL strSQL
    
    IncluirJustificativaConciliacaoOperacao = vntSequencialConciliacaoOperacao
    
    Exit Function
    
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirJustificativaConciliacaoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Promover a inclusão dos registros conciliados de operação e/ou mensagem na tabela:
'    TB_CNCL_OPER_ATIV
'--------------------------------------------------------------------------------------------------

Public Function IncluirConciliacaoOperacao(ByVal plngSequencialConciliacaoOperacao As Long, _
                                  Optional ByVal pvntSequenciaOperacao As Variant = vbNullString, _
                                  Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                                  Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString, _
                                  Optional ByVal plngQuantidadeConciliada As Long = 0, _
                                  Optional ByVal pintNumeroSequencialContrRepeticao As Integer = 1, _
                                  Optional ByVal pblnVerificaRegistroJaExiste As Boolean = False) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    If pblnVerificaRegistroJaExiste Then
        If ConsultaConciliacao(Val(pvntSequenciaOperacao), pstrNumeroControleIF, pstrDataRegistroMensagemSPB) <> vbNullString Then
            Exit Function
        End If
    End If
                                             
    strSQL = " INSERT INTO A8.TB_CNCL_OPER_ATIV (                                           " & vbNewLine & _
             "         NU_SEQU_CNCL_OPER_ATIV_MESG,                                         " & vbNewLine & _
             "         NU_SEQU_OPER_ATIV,                                                   " & vbNewLine & _
             "         NU_CTRL_IF,                                                          " & vbNewLine & _
             "         DH_REGT_MESG_SPB,                                                    " & vbNewLine & _
             "         QT_ATIV_MERC_CNCL,                                                   " & vbNewLine & _
             "         NU_SEQU_CNTR_REPE)                                                   " & vbNewLine & _
             " VALUES (" & plngSequencialConciliacaoOperacao & ",                           " & vbNewLine & _
             "         " & IIf(pvntSequenciaOperacao = vbNullString, _
                                    "NULL", _
                                    pvntSequenciaOperacao) & ",                             " & vbNewLine & _
             "         " & IIf(pstrNumeroControleIF = vbNullString, _
                                    "NULL", _
                                    "'" & pstrNumeroControleIF & "'") & ",                  " & vbNewLine & _
             "         " & IIf(pstrDataRegistroMensagemSPB = vbNullString, _
                                    "NULL", _
                                    fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB)) & ",  " & vbNewLine & _
             "         " & plngQuantidadeConciliada & ", " & _
             "         " & pintNumeroSequencialContrRepeticao & ")"
    
    fgExecuteSQL strSQL
    
    Exit Function
    
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirConciliacaoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Promover a validação de todas as considerações necessárias para a conciliação e, em caso de inconsistências detectadas, enviar um XML de retorno à camada de interface, contendo todos os itens encontrados.
'As validações estão descritas conforme abaixo:
'Quantidades:
'    1) Operação e Mensagem informadas e Títulos divergentes. Regra Impeditiva
'    2) Mair de uma operação e mensagem informadas. Regra Impeditiva
'    3) Mais de uma operação informada e nenhuma mensagem informada. Regra Impeditiva
'    4) Mais de uma mensagem informada e nenhuma operação informada. Regra Impeditiva
'    5) Mais de uma operação informada, apenas uma mensagem informada e quantidades divergentes. Regra Impeditiva
'    6) Mais de uma mensagem informada, apenas uma operação informada e quantidades divergentes. Regra Impeditiva
'    7) Justificativa não informada e quantidades divergentes. Regra de Aviso
'    8) Justificativa não informada e PUs divergentes. Regra de Aviso
'    9) Justificativa não informada, tipos de operações iguais a:
'       Compromissada Volta Conciliação
'       Termo Liquidação
'       Eventos de Resgate
'       Eventos de Amortização
'       Eventos de Juros e Nº comando divergente
'
'    Leilão Primário
'    Leilão Secundário e Nº comando original divergente
'    Regra de Aviso

'Valores:
'    1) Totais atuais modificados, exemplo: recebimento de novas operações e/ou mensagens, em relação aos totais informados no momento da solicitação de conciliação. Regra Impeditiva
'    2) Operações informadas e mensagens não informadas. Regra Impeditiva
'    3) Valores divergentes e justificativa não informada. Regra de Aviso
'--------------------------------------------------------------------------------------------------

Private Function flValidaConciliacao(ByVal pintTipoOperacao As Integer, _
                                     ByVal plngEmpresa As Long, _
                                     ByVal xmlLoteOperacaoMensagem As MSXML2.DOMDocument40) As String

'---------------------------------------------------------------------------------------------------------
'>>> Validação do XML de Conciliação
'---------------------------------------------------------------------------------------------------------
'    Os erros, se encontrados, serão retornados na própria TAG em forma de atributos:
'               EN      - Erro Negócio                      ou
'               EN_IMP  - Erro Negócio Impeditivo
'---------------------------------------------------------------------------------------------------------

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim blnExisteJustificativa                  As Boolean
Dim blnRetornaErro                          As Boolean
Dim strxmlRetornoComp                       As String
Dim lngQtdItemSelOperacao                   As Long
Dim lngQtdItemSelMensagem                   As Long
Dim lngQtdOperacao                          As Long
Dim lngQtdMensagem                          As Long
Dim blnQtdMesgOpOK                          As Boolean
Dim blnVlrMesgOpOK                          As Boolean
Dim blnLeilaoBMA                            As Boolean
Dim blnEventos                              As Boolean
Dim intTipoOperacao                         As Integer

On Error GoTo ErrorHandler
    
    blnLeilaoBMA = (pintTipoOperacao = enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA)
    
    'Verifica se a Justificativa e o comentário foram informados
    blnExisteJustificativa = _
        (xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/TipoJustificativa").length > 0 And _
         xmlLoteOperacaoMensagem.selectNodes("Repeat_Conciliacao/Grupo_Conciliacao/Comentario").length > 0)
         
    'Captura as quantidades de itens (Operação e Mensagem)
    lngQtdItemSelOperacao = _
        xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeItemSelOperacao").Text
        
    lngQtdItemSelMensagem = _
        xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeItemSelMensagem").Text

    If Not xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeTotalOperacao") Is Nothing Then
        lngQtdOperacao = _
            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeTotalOperacao").Text
    End If

    If Not xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeTotalMensagem") Is Nothing Then
        lngQtdMensagem = _
            xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/QuantidadeTotalMensagem").Text
    End If

    'Verifica se a conciliação é por VALOR...               (DESPESA)
    If pintTipoOperacao = enumTipoOperacaoLQS.DespesaSelic Then
        'Valida se a conciliação foi alterada,
        'comparando o XML inicial de operações com o XML a ser capturado abaixo
        
        '>>> Formata XML Filtro padrão ------------------------------------------------------------------------
        Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_Status", "Status", enumStatusOperacao.AConciliar)
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacaoRotinaAbertura", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacaoRotinaAbertura", "TipoOperacaoRotinaAbertura", "S")
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", pintTipoOperacao)
        
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
        
        'Filtro também implementado no formulário << frmConciliacaoOperacao >>

        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataIni", _
                                fgDtHrXML_To_Oracle(Format(Date, "YYYYMMDD") & "000000"))
                                
        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle("99991231235959"))
        '>>> --------------------------------------------------------------------------------------------------
        
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        strxmlRetornoComp = objOperacao.ObterValoresPorGrupoVeiculoLegal(xmlDomFiltros)
        Set objOperacao = Nothing
        Set xmlDomFiltros = Nothing
        
        If Replace(strxmlRetornoComp, vbNewLine, vbNullString) <> _
           Replace(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/ConciliacaoOperacao").xml, vbNewLine, vbNullString) Then
            
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "ConciliacaoOperacao", _
                            "EN_IMP", _
                            "Totais apresentados foram modificados, a conciliação deverá ser reenviada.")
            
            blnRetornaErro = True
        End If
        
        'Pelo menos uma Operação selecionada e Mensagem NÃO selecionada
        If lngQtdItemSelOperacao > 0 And lngQtdItemSelMensagem = 0 Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelOperacao", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação para operações de despesa sem mensagens associadas.")
            
            blnRetornaErro = True
        End If
        
        'Validações específicas: JUSTIFICATIVA não for informada
        If Not blnExisteJustificativa Then
            'Compara VALOR
            If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "ValorTotalMensagem", "ValorTotalOperacao") Then
                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                       "ValorTotalMensagem", _
                                       "EN", _
                                       "Valor Total das Operações divergente do Valor Total das Mensagens.")
                
                blnRetornaErro = True
            End If
        End If

    '...se não, a conciliação é por QUANTIDADE
    Else
        If lngQtdItemSelOperacao > 0 And lngQtdItemSelMensagem > 0 Then
            'Compara TÍTULO - Restrição Impeditiva
            If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "Titulo") Then
                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                "Titulo", _
                                "EN_IMP", _
                                "Identificador do Título não pode ser divergente.")

                blnRetornaErro = True
            End If

            'Compara Vencimento - Restrição Impeditiva
            If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "Vencimento") Then
                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                "Vencimento", _
                                "EN_IMP", _
                                "Data de Vencimento do Título não pode ser divergente.")

                blnRetornaErro = True
            End If

            intTipoOperacao = Val(xmlLoteOperacaoMensagem.selectSingleNode("Repeat_Conciliacao/Grupo_Conciliacao/TipoOperacao").Text)
            blnEventos = False

            If intTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
                intTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
                intTipoOperacao = enumTipoOperacaoLQS.EventosResgate Then

                blnEventos = True

                'Compara Conta do Veiculo Legal - Restrição Impeditiva
                If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "ContaCustodia") Then
                    Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                    "ContaCustodia", _
                                    "EN_IMP", _
                                    "Conta Selic do Veículo Legal não pode ser divergente.")

                    blnRetornaErro = True
                End If

            End If

        End If

        'Captura diferença de quantidades
        blnQtdMesgOpOK = _
            flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "QuantidadeTotalOperacao", "QuantidadeTotalMensagem")

        'Pikachu - 31/03/2005
        'Para Conciliacao de Eventos com 1 Oper e N Mensagens com Qauntidade não batida,
        'Se Conciliacao Finalizada estiver selecionada , liquidar operacao e mensagem
        If intTipoOperacao = enumTipoOperacaoLQS.EventosAmortização Or _
            intTipoOperacao = enumTipoOperacaoLQS.EventosJuros Or _
            intTipoOperacao = enumTipoOperacaoLQS.EventosResgate Then
            
            If Not blnQtdMesgOpOK Then
                If Not xmlLoteOperacaoMensagem.selectSingleNode("//ConciliacaoFinalizada") Is Nothing Then
                    If xmlLoteOperacaoMensagem.selectSingleNode("//ConciliacaoFinalizada").Text Then
                        blnQtdMesgOpOK = True
                    End If
                End If
            End If
        End If

        'Mais de uma operação selecionada e Mais de uma mensagem selecionada
        If lngQtdItemSelOperacao > 1 And lngQtdItemSelMensagem > 1 Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelOperacao", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação com seleções múltiplas de operação e mensagem.")

            blnRetornaErro = True

        'Mais de uma operação selecionada e Mensagem NÃO selecionada
        ElseIf lngQtdItemSelOperacao > 1 And lngQtdItemSelMensagem = 0 Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelOperacao", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação para múltiplas operações sem informar mensagem.")

            blnRetornaErro = True

        'Mais de uma mensagem selecionada e Operação NÃO selecionada
        ElseIf lngQtdItemSelMensagem > 1 And lngQtdItemSelOperacao = 0 Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelMensagem", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação para múltiplas mensagens sem informar operação.")

            blnRetornaErro = True

        'Mais de uma operação selecionada e Quantidade a conciliar NÃO batida
        ElseIf lngQtdItemSelOperacao > 1 And Not blnQtdMesgOpOK And Not blnLeilaoBMA Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelOperacao", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação para múltiplas operações com divergência de quantidades.")

            blnRetornaErro = True

        'Mais de uma mensagem selecionada e Quantidade a conciliar NÃO batida
        ElseIf lngQtdItemSelMensagem > 1 And Not blnQtdMesgOpOK Then
            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                            "QuantidadeItemSelMensagem", _
                            "EN_IMP", _
                            "Impossível efetuar a conciliação para múltiplas mensagens com divergência de quantidades.")

            blnRetornaErro = True

        End If
        
        If blnLeilaoBMA Then
            If lngQtdOperacao > lngQtdMensagem Then
                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                "QuantidadeItemSelOperacao", _
                                "EN_IMP", _
                                "Quantidade Total da Operação superior à Quantidade Total da Mensagem.")
                blnRetornaErro = True
            End If
        End If

        'Validações específicas: JUSTIFICATIVA não for informada
        If Not blnExisteJustificativa Then
            If Not blnLeilaoBMA And (lngQtdItemSelOperacao > 0 And lngQtdItemSelMensagem > 0) Or _
               (lngQtdItemSelOperacao = 1 And lngQtdItemSelMensagem = 0) Or _
               (lngQtdItemSelOperacao = 0 And lngQtdItemSelMensagem = 1) Then

                If lngQtdItemSelOperacao >= 1 And lngQtdItemSelMensagem >= 1 Then
                    'Compara QUANTIDADE TOTAL
                    If Not blnQtdMesgOpOK Then
                        Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                               "QuantidadeTotalOperacao", _
                                               "EN", _
                                               "Quantidade Total da Operação divergente da Quantidade Total da Mensagem.")
    
                        blnRetornaErro = True
                    End If
                Else
                    Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                           "QuantidadeTotalOperacao", _
                                           "EN", _
                                           "Obrigatório informar a justificativa desta conciliação.")

                    blnRetornaErro = True
                End If

            End If

            If lngQtdItemSelOperacao > 0 And lngQtdItemSelMensagem > 0 Then

                If blnEventos Then
                    'Compara o Valor Total de Mensagem e Operacao
                    If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "ValorTotalMensagem", "ValorTotalOperacao") Then
                        Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                               "ValorTotalMensagem", _
                                               "EN", _
                                               "Valor Total das Operações divergente do Valor Total das Mensagens.")

                        blnRetornaErro = True
                    End If

                Else

                    'Compara PU
                    If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "PU") Then
                        Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                               "PU", _
                                               "EN", _
                                               "PU da Operação divergente do PU da Mensagem.")

                        blnRetornaErro = True
                    End If
                End If

                'Validações específicas por Tipo de Operação                        (SEM JUSTIFICATIVA)
                Select Case pintTipoOperacao
                    Case enumTipoOperacaoLQS.CompromissadaVoltaConciliacao, _
                         enumTipoOperacaoLQS.TermoDataLiquidacaoCerta, _
                         enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta

                        'Compara COMANDO
                        If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "Comando") Then
                            Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                                   "Comando", _
                                                   "EN", _
                                                   "Número de Comando da Operação divergente do Número de Comando da Mensagem.")

                            blnRetornaErro = True
                        End If

                    Case enumTipoOperacaoLQS.LeilaoPrimario, _
                         enumTipoOperacaoLQS.LeilaoSecundario, _
                         enumTipoOperacaoLQS.LeilaoSecundarioCompra

                        'Compara COMANDO ORIGINAL
                        If xmlLoteOperacaoMensagem.documentElement.selectSingleNode("//Repeat_Operacao") Is Nothing Then
                            If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "ComandoOriginal", "Comando", True) Then
                                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                                       "ComandoOriginal", _
                                                       "EN", _
                                                       "Número de Comando Original da Operação divergente do Número de Comando da Mensagem.")

                                blnRetornaErro = True
                            End If
                        Else
                            'Qdo vem mais de 1 operação para conciliar com a mesma mensagem, as informações das mensagens vem antes
                            If Not flComparaAtributosConciliacao(xmlLoteOperacaoMensagem, "Comando", "ComandoOriginal", True) Then
                                Call fgAppendAttribute(xmlLoteOperacaoMensagem, _
                                                       "ComandoOriginal", _
                                                       "EN", _
                                                       "Número de Comando Original da Operação divergente do Número de Comando da Mensagem.")

                                blnRetornaErro = True
                            End If
                        End If
                End Select
            End If
        End If
    End If

    If blnRetornaErro Then
        flValidaConciliacao = xmlLoteOperacaoMensagem.xml
    Else
        flValidaConciliacao = vbNullString
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidaConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Promover a comparação de informações, inerentes à conciliação, fornecidos, tais como:
'    Valor
'    Título
'    Quantidade
'    PU
'    Nº de Comando   e   Nº de Comando Original
'--------------------------------------------------------------------------------------------------

Private Function flComparaAtributosConciliacao(ByVal xmlDomFiltro As MSXML2.DOMDocument40, _
                                               ByVal pstrAtributo As String, _
                                      Optional ByVal pstrAtributo2 As String = vbNullString, _
                                      Optional ByVal pblnAtributo2Mensagem As Boolean = False) As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim strAux                                  As String
Dim blnRetornaErro                          As Boolean
Dim vntConteudo1                            As Variant
Dim vntConteudo2                            As Variant
Dim blnSegundo                              As Boolean

On Error GoTo ErrorHandler

    If pstrAtributo2 = vbNullString Then
        For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo)
            If objDomNode.Text <> vbNullString Then
                If strAux <> vbNullString Then
                    If strAux <> objDomNode.Text Then
                        blnRetornaErro = True

                        Exit For
                    End If
                Else
                    strAux = objDomNode.Text
                End If
            End If
        Next
    ElseIf pblnAtributo2Mensagem Then
        If xmlDomFiltro.documentElement.selectSingleNode("//Repeat_Operacao") Is Nothing Then

            For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo)
                If objDomNode.Text <> vbNullString Then
                    vntConteudo1 = objDomNode.Text
    
                    Exit For
                End If
            Next

            blnSegundo = False
            For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo2)
                If objDomNode.Text <> vbNullString Then
                    If blnSegundo Then
                            vntConteudo2 = objDomNode.Text
    
                        Exit For
                    End If
                    blnSegundo = True
                End If
            Next

            If UCase(vntConteudo1) <> UCase(vntConteudo2) Then blnRetornaErro = True
        Else
            
            For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo)
                If objDomNode.Text <> vbNullString Then
                    vntConteudo1 = objDomNode.Text
    
                    Exit For
                End If
            Next

            For Each objDomNodeAux In xmlDomFiltro.selectNodes("//Repeat_Operacao")
                For Each objDomNode In objDomNodeAux.selectNodes(pstrAtributo2)
                    If objDomNode.Text <> vbNullString Then
                                vntConteudo2 = objDomNode.Text
        
                    End If
                Next

                If UCase(vntConteudo1) <> UCase(vntConteudo2) Then blnRetornaErro = True
                If blnRetornaErro Then
                    Exit For
                End If
            Next

        End If
    Else
        For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo)
            If objDomNode.Text <> vbNullString Then
                vntConteudo1 = objDomNode.Text

                Exit For
            End If
        Next

        For Each objDomNode In xmlDomFiltro.selectNodes("//" & pstrAtributo2)
            If objDomNode.Text <> vbNullString Then
                vntConteudo2 = objDomNode.Text

                Exit For
            End If
        Next

        If UCase(vntConteudo1) <> UCase(vntConteudo2) Then blnRetornaErro = True

    End If

    flComparaAtributosConciliacao = Not blnRetornaErro

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flComparaAtributosConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Consultar todos as conciliações cadastradas na tabela:
'    TB_CNCL_OPER_ATIV
'--------------------------------------------------------------------------------------------------

Public Function LerTodos(Optional ByVal pvntSequenciaConciliacao As Variant = 0, _
                         Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                         Optional ByVal pstrNumeroControleIF As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   A.NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV, " & vbCrLf & _
             "          A.NU_CTRL_IF, " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,  " & vbCrLf & _
             "          A.QT_ATIV_MERC_CNCL, " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,  " & vbCrLf & _
             "          B.CO_MESG_SPB " & vbCrLf & _
             " FROM     A8.TB_CNCL_OPER_ATIV A, " & vbCrLf & _
             "          A8.TB_MESG_RECB_ENVI_SPB B " & vbCrLf & _
             " WHERE    A.NU_SEQU_CNCL_OPER_ATIV_MESG = B.NU_SEQU_CNCL_OPER_ATIV_MESG " & vbCrLf
             
    If pvntSequenciaConciliacao > 0 Then
        strSQL = strSQL & " AND      A.NU_SEQU_CNCL_OPER_ATIV_MESG =   " & pvntSequenciaConciliacao & vbCrLf
    End If

    If pvntSequenciaOperacao > 0 Then
        strSQL = strSQL & " And      A.NU_SEQU_OPER_ATIV           =   " & pvntSequenciaOperacao & vbCrLf
    End If

    If pstrNumeroControleIF <> vbNullString Then
        strSQL = strSQL & " AND      A.NU_CTRL_IF                  =     B.NU_CTRL_IF " & vbCrLf
        strSQL = strSQL & " AND      A.DH_REGT_MESG_SPB            =     A.DH_REGT_MESG_SPB " & vbCrLf
        strSQL = strSQL & " AND      A.NU_SEQU_CNTR_REPE           =     B.NU_SEQU_CNTR_REPE " & vbCrLf
        strSQL = strSQL & " And      A.NU_CTRL_IF                  =    '" & pstrNumeroControleIF & "'" & vbCrLf
        strSQL = strSQL & " Order by A.NU_SEQU_OPER_ATIV                 " & vbCrLf
    End If

    LerTodos = fgQueryXMLLerTodos("Conciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Obter NU_SEQU_OPER_ATIV da tabela A8.TB_CNCL_OPER_ATIV
'--------------------------------------------------------------------------------------------------

Public Function LerSequenciaOperacao(ByVal pvntSequenciaConciliacao As Variant) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Select   DISTINCT(NU_SEQU_OPER_ATIV) " & vbCrLf & _
             "  From    A8.TB_CNCL_OPER_ATIV " & vbCrLf & _
             " Where    NU_SEQU_CNCL_OPER_ATIV_MESG =   " & pvntSequenciaConciliacao & vbCrLf

    LerSequenciaOperacao = fgQueryXMLLerTodos("Conciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerSequenciaOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Promover a inclusão de um novo registro na tabela:
'    TB_MESG_RECB_SPB_CNCL
'    para o controle de mensagens já conciliadas
'--------------------------------------------------------------------------------------------------

Public Function Incluir(ByVal pstrControleIF As String, ByVal pstrDataRegistroMsg As String, ByVal pstrDataOperacao As String, _
                        ByVal plngCodigoOperacaoSelic As Long, ByVal plngComandoOperacaoOrig As Long, ByVal pstrTituloSelic As String, _
                        ByVal pvntPrecoUnitario As Variant, ByVal pvntQuantidadeTitulo As Variant, ByVal pvntValorFinanceiro As Variant, _
               Optional ByVal plngSeqControleRepeticao As Long = 1, Optional ByVal pstrDataLiquidacao As String = vbNullString, _
               Optional ByVal pintIndicadorDebitoCredito As Integer = 0, Optional ByVal pstrDataVencimento As String = vbNullString, _
               Optional ByVal pvntCNPJContraparte As Variant = 0, Optional ByVal pstrNomeContraparte As String = vbNullString, _
               Optional ByVal plngContaCustVeiculoLegal As Long = 0, Optional ByVal pstrTipoInformacaoLDL As String = vbNullString, _
               Optional ByVal pvntCodigoTitularCustodiante As Variant = 0, Optional ByVal plngTipoNegociacaoBMA As Long = 0, _
               Optional ByVal pstrTaxaNegociada As String = "", Optional ByVal pstrISPBBancoLiquContraparte As String = "", _
               Optional ByVal plngTipoLiquidacao As Long = 0, Optional ByVal pstrDescricaoTitulo As String = "", _
               Optional ByVal pstrCodigoOperacaoCETIP As String = "", Optional ByVal pvntCodigoAgencia As Variant = 0, _
               Optional ByVal pvntNumeroCC As Variant = 0, Optional ByVal pstrDataRetornoOperacao As String = vbNullString, _
               Optional ByVal pvntPrazoDiasRetornoOperacao As Variant = 0, Optional ByVal pvntValorRetornoOperacao As Variant = 0, _
               Optional ByVal pvntIdentContraParteCamara As Variant = vbNullString, Optional ByVal pstrTipoPagamentoLDL As String = vbNullString, _
               Optional ByVal pstrCodigoMoedaEstrangeira As String = vbNullString, Optional ByVal pstrDataLiquidacaoMoedaEstrangeira As String = vbNullString, _
               Optional ByVal pstrCodigoPraca As String = vbNullString, Optional ByVal pstrCodigoSisbacen As String = vbNullString, _
               Optional ByVal pstrValorMoedaEstrangeira As String = vbNullString, Optional ByVal pstrIspbIfContraparte As String = vbNullString, _
               Optional ByVal pstrCanalSISBACENCorretora As String = vbNullString, Optional ByVal pstrCanalOperacaoInterbancaria As String = vbNullString, _
               Optional ByVal pvntNrIdefNegoBMC As Variant = 0, Optional ByVal pstrChACAM As String = vbNullString) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Insert Into A8.TB_MESG_RECB_SPB_CNCL (       " & vbCrLf & _
             "      NU_CTRL_IF,                             " & vbCrLf & _
             "      DH_REGT_MESG_SPB,                       " & vbCrLf & _
             "      NU_SEQU_CNTR_REPE,                      " & vbCrLf & _
             "      DT_OPER,                                " & vbCrLf & _
             "      CO_OPER_SELIC,                          " & vbCrLf & _
             "      NU_COMD_OPER_ORIG,                      " & vbCrLf & _
             "      NU_ATIV_MERC,                           " & vbCrLf & _
             "      PU_ATIV_MERC,                           " & vbCrLf & _
             "      QT_ATIV_MERC,                           " & vbCrLf & _
             "      VA_FINC,                                " & vbCrLf & _
             "      DT_LIQU,                                " & vbCrLf & _
             "      IN_OPER_DEBT_CRED,                      " & vbCrLf & _
             "      DT_VENC,                                " & vbCrLf & _
             "      CO_CNPJ_CNPT,                           " & vbCrLf & _
             "      NO_CNPT,                                " & vbCrLf & _
             "      CO_CNTA_CUTD_SELIC_VEIC_LEGA,           " & vbCrLf & _
             "      TP_INFO_LDL,                            " & vbCrLf & _
             "      CO_TITL_CUTD,                           " & vbCrLf & _
             "      TP_NEGO_BMA,                            " & vbCrLf & _
             "      PE_TAXA_NEGO,                           " & vbCrLf & _
             "      CO_ISPB_BANC_LIQU_CNPT,                 " & vbCrLf & _
             "      TP_LIQU_OPER_ATIV,                      " & vbCrLf & _
             "      CO_OPER_CETIP,                          " & vbCrLf & _
             "      DE_ATIV_MERC,                           " & vbCrLf
    
    strSQL = strSQL & _
             "      CO_AGEN_COTR,                           " & vbCrLf & _
             "      NU_CC_COTR,                             " & vbCrLf & _
             "      DT_OPER_ATIV_RETN,                      " & vbCrLf & _
             "      PZ_DIAS_RETN_OPER_ATIV,                 " & vbCrLf & _
             "      VA_OPER_ATIV_RETN,                      " & vbCrLf & _
             "      CO_CNPT_CAMR,                           " & vbCrLf & _
             "      TP_PAGT_LDL,                            " & vbCrLf & _
             "      CO_MOED_ESTR,                           " & vbCrLf & _
             "      DT_LIQU_OPER_ATIV_MOED_ESTR,            " & vbCrLf & _
             "      CO_PRAC,                                " & vbCrLf & _
             "      CO_CNTR_SISB,                           " & vbCrLf & _
             "      VA_MOED_ESTR,                           " & vbCrLf & _
             "      CO_ISPB_IF_CNPT,                        " & vbCrLf & _
             "      CO_SISB_COTR,                           " & vbCrLf & _
             "      CO_CNAL_OPER_INTE,                      " & vbCrLf & _
             "      NR_IDEF_NEGO_BMC,                       " & vbCrLf & _
             "      CD_ASSO_CAMB)                           "

    strSQL = strSQL & _
                    " Values (" & vbCrLf & _
                    "'" & pstrControleIF & "', " & vbCrLf & _
                    fgDtHrXML_To_Oracle(pstrDataRegistroMsg) & ", " & vbCrLf & _
                    plngSeqControleRepeticao & ", " & vbCrLf & _
                    fgDtXML_To_Oracle(pstrDataOperacao) & ", " & vbCrLf & _
                    plngCodigoOperacaoSelic & ", " & vbCrLf & _
                    IIf(plngComandoOperacaoOrig = 0, "Null", plngComandoOperacaoOrig) & ", " & vbCrLf & _
                    "'" & pstrTituloSelic & "', " & vbCrLf & _
                    fgVlrToDBServer(pvntPrecoUnitario) & ", " & vbCrLf & _
                    IIf(pvntQuantidadeTitulo = vbNullString, "Null", fgVlrToDBServer(pvntQuantidadeTitulo)) & ", " & vbCrLf & _
                    fgVlrToDBServer(pvntValorFinanceiro) & ", " & vbCrLf & _
                    IIf(pstrDataLiquidacao = vbNullString, "Null", fgDtXML_To_Oracle(pstrDataLiquidacao)) & ", " & vbCrLf
    
    strSQL = strSQL & _
                    IIf(pintIndicadorDebitoCredito = 0, "Null", pintIndicadorDebitoCredito) & ", " & vbCrLf & _
                    IIf(pstrDataVencimento = vbNullString, "Null", fgDtXML_To_Oracle(pstrDataVencimento)) & ", " & vbCrLf & _
                    IIf(pvntCNPJContraparte = 0, "Null", pvntCNPJContraparte) & ", " & vbCrLf & _
                    "'" & IIf(pstrNomeContraparte = vbNullString, "", pstrNomeContraparte) & "', " & vbCrLf & _
                    IIf(plngContaCustVeiculoLegal = 0, "Null", plngContaCustVeiculoLegal) & ", " & vbCrLf & _
                    "'" & IIf(pstrTipoInformacaoLDL = vbNullString, "", pstrTipoInformacaoLDL) & "', " & vbCrLf & _
                    IIf(pvntCodigoTitularCustodiante = 0, "Null", pvntCodigoTitularCustodiante) & ", " & vbCrLf & _
                    IIf(plngTipoNegociacaoBMA = 0, "Null", plngTipoNegociacaoBMA) & ", " & vbCrLf & _
                    IIf(pstrTaxaNegociada = vbNullString, "Null", fgVlrToDBServer(pstrTaxaNegociada)) & ", " & vbCrLf & _
                    IIf(pstrISPBBancoLiquContraparte = vbNullString, "Null", "'" & Trim(pstrISPBBancoLiquContraparte) & "'") & ", " & vbCrLf & _
                    IIf(plngTipoLiquidacao > 8, "Null", plngTipoLiquidacao) & ", " & vbCrLf & _
                    IIf(pstrCodigoOperacaoCETIP = vbNullString, "Null", "'" & pstrCodigoOperacaoCETIP & "'") & ", " & vbCrLf & _
                    IIf(pstrDescricaoTitulo = vbNullString, "Null", "'" & pstrDescricaoTitulo & "'") & ", " & vbCrLf & _
                    IIf(pvntCodigoAgencia = 0, "Null", pvntCodigoAgencia) & ", " & vbCrLf & _
                    IIf(pvntNumeroCC = 0, "Null", pvntNumeroCC) & ", " & vbCrLf & _
                    IIf(pstrDataRetornoOperacao = vbNullString, "Null", fgDtXML_To_Oracle(pstrDataRetornoOperacao)) & ", " & vbCrLf & _
                    IIf(pvntPrazoDiasRetornoOperacao = 0, "Null", fgVlrToDBServer(pvntPrazoDiasRetornoOperacao)) & ", " & vbCrLf & _
                    IIf(pvntValorRetornoOperacao = 0, "Null", fgVlrToDBServer(pvntValorRetornoOperacao)) & ", " & vbCrLf & _
                    IIf(pvntIdentContraParteCamara = vbNullString, "Null", pvntIdentContraParteCamara) & ", " & vbCrLf & _
                    IIf(pstrTipoPagamentoLDL = vbNullString, "Null", pstrTipoPagamentoLDL) & ", " & vbCrLf

    strSQL = strSQL & _
                    IIf(pstrCodigoMoedaEstrangeira = vbNullString, "Null", pstrCodigoMoedaEstrangeira) & ", " & vbCrLf & _
                    IIf(pstrDataLiquidacaoMoedaEstrangeira = vbNullString, "Null", fgDtXML_To_Oracle(pstrDataLiquidacaoMoedaEstrangeira)) & ", " & vbCrLf & _
                    IIf(pstrCodigoPraca = vbNullString, "Null", pstrCodigoPraca) & ", " & vbCrLf & _
                    IIf(pstrCodigoSisbacen = vbNullString, "Null", pstrCodigoSisbacen) & ", " & vbCrLf & _
                    IIf(pstrValorMoedaEstrangeira = vbNullString, "Null", fgVlrToDBServer(pstrValorMoedaEstrangeira)) & ", " & vbCrLf & _
                    "'" & IIf(pstrIspbIfContraparte = vbNullString, "", pstrIspbIfContraparte) & "', " & vbCrLf & _
                    IIf(pstrCanalSISBACENCorretora = vbNullString, "Null", fgVlrToDBServer(pstrCanalSISBACENCorretora)) & ", " & vbCrLf & _
                    "'" & IIf(pstrCanalOperacaoInterbancaria = vbNullString, "", pstrCanalOperacaoInterbancaria) & "', " & vbCrLf & _
                    IIf(pvntNrIdefNegoBMC = 0, "Null", pvntNrIdefNegoBMC) & ", " & vbCrLf & _
                    "'" & IIf(pstrChACAM = vbNullString, "", pstrChACAM) & "' " & vbCrLf & _
                    ")"

    Call fgExecuteSQL(strSQL)

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Controlar a distribuição da operação a ser conciliada.
'
'Verificar se a operção é:
'1. Compromissada de Volta com Conciliação
'   Executar o método flConciliarCompromissada;
'2. Termo na Data da Liquidação
'   Executar o método flConciliarTermoDataLiquidacao;
'3. Eventos de Amortização, Eventos de Juros ou Eventos de Resgate
'   Executar o método flConciliarEventosSelic;
'4. Leilão Primário ou Leilão Secundário
'   Executar o método flConciliarLeilao;
'--------------------------------------------------------------------------------------------------

Public Function ConciliarOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                         Optional ByRef plngJustificativa As Long = 0) As Boolean

Dim xmlTipos                                As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlTipos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlTipos.loadXML ObterTiposConciliacaoCETIP

    Select Case xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text
        Case enumTipoOperacaoLQS.CompromissadaVoltaConciliacao
            ConciliarOperacao = flConciliarCompromissada(xmlRemessa)
        Case enumTipoOperacaoLQS.TermoDataLiquidacaoCerta, enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta
            ConciliarOperacao = flConciliarTermoDataLiquidacao(xmlRemessa)
        Case enumTipoOperacaoLQS.EventosAmortização, enumTipoOperacaoLQS.EventosJuros, enumTipoOperacaoLQS.EventosResgate
            ConciliarOperacao = flConciliarEventosSelic(xmlRemessa)
        Case enumTipoOperacaoLQS.LeilaoPrimario, enumTipoOperacaoLQS.LeilaoSecundario, enumTipoOperacaoLQS.LeilaoSecundarioCompra
            ConciliarOperacao = flConciliarLeilao(xmlRemessa)
        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            ConciliarOperacao = flConciliarLeilao(xmlRemessa)
        Case enumTipoOperacaoLQS.DespesaSelic
            ConciliarOperacao = flConciliarDespesas(xmlRemessa)
        Case Else
            If (xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMA) _
                Or _
               (xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.BMC) _
                Or _
               (xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP _
                 And xmlTipos.selectNodes("//*[TP_OPER='" & xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text & "']").length > 0) _
                    Then
                ConciliarOperacao = flConciliarRegistroOperacao(xmlRemessa, plngJustificativa)
            ElseIf xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
                ConciliarOperacao = flConciliarCETIP(xmlRemessa)
            End If
    End Select

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Controlar a distribuição da mensagem para a conciliação.
'Verificar se a mensagem contém a tag "NumOpSELOr" e a mesma está com valor maior "0", se não atender encerrar a execução deste método;
'
'Executar o método clsOperacao.LerXMLOperacao;'
'
'Verificar se a operção é:
'1. Compromissada de Volta
'   Executar o método flConciliarCompromissada;
'2. Termo na Data da Liquidação
'   Executar o método flConciliarTermoDataLiquidacao;
'3. Eventos de Amortização, Eventos de Juros ou Eventos de Resgate
'   Executar o método flConciliarEventosSelic;
'4. Leilão Primário ou Leilão Secundário
'   Executar o método flConciliarLeilao;
'--------------------------------------------------------------------------------------------------

Public Function ConciliarMensagem(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                         Optional ByVal plngCodigoEmpresa As Long) As Boolean


Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim xmlControleProcessoOperacao             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlOperacaoXML                          As MSXML2.DOMDocument40
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim strTipoOperacao                         As String
Dim strOperacao                             As String
Dim strXML                                  As String

On Error GoTo ErrorHandler

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    If Val("0" & xmlMensagem.documentElement.selectSingleNode("//CodOpSEL").Text) = 0 Then
        Exit Function
    End If
    strTipoOperacao = objTipoOperacao.LerTodos(, xmlMensagem.documentElement.selectSingleNode("//CodOpSEL").Text)
    Set objTipoOperacao = Nothing
    'Se não achar a o tipo de operação sair pois é uma mensasgem que não concilia.
    If Trim(strTipoOperacao) = vbNullString Then
        Exit Function
    End If
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlTipoOperacao.loadXML strTipoOperacao

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    Select Case CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text)
        Case enumTipoOperacaoLQS.LeilaoPrimario, _
             enumTipoOperacaoLQS.LeilaoSecundario, _
             enumTipoOperacaoLQS.LeilaoSecundarioCompra
             
            strOperacao = objOperacao.LerTodos(, , , CLng(xmlMensagem.documentElement.selectSingleNode("//NumOpSEL").Text), _
                                               fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DtMovto").Text), _
                                               CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                                               , , plngCodigoEmpresa)

        Case enumTipoOperacaoLQS.EventosResgate, _
             enumTipoOperacaoLQS.EventosJuros, _
             enumTipoOperacaoLQS.EventosAmortização
             
            strOperacao = objOperacao.LerTodos(, , , , _
                                               fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DtMovto").Text), _
                                               CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                                               , xmlMensagem.documentElement.selectSingleNode("//IdentdTitSEL").Text, _
                                                plngCodigoEmpresa, _
                                                , fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DtVenc").Text))

        Case Else
            strOperacao = objOperacao.LerTodos(, , CLng(xmlMensagem.documentElement.selectSingleNode("//NumOpSEL").Text), _
                                               , fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DtMovto").Text), _
                                               CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                                               , , plngCodigoEmpresa)
    End Select

    'Se não achar a operação não tenta conciliar automáticamente.
    If Trim(strOperacao) = vbNullString Then
        Exit Function
    End If

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    xmlOperacao.loadXML strOperacao

    Set xmlOperacaoXML = objOperacao.LerXMLOperacao(xmlOperacao.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

    If CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.LiquidadaFisicamente Or _
       CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.BaixadaViaContingencia Or _
       CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Liquidada Or _
       CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.LiquidadaViaContingencia Then
        If CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosResgate Or _
            CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosJuros Or _
            CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosAmortização Then
            Exit Function
        Else
            '3038 - Operação já liquidada.
            lngCodigoErroNegocio = 3038
            GoTo ErrorHandler
        End If
    End If

    '***********************************************************************************
    'Ver dse a operação esta com status A CONCILIAR - Carlos 16/07/2004
    If CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusOperacao.AConciliar Then
        Exit Function
    End If

    'Verificar se deve ou não conciliar automático - Carlos 06/02/2004
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set xmlControleProcessoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    strXML = objProcessoOperacao.LerTodos("Confirmacao", _
                                          CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text), _
                                          enumIndicadorSimNao.Nao)

    If strXML = vbNullString Then
        Exit Function
    End If

    xmlControleProcessoOperacao.loadXML strXML

'    If CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Inicial Or _
'        CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.EmSer Or _
'        CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Cancelada Or _
'        CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.CanceladaOrigem Then
'        Exit Function
'    End If

    If xmlControleProcessoOperacao.documentElement.selectSingleNode("Grupo_ProcessoOperacao/IN_VERI_REGR_CNCL").Text = enumIndicadorSimNao.Sim Then
        Set objRegra = CreateObject("A8LQS.clsRegra")
        'Verifica regra para conciliação automática
        If Not objRegra.VerificarRegraAutomatica(xmlOperacaoXML, enumFuncaoSistema.Conciliar) Then
            Exit Function
        End If
        Set objRegra = Nothing
    Else
        Exit Function
    End If
    'Fim - Verificar se deve ou não conciliar automático - Carlos 06/02/2004
    '***********************************************************************************

    Select Case CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text)
        Case enumTipoOperacaoLQS.CompromissadaVoltaConciliacao
            ConciliarMensagem = flConciliarCompromissada(xmlOperacaoXML)
        Case enumTipoOperacaoLQS.TermoDataLiquidacaoCerta, enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta
            ConciliarMensagem = flConciliarTermoDataLiquidacao(xmlOperacaoXML)
        Case enumTipoOperacaoLQS.EventosAmortização, enumTipoOperacaoLQS.EventosJuros, enumTipoOperacaoLQS.EventosResgate
            ConciliarMensagem = flConciliarEventosSelic(xmlOperacaoXML)
        Case enumTipoOperacaoLQS.LeilaoPrimario, enumTipoOperacaoLQS.LeilaoSecundario, enumTipoOperacaoLQS.LeilaoSecundarioCompra
            ConciliarMensagem = flConciliarLeilao(xmlOperacaoXML)
        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            ConciliarMensagem = flConciliarLeilao(xmlOperacaoXML)
    End Select

    If (CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoCerta Or _
        CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.TermoDataLiquidacaoIncerta Or _
        CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosAmortização Or _
        CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosJuros Or _
        CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EventosResgate) And _
        CLng(xmlOperacaoXML.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConciliadaAutomatica Then
        'Função especifica de termo liquidação e eventos, pois não tem liberação para a mesma
        If xmlOperacaoXML.documentElement.selectSingleNode("TP_MESG_RETN_INTE") Is Nothing Then
            fgAppendNode xmlOperacaoXML, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.documentElement.selectSingleNode("//TP_MESG_RETN_INTE").Text
        Else
            xmlOperacaoXML.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text = xmlTipoOperacao.documentElement.selectSingleNode("//TP_MESG_RETN_INTE").Text
        End If

        objProcessoOperacao.LiquidarTermoVoltaEventos xmlOperacaoXML, xmlOperacaoXML

    ElseIf CLng(xmlTipoOperacao.documentElement.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaVoltaConciliacao Then
        If xmlOperacaoXML.documentElement.selectSingleNode("TP_MESG_RETN_INTE") Is Nothing Then
            fgAppendNode xmlOperacaoXML, "MESG", "TP_MESG_RETN_INTE", xmlTipoOperacao.documentElement.selectSingleNode("//TP_MESG_RETN_INTE").Text
        Else
            xmlOperacaoXML.documentElement.selectSingleNode("TP_MESG_RETN_INTE").Text = xmlTipoOperacao.documentElement.selectSingleNode("//TP_MESG_RETN_INTE").Text
        End If

        objProcessoOperacao.LiquidarComprVoltaConciliacao xmlOperacaoXML

    Else
        'Verificar se irá liberar automático
        If CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Conciliada Or _
            CLng(xmlOperacaoXML.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConciliadaAutomatica Then
            objProcessoOperacao.GerenciarChamadaViaTela xmlOperacaoXML.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, enumFuncionalidade.Liberacao, , True
        End If
    End If

    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoXML = Nothing
    Set xmlTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlOperacaoXML = Nothing
    Set xmlTipoOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Conciliar operações compromissadas.
'
'Executar o método clsMensagemEnviadaRecebida.ObterValoresConciliacao;'
'Verificar se encontrou a mensagem para conciliação (SEL1611):
'   1. Verificar se a quantidade de titulos e o valor do PU são iguais entre a mensagem e a operação;
'      1.1 Executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "LiquidadaFisicamente";
'      1.2 Executar o método clsMensagemEnviadaRecebida.AlterarStatus setanto o parâmetro status "Conciliada";
'   2. Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'
'Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'--------------------------------------------------------------------------------------------------

Private Function flConciliarCompromissada(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim blnSEL1611                              As Boolean

On Error GoTo ErrorHandler

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strXML = objMensagemEnviadaRecebida.ObterValoresConciliacao(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                                enumCodigoOperacaoSelic.RetornoCompromissosTitulosVencemDia, _
                                                                True, _
                                                                xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    If strXML <> vbNullString Then
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML strXML
        blnSEL1611 = True
    Else
        blnSEL1611 = False
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If blnSEL1611 Then
        If Val(xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//QT_ATIV_MERC").Text) And _
            Val(xmlRemessa.documentElement.selectSingleNode("PU_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//PU_ATIV_MERC").Text) And _
            Trim(xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text) = Trim(xmlMensagem.documentElement.selectSingleNode("//NU_ATIV_MERC").Text) And _
            fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text) = fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DT_VENC").Text) Then
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.Liquidada
                                                'enumStatusOperacao.LiquidadaFisicamente
            'Alterar status da mensagem
            objMensagemEnviadaRecebida.AlterarStatus xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.Conciliada, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text
            flGravarConciliacao xmlRemessa, xmlMensagem
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
            'xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.LiquidadaFisicamente
            flConciliarCompromissada = True
        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar
            flConciliarCompromissada = False
        End If
        objMensagemEnviadaRecebida.AlterarTipoBackOffice xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                         fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    Else
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AConciliar
        flConciliarCompromissada = False
    End If

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarCompromissada Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Conciliar operações de termo na data da conciliação.
'
'Executar o método clsMensagemEnviadaRecebida.ObterValoresConciliacao;'
'
'Verificar se encontrou a mensagem para conciliação (SEL1611):
'   1. Verificar se a quantidade de titulos e o valor do PU são iguais entre a mensagem e a operação;
'      1.1 Executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "ConciliadaAutomatica";
'      1.2 Executar o método clsMensagemEnviadaRecebida.AlterarStatus setanto o parâmetro status "Conciliada";
'   2. Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'   3. Executar o método clsMensagemEnviadaRecebida.AlterarTipoBackOffice;
'Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'--------------------------------------------------------------------------------------------------

Private Function flConciliarTermoDataLiquidacao(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim blnSEL1611                              As Boolean
Dim lngCodigoOperacaoSelic                  As Long

On Error GoTo ErrorHandler

    lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.LiquidacaoTermosDataLiquidacao
    
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strXML = objMensagemEnviadaRecebida.ObterValoresConciliacao(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                                lngCodigoOperacaoSelic, _
                                                                True, _
                                                                xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    If strXML <> vbNullString Then
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML strXML
        blnSEL1611 = True
    Else
        blnSEL1611 = False
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")

    If blnSEL1611 Then
        If Val(xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//QT_ATIV_MERC").Text) And _
            Val(xmlRemessa.documentElement.selectSingleNode("PU_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//PU_ATIV_MERC").Text) And _
            Trim(xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text) = Trim(xmlMensagem.documentElement.selectSingleNode("//NU_ATIV_MERC").Text) And _
            fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text) = fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DT_VENC").Text) Then
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.ConciliadaAutomatica
            'Alterar status da mensagem
            objMensagemEnviadaRecebida.AlterarStatus xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.Conciliada, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text
            flGravarConciliacao xmlRemessa, xmlMensagem
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica
            flConciliarTermoDataLiquidacao = True
        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar
            flConciliarTermoDataLiquidacao = False
        End If
        objMensagemEnviadaRecebida.AlterarTipoBackOffice xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                         fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    Else
        If xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.AConciliar Then
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar
        End If
        flConciliarTermoDataLiquidacao = False
    End If

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarTermoDataLiquidacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operçaões CETIP
'--------------------------------------------------------------------------------------------------

Private Function flConciliarCETIP(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

    On Error GoTo ErrorHandler

    flConciliarCETIP = GerenciarConciliacaoAutomaticaCETIP(xmlRemessa)

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de Despesas SELIC
'--------------------------------------------------------------------------------------------------

Private Function flConciliarDespesas(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objOperacao                             As A8LQS.clsOperacao

On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                        enumStatusOperacao.AConciliar
    flConciliarDespesas = False

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarDespesas Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operações com mensagens de registro ( BMA0002, CTP1002, BMC0005)
'--------------------------------------------------------------------------------------------------

Private Function flConciliarRegistroOperacao(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                             ByRef plngJustificativa As Long) As Boolean

Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strRetorno                              As String
Dim lngStatusOperacao                       As Long

On Error GoTo ErrorHandler

    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlConciliacao, "", "Repeat_Conciliacao", ""
    fgAppendNode xmlConciliacao, "Repeat_Conciliacao", "Grupo_Operacao", ""
    fgAppendNode xmlConciliacao, "Grupo_Operacao", "NU_SEQU_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    fgAppendNode xmlConciliacao, "Grupo_Operacao", "DH_ULTI_ATLZ", xmlRemessa.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
    fgAppendNode xmlConciliacao, "Grupo_Operacao", "LocalLiquidacao", xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
    fgAppendNode xmlConciliacao, "Grupo_Operacao", "TipoConfirmacao", "1"
    fgAppendNode xmlConciliacao, "Grupo_Operacao", "IgnoraGradeHorario", "0"

    strRetorno = ConciliarRegistroOperacao(enumAcaoConciliacao.BOConcordar, _
                                           xmlConciliacao, _
                                           , _
                                           True, _
                                           lngStatusOperacao, _
                                           xmlRemessa)

    If strRetorno = vbNullString Then
        flConciliarRegistroOperacao = True
        plngJustificativa = 0
        xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = lngStatusOperacao
    Else
        Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
        xmlErro.loadXML strRetorno
        If xmlErro.documentElement.selectSingleNode("/Erro/Grupo_ErrorInfo/CodigoJustificativa") Is Nothing Then
            plngJustificativa = enumJustificativa.RegraWorkflow
        Else
            plngJustificativa = CLng(xmlErro.documentElement.selectSingleNode("/Erro/Grupo_ErrorInfo/CodigoJustificativa").Text)
        End If
        Set xmlErro = Nothing
        flConciliarRegistroOperacao = False
    End If

    Set xmlConciliacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarRegistroOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Conciliar eventos Selic.
'
'Verificar o tipo de evento;
'EventosAmortização: Setar a variável local do Tipo da Rotina de Abertura com Amortizacao;
'EventosJuros: Setar a variável local do Tipo da Rotina de Abertura com Juros;
'EventosResgate: Setar a variável local do Tipo da Rotina de Abertura com Resgate;
'
'Verificar se o título está em carteira própria verificando se o valor da tag "CO_CNTA_CUTD_SELIC_VEIC_LEGA" é diferente de "0000" nas 4 penultimas posições;
'  1. executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'
'Executar o método clsMensagemEnviadaRecebida.ObterValoresConciliacao;'
'
'Verificar se encontrou a mensagem para conciliação (SEL1611):
'   1. Verificar se a quantidade de titulos e o valor do PU são iguais entre a mensagem e a operação;
'      1.1 Executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "ConciliadaAutomatica";
'      1.2 Executar o método clsMensagemEnviadaRecebida.AlterarStatus setanto o parâmetro status "Conciliada";
'   2. Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'   3. Executar o método clsMensagemEnviadaRecebida.AlterarTipoBackOffice;
'Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'--------------------------------------------------------------------------------------------------

Private Function flConciliarEventosSelic(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim blnSEL1611                              As Boolean
Dim lngCodigoOperacaoSelic                  As Long
Dim strConta                                As String

On Error GoTo ErrorHandler

    Select Case xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text
        Case enumTipoOperacaoLQS.EventosAmortização
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.Amortizacao
        Case enumTipoOperacaoLQS.EventosJuros
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.Juros
        Case enumTipoOperacaoLQS.EventosResgate
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.Resgate
    End Select

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    strConta = Mid(fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text, "0", 9, True), 1, 4)

    strXML = objMensagemEnviadaRecebida.ObterValoresConciliacao(xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                                                lngCodigoOperacaoSelic, _
                                                                False, _
                                                                xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text, _
                                                                strConta, _
                                                                xmlRemessa.documentElement.selectSingleNode("VA_OPER_ATIV").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text, _
                                                                xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text)

    If strXML <> vbNullString Then
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML strXML
        blnSEL1611 = True
    Else
        blnSEL1611 = False
    End If

    If blnSEL1611 Then
        If Val(xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//QT_ATIV_MERC").Text) And _
            fgVlrXml_To_Decimal(xmlRemessa.documentElement.selectSingleNode("VA_OPER_ATIV").Text) = fgVlrXml_To_Decimal(xmlMensagem.documentElement.selectSingleNode("//VA_FINC").Text) And _
            Trim(xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text) = Trim(xmlMensagem.documentElement.selectSingleNode("//NU_ATIV_MERC").Text) And _
            fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text) = fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DT_VENC").Text) Then

            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.ConciliadaAutomatica
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica

            'Alterar o Status da mensagem
            objMensagemEnviadaRecebida.AlterarStatus xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.Conciliada, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text

            flGravarConciliacao xmlRemessa, xmlMensagem

            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica
            flConciliarEventosSelic = True

            'Alterado para enviar o valor para o PJ a partir da mensagem SPB - Carlos 16/06/2004
            fgAppendNode xmlRemessa, "MESG", "NumeroControleIF", xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text

        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar
            flConciliarEventosSelic = False
        End If
        objMensagemEnviadaRecebida.AlterarTipoBackOffice xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                         fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    Else
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AConciliar
        flConciliarEventosSelic = False
    End If

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarEventosSelic Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Conciliar operação de leilão.

'Verificar de leilão;
'LeilaoPrimario: Setar a variável local do Tipo da Rotina de Abertura com LiquidacaoLeilaoVendaPrimario;
'LeilaoSecundario: Setar a variável local do Tipo da Rotina de Abertura com LiquidacaoLeilaoVendaSecundario;

'Executar o método clsMensagemEnviadaRecebida.ObterValoresConciliacao;
'
'Verificar se encontrou a mensagem para conciliação (SEL1611):
'   1. Verificar se a quantidade de titulos e o valor do PU são iguais entre a mensagem e a operação;
'      1.1 Executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "ConciliadaAutomatica";
'      1.2 Executar o método clsMensagemEnviadaRecebida.AlterarStatus setanto o parâmetro status "Conciliada";
'   2. Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'   3. Executar o método clsMensagemEnviadaRecebida.AlterarTipoBackOffice;
'Senão, executar o método clsOperacao.AlterarStatusOperIndisp setanto o parâmetro status "AConciliar";
'--------------------------------------------------------------------------------------------------

Private Function flConciliarLeilao(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim blnSEL1611                              As Boolean
Dim lngCodigoOperacaoSelic                  As Long
Dim lngStatusOperacao                       As enumStatusOperacao

On Error GoTo ErrorHandler

    lngStatusOperacao = enumStatusOperacao.ConciliadaAutomatica
    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
        Case enumTipoOperacaoLQS.LeilaoPrimario
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.LiquidacaoLeilaoVendaPrimario
        Case enumTipoOperacaoLQS.LeilaoSecundario
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.LiquidacaoLeilaoVendaSecundario
        Case enumTipoOperacaoLQS.LeilaoSecundarioCompra
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.LiquidacaoLeilaoCompraSecundario
        Case enumTipoOperacaoLQS.LeilaoVendaPrimarioBMA
            lngCodigoOperacaoSelic = enumCodigoOperacaoSelic.LiquidacaoLeilaoVendaPrimario
            lngStatusOperacao = enumStatusOperacao.ConcordanciaAutomatica
    End Select

    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    strXML = objMensagemEnviadaRecebida.ObterValoresConciliacao(xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text, _
                                                                Val(xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER_RETN").Text), _
                                                                lngCodigoOperacaoSelic, _
                                                                True, _
                                                                xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)

    If strXML <> vbNullString Then
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML strXML
        blnSEL1611 = True
    Else
        blnSEL1611 = False
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    If blnSEL1611 Then
        If Val(xmlRemessa.documentElement.selectSingleNode("QT_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//QT_ATIV_MERC").Text) And _
            Val(xmlRemessa.documentElement.selectSingleNode("PU_ATIV_MERC").Text) = Val(xmlMensagem.documentElement.selectSingleNode("//PU_ATIV_MERC").Text) And _
            Trim(xmlRemessa.documentElement.selectSingleNode("NU_ATIV_MERC").Text) = Trim(xmlMensagem.documentElement.selectSingleNode("//NU_ATIV_MERC").Text) And _
            fgDtXML_To_Date(xmlRemessa.documentElement.selectSingleNode("DT_VENC_ATIV").Text) = fgDtXML_To_Date(xmlMensagem.documentElement.selectSingleNode("//DT_VENC").Text) Then
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                lngStatusOperacao
            'Alterar o status da mensagem
            objMensagemEnviadaRecebida.AlterarStatus xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                     enumStatusMensagem.Conciliada, _
                                                     xmlMensagem.documentElement.selectSingleNode("//DH_ULTI_ATLZ").Text
            flGravarConciliacao xmlRemessa, xmlMensagem
            xmlRemessa.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConciliadaAutomatica
            flConciliarLeilao = True
        Else
            objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                enumStatusOperacao.AConciliar
            flConciliarLeilao = False
        End If
        objMensagemEnviadaRecebida.AlterarTipoBackOffice xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                                                         fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                                               xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    Else
        objOperacao.AlterarStatusOperIndisp xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.AConciliar
        flConciliarLeilao = False
    End If

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set objMensagemEnviadaRecebida = Nothing
    Set objOperacao = Nothing

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConciliarLeilao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Gravar o registro de conciliação.
'Executar o método IncluirJustificativaConciliacaoOperacao;
'Executar o método IncluirConciliacaoOperacao;
'--------------------------------------------------------------------------------------------------

Private Function flGravarConciliacao(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                     ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim lngSequencialJustificativa              As Long

On Error GoTo ErrorHandler

    lngSequencialJustificativa = IncluirJustificativaConciliacaoOperacao()

    IncluirConciliacaoOperacao lngSequencialJustificativa, _
                               xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                               xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text, _
                               xmlMensagem.documentElement.selectSingleNode("//DH_REGT_MESG_SPB").Text

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not ObjectContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flGravarConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Verificar a existência de mensagem de despesa na tabela A8.TB_MESG_RECB_ENVI_SPB, comparando empresa, comando da operação e o código da mensagem SPB.
'--------------------------------------------------------------------------------------------------

Private Function flVerificaMensagemDespesa(ByVal plngCodigoEmpresa As Long, _
                                           ByVal pstrNumeroComando As String, _
                                           ByVal pstrDataMensagem As String, _
                                           ByRef pblnExisteMensagemDespesa As Boolean, _
                                           ByRef pstrNumeroControleDespesa As String, _
                                           ByRef pdatDataMensagemDespesa As Date) As Boolean

Dim rsMesgDespesa                           As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_CTRL_IF,                                                                            " & vbNewLine & _
             "          DH_REGT_MESG_SPB,                                                                      " & vbNewLine & _
             "          CO_ULTI_SITU_PROC                                                                      " & vbNewLine & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB                                                               " & vbNewLine & _
             " WHERE    CO_EMPR                      =  " & plngCodigoEmpresa & "                              " & vbNewLine & _
             " AND      NU_COMD_OPER                 = '" & pstrNumeroComando & "'                             " & vbNewLine & _
             " AND      DH_RECB_ENVI_MESG_SPB        BETWEEN " & fgDtOracle_To_OracleBetween(fgDtXML_To_Oracle(Left$(pstrDataMensagem, 8))) & "  " & vbNewLine & _
             " AND      CO_MESG_SPB                  = '" & MESG_DESPESA & "'"
    
    Set rsMesgDespesa = fgQuerySQL(strSQL)
        
    pblnExisteMensagemDespesa = rsMesgDespesa.RecordCount
    If pblnExisteMensagemDespesa Then
        flVerificaMensagemDespesa = (rsMesgDespesa("CO_ULTI_SITU_PROC") = enumStatusMensagem.Encerrada)
        pstrNumeroControleDespesa = rsMesgDespesa("NU_CTRL_IF")
        pdatDataMensagemDespesa = rsMesgDespesa("DH_REGT_MESG_SPB")
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaMensagemDespesa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operações com a mensagem BMA0015 ( BMA informa Lançamentos a serem efetuados no SELIC )
'--------------------------------------------------------------------------------------------------

Public Function ConciliarBMA0015(ByVal pintTipoConciliacao As enumTipoConciliacao, _
                                 ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                 ByRef xmlMensagem As MSXML2.DOMDocument40) As String

Dim xmlErros                                As MSXML2.DOMDocument40
Dim strRet                                  As String

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlDetOperacao                          As MSXML2.DOMDocument40
Dim xmlDetMensagem                          As MSXML2.DOMDocument40
Dim xmlHistoricoMensagem                    As MSXML2.DOMDocument40
Dim xmlRepeticao                            As MSXML2.DOMDocument40

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem
Dim vntSequencialConciliacaoOperacao        As Variant
Dim blnIgnoraGradeHorario                   As Boolean

Dim blnPagtoEfetuadoContingencia            As Boolean

Dim dblQtdMensagem                          As Double
Dim dblSomaConciliacao                      As Double
Dim dblDiferencaConciliacao                 As Double
Dim intTipoRejeicao                         As enumTipoAcao
    
Dim vntSequencialConciliacao                As Variant

On Error GoTo ErrorHandler

    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")

    Call xmlDetMensagem.loadXML(xmlMensagem.xml)
    
    blnIgnoraGradeHorario = xmlMensagem.selectSingleNode("//IgnoraGradeHorario").Text = "1"
    
    strRet = flBuscarDetalhesOperacaoMensagem(pintTipoConciliacao, _
                                              xmlErros, _
                                              xmlDetOperacao, _
                                              xmlDetMensagem, _
                                              xmlOperacao, _
                                              xmlMensagem)
                                              
    'Apagar o Nó com erro de falta de operações
    If Not xmlErros.selectSingleNode("//Erro/Grupo_ErrorInfo[./Number='3064']") Is Nothing Then
        xmlErros.selectSingleNode("//Erro").removeChild _
                    xmlErros.selectSingleNode("//Erro/Grupo_ErrorInfo[./Number='3064']")
    End If
    
    If (pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaRejeitar Or _
        pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar) And _
        xmlErros.selectNodes("//Erro/Grupo_ErrorInfo[./Number!='3064']").length > 0 Then
       
        'Devolve o restante dos erros.
        ConciliarBMA0015 = xmlErros.xml
        Exit Function
    Else
        Call xmlErros.loadXML(vbNullString)
    End If
'   <Erro>
'- <Grupo_ErrorInfo>
'  <Number>3000</Number>
'  <Description>Erro não cadastrado
    
    If strRet <> vbNullString Then
        GoTo ErrorHandler
    Else
        'Verifica se os valores das operações continuam coerentes com os valores exibidos em tela.
        '=========================================================================================
        If pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaPagamentoContingencia Then
        
            dblSomaConciliacao = 0
            For Each objDomNodeOper In xmlDetOperacao.selectNodes("Repeat_DetalheOperacao/*")
                If objDomNodeOper.selectSingleNode("IN_OPER_DEBT_CRED").Text = "Débito" Then
                    dblSomaConciliacao = dblSomaConciliacao - fgVlrXml_To_Decimal(objDomNodeOper.selectSingleNode("QT_ATIV_MERC").Text)
                Else
                    dblSomaConciliacao = dblSomaConciliacao + fgVlrXml_To_Decimal(objDomNodeOper.selectSingleNode("QT_ATIV_MERC").Text)
                End If
            Next
        
            If xmlDetMensagem.selectSingleNode("//IN_OPER_DEBT_CRED").Text = "Débito" Then
                dblQtdMensagem = fgVlrXml_To_Decimal(xmlDetMensagem.selectSingleNode("//QT_ATIV_MERC").Text) * -1
            Else
                dblQtdMensagem = fgVlrXml_To_Decimal(xmlDetMensagem.selectSingleNode("//QT_ATIV_MERC").Text)
            End If
            
            dblDiferencaConciliacao = dblQtdMensagem - dblSomaConciliacao
        
            If pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar And dblDiferencaConciliacao <> 0 Then
                'Concordância não permitida. Diferença entre mensagem e operação, é diferente de Zero.
                lngCodigoErroNegocio = 3052
                GoTo ErrorHandler
        
            ElseIf pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar And dblDiferencaConciliacao = 0 Then
                'Discordância não permitida. Diferença entre mensagem e operação, é igual a Zero.
                lngCodigoErroNegocio = 3053
                GoTo ErrorHandler
                
            ElseIf pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaPagamentoContingencia And dblDiferencaConciliacao = 0 Then
                'Pagamento em contingência não permitido. Diferença entre mensagem e operação, é igual a Zero.
                lngCodigoErroNegocio = 4138
                GoTo ErrorHandler
                
            End If
            
        End If
        
        If pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaPagamentoContingencia Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaRejeitar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaRegularizar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberar Then
        
            'Verificar se a mensagem contém ação de liberação em contingência para os casos:
            '   - Pagamento em Contingência (não permitir nova libração em contingência)
            '   - Rejeição Adm Area (não permitir rejeição caso a mensagem já tenha sido paga)
            '   - Regularizar (verificar se a menasgem enconrtra-se em condição válida para regularização)
        
            Set xmlHistoricoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
            Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
            If Not xmlHistoricoMensagem.loadXML(objHistoricoMensagem.LerTodos(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                                           xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                                           xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)) Then
                'Não foi possivel obter a ação. Mensagem não possui histórico cadastrado'
                lngCodigoErroNegocio = 4134
                GoTo ErrorHandler
            End If
            
            blnPagtoEfetuadoContingencia = xmlHistoricoMensagem.selectNodes("//TP_ACAO_MESG_SPB[text()='" & enumTipoAcao.EnviadoPagamentoContingencia & "']").length > 0
            
            Set objHistoricoMensagem = Nothing
            Set xmlHistoricoMensagem = Nothing
            
        End If
        '=========================================================================================
        
        'Processamento da Mensagem Atual do For...Each
        '=========================================================================================
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
            
                vntSequencialConciliacaoOperacao = IncluirJustificativaConciliacaoOperacao()
                
                Call objMensagem.AlterarStatus( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.ConcordanciaBackoffice, _
                        xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, , , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , _
                        vntSequencialConciliacaoOperacao)
                        
            Case enumAcaoConciliacao.BODiscordar
            
                vntSequencialConciliacaoOperacao = IncluirJustificativaConciliacaoOperacao()
            
                Call objMensagem.AlterarStatus( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.DiscordanciaBackoffice, _
                        xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, , , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , _
                        vntSequencialConciliacaoOperacao)
                        
            Case enumAcaoConciliacao.AdmAreaLiberar
                If blnPagtoEfetuadoContingencia Then
                    'Pagamento não permitido. Pagamento já efetuado em contingência.
                    lngCodigoErroNegocio = 4193
                    GoTo ErrorHandler
                End If

                Call objMensagem.AlterarStatus( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.Liberada, _
                        xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, , , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Liberar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            Case enumAcaoConciliacao.AdmAreaRejeitar
                
                Call objMensagem.AlterarStatus( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.AConciliar, _
                        xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, , , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
            Case enumAcaoConciliacao.AdmAreaPagamentoContingencia
        
                If blnPagtoEfetuadoContingencia Then
                    'Pagamento em contingência não permitido. Pagamento já efetuado em contingência.
                    lngCodigoErroNegocio = 4136
                    GoTo ErrorHandler
                End If

                Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")

                'Grava ação de Pagamento em Contingencia
                Call objHistoricoMensagem.Incluir(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                  xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                  , _
                                                  enumTipoAcao.EnviadoPagamentoContingencia, , _
                                                  xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

                Set objHistoricoMensagem = Nothing

                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
                'Efetua o envio da mensagem SEL1023 para pagamento
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Liberar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

            Case enumAcaoConciliacao.AdmAreaRegularizar
                
                If Not blnPagtoEfetuadoContingencia Then
                    'Regularização não permitida pois não foi efetuado o pagamento em contingência.
                    lngCodigoErroNegocio = 4137
                    GoTo ErrorHandler
                End If
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
                If Not flAutorizaRegularizacao("BMA0015", vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, , , _
                        xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
        
        End Select
        
        'Processamento das Operações que compões a mensagem atual do For...Each
        '=========================================================================================
        For Each objDomNodeOper In xmlDetOperacao.selectNodes("Repeat_DetalheOperacao/*")
        
            Select Case pintAcaoConciliacao
                Case enumAcaoConciliacao.BOConcordar
                    
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacaoOperacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                            0, _
                            xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.ConcordanciaBackoffice, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text)
                    
                    Call objMensagem.AlterarSequenciaOperacao( _
                            xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            fgDtHrStr_To_DateTime(xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
                Case enumAcaoConciliacao.BODiscordar
                    
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacaoOperacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                            0, _
                            xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.DiscordanciaBackoffice, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text)
                    
                    Call objMensagem.AlterarSequenciaOperacao( _
                            xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            fgDtHrStr_To_DateTime(xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
                Case enumAcaoConciliacao.AdmAreaRejeitar
                    
                    If Val(objDomNodeOper.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.ConcordanciaBackoffice Then
                        intTipoRejeicao = enumTipoAcao.RejeicaoConcordancia
                    Else
                        intTipoRejeicao = enumTipoAcao.RejeicaoDiscordancia
                    End If
                
                    Call flExcluirConciliacao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Rejeitada, _
                            xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.MensagemRejeitada, _
                            intTipoRejeicao, _
                            , _
                            xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                                
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.AConciliar, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text)
                    
                Case enumAcaoConciliacao.AdmAreaLiberar
                
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liberada, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text)
                            
                Case enumAcaoConciliacao.AdmAreaPagamentoContingencia
                
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.DiscordanciaBackoffice, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, , _
                            enumTipoAcao.EnviadoPagamentoContingencia)
                
                Case enumAcaoConciliacao.AdmAreaRegularizar

                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liquidada, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text)

                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.RegularizarContingencia, _
                            blnIgnoraGradeHorario, _
                            True)

            End Select
        Next
                    
    End If

    ConciliarBMA0015 = xmlErros.xml
    
    Set xmlErros = Nothing
    Set xmlOperacao = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set xmlRepeticao = Nothing
    Set objProcessoOperacao = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlErros = Nothing
    Set xmlOperacao = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objHistoricoMensagem = Nothing
    Set xmlHistoricoMensagem = Nothing
    Set xmlRepeticao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarBMA0015 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Gerenciamento de conciliações de liquidação de operações com mensages (CETIP).
'--------------------------------------------------------------------------------------------------

Public Function GerenciarConciliacaoAutomaticaCETIP(Optional ByVal xmlOperacao As MSXML2.DOMDocument40, _
                                                    Optional ByVal xmlMensagem As MSXML2.DOMDocument40, _
                                                    Optional ByVal intNivelPerfil As Integer = 1) As Boolean

Dim strRet                                  As String
Dim intAcaoConciliacao                      As enumAcaoConciliacao
Dim intContPerfil                           As Integer

Dim xmlErros                                As MSXML2.DOMDocument40
Dim xmlDetOperacao                          As MSXML2.DOMDocument40
Dim xmlDetMensagem                          As MSXML2.DOMDocument40

Dim xmlAuxiliarProc                         As MSXML2.DOMDocument40
Dim xmlAuxiliarProc2                        As MSXML2.DOMDocument40
Dim intTipoConciliacao                      As enumTipoConciliacao
        
Dim objRegra                                As A8LQS.clsRegra
    
    On Error GoTo ErrorHandler
        
    intContPerfil = intNivelPerfil
    
VerificarPerfilAdmArea:
    GerenciarConciliacaoAutomaticaCETIP = False
    
    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAuxiliarProc = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlAuxiliarProc2 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlOperacao Is Nothing Then
        Call xmlAuxiliarProc.loadXML(xmlOperacao.xml)
    Else
        Call xmlAuxiliarProc.loadXML(xmlMensagem.xml)
    End If
    
    If Not xmlAuxiliarProc.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
        intTipoConciliacao = IIf(Val(xmlAuxiliarProc.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bruta, _
                                 enumTipoConciliacao.Bruta, _
                                 enumTipoConciliacao.Bilateral)
    Else
        intTipoConciliacao = enumTipoConciliacao.Bruta
    End If
    
    strRet = flBuscarDetalhesOperacaoMensagem(intTipoConciliacao, _
                                              xmlErros, _
                                              xmlDetOperacao, _
                                              xmlDetMensagem, _
                                              xmlOperacao, _
                                              xmlMensagem, _
                                              True)

    If strRet = vbNullString Then
        If xmlDetOperacao.xml <> vbNullString And xmlDetMensagem.xml <> vbNullString Then
            
            If Not xmlOperacao Is Nothing Then
                If intTipoConciliacao = enumTipoConciliacao.Bilateral And _
                   xmlAuxiliarProc.xml = xmlOperacao.xml Then
                   
                    Call xmlAuxiliarProc.loadXML(vbNullString)
                    Call xmlAuxiliarProc2.loadXML(xmlDetMensagem.xml)
                    
                    Call flBuscarDetalhesOperacaoMensagem(intTipoConciliacao, _
                                                          xmlErros, _
                                                          xmlDetOperacao, _
                                                          xmlDetMensagem, _
                                                          xmlAuxiliarProc, _
                                                          xmlAuxiliarProc2, _
                                                          True)
                                                              
                End If
            End If
            
            Set objRegra = CreateObject("A8LQS.clsRegra")
            
            If objRegra.VerificarRegraAutomatica2(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                  IIf(intContPerfil = 1, enumFuncaoSistema.Pagar, enumFuncaoSistema.LiberarPagamento)) Then
                
                If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0001" Or _
                   xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0003R2" Or _
                   xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0005R2" Or _
                   xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                
                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0001"
                            Select Case Val(xmlDetMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                                Case enumStatusMensagem.AConciliar
                                    If InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "CUSTO CETIP") = 0 Then
                                        intAcaoConciliacao = enumAcaoConciliacao.BOPagamento
                                        intContPerfil = 1
                                    Else
                                        GoTo Finaliza
                                    End If
                                Case enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficeAutomatico
                                    intAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia
                                    intContPerfil = 2
                                Case enumStatusMensagem.PagamentoBackoffice, enumStatusMensagem.PagamentoBackofficeAutomatico
                                    intAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento
                                    intContPerfil = 2
                            End Select
                            
                        Case "LTR0003R2", "LTR0005R2", "LTR0007"
                            Select Case Val(xmlDetMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
                                Case enumStatusMensagem.AConciliar
                                    intAcaoConciliacao = enumAcaoConciliacao.BOConcordar
                                    intContPerfil = 1
                                Case enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficeAutomatico
                                    intAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia
                                    intContPerfil = 2
                            End Select
                    
                    End Select
                            
                    strRet = ConciliarBrutaBilateral(intTipoConciliacao, _
                                                     intAcaoConciliacao, _
                                                     xmlDetMensagem, _
                                                     True, _
                                                     xmlDetOperacao)
                                                     
                    If strRet = vbNullString Then
                        GerenciarConciliacaoAutomaticaCETIP = True
                        intContPerfil = intContPerfil + 1
                        
                        If intContPerfil <= 2 Then GoTo VerificarPerfilAdmArea
                    Else
                        GerenciarConciliacaoAutomaticaCETIP = False
                    End If
                    
                End If
            End If
        End If
    End If
        
Finaliza:
    Set xmlErros = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlAuxiliarProc = Nothing
    Set xmlAuxiliarProc2 = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlErros = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlAuxiliarProc = Nothing
    Set xmlAuxiliarProc2 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarConciliacaoAutomaticaCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operações com modalidade Bruta e Bilateral.
'--------------------------------------------------------------------------------------------------

Public Function ConciliarBrutaBilateral(ByVal pintTipoConciliacao As enumTipoConciliacao, _
                                        ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                        ByRef xmlMensagem As MSXML2.DOMDocument40, _
                               Optional ByVal blnFluxoAutomatico As Boolean = False, _
                               Optional ByVal xmlOperacaoFluxoAutomatico As MSXML2.DOMDocument40, _
                               Optional ByVal blnIgnoraGradeHorario As Boolean = False) As String

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlDetOperacao                          As MSXML2.DOMDocument40
Dim xmlDetMensagem                          As MSXML2.DOMDocument40
Dim xmlDetMensagemAux                       As MSXML2.DOMDocument40
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim xmlNodeAtual                            As MSXML2.DOMDocument40
Dim xmlTipoLiberacao                        As MSXML2.DOMDocument40
Dim xmlErros                                As MSXML2.DOMDocument40

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim objProcessaMensagemBMA                  As A8LQS.clsProcessoMensagemBMA
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim objDomNodeISPB                          As MSXML2.IXMLDOMNode
Dim objDomNodeMsg                           As MSXML2.IXMLDOMNode

Dim blnAchouISPB                            As Boolean
Dim blnAcaoPermitida                        As Boolean
Dim blnEnvioLTR8                            As Boolean

Dim strMensagem                             As String
Dim vntSequencialConciliacao                As Variant
Dim strSeqOperacaoAnt                       As String
Dim strRet                                  As String
Dim strCondicao                             As String
Dim intTipoRejeicao                         As enumTipoAcao
Dim intStatusContraparteLTR0007             As Integer

Dim dblSomaOperacoes                        As Double
Dim dblSomaMensagens                        As Double
Dim dblDiferencaConciliacao                 As Double

Dim strCanalOperacaoInterna                 As String
Dim strDataHoraContraparteLTR0007           As String
    
    On Error GoTo ErrorHandler

    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetMensagemAux = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlNodeAtual = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    Set objProcessaMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    Call fgAppendNode(xmlDetMensagem, "", "Repeat_Mensagem", "")
    
    If Not blnFluxoAutomatico Then
        
        If xmlMensagem.selectSingleNode("//CONCILIACAO_FINANCEIRA_BILATERAL") Is Nothing Then
            For Each objDomNodeMsg In xmlMensagem.selectNodes("Repeat_Mensagem/*")
                
                Set xmlNodeAtual = CreateObject("MSXML2.DOMDocument.4.0")
                Call xmlNodeAtual.loadXML(objDomNodeMsg.xml)
                
                strRet = flBuscarDetalhesOperacaoMensagem(pintTipoConciliacao, _
                                                          xmlErros, _
                                                          xmlDetOperacao, _
                                                          xmlDetMensagemAux, _
                                                          xmlOperacao, _
                                                          xmlNodeAtual)
            
                Set xmlNodeAtual = Nothing
                
                If strRet <> vbNullString Then GoTo ErrorHandler
                
                'Apagar o Nó com erro de falta de operações
                If Not xmlErros.selectSingleNode("//Erro/Grupo_ErrorInfo[./Number='3064']") Is Nothing Then
                    xmlErros.selectSingleNode("//Erro").removeChild _
                                xmlErros.selectSingleNode("//Erro/Grupo_ErrorInfo[./Number='3064']")
                End If
                
                If (pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRejeitar Or _
                    pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Or _
                    pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar) And _
                    xmlErros.selectNodes("//Erro/Grupo_ErrorInfo[./Number!='3064']").length > 0 Then
                   
                    'Devolve o restante dos erros.
                    ConciliarBrutaBilateral = xmlErros.xml
                    Exit Function
                Else
                    Call xmlErros.loadXML(vbNullString)
                End If
                
                Call fgAppendXML(xmlDetMensagem, "Repeat_Mensagem", xmlDetMensagemAux.selectSingleNode("Repeat_DetalheMensagemCamara/Grupo_DetalheMensagemCamara").xml)
            
            Next
        Else
            Call xmlDetMensagem.loadXML(xmlMensagem.xml)
            Call xmlDetOperacao.loadXML(xmlMensagem.selectNodes("//Repeat_DetalheOperacao").Item(0).xml)
            
        End If
        
    Else
        Call fgAppendXML(xmlDetMensagem, "Repeat_Mensagem", xmlMensagem.selectSingleNode("Repeat_DetalheMensagemCamara/Grupo_DetalheMensagemCamara").xml)
        Call xmlDetOperacao.loadXML(xmlOperacaoFluxoAutomatico.xml)
        
    End If
    
    'Concilia valores das operações e mensagens quando o fluxo é automático
    '=========================================================================================
    If blnFluxoAutomatico Then
        If pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.BOPagamento Then
    
            dblSomaOperacoes = 0
            For Each objDomNodeOper In xmlDetOperacao.selectNodes("Repeat_DetalheOperacao/*")
                
                'No caso da Modalidade de Liquidação Multilateral / Mercado / Recebimento / Definitiva,
                'o valor da operação só entra na conciliação se o Código ISPB da operação pertencer ao rol de
                'Códigos ISPB passado pelo Form.
                If pintTipoConciliacao = enumTipoConciliacao.Bilateral And _
                   xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0005R2" Then
                    
                    blnAchouISPB = False
                    For Each objDomNodeISPB In xmlDetMensagem.selectNodes("//Grupo_DetalheMensagemCamara/*")
                        If objDomNodeISPB.baseName = "CO_ISPB_BANC_LIQU_CNPT" Then
                            If objDomNodeOper.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text = objDomNodeISPB.Text And _
                               Val(objDomNodeOper.selectSingleNode("TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bilateral Then
                                
                                blnAchouISPB = True
                                Exit For
                                
                            End If
                        End If
                    Next
                    
                Else
                    blnAchouISPB = True
                End If
                
                If blnAchouISPB Then
                    If objDomNodeOper.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                        dblSomaOperacoes = dblSomaOperacoes - fgVlrXml_To_Decimal(objDomNodeOper.selectSingleNode("VA_OPER_ATIV").Text)
                    Else
                        dblSomaOperacoes = dblSomaOperacoes + fgVlrXml_To_Decimal(objDomNodeOper.selectSingleNode("VA_OPER_ATIV").Text)
                    End If
                End If
            Next
        
            dblSomaMensagens = 0
            For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                If objDomNodeMsg.selectSingleNode("CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                    dblSomaMensagens = dblSomaMensagens - fgVlrXml_To_Decimal(objDomNodeMsg.selectSingleNode("VA_FINC").Text)
                Else
                    dblSomaMensagens = dblSomaMensagens + fgVlrXml_To_Decimal(objDomNodeMsg.selectSingleNode("VA_FINC").Text)
                End If
            Next
            
            dblDiferencaConciliacao = dblSomaMensagens - dblSomaOperacoes
        
            If (pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
                pintAcaoConciliacao = enumAcaoConciliacao.BOPagamento) And _
                dblDiferencaConciliacao <> 0 Then
                
                If Not blnFluxoAutomatico Then
                    'Concordância não permitida. Diferença entre mensagem e operação, é diferente de Zero.
                    lngCodigoErroNegocio = 3052
                    GoTo ErrorHandler
                    
                Else
                    ConciliarBrutaBilateral = "Conciliação automática não permitida."
                    Exit Function
                    
                End If
        
            ElseIf pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar And dblDiferencaConciliacao = 0 Then
                If Not blnFluxoAutomatico Then
                    'Discordância não permitida. Diferença entre mensagem e operação, é igual a Zero.
                    lngCodigoErroNegocio = 3053
                    GoTo ErrorHandler
                
                Else
                    ConciliarBrutaBilateral = "Conciliação automática não permitida."
                    Exit Function
                    
                End If
            
            End If
            
        End If
    End If
    '=========================================================================================
        
    'Verifica se a ação solicitada, pode ou não ser executada
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoBACEN Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoSTR Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRejeitar Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Or _
       pintAcaoConciliacao = enumAcaoConciliacao.BODiscordar Then
    
        For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
            strRet = objHistMensagem.LerTodos(objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                              objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                              objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
    
            Call xmlHistorico.loadXML(strRet)
            
            Select Case pintAcaoConciliacao
                Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']"
                    lngCodigoErroNegocio = 3077
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                
                    If blnAcaoPermitida Then
                        strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                        lngCodigoErroNegocio = 3084
                        blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                    End If
                
                Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Discordancia & "']"
                    lngCodigoErroNegocio = 3083
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                
                    If blnAcaoPermitida Then
                        If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                            If Val(xmlDetMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Nao Then
                                lngCodigoErroNegocio = 3088
                                blnAcaoPermitida = False
                            End If
                        End If
                    End If
                
                    If blnAcaoPermitida Then
                        strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                        lngCodigoErroNegocio = 3084
                        blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                    End If
                
                Case enumAcaoConciliacao.AdmGeralPagamento, _
                     enumAcaoConciliacao.AdmGeralPagamentoBACEN, _
                     enumAcaoConciliacao.AdmGeralPagamentoSTR, _
                     enumAcaoConciliacao.AdmGeralPagamentoContingencia
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                    lngCodigoErroNegocio = 3084
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
                    If blnAcaoPermitida Then
                        strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0004Pagamento & "' or " & _
                                      "                   ../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0003Pagamento & "' or " & _
                                      "                   ../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaSTR0004Pagamento & "']"
                        lngCodigoErroNegocio = 3131
                        blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                    End If
                
                Case enumAcaoConciliacao.AdmGeralRejeitar
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']"
                    lngCodigoErroNegocio = 3085
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
                Case enumAcaoConciliacao.AdmGeralRegularizar
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']"
                    lngCodigoErroNegocio = 3086
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
                    If blnAcaoPermitida Then
                        strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                        lngCodigoErroNegocio = 3087
                        blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length > 0
                    End If
            
                Case enumAcaoConciliacao.BODiscordar
                    blnAcaoPermitida = True
                    If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                        If Val(xmlDetMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Nao Then
                            lngCodigoErroNegocio = 3088
                            blnAcaoPermitida = False
                        End If
                    End If
            
            End Select
            
            If Not blnAcaoPermitida Then
                GoTo ErrorHandler
            End If
        Next
    
    End If
    '=========================================================================================
    
    'Define ações a serem tomadas para LTR0007
    '=========================================================================================
    blnEnvioLTR8 = False
    
    If Not xmlDetMensagem.selectSingleNode("//IN_CONF_MESG_LTR") Is Nothing Then
        If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" And _
           Val(xmlDetMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Sim Then
        
            If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
               pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
                
                If xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                    blnEnvioLTR8 = True
                Else
                    Call flVerificarContraparteLTR0007(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                       xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                                       xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text, _
                                                       intStatusContraparteLTR0007)
                End If
            
            ElseIf pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Then
                
                If xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                    Call flVerificarContraparteLTR0007(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                       xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                                       xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text, _
                                                       intStatusContraparteLTR0007, _
                                                       strDataHoraContraparteLTR0007)
                End If
            
            End If
            
        End If
    End If
    '=========================================================================================
    
    If xmlMensagem.selectSingleNode("//CONCILIACAO_FINANCEIRA_BILATERAL") Is Nothing Then
        If Not xmlDetOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing And _
           Not xmlDetMensagem.selectSingleNode("//CO_MESG_SPB") Is Nothing Then
            If Val(xmlDetOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bilateral And _
               xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0001" Then
                strCanalOperacaoInterna = "B"
            Else
                strCanalOperacaoInterna = "N"
            End If
        Else
            strCanalOperacaoInterna = "N"
        End If
    Else
        strCanalOperacaoInterna = "B"
    End If
    
    'Processamento das Operações que compõe a mensagem
    '=========================================================================================
    strSeqOperacaoAnt = vbNullString
    
    For Each objDomNodeOper In xmlDetOperacao.selectNodes("Repeat_DetalheOperacao/*")
    
        If pintTipoConciliacao = enumTipoConciliacao.Bilateral And _
           xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0005R2" Then
        
            blnAchouISPB = False
            For Each objDomNodeISPB In xmlDetMensagem.selectNodes("//Grupo_DetalheMensagemCamara/*")
                If objDomNodeISPB.baseName = "CO_ISPB_BANC_LIQU_CNPT" Then
                    If objDomNodeOper.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text = objDomNodeISPB.Text And _
                       Val(objDomNodeOper.selectSingleNode("TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bilateral Then
                        
                        blnAchouISPB = True
                        Exit For
                        
                    End If
                End If
            Next
            
        Else
            blnAchouISPB = True
        End If
        
        If blnAchouISPB Then
        
            Call objOperacao.Alterar(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, , , , , , _
                                     strCanalOperacaoInterna)
                    
            Select Case pintAcaoConciliacao
                Case enumAcaoConciliacao.BOConcordar
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            IIf(Not blnFluxoAutomatico, enumStatusOperacao.ConcordanciaBackoffice, enumStatusOperacao.ConcordanciaBackofficeAutomatico), _
                            vbNullString, , , _
                            False, , , _
                            vntSequencialConciliacao)
                            
                Case enumAcaoConciliacao.BODiscordar
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                        
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.DiscordanciaBackoffice, _
                            vbNullString, , , _
                            False, , , _
                            vntSequencialConciliacao)
                    
                Case enumAcaoConciliacao.AdmGeralRejeitar
                    If Val(objDomNodeOper.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.DiscordanciaBackoffice Then
                        intTipoRejeicao = enumTipoAcao.RejeicaoDiscordancia
                    Else
                        intTipoRejeicao = enumTipoAcao.RejeicaoConcordancia
                    End If
                
                    Call flExcluirConciliacao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Rejeitada, _
                            , _
                            , _
                            , _
                            intTipoRejeicao)
                                
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Registrada, _
                            vbNullString, , , _
                            False)
                    
                Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0007"
                            If Val(xmlDetMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Sim Then
                               
                                If blnEnvioLTR8 Then
                                    Call objOperacao.AlterarStatusAcao( _
                                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            IIf(Not blnFluxoAutomatico, enumStatusOperacao.ConcordanciaLib, enumStatusOperacao.ConcordanciaLibAuto), _
                                            vbNullString, , , _
                                            False, , , _
                                            vntSequencialConciliacao)
                                Else
                                    If intStatusContraparteLTR0007 = enumStatusMensagem.MensagemLiquidada Then
                                        If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                            strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                            Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                                        End If
                                        
                                        Call objProcessaMensagemBMA.GerenciarChamada( _
                                                xmlOperacao, _
                                                Nothing, _
                                                enumStatusOperacao.Liquidada, _
                                                0, _
                                                "", _
                                                enumIndicadorSimNao.Nao, _
                                                "", _
                                                False, _
                                                True)
                                    
                                    Else
                                        Call objOperacao.AlterarStatusAcao( _
                                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                IIf(Not blnFluxoAutomatico, enumStatusOperacao.Liberada, enumStatusOperacao.LiberadaAutomatica), _
                                                vbNullString, , , _
                                                False, , , _
                                                vntSequencialConciliacao)
                                    End If
                                End If
                                    
                            Else
                                        
                                Call objOperacao.AlterarStatusAcao( _
                                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                        enumStatusOperacao.Liquidada, _
                                        vbNullString, , , _
                                        False, , , _
                                        vntSequencialConciliacao)
                                        
                                If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                    strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                    Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                                End If
                                
                                Call objFluxoProcessamento.DisponibilizarLancamentoCC(xmlOperacao)
                                Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
                                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, _
                                                                                     enumTipoMovimento.RealizadoConfirmado)
                                            
                            End If
                    
                        Case "LTR0003R2", "LTR0005R2"
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumStatusOperacao.Liquidada, _
                                    vbNullString, , , _
                                    False, , , _
                                    vntSequencialConciliacao)
                                    
                            If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                            End If
                            
                            Call objFluxoProcessamento.DisponibilizarLancamentoCC(xmlOperacao)
                            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
                            
                            If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0005R2" Then
                                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, _
                                                                                     enumTipoMovimento.RealizadoConfirmado)
                            End If
                    End Select
                    
                Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0007", "LTR0001"
                               
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumStatusOperacao.DiscordanciaLib, _
                                    vbNullString, , , _
                                    False, , , _
                                    vntSequencialConciliacao)
                                        
                    End Select
                    
                Case enumAcaoConciliacao.AdmGeralPagamento, _
                     enumAcaoConciliacao.AdmGeralPagamentoSTR, _
                     enumAcaoConciliacao.AdmGeralPagamentoBACEN
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0001"
            
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    IIf(Not blnFluxoAutomatico, enumStatusOperacao.PagamentoLib, enumStatusOperacao.PagamentoLiberadoAutomatico), _
                                    vbNullString, , , _
                                    False, , , _
                                    vntSequencialConciliacao)
                                        
                            strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                    
                    End Select
                    
                Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0001"
            
                            If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                            End If
                            
                            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, _
                                                                                 enumTipoMovimento.RealizadoSolicitado)
                                        
                    End Select
                    
                Case enumAcaoConciliacao.BOPagamento
                    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                        , _
                                                        , _
                                                        objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                        blnFluxoAutomatico, _
                                                        False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                    Next

                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0001"
            
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    IIf(Not blnFluxoAutomatico, enumStatusOperacao.PagamentoBackoffice, enumStatusOperacao.PagamentoBackofficeAutomatico), _
                                    vbNullString, , , _
                                    False, , , _
                                    vntSequencialConciliacao)
                                    
                    End Select
                    
                Case enumAcaoConciliacao.AdmGeralRegularizar
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liquidada, _
                            vbNullString, , , _
                            False)
                            
                    Select Case xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text
                        Case "LTR0001"
            
                            If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                            End If
                            
                            Call objFluxoProcessamento.DisponibilizarLancamentoCC(xmlOperacao)
                            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
                            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.RealizadoConfirmado)
                                        
                    End Select
                    
            End Select
            
        End If
                
    Next

    'Processamento das Mensagens
    '========================================================================
    For Each objDomNodeMsg In xmlDetMensagem.selectNodes("Repeat_Mensagem/*")
        
        Call objMensagem.AlterarTabelaConciliacao( _
                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                strCanalOperacaoInterna)
                
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        IIf(Not blnFluxoAutomatico, enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficeAutomatico), _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                        Not blnFluxoAutomatico)
                        
                Call objMensagem.AlterarSequenciaConciliacao( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                        vntSequencialConciliacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        
            Case enumAcaoConciliacao.BODiscordar
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.DiscordanciaBackoffice, _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                        Not blnFluxoAutomatico)
                        
                If IsEmpty(vntSequencialConciliacao) Then
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                    objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                    objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                    objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                                    , _
                                                    , _
                                                    objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, _
                                                    blnFluxoAutomatico, _
                                                    False)
                End If
                
                Call objMensagem.AlterarSequenciaConciliacao( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                        vntSequencialConciliacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        
            Case enumAcaoConciliacao.AdmGeralRejeitar
                Call flExcluirConciliacao( _
                        , _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemRejeitada, _
                        intTipoRejeicao, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.AConciliar, _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                        Not blnFluxoAutomatico)
                        
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                Select Case objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text
                    Case "LTR0001"
                        Call objHistMensagem.Incluir( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                Val(objDomNodeMsg.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                                enumTipoAcao.EnviadaLTR0002Concordancia, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                        
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                        
                        Call objMensagem.AlterarSequenciaConciliacao( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                vntSequencialConciliacao, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                enumTipoAcao.EnviadaLTR0002Concordancia)
                
                        Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                        Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                        
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                enumFuncaoSistema.Concordar, _
                                blnIgnoraGradeHorario, _
                                True, _
                                True, _
                                xmlTipoLiberacao, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                    
                        Set xmlTipoLiberacao = Nothing
                        
                    Case "LTR0007"
                        If Val(objDomNodeMsg.selectSingleNode("IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Sim Then
                            If blnEnvioLTR8 Then
                                Call objMensagem.AlterarStatus( _
                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                        IIf(Not blnFluxoAutomatico, enumStatusMensagem.ConcordanciaLib, enumStatusMensagem.ConcordanciaLibAuto), _
                                        vbNullString, _
                                        , _
                                        False, _
                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                        Not blnFluxoAutomatico)
                                
                                Call objHistMensagem.Incluir( _
                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                        Val(objDomNodeMsg.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                                        enumTipoAcao.EnviadaLTR0008Concordancia, _
                                        , _
                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                                
                                Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                                Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                                
                                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                        enumFuncaoSistema.Concordar, _
                                        blnIgnoraGradeHorario, _
                                        True, _
                                        True, _
                                        xmlTipoLiberacao, _
                                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , _
                                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text)
                            
                                Set xmlTipoLiberacao = Nothing
                            
                            Else
                                If intStatusContraparteLTR0007 = enumStatusMensagem.MensagemLiquidada Then
                                    Call objMensagem.AlterarStatus( _
                                            objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                            objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                            enumStatusMensagem.MensagemLiquidada, _
                                            vbNullString, _
                                            , _
                                            False, _
                                            objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                            Not blnFluxoAutomatico)
                                Else
                                    Call objMensagem.AlterarStatus( _
                                            objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                            objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                            enumStatusMensagem.Liberada, _
                                            vbNullString, _
                                            , _
                                            False, _
                                            objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                            Not blnFluxoAutomatico)
                                End If
                                
                            End If
                        
                        Else
                            Call objMensagem.AlterarStatus( _
                                    objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                    objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                    enumStatusMensagem.MensagemLiquidada, _
                                    vbNullString, _
                                    , _
                                    False, _
                                    objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                    Not blnFluxoAutomatico)
                                
                        End If
                        
                    Case "LTR0003R2", "LTR0005R2"
                        Call objMensagem.AlterarStatus( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                enumStatusMensagem.MensagemLiquidada, _
                                vbNullString, _
                                , _
                                False, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                Not blnFluxoAutomatico)
                                
                End Select
        
            Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                Select Case objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text
                    Case "LTR0001"
                        Call objMensagem.AlterarStatus( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                enumStatusMensagem.DiscordanciaLib, _
                                vbNullString, _
                                , _
                                False, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                Not blnFluxoAutomatico)
                                
                        Call objHistMensagem.Incluir( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                Val(objDomNodeMsg.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                                enumTipoAcao.EnviadaLTR0002Discordancia, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                        
                        vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                False)
                
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                , _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                True)
                        
                        Call objMensagem.AlterarSequenciaConciliacao( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                vntSequencialConciliacao, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                enumTipoAcao.EnviadaLTR0002Discordancia)
                
                        Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                        Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "D")
                        
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                enumFuncaoSistema.Discordar, _
                                blnIgnoraGradeHorario, _
                                True, _
                                True, _
                                xmlTipoLiberacao, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                                
                        Set xmlTipoLiberacao = Nothing
                        
                    Case "LTR0007"
                        Call objMensagem.AlterarStatus( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                enumStatusMensagem.DiscordanciaLib, _
                                vbNullString, _
                                , _
                                False, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , , , , _
                                Not blnFluxoAutomatico)
                                
                        If blnEnvioLTR8 Then
                            Call objHistMensagem.Incluir( _
                                    objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                    objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                    Val(objDomNodeMsg.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                                    enumTipoAcao.EnviadaLTR0008Discordancia, _
                                    , _
                                    objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                            
                            Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                            Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "D")
                            
                            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                    enumFuncaoSistema.Discordar, _
                                    blnIgnoraGradeHorario, _
                                    True, _
                                    True, _
                                    xmlTipoLiberacao, _
                                    objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , _
                                    objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text)
                        
                            Set xmlTipoLiberacao = Nothing
                        End If
                            
                End Select
        
            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                Call objHistMensagem.Incluir( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeMsg.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoPagamentoContingencia, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        True)
        
                Call objMensagem.AlterarSequenciaConciliacao( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                        vntSequencialConciliacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        enumTipoAcao.EnviadoPagamentoContingencia)
                
                
                If strSeqOperacaoAnt <> vbNullString Then
                    Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                    Call fgAppendNode(xmlOperacao, "MESG", "TP_CONF_DIVG", "C")
                End If
                        
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.PagarContingencia, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        xmlOperacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text)
                    
            Case enumAcaoConciliacao.AdmGeralPagamento
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , _
                        blnFluxoAutomatico, _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        True)
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        IIf(Not blnFluxoAutomatico, enumStatusMensagem.PagamentoLib, enumStatusMensagem.PagamentoLibAutomatico), _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , _
                        enumTipoAcao.EnviadaLTR0004Pagamento, _
                        vntSequencialConciliacao, _
                        True, , , , _
                        Not blnFluxoAutomatico)
                        
                Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Pagar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        xmlOperacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text)
    
            Case enumAcaoConciliacao.AdmGeralPagamentoSTR
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , _
                        blnFluxoAutomatico, _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        True)
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        IIf(Not blnFluxoAutomatico, enumStatusMensagem.PagamentoLib, enumStatusMensagem.PagamentoLibAutomatico), _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , _
                        enumTipoAcao.EnviadaSTR0004Pagamento, _
                        vntSequencialConciliacao, _
                        True, , , , _
                        Not blnFluxoAutomatico)
                        
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.PagarSTR, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        xmlMensagem, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
    
            Case enumAcaoConciliacao.AdmGeralPagamentoBACEN
                If objDomNodeMsg.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text = vbNullString Then
                    lngCodigoErroNegocio = 3089
                    GoTo ErrorHandler
                End If
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , _
                        blnFluxoAutomatico, _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        True)
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        IIf(Not blnFluxoAutomatico, enumStatusMensagem.PagamentoLib, enumStatusMensagem.PagamentoLibAutomatico), _
                        vbNullString, _
                        , _
                        False, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , _
                        enumTipoAcao.EnviadaLTR0003Pagamento, _
                        vntSequencialConciliacao, _
                        True, , , , _
                        Not blnFluxoAutomatico)
                        
                Set xmlOperacao = objOperacao.LerXMLOperacao(strSeqOperacaoAnt)
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.PagarBACEN, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        xmlOperacao, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                        
            Case enumAcaoConciliacao.BOPagamento
                Select Case objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text
                    Case "LTR0001"
        
                        Call objMensagem.AlterarStatus( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                IIf(Not blnFluxoAutomatico, enumStatusMensagem.PagamentoBackoffice, enumStatusMensagem.PagamentoBackofficeAutomatico), _
                                vbNullString, , _
                                False, , , , , , , , , _
                                Not blnFluxoAutomatico)
                                
                        Call objMensagem.AlterarSequenciaConciliacao( _
                                objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                fgDtHrStr_To_DateTime(objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                vntSequencialConciliacao, _
                                objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                
                End Select
                
            Case enumAcaoConciliacao.AdmGeralRegularizar
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeMsg.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                If Not flAutorizaRegularizacao(objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text, vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeMsg.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        vbNullString, , _
                        False, , , , , , , , , _
                        Not blnFluxoAutomatico)
        
                Select Case objDomNodeMsg.selectSingleNode("CO_MESG_SPB").Text
                    Case "LTR0007"
                        
                        If intStatusContraparteLTR0007 = enumStatusMensagem.Liberada Then
                            Call objMensagem.AlterarStatus( _
                                    objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                    strDataHoraContraparteLTR0007, _
                                    enumStatusMensagem.MensagemLiquidada, _
                                    vbNullString, , _
                                    False, , , , , , , , , _
                                    Not blnFluxoAutomatico)
                        
                            Call xmlOperacao.loadXML(ConsultaConciliacao(, _
                                                                         objDomNodeMsg.selectSingleNode("NU_CTRL_IF").Text, _
                                                                         strDataHoraContraparteLTR0007))
                                                                         
                            strSeqOperacaoAnt = vbNullString
                            
                            For Each objDomNodeOper In xmlOperacao.selectNodes("Repeat_ConsultaConciliacao/*")
                            
                                If strSeqOperacaoAnt <> objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text Then
                                    strSeqOperacaoAnt = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                                    
                                    Call objOperacao.AlterarStatusAcao( _
                                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            enumStatusOperacao.Liquidada, _
                                            vbNullString, , , _
                                            False)
                                                
                                End If
                            
                            Next
                            
                        End If
        
                End Select
                
        End Select
    Next
    '========================================================================
    
    If xmlErros.selectNodes("/Grupo/*").length > 0 Then
        ConciliarBrutaBilateral = xmlErros.xml
    Else
        ConciliarBrutaBilateral = ""
        
        If Not xmlMensagem.selectSingleNode("//CONCILIACAO_FINANCEIRA_BILATERAL") Is Nothing Then
            If Not xmlMensagem.selectSingleNode("//Grupo_Envio") Is Nothing Then
                Call fgAppendNode(xmlMensagem, "Grupo_Envio", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Bilateral)
            End If
        End If
            
        If (pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
            pintAcaoConciliacao = enumAcaoConciliacao.BOPagamento) And _
            Not blnFluxoAutomatico Then

            Call GerenciarConciliacaoAutomaticaCETIP(, xmlDetMensagem, 2)

        End If
    End If
    
    Set xmlOperacao = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlDetMensagemAux = Nothing
    Set xmlHistorico = Nothing
    Set xmlNodeAtual = Nothing
    Set xmlTipoLiberacao = Nothing
    Set xmlErros = Nothing
    
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing
    Set objProcessaMensagemBMA = Nothing
    Set objFluxoProcessamento = Nothing

    Exit Function

ErrorHandler:
    Set xmlOperacao = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlDetMensagemAux = Nothing
    Set xmlHistorico = Nothing
    Set xmlNodeAtual = Nothing
    Set xmlTipoLiberacao = Nothing
    Set xmlErros = Nothing
    
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing
    Set objProcessaMensagemBMA = Nothing
    Set objFluxoProcessamento = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarBrutaBilateral Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de liquidação multilateral CBLC
'--------------------------------------------------------------------------------------------------

Public Function LiquidarMultilateralCBLC(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                         ByRef objDomNodeProc As MSXML2.IXMLDOMNode, _
                                         ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem

Dim xmlTipoLiberacao                        As MSXML2.DOMDocument40
Dim xmlHistorico                            As MSXML2.DOMDocument40

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim vntSequencialConciliacao                As Variant

Dim strCondicao                             As String
Dim blnAcaoPermitida                        As Boolean

    On Error GoTo ErrorHandler

    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Verifica se a ação solicitada, pode ou não ser executada
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizarRecebimento Then
    
        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                           objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text))
        
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']"
                lngCodigoErroNegocio = 3077
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
                If blnAcaoPermitida Then
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                    lngCodigoErroNegocio = 3084
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                End If
            
            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
        
            Case enumAcaoConciliacao.AdmGeralRegularizar
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']"
                lngCodigoErroNegocio = 3086
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
        
                If blnAcaoPermitida Then
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                    lngCodigoErroNegocio = 3087
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length > 0
                End If
            
            Case enumAcaoConciliacao.AdmGeralPagamento
                
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
            Case enumAcaoConciliacao.AdmGeralRegularizarRecebimento
                
                ' Verificar se existe LDL0005R2
                
                If Not flExisteLDL0005R2(objDomNodeProc) Then
                    Err.Raise "3086", "LiquidacaoCBCL", "LDL0005R2 não existe."
                Else
                    blnAcaoPermitida = True
                End If
            
            
        End Select
        
        If Not blnAcaoPermitida Then
            GoTo ErrorHandler
        End If
    
    End If
    
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    '=========================================================================================
    
    'Processamento das Operações
    '=========================================================================================
    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaBackoffice, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.ConcordanciaPrevia, _
                        blnIgnoraGradeHorario, _
                        True)
    
            Case enumAcaoConciliacao.AdmAreaLiberar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaAdmArea, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.Concordancia, _
                        blnIgnoraGradeHorario, _
                        True)
    
            Case enumAcaoConciliacao.AdmGeralRejeitar
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaBackoffice, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.RejeicaoConcordancia, _
                        blnIgnoraGradeHorario, _
                        True)
    
            Case enumAcaoConciliacao.AdmGeralPagamento
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.PagamentoLib, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

            Case enumAcaoConciliacao.AdmGeralRegularizar
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Liquidada, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.LiberacaoPagamento, _
                        blnIgnoraGradeHorario, _
                        True)
    
            Case enumAcaoConciliacao.AdmGeralRecebimento
                                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.RecebimentoLib, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)
        
            Case enumAcaoConciliacao.AdmGeralRegularizarRecebimento
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Liquidada, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.RegularizarContingencia, _
                        blnIgnoraGradeHorario, _
                        True)


        End Select
    Next

    'Processamento da mensagem
    '=========================================================================================
    If Not objDomNodeProc.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        
        ' KIDA - CBLC - 23/09/2008
        If pintAcaoConciliacao <> enumAcaoConciliacao.AdmGeralEnviarConcordancia And _
           pintAcaoConciliacao <> enumAcaoConciliacao.AdmGeralRegularizarRecebimento Then
        
            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                                           objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                                           objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                                                                           False)
    
            For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                Call IncluirConciliacaoOperacao(vntSequencialConciliacao, _
                                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                0)
            Next
    
            Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                        fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                                        vntSequencialConciliacao, _
                                                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        End If
        
        Select Case pintAcaoConciliacao
            
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadaLDL0003Concordancia, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                Call fgAppendNode(xmlTipoLiberacao, "", "CO_LOCA_LIQU", enumLocalLiquidacao.CLBCAcoes)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Concordar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        xmlTipoLiberacao, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoPagamentoContingencia, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.PagarContingencia, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

            Case enumAcaoConciliacao.AdmGeralPagamento
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoPagamento, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Pagar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.PagamentoLib, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.EnviadoPagamento)

            Case enumAcaoConciliacao.AdmGeralRecebimento
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoRecebimento, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)

                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.RecebimentoLib, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.EnviadoPagamento)

            Case enumAcaoConciliacao.AdmGeralRegularizar
            
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                If Not flAutorizaRegularizacao("LDL0001", vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.Liquidacao)

            Case enumAcaoConciliacao.AdmGeralRegularizar
            
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                If Not flAutorizaRegularizacao("LDL0001", vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.Liquidacao)

            Case enumAcaoConciliacao.AdmGeralRegularizarRecebimento
                
                'Liquidar LDL0001 ---------------------
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        "", _
                        False, _
                        0)
                        
                'Liquidar LDL0005R2 E OPERAÇÕES ---------------------
                
                Call flRegularizarRecebimentoMultiCBLC(objDomNodeProc)
                
                

        End Select

    End If

    LiquidarMultilateralCBLC = vbNullString

    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlTipoLiberacao = Nothing
    Set xmlHistorico = Nothing
    
    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlTipoLiberacao = Nothing
    Set xmlHistorico = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarMultilateralCBLC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



'--------------------------------------------------------------------------------------------------
'Processo de liquidação multilateral - BMF Derivativa
'--------------------------------------------------------------------------------------------------

Public Function LiquidarMultilateralBMF(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                        ByRef objDomNodeProc As MSXML2.IXMLDOMNode, _
                                        ByVal xmlRepeticao As MSXML2.DOMDocument40, _
                                        ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim objJustificativaArea                    As A8LQS.clsJustificativaArea

Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim xmlMsgEnvio                             As MSXML2.DOMDocument40
Dim xmlRemessaOriginal                      As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim xmlTipoLiberacao                        As MSXML2.DOMDocument40

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim vntSequencialConciliacao                As Variant
Dim vntSequencialOperacao                   As Variant

Dim strCondicao                             As String
Dim blnAcaoPermitida                        As Boolean
Dim blnExisteLTR0005R2                      As Boolean
Dim blnLeuLTR0005R2                         As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    Set objJustificativaArea = CreateObject("A8LQS.clsJustificativaArea")

    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Verifica se a ação solicitada, pode ou não ser executada
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarDiscordancia Then
    
        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                           objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           "1"))
        
        blnAcaoPermitida = True
        
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmGeralPagamento, _
                 enumAcaoConciliacao.AdmGeralPagamentoContingencia
                
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
        
            Case enumAcaoConciliacao.AdmGeralRegularizar
                 
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3087
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length > 0
                
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                 enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = Not xmlHistorico.selectNodes(strCondicao).length > 0
        
        End Select
        
        If Not blnAcaoPermitida Then
            GoTo ErrorHandler
        End If
    
    End If
    '=========================================================================================
    
    blnLeuLTR0005R2 = False
    blnExisteLTR0005R2 = False
    
    'Processamento das Operações
    '=========================================================================================
    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
        
        vntSequencialOperacao = objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text
        
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaBackoffice, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.ConcordanciaPrevia, _
                        blnIgnoraGradeHorario, _
                        True)
    '=========================================================================================
            Case enumAcaoConciliacao.AdmAreaLiberar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaAdmArea, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Set xmlRemessaOriginal = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.Concordancia, _
                        blnIgnoraGradeHorario, _
                        True, _
                        , _
                        xmlRemessaOriginal)
    
                If Val(xmlRemessaOriginal.selectSingleNode("//CO_FORM_LIQU").Text) = enumFormaLiquidacao.Str Then
                    Set xmlMsgEnvio = CreateObject("MSXML2.DOMDocument.4.0")
    
                    Call xmlMsgEnvio.loadXML(flMontarMsgEnvioSTR0004(objDomNodeProc))
    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            0, _
                            enumFuncionalidade.LiberacaoMensagemSTR, _
                            blnIgnoraGradeHorario, _
                            False, _
                            False, _
                            xmlMsgEnvio)
    
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                    xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                                                    xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text, _
                                                    1)
    
                    Call objMensagem.AlterarSequenciaConciliacao(xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                                           fgDtHrStr_To_DateTime(xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text), _
                                           vntSequencialConciliacao, _
                                           1)
    
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                            xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text, _
                            0)
                    
                End If
    
                Set xmlRemessaOriginal = Nothing
                Set xmlMsgEnvio = Nothing
    '=========================================================================================
            Case enumAcaoConciliacao.AdmGeralRejeitar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.ConcordanciaBackoffice, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.RejeicaoConcordancia, _
                        blnIgnoraGradeHorario, _
                        True)
        
                Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
                
                Call xmlLeitura.loadXML(ConsultaConciliacao(CLng(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text), _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            vbNullString, _
                                                            0, _
                                                            0))
                
                If xmlLeitura.xml <> vbNullString Then
                    If xmlLeitura.selectSingleNode("//CO_MESG_SPB").Text = "STR0004" Then
                        
                        Call objMensagem.AlterarStatus( _
                                xmlLeitura.selectSingleNode("//NU_CTRL_IF").Text, _
                                xmlLeitura.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                enumStatusMensagem.MensagemCancelada, _
                                xmlLeitura.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                , _
                                False)
                                
                    End If
                End If
                
                Set xmlLeitura = Nothing
    '=========================================================================================
            Case enumAcaoConciliacao.AdmGeralPagamento
                                
                If objDomNodeProc.selectSingleNode("TP_INFO_LDL").Text = "P" Then
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.ConcordanciaAdmArea, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            enumTipoAcao.PreviaLiquidada, _
                            True)
                Else
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.PagamentoLib, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            0, _
                            True)
                
                
                End If
    '=========================================================================================
            Case enumAcaoConciliacao.AdmGeralRecebimento
                
                If objDomNodeProc.selectSingleNode("TP_INFO_LDL").Text = "P" Then
                    
                    'KIDA - 24/08/2009
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.ConcordanciaAdmArea, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            0, _
                            enumTipoAcao.PreviaLiquidada, _
                            True)
                
                    
                    If Trim(objDomNodeProc.selectSingleNode("NU_CTRL_IF_LDL1001").Text) <> vbNullString Then
                        'Alterar Situação LDL1001
                        Call objMensagem.AlterarStatus( _
                                  objDomNodeProc.selectSingleNode("NU_CTRL_IF_LDL1001").Text, _
                                  objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB_LDL1001").Text, _
                                  enumStatusMensagem.Conciliada, _
                                  "", "", False, 0, "", enumTipoAcao.EnviadoPagamento)
                    End If
                    
                Else
                
                    If fgVlrXml_To_Decimal(objDomNodeProc.selectSingleNode("VA_FINC").Text) = 0 Then
                    
                        Call objOperacao.AlterarStatusAcao( _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                enumStatusOperacao.Liquidada, _
                                objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                , _
                                enumTipoAcao.Liquidacao, _
                                True)
                                
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                enumFuncionalidade.LiberacaoRecebimento, _
                                blnIgnoraGradeHorario, _
                                True)
                                
                    Else
                
                        If Not blnLeuLTR0005R2 Then
                            blnLeuLTR0005R2 = True
                            
                            If objDomNodeProc.selectSingleNode("NU_CTRL_CAMR") Is Nothing Then
                                strCondicao = ""
                            Else
                                strCondicao = "" 'objDomNodeProc.selectSingleNode("NU_CTRL_CAMR").Text
                            End If
                            
                            Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
                            
                            'KIDA 11/12/2009
                            'RATS 696
                            Call xmlLeitura.loadXML(objMensagem.LerTodos("", _
                                                                         0, _
                                                                         strCondicao, _
                                                                         "LDL0005R2", _
                                                                         enumLocalLiquidacao.BMD, _
                                                                         Mid(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, 1, 8)))
                        
                            blnExisteLTR0005R2 = IIf(xmlLeitura.xml <> vbNullString, True, False)
                            'Set xmlLeitura = Nothing
                        End If
                        
                        If blnExisteLTR0005R2 Then
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumStatusOperacao.Liquidada, _
                                    objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    enumTipoAcao.Liquidacao, _
                                    True)
                                    
                            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumFuncionalidade.LiberacaoRecebimento, _
                                    blnIgnoraGradeHorario, _
                                    True)
                    
                        Else
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumStatusOperacao.RecebimentoLib, _
                                    objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    enumTipoAcao.EnviadoRecebimento, _
                                    True)
                                    
                        End If
                        
                    End If

                End If
    '=========================================================================================
            Case enumAcaoConciliacao.AdmGeralRegularizar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Liquidada, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.LiberacaoPagamento, _
                        blnIgnoraGradeHorario, _
                        True)

        End Select
    Next

    'Inclusão de justificativa por área no caso de divergência entre valores
    '=========================================================================================
    If Not objDomNodeProc.selectSingleNode("DE_JUST") Is Nothing Then
        If objJustificativaArea.LerTodos(objDomNodeProc.selectSingleNode("TP_BKOF").Text, _
                                         objDomNodeProc.selectSingleNode("CO_EMPR").Text, _
                                         objDomNodeProc.selectSingleNode("DT_LIQU").Text) = vbNullString Then

            Call objJustificativaArea.Incluir(objDomNodeProc.selectSingleNode("TP_BKOF").Text, _
                                              objDomNodeProc.selectSingleNode("CO_EMPR").Text, _
                                              objDomNodeProc.selectSingleNode("DT_LIQU").Text, _
                                              objDomNodeProc.selectSingleNode("DE_JUST").Text)
                                              
        Else
        
            Call objJustificativaArea.Alterar(objDomNodeProc.selectSingleNode("TP_BKOF").Text, _
                                              objDomNodeProc.selectSingleNode("CO_EMPR").Text, _
                                              objDomNodeProc.selectSingleNode("DT_LIQU").Text, _
                                              objDomNodeProc.selectSingleNode("DE_JUST").Text)
        
        End If
    End If
    
    'Processamento da mensagem
    '=========================================================================================
    If Not objDomNodeProc.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        
        Select Case pintAcaoConciliacao
        '=========================================================================================
            Case enumAcaoConciliacao.AdmGeralRecebimento
                
                If objDomNodeProc.selectSingleNode("TP_INFO_LDL").Text = "P" Then
                    
                    Call objMensagem.AlterarStatus( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.MensagemLiquidada, _
                            objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            "", _
                            False, _
                            0, _
                            "", _
                            enumTipoAcao.Liquidacao)
                    
                    'KIDA - 24/08/2009
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                    1)
            
                    Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                           fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                           vntSequencialConciliacao, _
                                           0, 1)
                                           
                    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, 1)
                    Next
        '=========================================================================================
                            
                            
                Else
                
                    If fgVlrXml_To_Decimal(objDomNodeProc.selectSingleNode("VA_FINC").Text) = 0 Then
                        Call objMensagem.AlterarStatus( _
                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                enumStatusMensagem.MensagemLiquidada, _
                                objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                , _
                                False, _
                                0, _
                                , _
                                enumTipoAcao.Liquidacao)
                                
                    Else
                        
                        If objDomNodeProc.selectSingleNode("IN_ULTI_MESG").Text = "S" Then
                            If blnExisteLTR0005R2 Then
                                Call objMensagem.AlterarStatus( _
                                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                        enumStatusMensagem.MensagemLiquidada, _
                                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                        , _
                                        False, _
                                        1, _
                                        , _
                                        enumTipoAcao.Liquidacao)
                                
                                'LIQUIDAR LDL0005R2
                                Call objMensagem.AlterarStatus( _
                                        xmlLeitura.selectSingleNode("//NU_CTRL_IF").Text, _
                                        xmlLeitura.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                        enumStatusMensagem.MensagemLiquidada, _
                                        xmlLeitura.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                        , _
                                        False, _
                                        0, _
                                        , _
                                        enumTipoAcao.Liquidacao)
                                        
                            Else
                                Call objMensagem.AlterarStatus( _
                                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                        enumStatusMensagem.RecebimentoLib, _
                                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                        , _
                                        False, _
                                        1, _
                                        , _
                                        enumTipoAcao.EnviadoRecebimento)
                                    
                            End If
                        End If
                        
                        If blnExisteLTR0005R2 Then
                            Call objMensagem.AlterarStatus( _
                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                    enumStatusMensagem.MensagemLiquidada, _
                                    objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    False, _
                                    objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                    , _
                                    enumTipoAcao.Liquidacao)
                                                                    
                                                                    
                        Else
                            Call objMensagem.AlterarStatus( _
                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                    enumStatusMensagem.RecebimentoLib, _
                                    objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    False, _
                                    objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                                    , _
                                    enumTipoAcao.EnviadoRecebimento)
                                
                        End If
                        
                    End If
                
                End If
'---------------------------------------------------------------------------------------------------
            Case enumAcaoConciliacao.AdmGeralPagamento
                
                If fgVlrXml_To_Decimal(objDomNodeProc.selectSingleNode("VA_FINC").Text) <> 0 Then
                    Call objMensagem.AlterarStatus( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.PagamentoLib, _
                            objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            False, _
                            0, _
                            , _
                            enumTipoAcao.EnviadoPagamento)
                Else
                    Call objMensagem.AlterarStatus( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.MensagemLiquidada, _
                            objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            False, _
                            0, _
                            , _
                            enumTipoAcao.Liquidacao)
                
                
                End If
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                1)
                
                Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                       fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                       vntSequencialConciliacao, _
                                       1)
                
                If objDomNodeProc.selectSingleNode("TP_INFO_LDL").Text <> "P" Then
                    
                    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                        Call IncluirConciliacaoOperacao( _
                                vntSequencialConciliacao, _
                                objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                0, 1)
                    Next
                Else
                    If Trim(objDomNodeProc.selectSingleNode("NU_CTRL_IF_LDL1001").Text) <> vbNullString Then
                        'Alterar Situação LDL1001
                        Call objMensagem.AlterarStatus( _
                                  objDomNodeProc.selectSingleNode("NU_CTRL_IF_LDL1001").Text, _
                                  objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB_LDL1001").Text, _
                                  enumStatusMensagem.Conciliada, _
                                  "", "", False, 0, "", enumTipoAcao.EnviadoPagamento)
                    End If
                End If
                
                If objDomNodeProc.selectSingleNode("IN_ULTI_MESG").Text = "S" Then
                    
                    If fgVlrXml_To_Decimal(objDomNodeProc.selectSingleNode("VA_FINC").Text) <> 0 Then
                        If objDomNodeProc.selectSingleNode("VA_OPER_MESG_BATIDOS").Text = "S" Then
                        
                            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                    enumFuncaoSistema.Pagar, _
                                    blnIgnoraGradeHorario, _
                                    True, _
                                    True, _
                                    xmlRepeticao, _
                                    "1")
                        Else
                            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                    enumFuncaoSistema.PagarContingencia, _
                                    blnIgnoraGradeHorario, _
                                    True, _
                                    True, _
                                    xmlRepeticao, _
                                    "1")
                        
                        
                        End If
                    End If
                End If
                
'---------------------------------------------------------------------------------------------------
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                
                If objDomNodeProc.selectSingleNode("IN_ULTI_MESG").Text = "S" Then
                    
                    Call objHistMensagem.Incluir( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.AConciliar, _
                            enumTipoAcao.EnviadaLDL0003Concordancia, _
                            , _
                            "1")
                    
                    Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                    Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                    
                    'KIDA - BMF
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            enumFuncaoSistema.Concordar, _
                            blnIgnoraGradeHorario, _
                            True, _
                            True, _
                            xmlRepeticao, _
                            "1")
                            
                    Set xmlTipoLiberacao = Nothing
                    
                    
                End If
'---------------------------------------------------------------------------------------------------
            Case enumAcaoConciliacao.AdmGeralEnviarDiscordancia
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.AConciliar, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        0, _
                        , _
                        enumTipoAcao.Liberacao)
                        
                If objDomNodeProc.selectSingleNode("IN_ULTI_MESG").Text = "S" Then
                
                    Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                    Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "D")
                    
                    Call objHistMensagem.Incluir( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.AConciliar, _
                            enumTipoAcao.EnviadaLDL0003Discordancia, _
                            , _
                            "1")
                    
                    'KIDA - BMF
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            enumFuncaoSistema.Concordar, _
                            blnIgnoraGradeHorario, _
                            True, _
                            True, _
                            xmlRepeticao, _
                            "1")
                            
                    Set xmlTipoLiberacao = Nothing
                End If
'---------------------------------------------------------------------------------------------------
            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                1)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        vntSequencialOperacao, _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        
            
                If objDomNodeProc.selectSingleNode("IN_ULTI_MESG").Text = "S" Then
                    
                    Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                                 fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                                                 vntSequencialConciliacao, _
                                                                 0)
                    
                    Call objHistMensagem.Incluir( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.AConciliar, _
                            enumTipoAcao.EnviadoPagamentoContingencia, _
                            , _
                            "1")
                
                    'KIDA - BMF
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            enumFuncaoSistema.PagarContingencia, _
                            blnIgnoraGradeHorario, _
                            True, _
                            True, _
                            , _
                            "1")
                End If
                        
'---------------------------------------------------------------------------------------------------
            Case enumAcaoConciliacao.AdmGeralRegularizar
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
        
                If Not flAutorizaRegularizacao("LDL0001", vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        0, _
                        , _
                        enumTipoAcao.Liquidacao)
            
        End Select
    
    End If

    LiquidarMultilateralBMF = vbNullString

    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlHistorico = Nothing
    
    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlHistorico = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarMultilateralBMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LiquidarEventosCBLC(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                    ByRef objDomNodeProc As MSXML2.IXMLDOMNode, _
                                    ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem

Dim xmlHistorico                            As MSXML2.DOMDocument40

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim vntSequencialConciliacao                As Variant

Dim strCondicao                             As String
Dim blnAcaoPermitida                        As Boolean

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")

    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    
    'Verifica se a ação solicitada, pode ou não ser executada
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Then
    
        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                           objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                           1))
        
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmGeralPagamento, _
                 enumAcaoConciliacao.AdmGeralPagamentoContingencia
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
        
            Case enumAcaoConciliacao.AdmGeralRegularizar
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3087
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length > 0
        
        End Select
        
        If Not blnAcaoPermitida Then
            GoTo ErrorHandler
        End If
    
    End If
    '=========================================================================================
    
    'Processamento das Operações
    '=========================================================================================
    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmAreaLiberar
                
                If objDomNodeProc.selectSingleNode("CO_MESG_SPB").Text = "LDL0007" Then
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.ConcordanciaAdmArea, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            , _
                            True)
    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.LiberacaoPagamento, _
                            blnIgnoraGradeHorario, _
                            True)
                
                Else
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liquidada, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            , _
                            True)
    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.LiberacaoRecebimento, _
                            blnIgnoraGradeHorario, _
                            True)
    
                End If
                
            Case enumAcaoConciliacao.AdmGeralPagamento
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.PagamentoLib, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

            Case enumAcaoConciliacao.AdmGeralRegularizar
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Liquidada, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.RegularizarContingencia, _
                        blnIgnoraGradeHorario, _
                        True)

        End Select
    Next

    'Processamento da mensagem
    '=========================================================================================
    If Not objDomNodeProc.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmAreaLiberar
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                1, , , , , _
                                                False)
        
                Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                       fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                       vntSequencialConciliacao, _
                                       0)
            
                For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            0)
                Next
        
                If objDomNodeProc.selectSingleNode("CO_MESG_SPB").Text = "LDL0007" Then
                    Call objMensagem.AlterarStatus( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.ConcordanciaAdmArea, _
                            objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            False, _
                            0, _
                            , _
                            enumTipoAcao.EnviadoPagamento)
                            
                Else
                    Call objMensagem.AlterarStatus( _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            enumStatusMensagem.MensagemLiquidada, _
                            objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            False, _
                            0, _
                            , _
                            enumTipoAcao.EnviadoPagamento)
                            
                End If
                        
            Case enumAcaoConciliacao.AdmGeralPagamentoContingencia
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                1, , , , , _
                                                False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        , _
                        1)
        
                Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                       fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                       vntSequencialConciliacao, _
                                       0)
            
                For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            0)
                Next
        
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoPagamentoContingencia, _
                        , _
                        1)
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.PagarContingencia, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , _
                        1)
                    
            Case enumAcaoConciliacao.AdmGeralPagamento
                Call objHistMensagem.Incluir( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        Val(objDomNodeProc.selectSingleNode("CO_ULTI_SITU_PROC").Text), _
                        enumTipoAcao.EnviadoPagamento, _
                        , _
                        1)
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        enumFuncaoSistema.Pagar, _
                        blnIgnoraGradeHorario, _
                        True, _
                        True, _
                        , _
                        1)
                    
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.PagamentoLib, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        0, _
                        , _
                        enumTipoAcao.EnviadoPagamento)
                        
            Case enumAcaoConciliacao.AdmGeralRegularizar
                
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        1, , , , , _
                        False)
        
                If Not flAutorizaRegularizacao(objDomNodeProc.selectSingleNode("CO_MESG_SPB").Text, vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        False, _
                        0, _
                        , _
                        enumTipoAcao.Liquidacao)
            
        End Select
        
    End If

    LiquidarEventosCBLC = vbNullString

    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlHistorico = Nothing
    
    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objHistMensagem = Nothing

    Set xmlHistorico = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarEventosCBLC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de liberação de registro de Roda de Dólar Pronto
'--------------------------------------------------------------------------------------------------

Public Function ConciliarRodaDolarPronto(ByVal intAcaoConciliacao As enumAcaoConciliacao, _
                                         ByRef xmlProcessamento As MSXML2.DOMDocument40, _
                                         ByVal blnFluxoAutomatico As Boolean, _
                                Optional ByRef blnGerarAlerta As Boolean = False) As String

Dim objMensagemEnvRec                       As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlErro                                 As MSXML2.DOMDocument40
Dim xmlRepeticao                            As MSXML2.DOMDocument40
Dim xmlMsgBMC0012                           As MSXML2.DOMDocument40
Dim xmlMsgBMC0010                           As MSXML2.DOMDocument40
Dim xmlMsgBMC0010XML                        As MSXML2.DOMDocument40
Dim xmlMsgBMC0013                           As MSXML2.DOMDocument40

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim blnConciliou                            As Boolean
Dim intProximaAcao                          As enumAcaoConciliacao
Dim vntSequencialConciliacao                As Variant
Dim blnPermiteRegularizacao                 As Boolean
Dim strValorOriginalOperacao                As String
    
    On Error GoTo ErrorHandler

    ConciliarRodaDolarPronto = "Conciliação Inicia com Erro como Default"
    
    Set objMensagemEnvRec = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objRegra = CreateObject("A8LQS.clsRegra")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
    
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgBMC0012 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgBMC0010 = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgBMC0010XML = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgBMC0013 = CreateObject("MSXML2.DOMDocument.4.0")
    
    If intAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
       intAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberarContingencia Or _
       intAcaoConciliacao = enumAcaoConciliacao.AdmAreaRejeitar Then
        
        If blnFluxoAutomatico Or _
           intAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberarContingencia Or _
           intAcaoConciliacao = enumAcaoConciliacao.AdmAreaRejeitar Then
            
            Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
            Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", xmlProcessamento.selectSingleNode("//CO_EMPR").Text)
                
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
            Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.BMC)
                
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
            Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))))
            
            If Not xmlProcessamento.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
                Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.EmSer)
                Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusOperacao.ManualEmSer)
            
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LayoutEntrada", "")
                Call fgAppendNode(xmlFiltros, "Grupo_LayoutEntrada", "LayoutEntrada", "138")
                    
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NumeroControleLTR", "")
                Call fgAppendNode(xmlFiltros, "Grupo_NumeroControleLTR", "NumeroControleLTR", "NULL")
                    
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_OrderBy", "")
                Call fgAppendNode(xmlFiltros, "Grupo_OrderBy", "OrderBy", "A.NU_SEQU_OPER_ATIV")
                    
                Set xmlMensagem = objMensagemEnvRec.LerXMLMensagem(xmlProcessamento.selectSingleNode("//NU_CTRL_IF").Text, _
                                                                   xmlProcessamento.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            
                Call xmlOperacao.loadXML(objOperacao.ObterDetalheOperacao(xmlFiltros, True))
                
                blnConciliou = False
                    
                If xmlOperacao.xml <> vbNullString Then
                    If blnFluxoAutomatico Then
                        If Not objRegra.VerificarRegraAutomatica2(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, enumFuncaoSistema.Conciliar) Then
                            GoTo ExitSub
                        End If
                    End If
                    
                    For Each objDomNode In xmlOperacao.selectNodes("Repeat_DetalheOperacao/*")
                        If flCompararCamposRodaDolar(xmlMensagem.xml, objDomNode.xml) = vbNullString Then
                            If blnFluxoAutomatico Then
                                Call objOperacao.Alterar(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, , , _
                                                         xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text)
                                
                                Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                                blnConciliou = True
                                Exit For
                            Else
                                'Existe pelo menos uma operação em condições de conciliação com esta BMC0011.
                                lngCodigoErroNegocio = 4246
                                GoTo ErrorHandler
                            End If
                        End If
                    Next
                Else
                    If blnFluxoAutomatico Then
                        blnGerarAlerta = True
                    End If
                End If
            
                If blnFluxoAutomatico Then
                    If Not blnConciliou Then GoTo ExitSub
                End If
                    
            ElseIf Not xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
                Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
            
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
                Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "BMC0011")
                
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_NumeroSeqOperacao", "")
                Call fgAppendNode(xmlFiltros, "Grupo_NumeroSeqOperacao", "NumeroSeqOperacao", "NULL")
                    
                Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_OrderBy", "")
                Call fgAppendNode(xmlFiltros, "Grupo_OrderBy", "OrderBy", "A.DH_REGT_MESG_SPB")
                    
                Set xmlOperacao = objOperacao.LerXMLOperacao(xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                
                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.Previsto)
                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.Previsto)
                
                If Not objRegra.VerificarRegraAutomatica2(xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, enumFuncaoSistema.Conciliar) Then
                    GoTo ExitSub
                End If
            
                Call xmlMensagem.loadXML(objMensagemSLCC.ObterDetalheMensagem(xmlFiltros, True))
                
                blnConciliou = False
                    
                If xmlMensagem.xml <> vbNullString Then
                    For Each objDomNode In xmlMensagem.selectNodes("Repeat_DetalheMensagem/*")
                        If flCompararCamposRodaDolar(objDomNode.xml, xmlOperacao.xml) = vbNullString Then
                            
                            Call objMensagemEnvRec.AlterarSequenciaOperacao( _
                                    objDomNode.selectSingleNode("NU_CTRL_IF").Text, _
                                    fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                    xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                    objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                            
                            Set xmlMensagem = objMensagemEnvRec.LerXMLMensagem(objDomNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                                               objDomNode.selectSingleNode("NU_SEQU_CNTR_REPE").Text)
                            blnConciliou = True
                            Exit For
                        End If
                    Next
                End If
            
                If Not blnConciliou Then GoTo ExitSub
                    
            End If
            
        Else
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            Set xmlMensagem = objMensagemEnvRec.LerXMLMensagem(xmlProcessamento.selectSingleNode("//NU_CTRL_IF").Text, _
                                                               xmlProcessamento.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
            Call xmlErro.loadXML(flCompararCamposRodaDolar(xmlMensagem.xml, xmlOperacao.xml))
            
            If xmlErro.xml <> vbNullString Then
                ConciliarRodaDolarPronto = xmlErro.xml
                GoTo ExitSub
            End If
        
        End If
                    
    Else
    
        Set xmlMensagem = objMensagemEnvRec.LerXMLMensagem(xmlProcessamento.selectSingleNode("//NU_CTRL_IF").Text, _
                                                           xmlProcessamento.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
        If intAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberar Or _
           intAcaoConciliacao = enumAcaoConciliacao.AdmAreaRegularizar Then
            
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlOperacao.loadXML(ConsultaConciliacao(, xmlProcessamento.selectSingleNode("//NU_CTRL_IF").Text, _
                                                           xmlProcessamento.selectSingleNode("//DH_REGT_MESG_SPB").Text))
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
        
        End If
    
    End If
    
    intProximaAcao = 0
    
    'Processamento de Operações e Mensagens
    '=========================================================================================
    Select Case intAcaoConciliacao
        Case enumAcaoConciliacao.BOConcordar
            
            intProximaAcao = enumAcaoConciliacao.AdmAreaLiberar
            
            If Not xmlMensagem.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC") Is Nothing Then
                If Val(xmlMensagem.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) = enumTipoAcao.RegistroContingencia Then
                    intProximaAcao = 0
                ElseIf Val(xmlMensagem.selectSingleNode("//TP_ACAO_MESG_SPB_EXEC").Text) = enumTipoAcao.DiscordanciaAdmBO Then
                    'DISCORDÂNCIA já foi enviada para esta mensagem.
                    Call fgAdicionaErro(xmlErro, 3083)
                    ConciliarRodaDolarPronto = xmlErro.xml
                    GoTo ExitSub
                End If
            End If

            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                            xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                            xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                            , _
                                            , _
                                            xmlMensagem.selectSingleNode("//CO_MESG").Text, _
                                            blnFluxoAutomatico)
    
            Call IncluirConciliacaoOperacao( _
                    vntSequencialConciliacao, _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                    True)
            
            Call objOperacao.AlterarStatusAcao( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    IIf(blnFluxoAutomatico, enumStatusOperacao.ConciliadaAutomatica, enumStatusOperacao.Conciliada), _
                    vbNullString, , , _
                    False)

            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    IIf(blnFluxoAutomatico, enumStatusMensagem.ConciliadaAutomatica, enumStatusMensagem.Conciliada), _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , _
                    False, _
                    False)
                    
        Case enumAcaoConciliacao.AdmAreaLiberar
            
            Call objOperacao.AlterarStatusAcao( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    IIf(blnFluxoAutomatico, enumStatusOperacao.LiberadaAutomatica, enumStatusOperacao.Liberada), _
                    vbNullString, , , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text)

            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , _
                    enumTipoAcao.EnviadaBMC0012)

            Call xmlRepeticao.loadXML("<Grupo_TagsObrigatorias>" & _
                                      "<CO_SITU_ACTE>" & enumSituacaoAceite.Aceito & "</CO_SITU_ACTE>" & _
                                      "</Grupo_TagsObrigatorias>")
            
            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    enumFuncaoSistema.Liberar, _
                    True, , _
                    True, _
                    xmlRepeticao, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
        
        Case enumAcaoConciliacao.AdmAreaRejeitar
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , _
                    enumTipoAcao.DiscordanciaAdmBO)

            Call xmlRepeticao.loadXML("<Grupo_TagsObrigatorias>" & _
                                      "<CO_SITU_ACTE>" & enumSituacaoAceite.Rejeitado & "</CO_SITU_ACTE>" & _
                                      "</Grupo_TagsObrigatorias>")
            
            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    enumFuncaoSistema.Discordar, _
                    True, , _
                    True, _
                    xmlRepeticao, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        Case enumAcaoConciliacao.AdmAreaLiberarContingencia
            
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , _
                    enumTipoAcao.RegistroContingencia)

            Call xmlRepeticao.loadXML("<Grupo_TagsObrigatorias>" & _
                                      "<CO_SITU_ACTE>" & enumSituacaoAceite.Aceito & "</CO_SITU_ACTE>" & _
                                      "</Grupo_TagsObrigatorias>")
            
            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    enumFuncaoSistema.PagarContingencia, _
                    True, , _
                    True, _
                    xmlRepeticao, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

        Case enumAcaoConciliacao.AdmAreaRegularizar
            
            Set xmlMsgBMC0012 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0012R1", _
                                                                       xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text)
            
            blnPermiteRegularizacao = True
            
            If xmlMsgBMC0012.xml = vbNullString Then
                blnPermiteRegularizacao = False
            ElseIf xmlMsgBMC0012.selectSingleNode("//SitLancBMC") Is Nothing Then
                blnPermiteRegularizacao = False
            ElseIf xmlMsgBMC0012.selectSingleNode("//SitLancBMC").Text <> "301" Then
                blnPermiteRegularizacao = False
            End If
                
            If Not blnPermiteRegularizacao Then
                'Não foi possível encontrar a resposta da mensagem BMC0012 com Situação de Lançamento = 301 (Confirmação aceita pela câmara).
                lngCodigoErroNegocio = 4251
                GoTo ErrorHandler
            End If
            
            Call objOperacao.AlterarStatusAcao( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    enumStatusOperacao.Registrada, _
                    vbNullString, , , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text)
    
            Call objMensagemEnvRec.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text, _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , _
                    enumTipoAcao.EnviadaBMC0012)

            Call objMensagemEnvRec.AlterarSequenciaOperacao( _
                    xmlMsgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                    fgDtHrStr_To_DateTime(xmlMsgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    xmlMsgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            
            Set xmlMsgBMC0012 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0012", _
                                                                       xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text)
            
            If xmlMsgBMC0012.xml <> vbNullString Then
                Call objMensagemEnvRec.AlterarStatus( _
                        xmlMsgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                        xmlMsgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.Registrada, _
                        vbNullString, _
                        , _
                        False, _
                        xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            
                Call objMensagemEnvRec.AlterarSequenciaOperacao( _
                        xmlMsgBMC0012.selectSingleNode("//NU_CTRL_IF").Text, _
                        fgDtHrStr_To_DateTime(xmlMsgBMC0012.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                        xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                        xmlMsgBMC0012.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If

            strValorOriginalOperacao = xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text
            
            Call xmlFiltros.loadXML(vbNullString)
            Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
            Call fgAppendNode(xmlFiltros, "Grupo_BancoLiquidante", "BancoLiquidante", xmlProcessamento.selectSingleNode("//CO_EMPR").Text)
                
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
            Call fgAppendNode(xmlFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.BMC)
                
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Data", "")
            Call fgAppendNode(xmlFiltros, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux))))
            
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.EnviadaBUS)
            Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", enumStatusMensagem.MensagemCancelada)
        
            Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
            Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", "BMC0010")
            
            Call xmlMsgBMC0010.loadXML(objMensagemSLCC.ObterDetalheMensagem(xmlFiltros, True))
            
            If xmlMsgBMC0010.xml <> vbNullString Then
                For Each objDomNode In xmlMsgBMC0010.selectNodes("Repeat_DetalheMensagem/*")
                    
                    Set xmlMsgBMC0010XML = objMensagemEnvRec.LerXMLMensagem(objDomNode.selectSingleNode("NU_CTRL_IF").Text, , , _
                                                                            "BMC0010")
                    
                    If xmlMensagem.selectSingleNode("//NU_CTRL_CAMR").Text = xmlMsgBMC0010XML.selectSingleNode("//NumCtrlBMCOr").Text Then
                        xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = xmlMsgBMC0010XML.selectSingleNode("//VA_FINC").Text
                        If xmlOperacao.selectSingleNode("//ENVIAR_PREV") Is Nothing Then
                            Call fgAppendNode(xmlOperacao, "MESG", "ENVIAR_PREV", enumIndicadorSimNao.Sim)
                        Else
                            xmlOperacao.selectSingleNode("//ENVIAR_PREV").Text = enumIndicadorSimNao.Sim
                        End If
                        
                        Call objMensagemEnvRec.AlterarSequenciaOperacao( _
                                xmlMsgBMC0010XML.selectSingleNode("//NU_CTRL_IF").Text, _
                                fgDtHrStr_To_DateTime(xmlMsgBMC0010XML.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                                xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                xmlMsgBMC0010XML.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                        
                        Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.Realizado)
                    
                        Set xmlMsgBMC0013 = objMensagemEnvRec.LerXMLMensagem(, , , "BMC0013", _
                                                                                   xmlMsgBMC0010XML.selectSingleNode("//NU_CTRL_CAMR").Text)
            
                        If xmlMsgBMC0013.xml <> vbNullString Then
                            If Val(xmlMsgBMC0013.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusMensagem.MensagemInconsistente Then
                                Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.EstornoRealizado)
                            
                                Call objMensagemEnvRec.AlterarSequenciaOperacao( _
                                        xmlMsgBMC0013.selectSingleNode("//NU_CTRL_IF").Text, _
                                        fgDtHrStr_To_DateTime(xmlMsgBMC0013.selectSingleNode("//DH_REGT_MESG_SPB").Text), _
                                        xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                        xmlMsgBMC0013.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            End If
                        End If
                    End If
                Next
            End If
        
            xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text = strValorOriginalOperacao
    
    End Select

    If intProximaAcao <> 0 Then
        If objRegra.VerificarRegraAutomatica2(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, enumFuncaoSistema.Liberar) Then
            If xmlOperacao.selectSingleNode("NU_CTRL_IF") Is Nothing Then
                Call fgAppendNode(xmlOperacao, "MESG", "NU_CTRL_IF", xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text)
                Call fgAppendNode(xmlOperacao, "MESG", "DH_REGT_MESG_SPB", xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text)
                Call fgAppendNode(xmlOperacao, "MESG", "NU_SEQU_CNTR_REPE", xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            End If
            
            Call ConciliarRodaDolarPronto(intProximaAcao, xmlOperacao, True)
        End If
    End If
    
    ConciliarRodaDolarPronto = vbNullString
    
ExitSub:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objMensagemSLCC = Nothing
    Set objProcessoOperacao = Nothing
    Set objRegra = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    Set xmlErro = Nothing
    Set xmlRepeticao = Nothing
    Set xmlFiltros = Nothing
    Set xmlMsgBMC0012 = Nothing
    Set xmlMsgBMC0010 = Nothing
    Set xmlMsgBMC0010XML = Nothing
    Set xmlMsgBMC0013 = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagemEnvRec = Nothing
    Set objOperacao = Nothing
    Set objMensagemSLCC = Nothing
    Set objProcessoOperacao = Nothing
    Set objRegra = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    Set xmlErro = Nothing
    Set xmlRepeticao = Nothing
    Set xmlFiltros = Nothing
    Set xmlMsgBMC0012 = Nothing
    Set xmlMsgBMC0010 = Nothing
    Set xmlMsgBMC0010XML = Nothing
    Set xmlMsgBMC0013 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarRodaDolarPronto Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Comparar Campos da BMC0011 com a operação registro roda dólar pronto
'--------------------------------------------------------------------------------------------------

Private Function flCompararCamposRodaDolar(ByVal strXMLMensagem As String, _
                                           ByVal strXMLOperacao As String) As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlErros                                As MSXML2.DOMDocument40

Dim intDebCredMesg                          As Integer
Dim intDebCredOper                          As Integer
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlMensagem.loadXML("<Grupo_Mensagem>" & strXMLMensagem & "</Grupo_Mensagem>")
    Call xmlOperacao.loadXML("<Grupo_Operacao>" & strXMLOperacao & "</Grupo_Operacao>")
    
    If Not xmlMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED") Is Nothing Then
        intDebCredMesg = xmlMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text
    Else
        intDebCredMesg = xmlMensagem.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    End If
    
    If Not xmlOperacao.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED") Is Nothing Then
        intDebCredOper = xmlOperacao.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text
    Else
        intDebCredOper = xmlOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text
    End If
    
    If intDebCredMesg <> intDebCredOper Then
        fgAdicionaErro xmlErros, 3067, enumJustificativa.DebitoCredito
    End If
              
    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//CO_PRAC").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//CO_PRAC").Text) Then
        fgAdicionaErro xmlErros, 3133, enumJustificativa.CodigoPraca
    End If
    
    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//CO_MOED_ESTR").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//CO_MOED_ESTR").Text) Then
        fgAdicionaErro xmlErros, 3134, enumJustificativa.MoedaEstrangeira
    End If
    
    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//PE_TAXA_NEGO").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//PE_TAXA_NEGO").Text) Then
        fgAdicionaErro xmlErros, 3074, enumJustificativa.TaxaNegociada
    End If

    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//VA_FINC").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//VA_OPER_ATIV").Text) Then
        fgAdicionaErro xmlErros, 3072, enumJustificativa.ValorFinanceiro
    End If

    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//VA_MOED_ESTR").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//VA_MOED_ESTR").Text) Then
        fgAdicionaErro xmlErros, 3135, enumJustificativa.ValorMoedaEstrangeira
    End If

    If xmlMensagem.selectSingleNode("//DT_LIQU").Text <> xmlOperacao.selectSingleNode("//DT_LIQU_OPER_ATIV").Text Then
        fgAdicionaErro xmlErros, 3069, enumJustificativa.DataLiquidacao
    End If

    If fgVlrXml_To_Decimal(xmlMensagem.selectSingleNode("//CO_SISB_COTR").Text) <> fgVlrXml_To_Decimal(xmlOperacao.selectSingleNode("//CO_SISB_COTR").Text) Then
        fgAdicionaErro xmlErros, 4245, enumJustificativa.CanalSISBACENCorretora
    End If

    If xmlErros.selectNodes("//Grupo_ErrorInfo").length > 0 Then
        flCompararCamposRodaDolar = xmlErros.xml
    Else
        flCompararCamposRodaDolar = vbNullString
    End If
    
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    Set xmlErros = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    Set xmlErros = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flCompararCamposRodaDolar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LiquidarBrutaCBLC(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                  ByRef xmlProcessamento As MSXML2.DOMDocument40, _
                         Optional ByVal blnFluxoAutomatico As Boolean = False, _
                         Optional ByVal blnIgnoraGradeHorario As Boolean = False) As String

Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objOperacao                             As A8LQS.clsOperacao
Dim objHistMensagem                         As A8LQS.clsHistSituacaoMensagem
Dim objProcessaMensagemBMA                  As A8LQS.clsProcessoMensagemBMA
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim xmlTipoLiberacao                        As MSXML2.DOMDocument40

Dim blnAcaoPermitida                        As Boolean
Dim blnEnvioLTR8                            As Boolean
Dim blnPagtoCont                            As Boolean
Dim vntSequencialConciliacao                As Variant

Dim strCondicao                             As String
Dim intStatusContraparteLTR0007             As Integer
Dim strDataHoraXML                          As String
    
On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
        Set xmlOperacao = objOperacao.LerXMLOperacao(xmlProcessamento.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    End If

    Set xmlMensagem = objMensagem.LerXMLMensagem(xmlProcessamento.selectSingleNode("//NU_CTRL_IF").Text, _
                                                 1, _
                                                 xmlProcessamento.selectSingleNode("//IN_OPER_DEBT_CRED").Text)

    If xmlMensagem.xml = vbNullString Then
        '3062 - Mensagem não encontrada
        lngCodigoErroNegocio = 3062
        GoTo ErrorHandler
    End If

    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    Set objProcessaMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    strDataHoraXML = xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text

    'Verifica se a ação solicitada, pode ou não ser executada
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamento Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Or _
       pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralRegularizar Then
    
'        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
 '                                                          strDataHoraXML, _
  '                                                         xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text))
        
        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                           "", _
                                                           xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text))
        
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']"
                lngCodigoErroNegocio = 3077
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
            
                If blnAcaoPermitida Then
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                    lngCodigoErroNegocio = 3084
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                End If
            
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                blnPagtoCont = xmlHistorico.selectNodes(strCondicao).length > 0
            
            Case enumAcaoConciliacao.AdmGeralPagamento, _
                 enumAcaoConciliacao.AdmGeralPagamentoContingencia
                
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                lngCodigoErroNegocio = 3084
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
                blnPagtoCont = xmlHistorico.selectNodes(strCondicao).length > 0
        
            Case enumAcaoConciliacao.AdmGeralRegularizar
                strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLTR0002Concordancia & "']"
                lngCodigoErroNegocio = 3086
                blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length = 0
        
                If blnAcaoPermitida Then
                    strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
                    lngCodigoErroNegocio = 3087
                    blnAcaoPermitida = xmlHistorico.selectNodes(strCondicao).length > 0
                End If
        
        End Select
        
        If Not blnAcaoPermitida Then
            GoTo ErrorHandler
        End If
    
    ElseIf pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Then
    
        Call xmlHistorico.loadXML(objHistMensagem.LerTodos(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                           "", _
                                                           xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text))
    
        strCondicao = "//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']"
        blnPagtoCont = xmlHistorico.selectNodes(strCondicao).length > 0
    
    End If
    '=========================================================================================
    
    'Define ações a serem tomadas para LTR0007
    '=========================================================================================
'    blnEnvioLTR8 = False
    
    If xmlMensagem.selectSingleNode("//CO_MESG").Text = "LTR0007" And _
       Val(xmlMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Sim Then
    
        If pintAcaoConciliacao = enumAcaoConciliacao.BOConcordar Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia Or _
           pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralPagamentoContingencia Then
            
            If xmlProcessamento.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                blnEnvioLTR8 = True
            Else
                Call flVerificarContraparteLTR0007(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                   xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                                   xmlProcessamento.selectSingleNode("//IN_OPER_DEBT_CRED").Text, _
                                                   intStatusContraparteLTR0007)
                
                If intStatusContraparteLTR0007 = enumStatusMensagem.MensagemLiquidada Then
                
                    pintAcaoConciliacao = enumAcaoConciliacao.AdmGeralEnviarConcordancia
                
                End If
                
            End If
        
        End If
        
    End If
    '=========================================================================================
    
    'Processamento de Operações e Mensagens
    '=========================================================================================
    Select Case pintAcaoConciliacao
        Case enumAcaoConciliacao.BOConcordar
            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                            xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                            strDataHoraXML, _
                                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                             , _
                                            , _
                                            xmlMensagem.selectSingleNode("//CO_MESG").Text, _
                                            blnFluxoAutomatico, _
                                            False)
    
            Call IncluirConciliacaoOperacao( _
                    vntSequencialConciliacao, _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    strDataHoraXML, _
                    0, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

            Call objOperacao.AlterarStatusAcao( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    IIf(Not blnFluxoAutomatico, enumStatusOperacao.ConcordanciaBackoffice, enumStatusOperacao.ConcordanciaBackofficeAutomatico), _
                    xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)
                    
            Call objMensagem.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    strDataHoraXML, _
                    IIf(Not blnFluxoAutomatico, enumStatusMensagem.ConcordanciaBackoffice, enumStatusMensagem.ConcordanciaBackofficeAutomatico), _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                                
            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    enumFuncionalidade.Concordancia, _
                    True, _
                    True, _
                    False)
            
        Case enumAcaoConciliacao.AdmGeralEnviarConcordancia
            Select Case xmlMensagem.selectSingleNode("//CO_MESG").Text
                Case "LTR0001"
                    Call objHistMensagem.Incluir( _
                            xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            strDataHoraXML, _
                            Val(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text), _
                            enumTipoAcao.EnviadaLTR0002Concordancia, _
                            , _
                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    
                    Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                    Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            enumFuncaoSistema.Concordar, _
                            blnIgnoraGradeHorario, _
                            True, _
                            True, _
                            xmlTipoLiberacao, _
                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.Confirmacao, _
                            blnIgnoraGradeHorario, _
                            False)

                    Set xmlTipoLiberacao = Nothing
                    
                Case "LTR0005R2"
                    Call objOperacao.AlterarStatusAcao( _
                            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liquidada, _
                            xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)

                    Call objMensagem.AlterarStatus( _
                            xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                            strDataHoraXML, _
                            enumStatusMensagem.MensagemLiquidada, _
                            vbNullString, _
                            , _
                            False, _
                            xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.LiberacaoRecebimento, _
                            blnIgnoraGradeHorario, _
                            False)

                Case "LTR0007"
                    If Val(xmlMensagem.selectSingleNode("//IN_CONF_MESG_LTR").Text) = enumIndicadorSimNao.Sim Then
                       
                        If blnEnvioLTR8 Then
                            Call objOperacao.AlterarStatusAcao( _
                                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                    IIf(Not blnFluxoAutomatico, enumStatusOperacao.ConcordanciaLib, enumStatusOperacao.ConcordanciaLibAuto), _
                                    xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)
                        
                            Call objMensagem.AlterarStatus( _
                                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                    strDataHoraXML, _
                                    IIf(Not blnFluxoAutomatico, enumStatusMensagem.ConcordanciaLib, enumStatusMensagem.ConcordanciaLibAuto), _
                                    vbNullString, _
                                    , _
                                    False, _
                                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            
                            Call objHistMensagem.Incluir( _
                                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                    strDataHoraXML, _
                                    Val(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text), _
                                    enumTipoAcao.EnviadaLTR0008Concordancia, _
                                    , _
                                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            
                            Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
                            Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                            
                            
                            'KIDA - CBLC - 19/09/2008
                            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                    enumFuncaoSistema.Concordar, _
                                    blnIgnoraGradeHorario, _
                                    True, _
                                    True, _
                                    xmlTipoLiberacao, _
                                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, _
                                    vbNullString, _
                                    False, _
                                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text)
                        
                            Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.RealizadoSolicitado)

                            Set xmlTipoLiberacao = Nothing
                        
                        Else
                            If intStatusContraparteLTR0007 = enumStatusMensagem.MensagemLiquidada Then
                                Call objProcessaMensagemBMA.GerenciarChamada( _
                                        xmlOperacao, _
                                        Nothing, _
                                        enumStatusOperacao.Liquidada, _
                                        0, _
                                        "", _
                                        enumIndicadorSimNao.Nao, _
                                        "", _
                                        False, _
                                        True)
                            
                                Call objMensagem.AlterarStatus( _
                                        xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                        strDataHoraXML, _
                                        enumStatusMensagem.MensagemLiquidada, _
                                        vbNullString, _
                                        , _
                                        False, _
                                        xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                                
                                If blnPagtoCont Then
                                    Call objFluxoProcessamento.EnviarMovimentoPJ(xmlOperacao, enumTipoMovimentoPJ.Realizado)
                                End If
                                
                            Else
                                Call objOperacao.AlterarStatusAcao( _
                                        xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                        IIf(Not blnFluxoAutomatico, enumStatusOperacao.Liberada, enumStatusOperacao.LiberadaAutomatica), _
                                        xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)
                            
                                Call objMensagem.AlterarStatus( _
                                        xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                        strDataHoraXML, _
                                        enumStatusMensagem.Liberada, _
                                        vbNullString, _
                                        , _
                                        False, _
                                        xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            
                            End If
                        End If
                            
                    Else
                                
                        Call objOperacao.AlterarStatusAcao( _
                                xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                enumStatusOperacao.Liquidada, _
                                xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)
                                
                        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
                        
                        Call objFluxoProcessamento.DisponibilizarLancamentoCC(xmlOperacao)
                        Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, _
                                                                             enumTipoMovimento.RealizadoConfirmado)
                                    
                        Call objMensagem.AlterarStatus( _
                                xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                strDataHoraXML, _
                                enumStatusMensagem.MensagemLiquidada, _
                                vbNullString, _
                                , _
                                False, _
                                xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                            
                    End If
            
            End Select
            
        Case enumAcaoConciliacao.AdmGeralPagamento
            
            Call objOperacao.AlterarStatusAcao( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    IIf(Not blnFluxoAutomatico, enumStatusOperacao.PagamentoLib, enumStatusOperacao.PagamentoLiberadoAutomatico), _
                    xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , , False)
                                
            Call objMensagem.AlterarStatus( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    strDataHoraXML, _
                    IIf(Not blnFluxoAutomatico, enumStatusMensagem.PagamentoLib, enumStatusMensagem.PagamentoLibAutomatico), _
                    vbNullString, _
                    , _
                    False, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    enumFuncaoSistema.Pagar, _
                    blnIgnoraGradeHorario, _
                    True, _
                    True, _
                    , _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                    enumFuncionalidade.LiberacaoPagamento, _
                    blnIgnoraGradeHorario, _
                    False)

        Case enumAcaoConciliacao.AdmGeralPagamentoContingencia

            'KIDA - CBLC - 07/10/2008
            'If xmlOperacao.xml <> vbNullString Then
             '   Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.RealizadoSolicitado)
            'End If
                            
            Call objHistMensagem.Incluir( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    strDataHoraXML, _
                    Val(xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text), _
                    enumTipoAcao.EnviadoPagamentoContingencia, _
                    , _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
            
            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , , , _
                    False)
    
            Call IncluirConciliacaoOperacao( _
                    vntSequencialConciliacao, _
                    , _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    , _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
    
            Set xmlTipoLiberacao = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlTipoLiberacao, "", "TP_CONF_DIVG", "C")
                    
            Call objProcessoOperacao.GerenciarChamadaViaTela( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    enumFuncaoSistema.PagarContingencia, _
                    blnIgnoraGradeHorario, _
                    True, _
                    True, _
                    xmlTipoLiberacao, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
            Set xmlTipoLiberacao = Nothing
        
        Case enumAcaoConciliacao.AdmGeralRegularizar
            
            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                    xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                    xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                    xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , , , _
                    False)
            
            If xmlMensagem.selectSingleNode("//CO_MESG").Text = "LTR0007" Then
            
                Call flRegularizarBilateralBookTransferCBLC(xmlOperacao, xmlMensagem, vntSequencialConciliacao)
            
            Else
                If Not flAutorizaRegularizacaoCBLC(xmlMensagem.selectSingleNode("//CO_MESG").Text, vntSequencialConciliacao) Then
                    'Regularização não permitida pois a mensagem de pagamento ainda não está liquidada.
                    lngCodigoErroNegocio = 4237
                    GoTo ErrorHandler
                End If
                
                'KIDA - CBLC - 07/10/2008
                Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.RealizadoSolicitado)






                
                Call objMensagem.AlterarStatus( _
                        xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                        strDataHoraXML, _
                        enumStatusMensagem.MensagemLiquidada, _
                        vbNullString, _
                        vbNullString, _
                        False)
                
                Call objProcessaMensagemBMA.GerenciarChamada(xmlOperacao, _
                                                             xmlMensagem, _
                                                             enumStatusOperacao.Liquidada, _
                                                             0, _
                                                             "", _
                                                             enumIndicadorSimNao.Nao, _
                                                             "")
            End If
                                
                                
                                
    End Select

    LiquidarBrutaCBLC = vbNullString
    
    Set objMensagem = Nothing
    Set objOperacao = Nothing
    Set objHistMensagem = Nothing
    Set objProcessaMensagemBMA = Nothing
    Set objProcessoOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlHistorico = Nothing
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function

ErrorHandler:
    Set objMensagem = Nothing
    Set objOperacao = Nothing
    Set objHistMensagem = Nothing
    Set objProcessaMensagemBMA = Nothing
    Set objProcessoOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    
    Set xmlHistorico = Nothing
    Set xmlOperacao = Nothing
    Set xmlMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarBrutaCBLC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'KIDA - CBLC - 10/10/2008
Private Function flAutorizaRegularizacaoCBLC(ByVal pstrCodigoMensagemSPB As String, _
                                             ByVal pvntSequenciaConciliacaoMensagemOperacao As Variant) As Boolean
                                         
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objMensagem                             As A8LQS.clsMensagemSLCC
Dim strCodMesgEnviada                       As String
Dim intStatusMensagem                       As Integer


Dim strSQL                                  As String
Dim rsSQL                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    Select Case pstrCodigoMensagemSPB
        Case "BMA0015"
            strCodMesgEnviada = "SEL1023"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "BMC0101"
            strCodMesgEnviada = "BMC0102"
            intStatusMensagem = enumStatusMensagem.Confirmada
        Case "LDL0001"
            strCodMesgEnviada = "LDL0004"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LDL0007"
            strCodMesgEnviada = "LDL0008"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LTR0001"
            strCodMesgEnviada = "LTR0004"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LTR0007"
            strCodMesgEnviada = "LTR0008"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
    End Select
    
    
    If pstrCodigoMensagemSPB = "LTR0007" Then
    
        
        strSQL = " SELECT    B.CO_MESG_SPB, " & vbCrLf & _
                 "           B.CO_ULTI_SITU_PROC " & vbCrLf & _
                 "   FROM    A8.TB_MESG_RECB_ENVI_SPB A, " & vbCrLf & _
                 "           A8.TB_MESG_RECB_ENVI_SPB B  " & vbCrLf & _
                 "  Where    a.NU_SEQU_CNCL_OPER_ATIV_MESG = " & pvntSequenciaConciliacaoMensagemOperacao & vbCrLf & _
                 "    AND    A.NU_CTRL_CAMR  = B.NU_CTRL_CAMR " & vbCrLf & _
                 "    AND    B.CO_MESG_SPB = 'LTR0008' "

        
        Set rsSQL = fgQuerySQL(strSQL)
        
        If rsSQL.EOF Then
            flAutorizaRegularizacaoCBLC = False
        Else
            If rsSQL("CO_ULTI_SITU_PROC") = enumStatusMensagem.MensagemLiquidada Then
                flAutorizaRegularizacaoCBLC = True
            Else
                flAutorizaRegularizacaoCBLC = False
            End If
        End If
    
    Else
        Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
        Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMensagem = CreateObject("A8LQS.clsMensagemSLCC")
        
        Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
        Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", intStatusMensagem)
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
        Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", strCodMesgEnviada)
        
        Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SequencialConciliacao", "")
        Call fgAppendNode(xmlFiltros, "Grupo_SequencialConciliacao", "SequencialConciliacao", pvntSequenciaConciliacaoMensagemOperacao)
        
        Call xmlLeitura.loadXML(objMensagem.ObterDetalheMensagem(xmlFiltros))
        flAutorizaRegularizacaoCBLC = IIf(xmlLeitura.xml = vbNullString, False, True)
    End If
    
    Set xmlFiltros = Nothing
    Set xmlLeitura = Nothing
    Set objMensagem = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlFiltros = Nothing
    Set xmlLeitura = Nothing
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAutorizaRegularizacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'KIDA - CBLC
Private Function flRegularizarBilateralBookTransferCBLC(ByVal pxmlOperacao As MSXML2.DOMDocument40, _
                                                        ByVal pxmlMensagem As MSXML2.DOMDocument40, _
                                                        ByVal pvntSequencialConciliacao As Variant)


Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objMensagemBMA                          As A8LQS.clsProcessoMensagemBMA
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento


Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim vntSequenciaOperacao                    As Variant

On Error GoTo ErrorHandler
        
    
    If pvntSequencialConciliacao = 0 Then Exit Function
    
    '-------------------------------------------------------------------------------------------------------------
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    
    If xmlCociliacao.loadXML(objConciliacao.LerTodos(pvntSequencialConciliacao)) Then
    
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
        Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        
        For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
             
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
            If vntSequenciaOperacao <> "0" Then
                
                Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                    
                If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada Or _
                  xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBackoffice Then
                
                    
                    'KIDA - CBLC - 07/10/2008
                    Call objFluxoProcessamento.EnviarMovimentoSubReserva(xmlOperacao, enumTipoMovimento.RealizadoSolicitado)
                    
                    objMensagemBMA.GerenciarChamada xmlOperacao, _
                                                    pxmlMensagem, _
                                                    enumStatusOperacao.Liquidada, _
                                                    0, _
                                                    0, _
                                                    enumIndicadorSimNao.Nao, _
                                                    ""
                                                    
                    Call objMesgRecbEnviSPB.AlterarStatus(xmlNode.selectSingleNode("NU_CTRL_IF").Text, _
                                                          xmlNode.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                          enumStatusMensagem.MensagemLiquidada, _
                                                          "", _
                                                          "", _
                                                          False, _
                                                          1, _
                                                          "")
                                                    
                End If
            End If
            
            Set xmlOperacao = Nothing
        Next
        
        Set objOperacao = Nothing
        Set objFluxoProcessamento = Nothing
        Set objMesgRecbEnviSPB = Nothing
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    Set objFluxoProcessamento = Nothing
    Set objMesgRecbEnviSPB = Nothing
    
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemR1Pagamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




'--------------------------------------------------------------------------------------------------
'Processo de liquidação com Corretoras.
'--------------------------------------------------------------------------------------------------

Public Function LiquidarCorretoras(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                   ByRef objDomNodeProc As MSXML2.IXMLDOMNode, _
                                   ByVal blnIgnoraGradeHorario As Boolean, _
                                   Optional ByVal pblnLiquidacaoPorLote As Boolean = False) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objHistOperacao                         As A8LQS.clsHistSituacaoOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlMsgEnvio                             As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim vntSequencialConciliacao                As Variant
Dim lngSequencialJustificativa              As Long
Dim lngStatusOperacao                       As Long
Dim lngTipoBackOfficeLote                   As Long

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objHistOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    'Processamento das Operações
    '=========================================================================================
    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
            
                If pblnLiquidacaoPorLote = True Then
                    lngStatusOperacao = enumStatusOperacao.ConcordanciaBackofficeAutomatico
                Else
                    lngStatusOperacao = enumStatusOperacao.ConcordanciaBackoffice
                End If
            
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        lngStatusOperacao, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        enumTipoAcao.ConcordanciaEmSer, _
                        True)
                        
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.Confirmacao, _
                        blnIgnoraGradeHorario, _
                        True)
            
            Case enumAcaoConciliacao.AdmAreaRejeitar
                Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                
                lngStatusOperacao = enumStatusOperacao.Rejeitada
'                lngStatusOperacao = enumStatusOperacao.EmSer

'                If Not xmlOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
'                    If Val(xmlOperacao.selectSingleNode("//IN_ENTR_MANU").Text) = enumIndicadorSimNao.sim Then
'                        lngStatusOperacao = enumStatusOperacao.ManualEmSer
'                    End If
'                End If
                Set xmlOperacao = Nothing
                
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        lngStatusOperacao, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        enumTipoAcao.RejeicaoConcordanciaEmSer, _
                        True)
                
                Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumFuncionalidade.RejeicaoConcordancia, _
                        blnIgnoraGradeHorario, _
                        True)
            
            Case enumAcaoConciliacao.AdmAreaLiberar, enumAcaoConciliacao.AdmAreaLiberarContingencia
                'Pagamento
                If objDomNodeOper.selectSingleNode("ID_PGTO_RECB").Text = "P" Then
                    
                    If enumAcaoConciliacao.AdmAreaLiberarContingencia = True Then
                        lngSequencialJustificativa = IncluirJustificativaConciliacaoOperacao( _
                                            Val(objDomNodeProc.selectSingleNode("TipoJustificativa").Text), _
                                            objDomNodeProc.selectSingleNode("Comentario").Text)
                    End If
                    
                    Select Case Val(objDomNodeProc.selectSingleNode("TP_IF_CRED_DEBT").Text)
                        Case 1
                        
                            If pblnLiquidacaoPorLote = True Then
                                lngStatusOperacao = enumStatusOperacao.LiberadaAutomatica
                            Else
                                lngStatusOperacao = enumStatusOperacao.Liberada
                            End If
                        
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    lngStatusOperacao, _
                                    objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    enumTipoAcao.EnviadaSTR0007, _
                                    True)
                        Case 2
                            Call objOperacao.AlterarStatusAcao( _
                                    objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                    enumStatusOperacao.Liquidada, _
                                    objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                    , _
                                    enumTipoAcao.Liquidacao, _
                                    True)
                                
                    End Select
                
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.Liberacao, _
                            blnIgnoraGradeHorario, _
                            False)
            
                'Recebimento
                Else
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.Liquidada, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            enumTipoAcao.Liquidacao, _
                            True)
                                    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.LiberacaoRecebimento, _
                            blnIgnoraGradeHorario, _
                            False)
                
                End If

                
        End Select
        'Se for processo de Liberação das Corretoras por Lote
        If pblnLiquidacaoPorLote Then
            
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                
            If Not xmlOperacao.selectSingleNode("//CO_VEIC_LEGA") Is Nothing And _
               Not xmlOperacao.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then
                'Busca o Tipo de BackOffice por Veiculo Legal e Sigla de Sistema
                lngTipoBackOfficeLote = fgObterTipoBackOffice(xmlOperacao.selectSingleNode("//CO_VEIC_LEGA").Text, xmlOperacao.selectSingleNode("//SG_SIST_ORIG").Text)
            End If
            Set xmlOperacao = Nothing
            
        End If
    Next
    

    'Envio de mensagens a partir do agrupamento de operações, e processamento da mensagem
    '=========================================================================================
    If pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberar Or pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
        'Pagamento
        If objDomNodeProc.selectSingleNode("ID_PGTO_RECB_GRUP").Text = "P" Then
            Set xmlMsgEnvio = CreateObject("MSXML2.DOMDocument.4.0")
            Call xmlMsgEnvio.loadXML(flMontarMsgEnvioLiquidacaoCorretoras(objDomNodeProc, objOperacao))
            If xmlMsgEnvio.xml <> vbNullString Then
                If pblnLiquidacaoPorLote = True Then
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        0, _
                        enumFuncionalidade.LiberacaoMensagemSTR, _
                        blnIgnoraGradeHorario, _
                        False, _
                        False, _
                        xmlMsgEnvio, _
                        , , , , , True, lngTipoBackOfficeLote)
                    
                 
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                   xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                                                   xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text, _
                                                   1, , , , True, False)
                    
                 
                 Else
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                        0, _
                        enumFuncionalidade.LiberacaoMensagemSTR, _
                        blnIgnoraGradeHorario, _
                        False, _
                        False, _
                        xmlMsgEnvio)
                 
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                   xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                                                   xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text, _
                                                   1)
                 End If

                For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                            xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text, _
                            0)
                
                    Call objOperacao.Alterar(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                             , , , , _
                                             vntSequencialConciliacao)
                Next

                Call objMensagem.AlterarSequenciaConciliacao(xmlMsgEnvio.selectSingleNode("//NU_CTRL_IF").Text, _
                                       fgDtHrStr_To_DateTime(xmlMsgEnvio.selectSingleNode("//DH_INCLUSAO").Text), _
                                       vntSequencialConciliacao, _
                                       1)
            End If
            
        'Recebimento
        Else
            If Not objDomNodeProc.selectSingleNode("NU_CTRL_IF") Is Nothing Then
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.MensagemLiquidada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        True, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.Liquidacao)
                If pblnLiquidacaoPorLote = True Then
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                    1, , , , True, False)
                
                Else
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                                                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                                                    1)
                End If

                For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                            objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                            0)
                
                    Call objOperacao.Alterar(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                             , , , , _
                                             vntSequencialConciliacao)
                Next

                Call objMensagem.AlterarSequenciaConciliacao(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                                       fgDtHrStr_To_DateTime(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text), _
                                       vntSequencialConciliacao, _
                                       1)
            End If
        
        End If
    
    ElseIf pintAcaoConciliacao = enumAcaoConciliacao.AdmAreaLiberarContingencia Then
        If Not objDomNodeProc.selectSingleNode("NU_CTRL_IF") Is Nothing Then
            Call objMensagem.AlterarStatus( _
                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.MensagemLiquidada, _
                    objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                    , _
                    True, _
                    objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                    , _
                    enumTipoAcao.Liquidacao)
        End If
    
    End If
    
    LiquidarCorretoras = vbNullString
    
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objHistOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlMsgEnvio = Nothing
    
    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objHistOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set xmlMsgEnvio = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operações compromissadas genéricas BMA
'--------------------------------------------------------------------------------------------------

Public Function ConciliarCompromissadaGenerica(ByVal pintAcaoConciliacao As enumAcaoConciliacao, _
                                               ByRef objDomNodeProc As MSXML2.IXMLDOMNode, _
                                               ByVal blnIgnoraGradeHorario As Boolean) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objRegra                                As A8LQS.clsRegra

Dim objDomNodeOper                          As MSXML2.IXMLDOMNode
Dim vntSequencialConciliacao                As Variant
Dim blnRegraAutomatica                      As Boolean

On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objRegra = CreateObject("A8LQS.clsRegra")

    'Processamento das Operações
    '=========================================================================================
    For Each objDomNodeOper In objDomNodeProc.selectNodes("Repeat_Operacao/*")
        Select Case pintAcaoConciliacao
            Case enumAcaoConciliacao.BOConcordar
                
                blnRegraAutomatica = objRegra.VerificarRegraAutomatica2(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                                        enumFuncaoSistema.Liberar)
                
                If blnRegraAutomatica Then
                    
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.RegistradaAutomatica, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            , _
                            True)
    
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumFuncionalidade.LiberacaoRecebimento, _
                            blnIgnoraGradeHorario, _
                            True)
        
                Else
                
                    Call objOperacao.AlterarStatusAcao( _
                            objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                            enumStatusOperacao.ConcordanciaBackofficeBMA0013, _
                            objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                            , _
                            , _
                            True)
                
                End If
            
                Call objOperacao.GravarControleSPBOriginal(objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           objDomNodeProc.selectSingleNode("NU_CTRL_CAMR").Text)

            Case enumAcaoConciliacao.AdmAreaRejeitar
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.AConciliarBMA0013, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

            Case enumAcaoConciliacao.AdmAreaLiberar
                Call objOperacao.AlterarStatusAcao( _
                        objDomNodeOper.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                        enumStatusOperacao.Registrada, _
                        objDomNodeOper.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        , _
                        True)

        End Select
    Next

    'Processamento da mensagem
    '=========================================================================================
    Select Case pintAcaoConciliacao
        Case enumAcaoConciliacao.BOConcordar
            If blnRegraAutomatica Then
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.Conciliada, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        True, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.Liberacao)
            
            Else
                Call objMensagem.AlterarStatus( _
                        objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                        objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                        enumStatusMensagem.ConcordanciaBackoffice, _
                        objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                        , _
                        True, _
                        objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                        , _
                        enumTipoAcao.Concordancia)
                        
            End If

        Case enumAcaoConciliacao.AdmAreaRejeitar
            Call objMensagem.AlterarStatus( _
                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.AConciliar, _
                    objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                    , _
                    True, _
                    objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                    , _
                    enumTipoAcao.RejeicaoConcordancia)

        Case enumAcaoConciliacao.AdmAreaLiberar
            Call objMensagem.AlterarStatus( _
                    objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, _
                    objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text, _
                    enumStatusMensagem.Conciliada, _
                    objDomNodeProc.selectSingleNode("DH_ULTI_ATLZ").Text, _
                    , _
                    True, _
                    objDomNodeProc.selectSingleNode("NU_SEQU_CNTR_REPE").Text, _
                    , _
                    enumTipoAcao.Liberacao)

    End Select

    ConciliarCompromissadaGenerica = vbNullString

    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objRegra = Nothing

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objProcessoOperacao = Nothing
    Set objRegra = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarCompromissadaGenerica Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Obter detalhes de Operações e Mensagens para o processo de conciliação
'--------------------------------------------------------------------------------------------------

Function flBuscarDetalhesOperacaoMensagem(pintTipo As enumTipoConciliacao, _
                                          xmlErro As MSXML2.DOMDocument40, _
                                          xmlDetOperacao As MSXML2.DOMDocument40, _
                                          xmlDetMensagem As MSXML2.DOMDocument40, _
                                          xmlOperacao As MSXML2.DOMDocument40, _
                                          xmlMensagem As MSXML2.DOMDocument40, _
                           Optional ByVal pblnAutomatica As Boolean = False, _
                           Optional ByVal pblnConcilacaoCancelamento As Boolean = False) As String

'Buscar detalhes de Mensagem e Operacao

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemSLCC

Dim xmlFiltro                               As MSXML2.DOMDocument40

Dim strOP_NU_SEQU_OPER_ATIV                 As String
Dim strOP_DH_ULTI_ATLZ                      As String
Dim strMSG_NU_CTRL_IF                       As String
Dim strMSG_DH_REGT_MESG_SPB                 As String
Dim strMSG_NU_SEQU_CNTR_REPE                As String
Dim strMSG_DH_ULTI_ATLZ                     As String

Dim strErro                                 As String
Dim strParametros                           As String
Dim strRet                                  As String

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagem = CreateObject("A8LQS.clsMensagemSLCC")
        
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")

    'Verifica quais parametros foram passados
    'Se passou somente um parametro, veio do fluxo automatico de conciliadao
    'Se passaou ambos, eh chamada via tela, comandada por um usuario
    strParametros = ""
    If Not xmlOperacao Is Nothing Then
        If xmlOperacao.xml <> "" Then
            strParametros = strParametros & "O"
        End If
    End If
    If Not xmlMensagem Is Nothing Then
        If xmlMensagem.xml <> "" Then
            strParametros = strParametros & "M"
        End If
    End If
    
    If strParametros = "" Then
        'Operação e Mensagem não informados
        fgAdicionaErro xmlErro, 3060, enumJustificativa.MensagemOperacaoNaoInformado
        Exit Function
    Else

        'Carrega OPERACAO, se tiver sido passada
        If strParametros = "O" Or strParametros = "OM" Then
            strOP_NU_SEQU_OPER_ATIV = xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text
            strOP_DH_ULTI_ATLZ = xmlOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text
            
            xmlFiltro.Load ""
            fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
            
            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroOperacao", ""
            fgAppendNode xmlFiltro, "Grupo_NumeroOperacao", "NumOp", strOP_NU_SEQU_OPER_ATIV

            strRet = objOperacao.ObterDetalheOperacao(xmlFiltro, pblnAutomatica)
            If strRet = "" Then
                'operação não encontrada
                fgAdicionaErro xmlErro, 3061, enumJustificativa.OperacaoNaoEncontrada
                Exit Function
            Else
                xmlDetOperacao.loadXML strRet
                
                If (pintTipo <> enumTipoConciliacao.Bruta And _
                    pintTipo <> enumTipoConciliacao.Bilateral) Or _
                    Not pblnAutomatica Then
                   
                    If xmlDetOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text <> strOP_DH_ULTI_ATLZ Then
                        'operação já alterada
                        fgAdicionaErro xmlErro, 3013, enumJustificativa.OperacaoJaAlterada
                        Exit Function
                    End If
                    
                End If
            End If
        End If

        'Carrega MENSAGEM, se tiver sido passada
        If strParametros = "M" Or strParametros = "OM" Then
            strMSG_NU_CTRL_IF = xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text
            strMSG_DH_REGT_MESG_SPB = xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text
            strMSG_NU_SEQU_CNTR_REPE = xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text
            
            If Not xmlMensagem.selectSingleNode("//Grupo_Mensagem/DH_ULTI_ATLZ") Is Nothing Then
                strMSG_DH_ULTI_ATLZ = xmlMensagem.selectSingleNode("//Grupo_Mensagem/DH_ULTI_ATLZ").Text
            End If
            
            'Monta filtro
            xmlFiltro.Load ""
            fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Chave", ""
            fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_CTRL_IF", strMSG_NU_CTRL_IF
            fgAppendAttribute xmlFiltro, "Grupo_Chave", "DH_REGT_MESG_SPB", fgDtHrXML_To_Oracle(strMSG_DH_REGT_MESG_SPB)
            fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_SEQU_CNTR_REPE", strMSG_NU_SEQU_CNTR_REPE
            
            'Verifica se é Despesa CETIP
            If Not xmlMensagem.selectSingleNode("//NU_COMD_OPER") Is Nothing Then
                If InStr(1, xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text, "CUSTO CETIP") > 0 Then
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                    fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumeroComando", xmlMensagem.selectSingleNode("//NU_COMD_OPER").Text
                    
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_SegregaBackOffice", ""
                    fgAppendNode xmlFiltro, "Grupo_SegregaBackOffice", "Segrega", "False"
                End If
            End If
                    
            strRet = objMensagem.ObterDetalheMensagemCamara(xmlFiltro, pblnAutomatica)
            
            If strRet = "" Then
                fgAdicionaErro xmlErro, 3062, enumJustificativa.MensagemNaoEncontrada
                Exit Function
            Else
                xmlDetMensagem.loadXML strRet
                
                If (pintTipo <> enumTipoConciliacao.Bruta And _
                    pintTipo <> enumTipoConciliacao.Bilateral) Or _
                    Not pblnAutomatica Then
                   
                    If xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text <> strMSG_DH_ULTI_ATLZ Then
                        fgAdicionaErro xmlErro, 3014, enumJustificativa.MensagemJaAlterada
                        Exit Function
                    End If
                
                End If
            End If
        End If

        If strParametros = "O" Then
            'So passou operacao, entao procura a(s) mensagem(ens)
            
            xmlFiltro.loadXML ""
            fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
                        
            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_DataOperacao", ""
            fgAppendNode xmlFiltro, "Grupo_DataOperacao", "DataOperacao", fgDtXML_To_Oracle(xmlDetOperacao.selectSingleNode("//DT_OPER_ATIV").Text)
                                   
            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Status", ""
            
            If (pintTipo <> enumTipoConciliacao.Bruta And _
                pintTipo <> enumTipoConciliacao.Bilateral) Or _
                Not pblnAutomatica Then
               
                'Registro de Operações
               
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusMensagem.AConciliar
            
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", ""
                fgAppendNode xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", xmlDetOperacao.selectSingleNode("//CO_EMPR").Text
                
                If CLng(xmlDetOperacao.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMC Then
                    If CLng(xmlDetOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao _
                    Or CLng(xmlDetOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica _
                    Or CLng(xmlDetOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroOperacoesBMC Then
                        'Tipo de Negociacao 1 e 2
                        If CLng(xmlDetOperacao.selectSingleNode("//TP_NEGO").Text) = 1 _
                        Or CLng(xmlDetOperacao.selectSingleNode("//TP_NEGO").Text) = 2 Then
                            If CLng(xmlDetOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.AConciliarRegistro Then
                                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentNegocioBMC", ""
                                fgAppendNode xmlFiltro, "Grupo_IdentNegocioBMC", "NrIdefNegoBMC", xmlDetOperacao.selectSingleNode("//NR_IDEF_NEGO_BMC").Text
                            Else
                                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ChACAM", ""
                                fgAppendNode xmlFiltro, "Grupo_ChACAM", "ChACAM", xmlDetOperacao.selectSingleNode("//CD_ASSO_CAMB").Text
                            End If
                        'Tipo de Negociacao 3
                        ElseIf CLng(xmlDetOperacao.selectSingleNode("//TP_NEGO").Text) = 3 Then
                            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ContratacaoSISBACEN", ""
                            fgAppendNode xmlFiltro, "Grupo_ContratacaoSISBACEN", "CntrSISBACEN", xmlDetOperacao.selectSingleNode("//CO_CNTR_SISB").Text
                        End If
                    End If
                Else
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                    If pblnConcilacaoCancelamento Then
                        fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", xmlOperacao.selectSingleNode("//NU_COMD_OPER").Text
                    Else
                        fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text
                    End If
                End If
            Else
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", ""
                fgAppendNode xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", xmlDetOperacao.selectSingleNode("//CO_EMPR").Text
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_LocalLiquidacao", ""
                fgAppendNode xmlFiltro, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP
                
                If xmlDetOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Registrada Or _
                   xmlDetOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.RegistradaAutomatica Then
                    
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusMensagem.AConciliar
                
                ElseIf xmlDetOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.ConcordanciaBackofficeAutomatico Then
                    
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusMensagem.ConcordanciaBackofficeAutomatico
                
                ElseIf xmlDetOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.PagamentoBackofficeAutomatico Then
                    
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusMensagem.PagamentoBackofficeAutomatico
                
                End If
            
                If pintTipo = enumTipoConciliacao.Bruta Then
                    If xmlDetOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DespesasCETIP Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_LayoutEntrada", ""
                        fgAppendNode xmlFiltro, "Grupo_LayoutEntrada", "LayoutEntrada", "88"
                            
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_SegregaBackOffice", ""
                        fgAppendNode xmlFiltro, "Grupo_SegregaBackOffice", "Segrega", "False"
                        
                    ElseIf xmlDetOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_LayoutEntrada", ""
                        fgAppendNode xmlFiltro, "Grupo_LayoutEntrada", "LayoutEntrada", "80"
                            
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                        If pblnConcilacaoCancelamento Then
                            fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", "SPR#" & xmlOperacao.selectSingleNode("//NU_COMD_OPER").Text
                        Else
                            fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", "SPR#" & xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text
                        End If
                    
                    ElseIf xmlDetOperacao.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_LayoutEntrada", ""
                        fgAppendNode xmlFiltro, "Grupo_LayoutEntrada", "LayoutEntrada", "80"
                            
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                        If pblnConcilacaoCancelamento Then
                            fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", "TERMO#" & xmlOperacao.selectSingleNode("//NU_COMD_OPER").Text
                        Else
                            fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumCMD", "TERMO#" & xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text
                        End If
                    
                    Else
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroControleLTR", ""
                        fgAppendNode xmlFiltro, "Grupo_NumeroControleLTR", "NumeroControleLTR", xmlDetOperacao.selectSingleNode("//NU_CTRL_MESG_SPB_ORIG").Text
                    
                    End If
                    
                    If xmlDetOperacao.selectSingleNode("//TP_OPER").Text <> enumTipoOperacaoLQS.DespesasCETIP Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_DebitoCredito", ""
                        If Val(xmlDetOperacao.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                            fgAppendNode xmlFiltro, "Grupo_DebitoCredito", "DebitoCredito", enumTipoDebitoCredito.Debito
                        Else
                            fgAppendNode xmlFiltro, "Grupo_DebitoCredito", "DebitoCredito", enumTipoDebitoCredito.Credito
                        End If
                    End If
                    
                Else
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentificadorPartCamara", ""
                    fgAppendNode xmlFiltro, "Grupo_IdentificadorPartCamara", "IdentificadorPartCamara", xmlDetOperacao.selectSingleNode("//CO_PARP_CAMR").Text
                    
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CNPJ", ""
                    fgAppendNode xmlFiltro, "Grupo_CNPJ", "CNPJ", xmlDetOperacao.selectSingleNode("//CO_CNPJ_VEIC_LEGA").Text
                    
                    If Val(xmlDetOperacao.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Interno Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Mensagem", ""
                        fgAppendNode xmlFiltro, "Grupo_Mensagem", "Mensagem", "LTR0007"
                        
                        'RATS 586/001
                        'Problema produção - Cassiano - 07/11/2006
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentificadorPartCamaraContraparte", ""
                        fgAppendNode xmlFiltro, "Grupo_IdentificadorPartCamaraContraparte", "IdentificadorPartCamaraContraparte", xmlDetOperacao.selectSingleNode("//CO_CNPT_CAMR").Text
                    Else
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Mensagem", ""
                        fgAppendNode xmlFiltro, "Grupo_Mensagem", "Mensagem", "LTR0001"
                        fgAppendNode xmlFiltro, "Grupo_Mensagem", "Mensagem", "LTR0005R2"
                    End If
                
                End If
            
            End If
                        
            strRet = objMensagem.ObterDetalheMensagemCamara(xmlFiltro, pblnAutomatica)
            
            If strRet = "" Then
                If pintTipo = enumTipoConciliacao.Bilateral And Val(xmlDetOperacao.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Interno Then
                    fgRemoveNode xmlFiltro, "Grupo_IdentificadorPartCamaraContraparte"
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CNPJContraparte", ""
                    fgAppendNode xmlFiltro, "Grupo_CNPJContraparte", "CNPJContraparte", xmlDetOperacao.selectSingleNode("//CO_CNPJ_CNPT").Text
                
                    strRet = objMensagem.ObterDetalheMensagemCamara(xmlFiltro, pblnAutomatica)
                    
                    If strRet = "" Then
                        'não foi possível encontrar a mensagem pela operação
                        fgAdicionaErro xmlErro, 3063, enumJustificativa.MensagemNaoEncontrada
                        Exit Function
                    Else
                        xmlDetMensagem.loadXML strRet
                    End If
                Else
                    'não foi possível encontrar a mensagem pela operação
                    fgAdicionaErro xmlErro, 3063, enumJustificativa.MensagemNaoEncontrada
                    Exit Function
                End If
            Else
                xmlDetMensagem.loadXML strRet
            End If
        End If

        If strParametros = "M" Then
            'So passou mensagem, entao procura a(s) operacao(oes)
            
            xmlFiltro.loadXML ""
            fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
            
            If pintTipo = enumTipoConciliacao.Registro Then
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", ""
                fgAppendNode xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", xmlDetMensagem.selectSingleNode("//CO_EMPR").Text
                
                If CLng(xmlDetMensagem.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.BMC Then
                    If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0011" Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
                        fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", "121"
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CanalSISBACENCorretora", ""
                        fgAppendNode xmlFiltro, "Grupo_CanalSISBACENCorretora", "CanalSISBACENCorretora", xmlDetMensagem.selectSingleNode("//CO_SISB_COTR").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NaturezaMovimento", ""
                        If Val(xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                            fgAppendNode xmlFiltro, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Credito
                        Else
                            fgAppendNode xmlFiltro, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Debito
                        End If
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CodigoPraca", ""
                        fgAppendNode xmlFiltro, "Grupo_CodigoPraca", "CodigoPraca", xmlDetMensagem.selectSingleNode("//CO_PRAC").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CodigoMoedaEstrangeira", ""
                        fgAppendNode xmlFiltro, "Grupo_CodigoMoedaEstrangeira", "CodigoMoedaEstrangeira", xmlDetMensagem.selectSingleNode("//CO_MOED_ESTR").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ValorMoedaEstrangeira", ""
                        fgAppendNode xmlFiltro, "Grupo_ValorMoedaEstrangeira", "ValorMoedaEstrangeira", xmlDetMensagem.selectSingleNode("//VA_MOED_ESTR").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TaxaNegociacao", ""
                        fgAppendNode xmlFiltro, "Grupo_TaxaNegociacao", "TaxaNegociacao", xmlDetMensagem.selectSingleNode("//PE_TAXA_NEGO").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ValorFinanceiro", ""
                        fgAppendNode xmlFiltro, "Grupo_ValorFinanceiro", "ValorFinanceiro", xmlDetMensagem.selectSingleNode("//VA_FINC").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_DataLiquidacaoMoedaEstr", ""
                        fgAppendNode xmlFiltro, "Grupo_DataLiquidacaoMoedaEstr", "DataLiquidacaoMoedaEstr", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_LIQU_OPER_ATIV_MOED_ESTR").Text)
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_OrderBy", ""
                        fgAppendNode xmlFiltro, "Grupo_OrderBy", "OrderBy", "NU_SEQU_OPER_ATIV"

                    ElseIf xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0015" Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentNegocioBMC", ""
                        fgAppendNode xmlFiltro, "Grupo_IdentNegocioBMC", "NrIdefNegoBMC", xmlDetMensagem.selectSingleNode("//NR_IDEF_NEGO_BMC").Text

                    ElseIf xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0005" Then
'                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ChACAM", ""
'                        fgAppendNode xmlFiltro, "Grupo_ChACAM", "ChACAM", xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text
'
'                    Else
'                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ContratacaoSISBACEN", ""
'                        fgAppendNode xmlFiltro, "Grupo_ContratacaoSISBACEN", "CntrSISBACEN", xmlDetMensagem.selectSingleNode("//CO_CNTR_SISB").Text
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ChACAMContratacaoSISBACEN", ""
                        fgAppendNode xmlFiltro, "Grupo_ChACAMContratacaoSISBACEN", "ChACAM", xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text
                        fgAppendNode xmlFiltro, "Grupo_ChACAMContratacaoSISBACEN", "CntrSISBACEN", xmlDetMensagem.selectSingleNode("//CO_CNTR_SISB").Text
                    
                    End If
                Else
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                    fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumeroCMD", xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text
                End If
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Data", ""
                fgAppendNode xmlFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_OPER").Text)
                fgAppendNode xmlFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_OPER").Text)
            
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Status", ""
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.AConciliar
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.AConciliarRegistro
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.AConciliarAceite
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ManualEmSer
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.EmSer
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.Liberada
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.LiberadaAutomatica
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.Confirmada
                fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.Respondida
            
            ElseIf pintTipo = enumTipoConciliacao.Bruta Or pintTipo = enumTipoConciliacao.Bilateral Then
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", ""
                fgAppendNode xmlFiltro, "Grupo_BancoLiquidante", "BancoLiquidante", xmlDetMensagem.selectSingleNode("//CO_EMPR").Text
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_LocalLiquidacao", ""
                fgAppendNode xmlFiltro, "Grupo_LocalLiquidacao", "LocalLiquidacao", enumLocalLiquidacao.CETIP
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Data", ""
                fgAppendNode xmlFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_OPER").Text)
                fgAppendNode xmlFiltro, "Grupo_Data", "DataFim", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_OPER").Text)
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Status", ""
                If xmlDetMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusMensagem.AConciliar Then
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.Registrada
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica
                Else
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaBackoffice
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.PagamentoBackoffice
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.Registrada
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.RegistradaAutomatica
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackofficeAutomatico
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.PagamentoBackofficeAutomatico
                End If
                
                If pintTipo = enumTipoConciliacao.Bruta Then
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoLiquidacao", ""
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoContraparte", ""
                    
                    If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                        fgAppendNode xmlFiltro, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta
                        fgAppendNode xmlFiltro, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NaturezaMovimento", ""
                        If Val(xmlDetMensagem.selectSingleNode("//CAMPO_IN_OPER_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                            fgAppendNode xmlFiltro, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Credito
                        Else
                            fgAppendNode xmlFiltro, "Grupo_NaturezaMovimento", "NaturezaMovimento", enumTipoDebitoCredito.Debito
                        End If
                    Else
                        fgAppendNode xmlFiltro, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bruta
                        fgAppendNode xmlFiltro, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo
                    End If
                
                    If InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "SPR#") <> 0 Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
                        fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                        fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumeroComando", Mid$(xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "#") + 1)
                    
                    ElseIf InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "TERMO#") <> 0 Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
                        fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.AntecipacaoResgateContratoTERMO
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroComando", ""
                        fgAppendNode xmlFiltro, "Grupo_NumeroComando", "NumeroComando", Mid$(xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "#") + 1)
                    
                    ElseIf InStr(1, xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text, "CUSTO CETIP") > 0 Then
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
                        fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.DespesasCETIP
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_SegregaBackOffice", ""
                        fgAppendNode xmlFiltro, "Grupo_SegregaBackOffice", "Segrega", "False"
                        
                    Else
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_NumeroControleLTR", ""
                        fgAppendNode xmlFiltro, "Grupo_NumeroControleLTR", "NumeroControleLTR", xmlDetMensagem.selectSingleNode("//NU_CTRL_CAMR").Text
                    
                    End If
                
                Else
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoLiquidacao", ""
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoContraparte", ""
                    
                    If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0007" Then
                        fgAppendNode xmlFiltro, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral
                        fgAppendNode xmlFiltro, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Interno
                        
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CNPJ", ""
                        fgAppendNode xmlFiltro, "Grupo_CNPJ", "CNPJ", xmlDetMensagem.selectSingleNode("//CO_CNPJ_VEIC_LEGA").Text
                    
                        'RATS 586/001
                        'Problema produção - Cassiano - 07/11/2006
                        If Val(xmlDetMensagem.selectSingleNode("//CO_CNPT_CAMR").Text) = 0 Then
                            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_CNPJContraparte", ""
                            fgAppendNode xmlFiltro, "Grupo_CNPJContraparte", "CNPJContraparte", xmlDetMensagem.selectSingleNode("//CO_CNPJ_CNPT").Text
                        Else
                            fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentificadorPartCamaraContraparte", ""
                            fgAppendNode xmlFiltro, "Grupo_IdentificadorPartCamaraContraparte", "IdentificadorPartCamaraContraparte", xmlDetMensagem.selectSingleNode("//CO_CNPT_CAMR").Text
                        End If
                    
                    ElseIf xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0001" Then
                        fgAppendNode xmlFiltro, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral
                        fgAppendNode xmlFiltro, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo
                        
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentificadorPartCamara", ""
                        fgAppendNode xmlFiltro, "Grupo_IdentificadorPartCamara", "IdentificadorPartCamara", xmlDetMensagem.selectSingleNode("//CO_PARP_CAMR").Text
                        
                    ElseIf xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "LTR0005R2" Then
                        fgAppendNode xmlFiltro, "Grupo_TipoLiquidacao", "TipoLiquidacao", enumTipoLiquidacao.Bilateral
                        fgAppendNode xmlFiltro, "Grupo_TipoContraparte", "TipoContraparte", enumTipoContraparte.Externo
                    
                        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_IdentificadorPartCamara", ""
                        fgAppendNode xmlFiltro, "Grupo_IdentificadorPartCamara", "IdentificadorPartCamara", xmlDetMensagem.selectSingleNode("//CO_PARP_CAMR").Text
                        
                    End If
                    
                End If
                    
            ElseIf pintTipo = enumTipoConciliacao.MultilateralTitulos Then
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Titulo", ""
                fgAppendNode xmlFiltro, "Grupo_Titulo", "Titulo", xmlDetMensagem.selectSingleNode("//NU_ATIV_MERC").Text
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_ContaSELIC", ""
                fgAppendNode xmlFiltro, "Grupo_ContaSELIC", "ContaSELIC", xmlDetMensagem.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_DataVencimento", ""
                fgAppendNode xmlFiltro, "Grupo_DataVencimento", "DataIni", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_VENC").Text)
                fgAppendNode xmlFiltro, "Grupo_DataVencimento", "DataFim", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_VENC").Text)
                
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Data", ""
                fgAppendNode xmlFiltro, "Grupo_Data", "DataIni", fgDtXML_To_Oracle(xmlDetMensagem.selectSingleNode("//DT_OPER").Text)
            
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_TipoOperacao", ""
                fgAppendNode xmlFiltro, "Grupo_TipoOperacao", "TipoOperacao", enumTipoOperacaoLQS.LiquidacaoFisicaOperacaoBMA
            
                fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Status", ""
                If xmlDetMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusMensagem.AConciliar Then
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.AConciliar
                Else
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.ConcordanciaBackoffice
                    fgAppendNode xmlFiltro, "Grupo_Status", "Status", enumStatusOperacao.DiscordanciaBackoffice
                End If
            
            End If
            
            strRet = objOperacao.ObterDetalheOperacao(xmlFiltro, pblnAutomatica)
            
            If strRet = "" Then
               'não foi possível encontrar a operação pela mensagem
                fgAdicionaErro xmlErro, 3064, enumJustificativa.OperacaoNaoEncontrada
                Exit Function
            Else
                xmlDetOperacao.loadXML strRet
            End If
        End If

    End If
    
    flBuscarDetalhesOperacaoMensagem = strErro

End Function

'Verifica dados da contraparte de uma LTR0007.

Private Function flVerificarContraparteLTR0007(ByVal strNU_CTRL_IF As String, _
                                               ByVal strNU_SEQU_CNTR_REPE As String, _
                                               ByVal strIN_OPER_DEBT_CRED As String, _
                                               ByRef intStatusContraparte As Integer, _
                                      Optional ByRef strDataHoraContraparte As String = vbNullString) As Long

Dim strSQL                                  As String
Dim rsLeitura                               As ADODB.Recordset
Dim intDebCredCnpt                          As Integer

    On Error GoTo ErrorHandler
    
    intDebCredCnpt = IIf(Val(strIN_OPER_DEBT_CRED) = enumTipoDebitoCredito.Credito, _
                                                     enumTipoDebitoCredito.Debito, _
                                                     enumTipoDebitoCredito.Credito)

    Set rsLeitura = CreateObject("ADODB.Recordset")

    strSQL = " SELECT A.CO_ULTI_SITU_PROC, A.DH_REGT_MESG_SPB                   " & vbNewLine & _
             " FROM   A8.TB_MESG_RECB_ENVI_SPB A,                               " & vbNewLine & _
             "        A8.TB_MESG_RECB_SPB_CNCL B                                " & vbNewLine & _
             " WHERE  A.NU_CTRL_IF           =  B.NU_CTRL_IF                    " & vbNewLine & _
             " AND    A.NU_SEQU_CNTR_REPE    =  B.NU_SEQU_CNTR_REPE             " & vbNewLine & _
             " AND    A.DH_REGT_MESG_SPB     =  B.DH_REGT_MESG_SPB              " & vbNewLine & _
             " AND    B.NU_CTRL_IF           = '" & strNU_CTRL_IF & "'          " & vbNewLine & _
             " AND    B.NU_SEQU_CNTR_REPE    =  " & strNU_SEQU_CNTR_REPE & "    " & vbNewLine & _
             " AND    B.IN_OPER_DEBT_CRED    =  " & intDebCredCnpt & "          " & vbNewLine & _
             " AND    A.CO_MESG_SPB          = 'LTR0007'                        "

    Set rsLeitura = fgQuerySQL(strSQL)
    
    If Not rsLeitura.EOF Then
        intStatusContraparte = rsLeitura!CO_ULTI_SITU_PROC
        strDataHoraContraparte = fgDtHr_To_Xml(rsLeitura!DH_REGT_MESG_SPB)
    End If
    
    Set rsLeitura = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set rsLeitura = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarContraparteLTR0007 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Utilizada na Conciliação Bruta / Bilateral. Quando uma condorcância do backoffice é rejeitada pelo administrador
' da área, o número de conciliação deve ser excluído da base de dados.

Private Function flExcluirConciliacao(Optional ByVal pvntSequenciaOperacao As Variant = vbNullString, _
                                      Optional ByVal pintStatusOperacao As enumStatusOperacao, _
                                      Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                                      Optional ByVal pstrDTRegistroMensagemSPB As String, _
                                      Optional ByVal pintStatusMensagem As enumStatusMensagem, _
                                      Optional ByVal pintTipoAcao As Integer = 0, _
                                      Optional ByVal plngJustificativa As Long = 0, _
                                      Optional ByVal plngNumeroSequenciaContadorRepeticao As Long = 1) As Boolean

Dim objHistoricoOperacao                    As A8LQS.clsHistSituacaoOperacao
Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    If pvntSequenciaOperacao <> vbNullString Then
        strSQL = " DELETE FROM A8.TB_CNCL_OPER_ATIV       " & vbNewLine & _
                 " WHERE       NU_SEQU_OPER_ATIV        = " & pvntSequenciaOperacao
    
        Call fgExecuteSQL(strSQL)
        
        'Armazenar o Histórico da Situacao da Operação
        Set objHistoricoOperacao = CreateObject("A8LQS.clsHistSituacaoOperacao")
        objHistoricoOperacao.Incluir pvntSequenciaOperacao, _
                                     pintStatusOperacao, _
                                     plngJustificativa, _
                                     pintTipoAcao
    End If
    
    If pstrNumeroControleIF <> vbNullString Then
        'Armazenar o Histórico da Situacao da Mensagem
        Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        objHistoricoMensagem.Incluir pstrNumeroControleIF, _
                                     pstrDTRegistroMensagemSPB, _
                                     pintStatusMensagem, _
                                     pintTipoAcao, _
                                     , _
                                     plngNumeroSequenciaContadorRepeticao
                                     
        If Not ObjectContext Is Nothing Then
            ObjectContext.SetComplete
        End If
    End If

    Set objHistoricoOperacao = Nothing
    Set objHistoricoMensagem = Nothing
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Processo de conciliação de operações com mensagens ( BMA0002, CTP1002, BMC0005 e BMC0015 )
'--------------------------------------------------------------------------------------------------

Public Function ConciliarRegistroOperacao(ByVal pintAcao As Integer, _
                                 Optional ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                 Optional ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                 Optional ByVal pblnAutomatica As Boolean = True, _
                                 Optional ByRef plngStatusOperacao As Long = 0, _
                                 Optional ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                 Optional ByVal pblnConcilacaoCancelamento As Boolean = False) As String

    '//////////////////////////////////////////////////////////////////////////////////////////////
    '//
    '//     Realiza a conciliação de operações e mensagens (1 pra 1)
    '//
    '//     Se for passado somente operação, deve procurar msg correspondente, e vice-versa.
    '//
    '//     Retorno: OK (string em branco), ou ERROS (string XML com colecao de erros)
    '//
    '//     DRU´s:  BMA - Registro de Operacoes Definitivas
    '//             CETIP - Registro de Operacoes
    '//             BMC - Registro de Operacoes
    '//
    '//////////////////////////////////////////////////////////////////////////////////////////////

    '////////////////////////////////
    '(*) Campos para a busca Operação e Mensagem    (BMA0002)
    '////////////////////////////////
    '    Tag Mensagem                                   Tabela Mensagem                 Tabela Operação
    '-------------------------------------------------------------------------------------------------------------
    '(*) NumOpNegcBMA                                   NU_COMD_OPER                    NU_COMD_OPER
    '(*) DtMovto                                        DT_OPER                         DT_OPER_ATIV
   
    '////////////////////////////////
    '// Campos para ...                             (CTP1002)
    '////////////////////////////////
    '    Tag Mensagem                                   Tabela Mensagem                 Tabela Operação
    '-------------------------------------------------------------------------------------------------------------
    '(*) NumOpCTP                                       NU_COMD_OPER                    NU_COMD_OPER
    '(*) DtMovto                                        DT_OPER                         DT_OPER_ATIV
    
    '////////////////////////////////
    '// Campos para ...                             (BMC0005)
    '////////////////////////////////
    '    Tag Mensagem                                   Tabela Mensagem                 Tabela Operação
    '-------------------------------------------------------------------------------------------------------------
    '(*) CodContrSISBACEN                               CO_CNTR_SISB                    CO_CNTR_SISB
    '(*) DtMovto                                        DT_OPER                         DT_OPER_ATIV
    
Dim vntPasso                                As Variant
Dim xmlErros                                As MSXML2.DOMDocument40

Dim xmlDetOperacao                          As MSXML2.DOMDocument40
Dim xmlDetMensagem                          As MSXML2.DOMDocument40
Dim xmlMensagemRecb                         As MSXML2.DOMDocument40

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objPO                                   As A8LQS.clsProcessoOperacao
Dim objRegra                                As A8LQS.clsRegra
Dim blnPassouOperacao                       As Boolean
Dim blnPassouMensagem                       As Boolean
Dim blnSoPassouOperacao                     As Boolean
Dim blnSoPassouMensagem                     As Boolean
Dim blnVerificaCamposOpcionais              As Boolean
Dim lngTipoOperacao                         As enumTipoOperacaoLQS

Dim strRet                                  As String
Dim blnAux                                  As Boolean
Dim lngStatus                               As Long
Dim lngStatusMensagem                       As Long
Dim lngFuncaoSistema                        As Long
Dim lngTipoAcao                             As Long
Dim blnCancelamentoCETIP                    As Boolean      'se é uma operação de Cancelamento CETIP (CO_OPER_CETIP >= 100)
Dim blnEntradaManual                        As Boolean
Dim blnBMC0005BMC0015                       As Boolean
Dim lngSequencialConciliacaoOperacao        As Long         'Para guardar o sequencia gerado na tabela de conciliacao

On Error GoTo ErrorHandler

    ConciliarRegistroOperacao = ""
    blnVerificaCamposOpcionais = True
    blnCancelamentoCETIP = False
    blnEntradaManual = False
    blnBMC0005BMC0015 = False
    lngTipoAcao = 0

    vntPasso = 10

    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")

    Set xmlDetOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDetMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    'se não passou os xmls, inicializa mesmo assim
    blnPassouOperacao = True
    blnPassouMensagem = True

    vntPasso = 20
    If xmlOperacao Is Nothing Then
        Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        blnPassouOperacao = False
    Else
        If xmlOperacao.selectNodes("//NaoVerificaCampoOpcionais").length > 0 Then
            blnVerificaCamposOpcionais = False
        End If
    End If
    If xmlMensagem Is Nothing Then
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        blnPassouMensagem = False
    End If

    blnSoPassouMensagem = (blnPassouMensagem And (Not blnPassouOperacao))
    blnSoPassouOperacao = (blnPassouOperacao And (Not blnPassouMensagem))

    vntPasso = 30
    strRet = flBuscarDetalhesOperacaoMensagem(enumTipoConciliacao.Registro, xmlErros, xmlDetOperacao, xmlDetMensagem, xmlOperacao, xmlMensagem, pblnAutomatica, pblnConcilacaoCancelamento)

    'Carlos - Para a conciliacao de cancelamento não considera os dados da tabela e sim do recebido do legado
    If pblnConcilacaoCancelamento Then
        xmlDetOperacao.loadXML xmlOperacao.xml
    End If

    vntPasso = 40
    If xmlErros.selectNodes("//Grupo_ErrorInfo").length > 0 Then
        'Erros na busca dos XML´s
        ConciliarRegistroOperacao = xmlErros.xml
        Exit Function
    End If
    
    'Se passou somente mensagem e estah no fluxo automatico, verifica se esta com fluxo auto p/ operacao
    If pblnAutomatica And blnSoPassouMensagem Then
        Set objRegra = CreateObject("A8LQS.clsRegra")
        vntPasso = 50
        If Not objRegra.VerificarRegraAutomatica2(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                  enumFuncaoSistema.Conciliar) Then
            'Esta operacao nao esta com fluxo automatico
            ConciliarRegistroOperacao = ""
            Exit Function
        End If
    End If
    
    lngTipoOperacao = Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text)
    
    '///////////////////////////////////////////////////////////////////////////////////////////////
    '// Aqui já estamos com XML´s na mão, então faz a conciliação no "sentido bíblico da palavra"
    '///////////////////////////////////////////////////////////////////////////////////////////////
    
    vntPasso = 60
    If pintAcao = enumAcaoConciliacao.BOConcordar Then
        
        Select Case UCase(xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text)
            Case "BMA0002"
                vntPasso = 100
                If xmlDetMensagem.selectSingleNode("//CO_VEIC_LEGA").Text _
                <> xmlDetOperacao.selectSingleNode("//CO_VEIC_LEGA").Text Then
                    
                    fgAdicionaErro xmlErros, 3065, enumJustificativa.VeicluloLegal
                End If
                vntPasso = 110
                If xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text _
                <> xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text Then

                    fgAdicionaErro xmlErros, 3030, enumJustificativa.NumeroComando
                End If
                vntPasso = 120
                If xmlDetMensagem.selectSingleNode("//DT_OPER").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3066, enumJustificativa.DataOperacao
                End If
                vntPasso = 130
                If xmlDetMensagem.selectSingleNode("//IN_OPER_DEBT_CRED").Text _
                <> xmlDetOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text Then

                    fgAdicionaErro xmlErros, 3067, enumJustificativa.DebitoCredito
                End If
                vntPasso = 140
                If xmlDetMensagem.selectSingleNode("//NU_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//NU_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3029, enumJustificativa.IdentificadorTitulo
                End If
                vntPasso = 150
                If xmlDetMensagem.selectSingleNode("//DT_VENC").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_VENC_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3068, enumJustificativa.DataVencimentoTitulo
                End If
                vntPasso = 160
                If xmlDetMensagem.selectSingleNode("//DT_LIQU").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_LIQU_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3069, enumJustificativa.DataLiquidacao
                End If
                vntPasso = 170
                If xmlDetMensagem.selectSingleNode("//QT_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//QT_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3070, enumJustificativa.QuantidadeTitulo
                End If
                vntPasso = 180
                If xmlDetMensagem.selectSingleNode("//PU_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//PU_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3071, enumJustificativa.PU
                End If
                
                vntPasso = 190
                If xmlDetMensagem.selectSingleNode("//VA_FINC").Text _
                <> xmlDetOperacao.selectSingleNode("//VA_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3072, enumJustificativa.ValorFinanceiro
                End If
                vntPasso = 200
                If xmlDetMensagem.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text _
                <> xmlDetOperacao.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text Then

                    fgAdicionaErro xmlErros, 3073, enumJustificativa.ContaCustodia
                End If
                vntPasso = 210
                If xmlDetMensagem.selectSingleNode("//PE_TAXA_NEGO").Text _
                <> xmlDetOperacao.selectSingleNode("//PE_TAXA_NEGO").Text Then

                    fgAdicionaErro xmlErros, 3074, enumJustificativa.TaxaNegociada
                End If
                vntPasso = 220
                If xmlDetMensagem.selectSingleNode("//CO_TITL_CUTD").Text _
                <> xmlDetOperacao.selectSingleNode("//CO_TITL_CUTD").Text Then

                    fgAdicionaErro xmlErros, 3075, enumJustificativa.TitulaCustodiante
                End If

                'Bate o Tipo de Operação com a Tag "TpNegcBMA" (veja XLS de domínios do BACEN para
                vntPasso = 230
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                                    enumTipoOperacaoLQS.DefinitivaCobertaBMA, _
                                    enumTipoOperacaoLQS.DefinitivaDescobertaBMA) Then
                                    
                    If Not fgIN(Val(xmlDetMensagem.selectSingleNode("//TP_NEGO_BMA").Text), enumTipoNegociacaoBMA.DefinitivaD0) Then
                        fgAdicionaErro xmlErros, 3076, enumJustificativa.TipoNegociacao
                    End If
                End If
                vntPasso = 240
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                            enumTipoOperacaoLQS.OperacaoTermoCobertaBMA, _
                            enumTipoOperacaoLQS.OperacaoTermodesCobertaBMA) Then
                                    
                    If Not fgIN(Val(xmlDetMensagem.selectSingleNode("//TP_NEGO_BMA").Text), _
                                    enumTipoNegociacaoBMA.TermoLeilao, _
                                    enumTipoNegociacaoBMA.TermoPapelDecorridoComCorrecao, _
                                    enumTipoNegociacaoBMA.TermoPapelDecorridoSemCorrecao) Then
                                    
                        fgAdicionaErro xmlErros, 3076, enumJustificativa.TipoNegociacao
                    End If
                End If
                vntPasso = 250
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA, _
                                    enumTipoOperacaoLQS.CompromissadaGenericaATermo, _
                                    enumTipoOperacaoLQS.CompromissadaGenericaAVista, _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaTermo, _
                                    enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA, _
                                    enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA) Then
                                    
                    vntPasso = 260
                    If xmlDetMensagem.selectSingleNode("//DT_OPER_ATIV_RETN").Text _
                    <> xmlDetOperacao.selectSingleNode("//DT_OPER_ATIV_RETN").Text Then
                    
                        fgAdicionaErro xmlErros, 3106, enumJustificativa.DataRetorno
                    End If

                    vntPasso = 270
                    If xmlDetMensagem.selectSingleNode("//PZ_DIAS_RETN_OPER_ATIV").Text _
                    <> xmlDetOperacao.selectSingleNode("//PZ_DIAS_RETN_OPER_ATIV").Text Then
                    
                        fgAdicionaErro xmlErros, 3107, enumJustificativa.PrazoDiasRetorno
                    End If

                    vntPasso = 280
                    If xmlDetMensagem.selectSingleNode("//VA_OPER_ATIV_RETN").Text _
                    <> xmlDetOperacao.selectSingleNode("//VA_OPER_ATIV_RETN").Text Then
                    
                        fgAdicionaErro xmlErros, 3108, enumJustificativa.ValorRetorno
                    End If
                    
                End If
                
                vntPasso = 290
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaCobertaBMA, _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaDescobertaBMA, _
                                    enumTipoOperacaoLQS.CompromissadaGenericaATermo, _
                                    enumTipoOperacaoLQS.CompromissadaGenericaAVista, _
                                    enumTipoOperacaoLQS.CompromissadaEspecificaTermo) Then

                    If Not fgIN(Val(xmlDetMensagem.selectSingleNode("//TP_NEGO_BMA").Text), enumTipoNegociacaoBMA.Compromissada) Then
                        fgAdicionaErro xmlErros, 3076, enumJustificativa.TipoNegociacao
                    End If
                End If
                    
                vntPasso = 300
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                                    enumTipoOperacaoLQS.CompromissadaMigracaoVoltaBMA) Then

                    If Not fgIN(Val(xmlDetMensagem.selectSingleNode("//TP_NEGO_BMA").Text), enumTipoNegociacaoBMA.MigracaoVoltaCamara) Then
                        fgAdicionaErro xmlErros, 3076, enumJustificativa.TipoNegociacao
                    End If
                End If
                    
                vntPasso = 310
                If fgIN(Val(xmlDetOperacao.selectSingleNode("//TP_OPER").Text), _
                                    enumTipoOperacaoLQS.CompromissadaMigracaoIdaBMA) Then

                    If Not fgIN(Val(xmlDetMensagem.selectSingleNode("//TP_NEGO_BMA").Text), enumTipoNegociacaoBMA.MigracaoIdaCamara) Then
                        fgAdicionaErro xmlErros, 3076, enumJustificativa.TipoNegociacao
                    End If
                End If
                

            Case "CTP1002"
            
                'Se operacao de estorno CETIP, gravar TipoAcao de cancelamento
                If Val(xmlDetOperacao.selectSingleNode("//CO_OPER_CETIP").Text) >= 100 Then
                    blnCancelamentoCETIP = True
                    lngTipoAcao = enumTipoAcao.CancelamentoSolicitado
                End If
            
                'data da operacao
                If xmlDetMensagem.selectSingleNode("//DT_OPER").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3066, enumJustificativa.DataOperacao
                End If
            
                If xmlDetOperacao.selectSingleNode("//CO_PARP_CAMR") Is Nothing Then
                    If Not xmlDetOperacao.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                        If CLng("0" & xmlDetMensagem.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text) _
                        <> CLng("0" & xmlDetOperacao.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text) Then
                            
                            fgAdicionaErro xmlErros, 4223, enumJustificativa.IndentificadorParticipanteCamara
                        End If
                    End If
                Else
                    If CLng("0" & xmlDetMensagem.selectSingleNode("//CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text) _
                    <> CLng("0" & xmlDetOperacao.selectSingleNode("//CO_PARP_CAMR").Text) Then
                        
                        fgAdicionaErro xmlErros, 4223, enumJustificativa.IndentificadorParticipanteCamara
                    End If
                End If
                    
                If blnVerificaCamposOpcionais Then
                    'Identificador contraparte camara
                    If xmlDetMensagem.selectSingleNode("//CO_CNTA_CUTD_CNPT").Text _
                    <> xmlDetOperacao.selectSingleNode("//CO_CNTA_CUTD_SELIC_CNPT").Text Then
                        
                        fgAdicionaErro xmlErros, 3078, enumJustificativa.IndentificadorContraparteCamara
                    End If
                End If
                
                'Codigo operacao CETIP
                If xmlDetMensagem.selectSingleNode("//CO_OPER_CETIP").Text _
                <> xmlDetOperacao.selectSingleNode("//CO_OPER_CETIP").Text Then

                    fgAdicionaErro xmlErros, 3079, enumJustificativa.CodigoOperacaoCETIP
                End If

                'Subtipo ativo
                If xmlDetMensagem.selectSingleNode("//DE_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//DE_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3080, enumJustificativa.DescricaoAtivo
                End If
                
                'identificador titulo cetip
                If xmlDetMensagem.selectSingleNode("//NU_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//NU_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3029, enumJustificativa.IdentificadorTitulo
                End If
                            
                If xmlDetMensagem.selectSingleNode("//IN_OPER_DEBT_CRED").Text _
                <> xmlDetOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text Then

                    fgAdicionaErro xmlErros, 3067, enumJustificativa.DebitoCredito
                End If
                
                If xmlDetMensagem.selectSingleNode("//QT_ATIV_MERC").Text _
                <> xmlDetOperacao.selectSingleNode("//QT_ATIV_MERC").Text Then

                    fgAdicionaErro xmlErros, 3070, enumJustificativa.QuantidadeTitulo
                End If
                
                If xmlDetMensagem.selectSingleNode("//NU_COMD_OPER").Text _
                <> xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text Then

                    fgAdicionaErro xmlErros, 3030, enumJustificativa.NumeroComando
                End If
            
                If fgIN(lngTipoOperacao, _
                        enumTipoOperacaoLQS.MovInstrumentoFinanceiroConciliacao, _
                        enumTipoOperacaoLQS.DepositoFundoInvestimentoConciliacaoCETIP, _
                        enumTipoOperacaoLQS.OperacaoDefinitivaVendaOrdem) Then
                
                
                    If xmlDetMensagem.selectSingleNode("//TP_LIQU_OPER_ATIV").Text _
                    <> xmlDetOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text Then
    
                        fgAdicionaErro xmlErros, 3081, enumJustificativa.ModalidadeLiquidacao
                    End If
                End If

            Case "BMC0005"
                
                'Validacao por Tipo de Operacao / Operacao = 230 ou 231
                If lngTipoOperacao = enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega _
                Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Then
                    'Valida o Codigo Contratacao SISBACEN da BMC0005 com o Numero Operacao Cambio
                    If xmlDetMensagem.selectSingleNode("//CO_CNTR_SISB").Text _
                    <> xmlDetOperacao.selectSingleNode("//NU_COMD_OPER").Text Then

                        fgAdicionaErro xmlErros, 3136, enumJustificativa.CodigoContratacaoSISBACEN
                    End If
                'Operacao 234
                ElseIf lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                    'Valida o Chacam da BMC0005 com o da Operacao
                    If xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text _
                    <> xmlDetOperacao.selectSingleNode("//CD_ASSO_CAMB").Text Then

                        fgAdicionaErro xmlErros, 3137, enumJustificativa.CodigoAssociacaoCambio
                    End If
                'Outras Operacoes
                Else
                    'Tipo de Negociacao = 1 ou 2
                    If xmlDetOperacao.selectSingleNode("//TP_NEGO").Text = 1 _
                    Or xmlDetOperacao.selectSingleNode("//TP_NEGO").Text = 2 Then
                        'Valida o Chacam da BMC0015 com o da BMC0005
                        If xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text _
                        <> xmlDetOperacao.selectSingleNode("//CD_ASSO_CAMB").Text Then
    
                            fgAdicionaErro xmlErros, 3137, enumJustificativa.CodigoAssociacaoCambio
                        End If
                    'Tipo de Negociacao = 3
                    ElseIf xmlDetOperacao.selectSingleNode("//TP_NEGO").Text = 3 Then
                        'Valida o Codigo Contratacao SISBACEN da BMC0005 com o Codigo Contratacao SISBACEN da Operacao
                        If xmlDetMensagem.selectSingleNode("//CO_CNTR_SISB").Text _
                        <> xmlDetOperacao.selectSingleNode("//CO_CNTR_SISB").Text Then
    
                            fgAdicionaErro xmlErros, 3136, enumJustificativa.CodigoContratacaoSISBACEN
                        End If
                    End If
                End If
                
            Case "BMC0015"

                'Operacao 234
                If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                    'Valida o Chacam da BMC0015 com o da Operacao
                    If xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text _
                    <> xmlDetOperacao.selectSingleNode("//CD_ASSO_CAMB").Text Then

                        fgAdicionaErro xmlErros, 3137, enumJustificativa.CodigoAssociacaoCambio
                    End If
                Else
                    'Valida Numero Identificador Nogociacao BMC
                    If xmlDetMensagem.selectSingleNode("//NR_IDEF_NEGO_BMC").Text _
                    <> xmlDetOperacao.selectSingleNode("//NR_IDEF_NEGO_BMC").Text Then
    
                        fgAdicionaErro xmlErros, 4295, enumJustificativa.NumeroIdentNegociacaoBMC
                    Else
                        If Not xmlRemessa Is Nothing Then
                            If Not xmlRemessa.selectSingleNode("//MESG") Is Nothing Then
                                Call fgAppendNode(xmlRemessa, "MESG", "ENCONTROU_IDENT_NEG_BMC", vbNullString)
                            End If
                        End If
                    End If
                    
                    'Valida Codigo Moeda
                    If xmlDetMensagem.selectSingleNode("//CO_MOED_ESTR").Text _
                    <> xmlDetOperacao.selectSingleNode("//CO_MOED_ESTR").Text Then
    
                        fgAdicionaErro xmlErros, 3134, enumJustificativa.MoedaEstrangeira
                    End If
                End If

                If xmlDetMensagem.selectSingleNode("//DT_OPER").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3066, enumJustificativa.DataOperacao
                End If

                If xmlDetMensagem.selectSingleNode("//IN_OPER_DEBT_CRED").Text _
                <> xmlDetOperacao.selectSingleNode("//IN_OPER_DEBT_CRED").Text Then

                    fgAdicionaErro xmlErros, 3067, enumJustificativa.DebitoCredito
                End If

                If xmlDetMensagem.selectSingleNode("//PE_TAXA_NEGO").Text _
                <> xmlDetOperacao.selectSingleNode("//PE_TAXA_NEGO").Text Then

                    fgAdicionaErro xmlErros, 3074, enumJustificativa.TaxaNegociada
                End If

                If xmlDetMensagem.selectSingleNode("//VA_FINC").Text _
                <> xmlDetOperacao.selectSingleNode("//VA_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3072, enumJustificativa.ValorFinanceiro
                End If

                If xmlDetMensagem.selectSingleNode("//VA_MOED_ESTR").Text _
                <> xmlDetOperacao.selectSingleNode("//VA_MOED_ESTR").Text Then

                    fgAdicionaErro xmlErros, 3135, enumJustificativa.ValorMoedaEstrangeira
                End If

                If xmlDetMensagem.selectSingleNode("//DT_LIQU").Text _
                <> xmlDetOperacao.selectSingleNode("//DT_LIQU_OPER_ATIV").Text Then

                    fgAdicionaErro xmlErros, 3069, enumJustificativa.DataLiquidacao
                End If

            Case Else
                fgAdicionaErro xmlErros, 3082, enumJustificativa.TipoMensagem
        
        End Select
    
    End If
    
    If xmlErros.selectNodes("//Grupo_ErrorInfo").length > 0 Then
        
        'Deu erro, não faz nada com os status
        ConciliarRegistroOperacao = xmlErros.xml
        vntPasso = 250
        Exit Function
    
    Else
        vntPasso = 500
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
        Set objPO = CreateObject("A8LQS.clsProcessoOperacao")

        Select Case pintAcao
            Case enumAcaoConciliacao.BOConcordar
                'BACKOFFICE
                vntPasso = 600
                'Atualiza status do registro de operacao
                If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0005" Then
                    lngStatus = IIf(pblnAutomatica, enumStatusOperacao.ConcordanciaAceiteAutomatica, enumStatusOperacao.ConcordanciaAceite)
                Else
                    lngStatus = IIf(pblnAutomatica, enumStatusOperacao.ConcordanciaAutomatica, enumStatusOperacao.Concordancia)
                End If
                Call objOperacao.AlterarStatusAcao(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                   lngStatus, _
                                                   xmlDetOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                                   , _
                                                   lngTipoAcao)
                plngStatusOperacao = lngStatus
                vntPasso = 610
                'Atualiza status da mensagem
                lngStatusMensagem = enumStatusMensagem.Conciliada
                
                Call objMensagem.AlterarStatus(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                               xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                               lngStatusMensagem, _
                                               xmlDetMensagem.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                                               , _
                                               True, _
                                               xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                
                'Gera numero para ser gerado como chave das tabelas
                vntPasso = 620
                lngSequencialConciliacaoOperacao = IncluirJustificativaConciliacaoOperacao()
                
                vntPasso = 630
                Call IncluirConciliacaoOperacao(lngSequencialConciliacaoOperacao, _
                                                xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                xmlDetMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                                0, _
                                                xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                                                
                'Se BMA0002, grava conteúdo da tag 'NumCtrlBMAEspecc' no campo '???' da operacao
                vntPasso = 640
                If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMA0002" Then
                    vntPasso = 650
                    Set xmlMensagemRecb = objMensagem.LerXMLMensagem(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                                     xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
                    vntPasso = 660
                    If xmlMensagemRecb.xml <> vbNullString Then
                        objOperacao.GravarControleSPBOriginal xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                              xmlMensagemRecb.selectSingleNode("//NumCtrlBMAEspecc").Text
                        
                        If Not xmlRemessa Is Nothing Then
                            Call fgAppendNode(xmlRemessa, "MESG", "NU_CTRL_MESG_SPB_ORIG", xmlMensagemRecb.selectSingleNode("//NumCtrlBMAEspecc").Text)
                        End If
                    End If
                    
                End If
                
                'Processo Automatizado
                If xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0015" Then
                    
                    blnBMC0005BMC0015 = True
                    
                    'Se Tipo de Operacao 234 envia CAM0054
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        lngFuncaoSistema = enumFuncaoSistema.LiberarCAM0054
                        lngTipoAcao = enumTipoAcao.EnviadaCAM0054
                    Else
                        lngFuncaoSistema = enumFuncaoSistema.Liberar
                        lngTipoAcao = enumTipoAcao.EnviadaCAM0002
                    End If
                    
                    'Gerenciar Chamadas
                    blnAux = objPO.GerenciarChamadaViaTela(xmlDetMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                                           lngFuncaoSistema, , _
                                                           True, _
                                                           True, , _
                                                           xmlDetMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text, , , , _
                                                           xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

                    
                    'Le novamente a operacao para atualizar o status
                    Set xmlDetOperacao = objOperacao.LerXMLOperacao(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

                    Call objOperacao.AlterarStatusAcao(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                       enumStatusOperacao.LiberadaAutomatica, _
                                                       xmlDetOperacao.selectSingleNode("//DH_ULTI_ATLZ").Text, , _
                                                       lngTipoAcao)
                                                       
                    'Atualiza o Codigo de Associacao Cambio na Operacao
                    Call objOperacao.AtualizarChACAM(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                     xmlDetMensagem.selectSingleNode("//CD_ASSO_CAMB").Text)
                    
                ElseIf xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text = "BMC0005" Then
                    
                    blnBMC0005BMC0015 = True
                        
                    If lngStatus = enumStatusOperacao.ConcordanciaAceiteAutomatica Then
                        
                        If xmlDetOperacao.selectSingleNode("//TP_NEGO").Text <> 3 Or _
                           xmlOperacao.xml = vbNullString Then
                            
                            'Busca Indicador de Entrada Manual
                            Set xmlDetOperacao = objOperacao.LerXMLOperacao(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                            If Not xmlDetOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                                If Val(xmlDetOperacao.selectSingleNode("//IN_ENTR_MANU").Text) = enumIndicadorSimNao.Sim Then
                                    blnEntradaManual = True
                                End If
                            End If
    
                            If Not blnEntradaManual Or (Val(xmlDetOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.RegistroOperacaoInterbancaria) Then
                                blnAux = objPO.GerenciarChamadaViaTela(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                       enumFuncionalidade.Liberacao, _
                                                                       , _
                                                                       True, , , , , _
                                                                       pblnAutomatica)
                            End If
                            
                        End If
                        
                    Else
                        'Busca Indicador de Entrada Manual
                        Set xmlDetOperacao = objOperacao.LerXMLOperacao(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                        If Not xmlDetOperacao.selectSingleNode("//IN_ENTR_MANU") Is Nothing Then
                            If Val(xmlDetOperacao.selectSingleNode("//IN_ENTR_MANU").Text) = enumIndicadorSimNao.Sim Then
                                blnEntradaManual = True
                            End If
                        End If
                        
                        If Not blnEntradaManual Then
                            blnAux = objPO.GerenciarChamadaViaTela(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                   enumFuncionalidade.Conciliacao, , _
                                                                   True)
                        End If
                    End If
                End If
                
                If Not blnCancelamentoCETIP _
                And Not blnBMC0005BMC0015 Then
                    If (Not pblnAutomatica) _
                        Or (blnSoPassouMensagem) Then
                        'Se está vindo da tela, chama o fluxo automático para tentar liberar automático
                        'Se está vindo do fluxo auto, só chama se só veio mensagem
                        
                        'Chama gerenciador do fluxo automático
                        vntPasso = 670
                        If (lngStatus = enumStatusOperacao.ConcordanciaAutomatica Or _
                            lngStatus = enumStatusOperacao.ConciliadaAutomatica) And _
                            xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text <> "BMA0002" And _
                            xmlDetMensagem.selectSingleNode("//CO_MESG_SPB").Text <> "CTP1002" Then
                        
                            blnAux = objPO.GerenciarChamadaViaTela(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                   enumFuncionalidade.Liberacao, _
                                                                   , _
                                                                   True, , , , , _
                                                                   pblnAutomatica)
                        
                        Else
                        
                            blnAux = objPO.GerenciarChamadaViaTela(xmlDetOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                   enumFuncionalidade.Conciliacao, , _
                                                                   True)
                        End If
                    End If
                End If

        End Select
    
    End If
    
    If xmlErros.selectNodes("//Grupo_ErrorInfo").length > 0 Then
        'Mensagens de erro
        ConciliarRegistroOperacao = xmlErros.xml
    Else
        ConciliarRegistroOperacao = ""
    End If
    
    Set xmlErros = Nothing
    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlMensagemRecb = Nothing
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objPO = Nothing
    Set objRegra = Nothing
    
    Exit Function

ErrorHandler:

    Set xmlDetOperacao = Nothing
    Set xmlDetMensagem = Nothing
    Set xmlMensagemRecb = Nothing
    Set objOperacao = Nothing
    Set objMensagem = Nothing
    Set objPO = Nothing
    Set objRegra = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ConciliarRegistroOperacao Function (passo=" & vntPasso & ")", lngCodigoErroNegocio, intNumeroSequencialErro, xmlErros.xml)

    Set xmlErros = Nothing
    
    Exit Function
    Resume

End Function

'--------------------------------------------------------------------------------------------------
'Realiza a conciliação multilateral financeira
'--------------------------------------------------------------------------------------------------

Public Function LiquidarMultilateralFinanceira(ByVal plngPerfilAcesso As Integer, _
                                               ByVal pintAcao As Integer, _
                                               Optional xmlConciliacao As MSXML2.DOMDocument40) As String

    '//////////////////////////////////////////////////////////////////////////////////////////////
    '//                                                                                          //
    '//     Realiza a conciliação multilateral financeira                                        //
    '//                                                                                          //
    '//     Retorno: OK (string em branco), ou ERROS (string XML com colecao de erros)           //
    '//                                                                                          //
    '//////////////////////////////////////////////////////////////////////////////////////////////

Dim xmlErros                                As MSXML2.DOMDocument40
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim xmlHistorico                            As MSXML2.DOMDocument40
Dim xmlMensagens                            As MSXML2.DOMDocument40
Dim xmlRepeticao                            As MSXML2.DOMDocument40
Dim xmlMensagemMae                          As MSXML2.DOMDocument40
   
Dim strRet                                  As String
Dim blnAux                                  As Boolean

Dim objOperacao                             As A8LQS.clsOperacao
Dim objMsg                                  As A8LQS.clsMensagem
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objNodeMensagem                         As MSXML2.IXMLDOMNode
Dim objNode                                 As MSXML2.IXMLDOMNode

Dim vntATUALStatusMsgMae                    As Variant  'status ATUAL da mensagem mae (antes de alterá-lo, se for o caso)

Dim vntStatusOp                             As Variant  'status das operacoes
Dim vntStatusMsg                            As Variant  'status do item de repetição
Dim vntStatusMsgMae                         As Variant  'status da 'mensagem mãe' (NU_SEQU_CNTR_REPE='1')

Dim vntStatusOp_Anterior                    As Variant
Dim vntStatusMsg_Anterior                   As Variant

Dim vntAcaoMsg                              As Variant  'acao a gravar no item de repetição
Dim vntAcaoMsg2                             As Variant  '            ''                    , para gravar duar acoes
Dim vntAcaoMsgMae                           As Variant  'acao a gravar na 'mensagem mãe' (NU_SEQU_CNTR_REPE='1')

Dim vntAcaoOperacao                         As Variant  'Acao a ser executada na operação, passando para o
                                                        'GerenciarChamadaViaTela. Por exemplo, conciliacao,
                                                        'liberacao, liberacaoprevia, concordanciaprevia.
                                                        'É na tabela A8.TB_CTRL_PROC_OPER_ATIV que estas açoes
                                                        'mapeiam para quais mensagens internas são enviadas, como
                                                        'previsao A6, PJ ou Conta Corrente

Dim vntEnviarMensagem                       As Variant  'qual mensagem será enviada

Dim blnIgnoraGradeHorario                   As Boolean

'Controle de quais ações já foram executadas
Dim blnAcaoJaEnviarConcordancia             As Boolean
Dim blnAcaoJaEnviarDiscordancia             As Boolean
Dim blnAcaoJaPagamento                      As Boolean
Dim blnAcaoJaPagamentoContingencia          As Boolean

Dim blnMensagemMae                          As Boolean  'se a mensagem atual tem NU_SEQU_CNTR_REPE=1
Dim blnDisponibilizaCC                      As Boolean  'se tem que disponibilizar operacoes p/ ContaCorrente
Dim blnFimFluxoAdmArea                      As Boolean  'se o ADM AREA é o fim de fluxo
Dim blnMsgDefinitiva                        As Boolean

Dim strCodigoMensagem                       As String
Dim strRepeticao                            As String
Dim lngLocalLiquidacao                      As Long

Dim vntSequencialConciliacao                As Variant

Dim vntPasso                                As String   'para trace do fluxo de execucao do programa, em caso de erro

On Error GoTo ErrorHandler

    LiquidarMultilateralFinanceira = ""
    
    'Pega o sequencial de conciliação da mensagem mãe.
    vntSequencialConciliacao = Null

    vntPasso = 1
    Set xmlErros = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagens = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlRepeticao = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagemMae = CreateObject("MSXML2.DOMDocument.4.0")

    'inicializa xml de erros
    fgAppendNode xmlErros, "", "Grupo", ""

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMsg = CreateObject("A8LQS.clsMensagem")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    
    For Each objNodeMensagem In xmlConciliacao.selectNodes("//Grupo_Mensagem")
    
        vntPasso = 20
        
        If IsNull(vntSequencialConciliacao) Then
            vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                               objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                               objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, _
                                               True, _
                                               True)

        End If
        
        strCodigoMensagem = objNodeMensagem.selectSingleNode("./CO_MESG_SPB").Text
        blnMensagemMae = (Val(objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text) = 1) And (Left(strCodigoMensagem, 3) = "LDL")
        blnIgnoraGradeHorario = objNodeMensagem.selectNodes("./IgnoraGradeHorario").length > 0
        lngLocalLiquidacao = CLng("0" & objNodeMensagem.selectSingleNode("./CO_LOCA_LIQU").Text)
        
        Select Case strCodigoMensagem
            Case "LDL0001"
                blnMsgDefinitiva = (objNodeMensagem.selectSingleNode("./TP_INFO_LDL").Text = "D")
            Case "LDL0005R2"
                blnMsgDefinitiva = True
            Case "LDL0009R2"
                blnMsgDefinitiva = True
            Case "BMC0101"
                blnMsgDefinitiva = True
            Case "BMC0103"
                blnMsgDefinitiva = True
            Case "LDL0026R1"
                blnMsgDefinitiva = False
        End Select
        
        If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then

            blnFimFluxoAdmArea = True
            
            vntAcaoOperacao = fgDECODE(pintAcao, _
                                       enumAcaoConciliacao.BOConcordar, fgDECODE(strCodigoMensagem, "LDL0001", enumFuncionalidade.Concordancia, _
                                                                                                    "BMC0101", enumFuncionalidade.ConcordanciaBMC0101, _
                                                                                                    "BMC0103", enumFuncionalidade.ConcordanciaBMC0103, _
                                                                                                    "LDL0005R2", enumFuncionalidade.ConcordanciaLDL0005R2), _
                                       enumAcaoConciliacao.BODiscordar, enumFuncionalidade.Discordancia, _
                                       enumAcaoConciliacao.AdmAreaLiberar, enumFuncionalidade.LiberacaoPagamento, _
                                       enumAcaoConciliacao.AdmAreaRejeitar, enumFuncionalidade.RejeicaoConcordancia, _
                                       enumAcaoConciliacao.AdmAreaPagamentoContingencia, enumFuncionalidade.PagamentoContingencia, _
                                       enumAcaoConciliacao.AdmAreaRegularizar, enumFuncionalidade.RegularizarContingencia, _
                                       Null)
    
        Else
            If ((strCodigoMensagem = "LDL0001" And blnMsgDefinitiva = False) _
             Or (strCodigoMensagem = "LDL0026R1")) Then
                
                'PREVIAS
                blnFimFluxoAdmArea = True
                vntAcaoOperacao = fgDECODE(pintAcao, _
                                            enumAcaoConciliacao.BOConcordar, enumFuncionalidade.ConcordanciaPrevia, _
                                            enumAcaoConciliacao.AdmAreaLiberar, enumFuncionalidade.LiberacaoPrevia, _
                                            enumAcaoConciliacao.AdmAreaRejeitar, enumFuncionalidade.RejeicaoConcordanciaPrevia, _
                                            Null)
                              
            ElseIf (strCodigoMensagem = "LDL0005R2") _
                Or (strCodigoMensagem = "LDL0009R2") Then
    
                'DEFINITIVAS DE RECEBIMENTO (R2)
                blnFimFluxoAdmArea = True
                vntAcaoOperacao = fgDECODE(pintAcao, _
                                            enumAcaoConciliacao.BOConcordar, enumFuncionalidade.Concordancia, _
                                            enumAcaoConciliacao.AdmAreaLiberar, enumFuncionalidade.LiberacaoRecebimento, _
                                            enumAcaoConciliacao.AdmAreaRejeitar, enumFuncionalidade.RejeicaoConcordancia, _
                                            Null)
                              
            Else
                
                'DEFINITIVA DE PAGAMENTO
                blnFimFluxoAdmArea = False
                vntAcaoOperacao = fgDECODE(pintAcao, _
                                            enumAcaoConciliacao.AdmGeralRegularizar, enumFuncionalidade.RegularizarContingencia, _
                                            Null)
            
            End If
        End If
        
        blnDisponibilizaCC = blnFimFluxoAdmArea
        
        vntStatusOp = Null
        vntStatusMsg = Null
        vntStatusMsgMae = Null
        vntStatusOp_Anterior = Null
        vntStatusMsg_Anterior = Null
        
        vntAcaoMsg = Null
        vntAcaoMsgMae = Null
        
        blnAcaoJaEnviarConcordancia = False
        blnAcaoJaEnviarDiscordancia = False
        blnAcaoJaPagamento = False
        blnAcaoJaPagamentoContingencia = False
               
        'Status anterior da primeira mensagem filha
        vntPasso = 30
        vntStatusMsg_Anterior = Val(objNodeMensagem.selectSingleNode("./CO_ULTI_SITU_PROC").Text)  '"CO_ULTI_SITU_PROC[../NU_SEQU_CNTR_REPE != '1']").Text)
        
        If Not blnMensagemMae Then
            'Status anterior da primeira operacao
            vntPasso = 40
            vntStatusOp_Anterior = Null
            If objNodeMensagem.selectNodes("./Repeat_Operacoes/Grupo_DetalheOperacao/*").length > 0 Then
                vntStatusOp_Anterior = Val(objNodeMensagem.selectSingleNode("./Repeat_Operacoes/Grupo_DetalheOperacao/CO_ULTI_SITU_PROC").Text)
            End If
        End If
        
        'Configura próximo status das operacoes, detalhes de mensagem e 'mensagem mãe' (Null = Não alterar)
        vntPasso = 50
        
        If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
            '---------------
            'Regras para BMC
            '---------------
            
            'Descobre o status atual da mensagem mae
            vntATUALStatusMsgMae = Null
            If fgIN(pintAcao, _
                enumAcaoConciliacao.BOEnviarConcordancia, _
                enumAcaoConciliacao.AdmAreaLiberar, _
                enumAcaoConciliacao.AdmAreaPagamentoContingencia) Then
                
                xmlMensagemMae.loadXML objMsg.ObterMensagensAssociadas(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                                       objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                                       1)
            
                vntATUALStatusMsgMae = xmlMensagemMae.selectSingleNode("//CO_ULTI_SITU_PROC").Text
            End If
            
            vntStatusOp = fgDECODE(pintAcao, _
                enumAcaoConciliacao.BOConcordar, IIf(fgIN(strCodigoMensagem, "LDL0005R2", "BMC0103"), _
                                                          enumStatusOperacao.Liquidada, _
                                                          enumStatusOperacao.ConcordanciaBackoffice), _
                enumAcaoConciliacao.BODiscordar, enumStatusOperacao.AConciliar, _
                enumAcaoConciliacao.BORegularizar, enumStatusOperacao.DebitoMoedaEstrangeiraLiquidado, _
                enumAcaoConciliacao.BOEnviarConcordancia, Null, _
                enumAcaoConciliacao.BOEnviarConcordanciaContingencia, Null, _
                enumAcaoConciliacao.AdmAreaRejeitar, enumStatusOperacao.AConciliar, _
                enumAcaoConciliacao.AdmAreaLiberar, enumStatusOperacao.PagamentoLib, _
                enumAcaoConciliacao.AdmAreaPagamentoContingencia, Null, _
                enumAcaoConciliacao.AdmAreaRegularizar, enumStatusOperacao.DebitoMoedaNacionalLiquidado, _
                Null)
                
            vntPasso = 60
            vntStatusMsg = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, IIf(fgIN(strCodigoMensagem, "LDL0005R2", "BMC0103"), _
                                                         enumStatusMensagem.MensagemLiquidada, _
                                                         enumStatusMensagem.ConcordanciaBackoffice), _
                            enumAcaoConciliacao.BODiscordar, enumStatusMensagem.AConciliar, _
                            enumAcaoConciliacao.BORegularizar, enumStatusMensagem.MensagemLiquidada, _
                            enumAcaoConciliacao.BOEnviarConcordancia, Null, _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, objNodeMensagem.selectSingleNode("./CO_ULTI_SITU_PROC").Text, _
                            enumAcaoConciliacao.AdmAreaRejeitar, enumStatusMensagem.AConciliar, _
                            enumAcaoConciliacao.AdmAreaLiberar, enumStatusMensagem.PagamentoLib, _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, Null, _
                            enumAcaoConciliacao.AdmAreaRegularizar, enumStatusMensagem.MensagemLiquidada, _
                            Null)
            
            vntPasso = 70
            vntStatusMsgMae = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, fgDECODE(strCodigoMensagem, _
                                                                "LDL0005R2", enumStatusMensagem.MensagemLiquidada, _
                                                                "LDL0001", enumStatusMensagem.ConcordanciaBackoffice, _
                                                                Null), _
                            enumAcaoConciliacao.BODiscordar, enumStatusMensagem.AConciliar, _
                            enumAcaoConciliacao.BORegularizar, Null, _
                            enumAcaoConciliacao.BOEnviarConcordancia, vntATUALStatusMsgMae, _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, Null, _
                            enumAcaoConciliacao.AdmAreaRejeitar, enumStatusMensagem.AConciliar, _
                            enumAcaoConciliacao.AdmAreaLiberar, enumStatusMensagem.PagamentoLib, _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, vntATUALStatusMsgMae, _
                            enumAcaoConciliacao.AdmAreaRegularizar, enumStatusMensagem.MensagemLiquidada, _
                            Null)
            
            vntPasso = 80
            vntAcaoMsg = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, IIf(strCodigoMensagem = "BMC0101", enumTipoAcao.EnviadaBMC0102, 0), _
                            enumAcaoConciliacao.BORegularizar, 0, _
                            enumAcaoConciliacao.BOEnviarConcordancia, 0, _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, IIf(strCodigoMensagem = "BMC0101", enumTipoAcao.EnviadaBMC0102, 0), _
                            enumAcaoConciliacao.AdmAreaRejeitar, enumTipoAcao.RejeicaoConcordancia, _
                            enumAcaoConciliacao.AdmAreaLiberar, 0, _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, 0, _
                            enumAcaoConciliacao.AdmAreaRegularizar, 0, _
                            0)
            
            vntPasso = 85
            vntAcaoMsg2 = fgDECODE(pintAcao, enumAcaoConciliacao.BOEnviarConcordanciaContingencia, enumTipoAcao.EnviadaConfirmacaoContingencia, 0)
            
            vntPasso = 90
            vntAcaoMsgMae = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, 0, _
                            enumAcaoConciliacao.BORegularizar, 0, _
                            enumAcaoConciliacao.BOEnviarConcordancia, IIf(strCodigoMensagem = "LDL0001", enumTipoAcao.EnviadaLDL0003Concordancia, 0), _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, 0, _
                            enumAcaoConciliacao.AdmAreaRejeitar, 0, _
                            enumAcaoConciliacao.AdmAreaLiberar, enumTipoAcao.EnviadoPagamento, _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, enumTipoAcao.EnviadoPagamentoContingencia, _
                            enumAcaoConciliacao.AdmAreaRegularizar, 0, _
                            0)
            
            'Verifica qual mensagem deve ser enviada (Null=nenhuma)
            vntPasso = 100
            vntEnviarMensagem = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, IIf(strCodigoMensagem = "BMC0101", enumFuncaoSistema.Concordar, Null), _
                            enumAcaoConciliacao.BORegularizar, Null, _
                            enumAcaoConciliacao.BOEnviarConcordancia, IIf(strCodigoMensagem = "LDL0001", enumFuncaoSistema.Concordar, 0), _
                            enumAcaoConciliacao.BOEnviarConcordanciaContingencia, IIf(strCodigoMensagem = "BMC0101", enumFuncaoSistema.Concordar, 0), _
                            enumAcaoConciliacao.AdmAreaRejeitar, Null, _
                            enumAcaoConciliacao.AdmAreaLiberar, enumFuncaoSistema.Pagar, _
                            enumAcaoConciliacao.AdmAreaPagamentoContingencia, enumFuncaoSistema.PagarContingencia, _
                            enumAcaoConciliacao.AdmAreaRegularizar, Null, _
                            Null)
                            
        ElseIf fgIN(lngLocalLiquidacao, enumLocalLiquidacao.BMA, enumLocalLiquidacao.CETIP) Then
    
            vntStatusOp = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, IIf(blnMsgDefinitiva, _
                                                                    enumStatusOperacao.ConcordanciaBackoffice, _
                                                                    enumStatusOperacao.ConcordanciaBackofficePrevia), _
                            enumAcaoConciliacao.BODiscordar, enumStatusOperacao.DiscordanciaBackoffice, _
                            enumAcaoConciliacao.AdmAreaLiberar, IIf(blnFimFluxoAdmArea, _
                                                                    IIf(fgIN(strCodigoMensagem, "LDL0001", "LDL0026R1"), _
                                                                        enumStatusOperacao.Registrada, enumStatusOperacao.Liquidada), _
                                                                    IIf((vntStatusOp_Anterior = enumStatusOperacao.ConcordanciaBackoffice), _
                                                                        enumStatusOperacao.ConcordanciaAdmArea, enumStatusOperacao.DiscordanciaAdmArea)), _
                            enumAcaoConciliacao.AdmAreaRejeitar, fgDECODE(strCodigoMensagem, _
                                                                        "LDL0026R1", enumStatusOperacao.AConciliar, _
                                                                        "LDL0005R2", enumStatusOperacao.Registrada, _
                                                                        "LDL0009R2", enumStatusOperacao.AConciliar, _
                                                                        "LDL0001", IIf(blnMsgDefinitiva, enumStatusOperacao.Registrada, enumStatusOperacao.AConciliar)), _
                            enumAcaoConciliacao.AdmGeralRejeitar, IIf(vntStatusOp_Anterior = enumStatusOperacao.ConcordanciaAdmArea, _
                                                                    enumStatusOperacao.ConcordanciaBackoffice, _
                                                                    enumStatusOperacao.DiscordanciaBackoffice), _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, Null, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, Null, _
                            enumAcaoConciliacao.AdmGeralPagamento, enumStatusOperacao.PagamentoLib, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, Null, _
                            enumAcaoConciliacao.AdmGeralRegularizar, enumStatusOperacao.Liquidada, _
                            Null)
                
            vntPasso = 60
            vntStatusMsg = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, IIf(blnMsgDefinitiva, _
                                                                    enumStatusMensagem.ConcordanciaBackoffice, _
                                                                    enumStatusMensagem.ConcordanciaBackofficePrevia), _
                            enumAcaoConciliacao.BODiscordar, enumStatusMensagem.DiscordanciaBackoffice, _
                            enumAcaoConciliacao.AdmAreaLiberar, IIf(blnFimFluxoAdmArea, _
                                                                    IIf(fgIN(strCodigoMensagem, "LDL0001", "LDL0026R1"), _
                                                                        enumStatusMensagem.Confirmada, _
                                                                        enumStatusMensagem.MensagemLiquidada), _
                                                                    IIf(vntStatusMsg_Anterior = enumStatusMensagem.ConcordanciaBackoffice, _
                                                                        enumStatusMensagem.ConcordanciaAdmArea, _
                                                                        enumStatusMensagem.DiscordanciaAdmArea)), _
                            enumAcaoConciliacao.AdmAreaRejeitar, enumStatusMensagem.AConciliar, _
                            enumAcaoConciliacao.AdmGeralRejeitar, IIf(vntStatusMsg_Anterior = enumStatusMensagem.ConcordanciaAdmArea, _
                                                                    enumStatusMensagem.ConcordanciaBackoffice, _
                                                                    enumStatusMensagem.DiscordanciaBackoffice), _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, Null, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, Null, _
                            enumAcaoConciliacao.AdmGeralPagamento, enumStatusMensagem.PagamentoLib, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, Null, _
                            enumAcaoConciliacao.AdmGeralRegularizar, enumStatusMensagem.MensagemLiquidada, _
                            Null)
            
            vntPasso = 70
            vntStatusMsgMae = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, Null, _
                            enumAcaoConciliacao.BODiscordar, Null, _
                            enumAcaoConciliacao.AdmAreaLiberar, Null, _
                            enumAcaoConciliacao.AdmAreaRejeitar, Null, _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, objNodeMensagem.selectSingleNode("./CO_ULTI_SITU_PROC").Text, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, objNodeMensagem.selectSingleNode("./CO_ULTI_SITU_PROC").Text, _
                            enumAcaoConciliacao.AdmGeralPagamento, enumStatusMensagem.PagamentoLib, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, objNodeMensagem.selectSingleNode("./CO_ULTI_SITU_PROC").Text, _
                            enumAcaoConciliacao.AdmGeralRegularizar, enumStatusMensagem.MensagemLiquidada, _
                            Null)
            
            vntPasso = 80
            vntAcaoMsg = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, 0, _
                            enumAcaoConciliacao.BODiscordar, 0, _
                            enumAcaoConciliacao.AdmAreaLiberar, 0, _
                            enumAcaoConciliacao.AdmAreaRejeitar, IIf(vntStatusMsg_Anterior = enumStatusMensagem.ConcordanciaBackoffice, _
                                                                    enumTipoAcao.RejeicaoConcordancia, _
                                                                    enumTipoAcao.RejeicaoDiscordancia), _
                            enumAcaoConciliacao.AdmGeralRejeitar, IIf(vntStatusMsg_Anterior = enumStatusMensagem.ConcordanciaAdmArea, _
                                                                    enumTipoAcao.RejeicaoConcordanciaAdmArea, _
                                                                    enumTipoAcao.RejeicaoDiscordanciaAdmArea), _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, 0, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, 0, _
                            enumAcaoConciliacao.AdmGeralPagamento, 0, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, 0, _
                            enumAcaoConciliacao.AdmGeralRegularizar, 0, _
                            0)
            
            vntPasso = 90
            vntAcaoMsgMae = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, 0, _
                            enumAcaoConciliacao.BODiscordar, 0, _
                            enumAcaoConciliacao.AdmAreaLiberar, 0, _
                            enumAcaoConciliacao.AdmAreaRejeitar, 0, _
                            enumAcaoConciliacao.AdmGeralRejeitar, 0, _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, enumTipoAcao.EnviadaLDL0003Concordancia, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, enumTipoAcao.EnviadaLDL0003Discordancia, _
                            enumAcaoConciliacao.AdmGeralPagamento, enumTipoAcao.EnviadoPagamento, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, enumTipoAcao.EnviadoPagamentoContingencia, _
                            enumAcaoConciliacao.AdmGeralRegularizar, 0, _
                            0)
            
            'Verifica qual mensagem deve ser enviada (Null=nenhuma)
            vntPasso = 100
            vntEnviarMensagem = fgDECODE(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, Null, _
                            enumAcaoConciliacao.BODiscordar, Null, _
                            enumAcaoConciliacao.AdmAreaLiberar, Null, _
                            enumAcaoConciliacao.AdmAreaRejeitar, Null, _
                            enumAcaoConciliacao.AdmGeralEnviarConcordancia, enumFuncaoSistema.Concordar, _
                            enumAcaoConciliacao.AdmGeralEnviarDiscordancia, enumFuncaoSistema.Discordar, _
                            enumAcaoConciliacao.AdmGeralPagamento, enumFuncaoSistema.Pagar, _
                            enumAcaoConciliacao.AdmGeralPagamentoContingencia, enumFuncaoSistema.PagarContingencia, _
                            enumAcaoConciliacao.AdmGeralRegularizar, Null, _
                            Null)
    

        Else
            Err.Raise 1, , "Câmara Inválida: " & lngLocalLiquidacao
        End If

        Dim objHistMensagem                     As A8LQS.clsHistSituacaoMensagem
        Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        
        vntPasso = 110
        
        If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
        
            vntPasso = 120
            strRet = objHistMensagem.LerTodos(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                              objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                              "1")

            If strRet <> vbNullString Then
                vntPasso = 130
                xmlHistorico.loadXML strRet
                
                blnAcaoJaEnviarConcordancia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']").length > 0 _
                                           Or xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaBMC0102 & "']").length > 0
               
                blnAcaoJaPagamento = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamento & "']").length > 0
                blnAcaoJaPagamentoContingencia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']").length > 0
                
                If Not blnAcaoJaPagamentoContingencia Then
                    blnAcaoJaPagamentoContingencia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaConfirmacaoContingencia & "']").length > 0
                End If
            End If
        
        Else
            If blnMensagemMae Then
                vntPasso = 120
                strRet = objHistMensagem.LerTodos(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                  objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                  "1")
    
                If strRet <> vbNullString Then
                    vntPasso = 130
                    xmlHistorico.loadXML strRet
                    
                    blnAcaoJaEnviarConcordancia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Concordancia & "']").length > 0
                    blnAcaoJaEnviarDiscordancia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadaLDL0003Discordancia & "']").length > _
                                                  xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.RejeicaoDiscordancia & "']").length
                    blnAcaoJaPagamento = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamento & "']").length > 0
                    blnAcaoJaPagamentoContingencia = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "']").length > 0
                End If
            End If
        
        End If
        
        '/////////////////////////////////////////////////////////
        '// Consistencias
        '/////////////////////////////////////////////////////////
        vntPasso = 200
        
        If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
            
            If strCodigoMensagem = "BMC0101" Then
                If Not blnAcaoJaEnviarConcordancia Then
                    If fgIN(pintAcao, enumAcaoConciliacao.AdmAreaRegularizar) Then
                        Err.Raise 1, , "CONCORDÂNCIA ainda não foi enviada para esta mensagem."
                    End If
                End If
            Else
                If Not blnAcaoJaPagamentoContingencia Then
                    If fgIN(pintAcao, enumAcaoConciliacao.AdmAreaRegularizar) Then
                        Err.Raise 1, , "PAGAMENTO EM CONTINGÊNCIA ainda não foi enviado para esta mensagem."
                    End If
                End If
            End If
        
            If blnAcaoJaEnviarConcordancia Then
                If fgIN(pintAcao, _
                                                    enumAcaoConciliacao.BOEnviarConcordancia, _
                                                    enumAcaoConciliacao.BOEnviarConcordanciaContingencia, _
                                                    enumAcaoConciliacao.AdmAreaPagamentoContingencia, _
                                                    enumAcaoConciliacao.AdmAreaRejeitar, _
                                                    enumAcaoConciliacao.AdmAreaRegularizar) Then
                                                    
                    Err.Raise 1, , "CONCORDÂNCIA já foi enviada para esta mensagem."
                End If
            End If
        
            If blnAcaoJaPagamentoContingencia Then
                If Not fgIN(pintAcao, _
                                    enumAcaoConciliacao.AdmAreaRejeitar, _
                                    enumAcaoConciliacao.AdmAreaRegularizar, _
                                    enumAcaoConciliacao.BORegularizar) Then
                
                    Err.Raise 1, , "PAGAMENTO EM CONTINGÊNCIA já foi enviado para esta mensagem."
                End If
            End If
        
        
        Else
            If blnMensagemMae Then
                
                If Not blnAcaoJaPagamentoContingencia Then
                    If fgIN(pintAcao, enumAcaoConciliacao.AdmGeralRegularizar) Then
                        Err.Raise 1, , "PAGAMENTO EM CONTINGÊNCIA ainda não foi enviado para esta mensagem."
                    End If
                End If
                
                If blnAcaoJaEnviarConcordancia Then
                    If fgIN(pintAcao, _
                                                        enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                                                        enumAcaoConciliacao.AdmGeralEnviarDiscordancia, _
                                                        enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                                                        enumAcaoConciliacao.AdmGeralRejeitar, _
                                                        enumAcaoConciliacao.AdmGeralRegularizar) Then
                                                        
                        Err.Raise 1, , "CONCORDÂNCIA já foi enviada para esta mensagem."
                    End If
                End If
                
                If blnAcaoJaEnviarDiscordancia Then
                    If fgIN(pintAcao, _
                                                        enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                                                        enumAcaoConciliacao.AdmGeralEnviarDiscordancia, _
                                                        enumAcaoConciliacao.AdmGeralPagamento) Then
                                                        
                        Err.Raise 1, , "DISCORDÂNCIA já foi enviada para esta mensagem."
                    End If
                End If
                
                If blnAcaoJaPagamento Then
                    If fgIN(pintAcao, _
                                                enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                                                enumAcaoConciliacao.AdmGeralPagamento) Then
                                                
                        Err.Raise 1, , "PAGAMENTO já foi enviado para esta mensagem."
                    End If
                End If
                
                If blnAcaoJaPagamentoContingencia Then
                    If fgIN(pintAcao, _
                                        enumAcaoConciliacao.AdmGeralPagamento, _
                                        enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                                        enumAcaoConciliacao.AdmGeralRejeitar, _
                                        enumAcaoConciliacao.AdmGeralEnviarConcordancia, _
                                        enumAcaoConciliacao.AdmGeralEnviarDiscordancia) Then
                    
                        Err.Raise 1, , "PAGAMENTO EM CONTINGÊNCIA já foi enviado para esta mensagem."
                    End If
                End If
                
            End If
        End If
                    
        vntPasso = 250
        
        'Inclusão de conciliação para mensagem, e autorização para regularização
        Select Case pintAcao
            Case enumAcaoConciliacao.AdmAreaPagamentoContingencia, _
                 enumAcaoConciliacao.AdmGeralPagamentoContingencia, _
                 enumAcaoConciliacao.BOEnviarConcordanciaContingencia
                 
                vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                        objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                        objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                        objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, , , , , _
                        False)
        
                Call IncluirConciliacaoOperacao( _
                        vntSequencialConciliacao, _
                        , _
                        objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                        objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                        , _
                        objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text)
        
            Case enumAcaoConciliacao.AdmAreaRegularizar, _
                 enumAcaoConciliacao.AdmGeralRegularizar, _
                 enumAcaoConciliacao.BORegularizar
                 
                If Val(vntSequencialConciliacao) = 0 Then
                    vntSequencialConciliacao = flObterSequenciaConciliacaoMensagem( _
                            objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                            objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                            objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, , , , , _
                            False)
                
                    Call IncluirConciliacaoOperacao( _
                            vntSequencialConciliacao, _
                            , _
                            objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                            objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                            , _
                            objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text)
                End If
        
                If Not flAutorizaRegularizacao(strCodigoMensagem, vntSequencialConciliacao) Then
                    Err.Raise 1, , "Regularização não permitida pois a mensagem de pagamento ainda não está liquidada."
                End If
                
        End Select
        
        '////////////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Se passou consistencias, então 'muda status', 'grava ação', 'manda mensagens',
        '// 'disponibiliza CC'. (JAMBRAR)
        '//
        '////////////////////////////////////////////////////////////////////////////////////////////////
                    
        If blnMensagemMae Then
            'Atualiza status (e grava acao, se for o caso) da mensagem mãe
            If Not IsNull(vntStatusMsgMae) Then
                vntPasso = 300
                Call objMensagem.AlterarStatus(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                               objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                               vntStatusMsgMae, _
                                               objNodeMensagem.selectSingleNode("./DH_ULTI_ATLZ").Text, _
                                               , _
                                               False, _
                                               objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, _
                                               , _
                                               vntAcaoMsgMae)
            End If
            
            'Envia mensagens
            If Not IsNull(vntEnviarMensagem) Then
                vntPasso = 310
                
                If fgIN(pintAcao, enumAcaoConciliacao.AdmGeralEnviarConcordancia, enumAcaoConciliacao.AdmGeralEnviarDiscordancia) Then
                    flMontaRepeticaoLDL0003 pintAcao, _
                                          objNodeMensagem.selectSingleNode("./IN_OPER_DEBT_CRED").Text, _
                                          xmlConciliacao, _
                                          xmlRepeticao, _
                                          lngLocalLiquidacao
                    
                    vntPasso = 313
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                vntEnviarMensagem, _
                                blnIgnoraGradeHorario, _
                                , _
                                True, _
                                xmlRepeticao, _
                                objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text)
                Else
                    vntPasso = 315
                    Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                vntEnviarMensagem, _
                                blnIgnoraGradeHorario, _
                                , _
                                True, _
                                , _
                                objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text)
                
                End If
            End If
            
        Else
            
            If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
            
                'Na BMC, nunca mostra a mensagem mae na tela, pois o maior perfil que interage é o Administrador
                'de Area. Mesmo assim, em alguns casos se atualiza o status da mensagem mae, e enviar mensagens
                
                If Not IsNull(vntStatusMsgMae) Then
                    vntPasso = 300
                    Call objMensagem.AlterarStatus(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                   objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                   vntStatusMsgMae, _
                                                   objNodeMensagem.selectSingleNode("./DH_ULTI_ATLZ").Text, _
                                                   , _
                                                   False, _
                                                   1, _
                                                   , _
                                                   vntAcaoMsgMae)
                End If
                
                'Envia mensagens
                If Not IsNull(vntEnviarMensagem) Then
                    vntPasso = 310
                    
                    If fgIN(pintAcao, enumAcaoConciliacao.BOEnviarConcordancia) Then
                                   
                        flMontaRepeticaoLDL0003 pintAcao, _
                                                objNodeMensagem.selectSingleNode("./IN_OPER_DEBT_CRED").Text, _
                                                xmlConciliacao, _
                                                xmlRepeticao, _
                                                lngLocalLiquidacao
                        
                        vntPasso = 313
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                    vntEnviarMensagem, _
                                    blnIgnoraGradeHorario, _
                                    , _
                                    True, _
                                    xmlRepeticao, _
                                    1)
                                    
                    Else
                    
                        vntPasso = 315
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                    objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                    vntEnviarMensagem, _
                                    blnIgnoraGradeHorario, _
                                    , _
                                    True, _
                                    , _
                                    1)
                    
                    End If
                End If
            
            End If
            
            'Atualiza status das operacoes
            If Not IsNull(vntStatusOp) Or Not IsNull(vntAcaoOperacao) Then
                For Each objNode In objNodeMensagem.selectNodes("./Repeat_Operacoes/*")
                    If Not IsNull(vntStatusOp) Then
                        vntPasso = 400
                        Call objOperacao.AlterarStatusAcao(objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           vntStatusOp, _
                                                           objNode.selectSingleNode("DH_ULTI_ATLZ").Text)
                    End If
                    
                    If Not IsNull(vntAcaoOperacao) Then
                        vntPasso = 410
                        Call objProcessoOperacao.GerenciarChamadaViaTela( _
                                            objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            vntAcaoOperacao, _
                                            False, _
                                            , _
                                            False)
                    End If
                    
                    If fgIN(pintAcao, _
                            enumAcaoConciliacao.BOConcordar, _
                            enumAcaoConciliacao.BODiscordar, _
                            enumAcaoConciliacao.BORegularizar, _
                            enumAcaoConciliacao.AdmAreaRegularizar) Then  'Associa a mãe com as operações
                        
                        vntPasso = 430
                        Call IncluirConciliacaoOperacao(vntSequencialConciliacao, _
                                            objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                            objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                            , _
                                            1, _
                                            True)

                    End If
                    
                    If lngLocalLiquidacao = enumLocalLiquidacao.BMC Then
                        If vntStatusOp = enumStatusOperacao.AConciliar Then
                            If pintAcao = enumAcaoConciliacao.AdmAreaRejeitar Then
                                
                                'ADM AREA Rejeitou, excluir conciliacao
                                ExcluirConciliacaoOperacao , _
                                                           objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                           objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                           1
                                                            
                            End If
                        End If
                    Else
                    
                        If vntStatusOp = enumStatusOperacao.Registrada Then
                            If strCodigoMensagem = "LDL0001" And blnMsgDefinitiva Then
                                'Na LDL0001 definitiva, quando voltar para Registrada (REJEICAO), desassocias as operações da mensagem mãe
                                ExcluirConciliacaoOperacao , _
                                                           objNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                           objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                           objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                           1
                                                            
                            End If
                        End If
    
                    End If
                Next
            End If
             
             
            'Atualiza status (e grava acao, se for o caso) da mensagem (repeticao)
            If Not IsNull(vntStatusMsg) Then
                vntPasso = 500
                Call objMensagem.AlterarStatus(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                               objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                               vntStatusMsg, _
                                               objNodeMensagem.selectSingleNode("./DH_ULTI_ATLZ").Text, _
                                               , _
                                               False, _
                                               objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, _
                                               , _
                                               vntAcaoMsg)
                                               
                'Segunda alteração de status, caso grave mais de um TipoAcao
                If vntAcaoMsg2 <> 0 Then
                    Call objMensagem.AlterarStatus(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                   objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                   vntStatusMsg, _
                                                   objNodeMensagem.selectSingleNode("./DH_ULTI_ATLZ").Text, _
                                                   , _
                                                   False, _
                                                   objNodeMensagem.selectSingleNode("./NU_SEQU_CNTR_REPE").Text, _
                                                   , _
                                                   vntAcaoMsg2)
                End If
                                               
                                               
                If blnFimFluxoAdmArea And lngLocalLiquidacao <> enumLocalLiquidacao.BMC Then
                    'Atualiza o status da msg mae, se o status de todas as filhas já estiver igual ao desta
                                    
                    'Pega os status das mensagens
                    vntPasso = 510
                    xmlFiltro.Load ""
                    fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
                    fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Chave", ""
                    fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_CTRL_IF", objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text
                    fgAppendAttribute xmlFiltro, "Grupo_Chave", "DH_REGT_MESG_SPB", fgDtHrXML_To_Oracle(objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text)
                    fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_SEQU_CNTR_REPE", "" 'não passa nada, pega todas as repetições
                
                    strRet = objMensagemSLCC.ObterDetalheMensagemCamara(xmlFiltro)
                    vntPasso = 520
                    If xmlMensagens.loadXML(strRet) Then
                        'se não tem nenhuma irmã com o status diferente
                        If xmlMensagens.selectNodes("//NU_CTRL_IF[../NU_SEQU_CNTR_REPE!='1' and ../CO_ULTI_SITU_PROC!='" & vntStatusMsg & "']").length = 0 Then
                            vntPasso = 530
                            Call objMensagem.AlterarStatus(objNodeMensagem.selectSingleNode("./NU_CTRL_IF").Text, _
                                                objNodeMensagem.selectSingleNode("./DH_REGT_MESG_SPB").Text, _
                                                vntStatusMsg, _
                                                "", _
                                                , _
                                                False, _
                                                1)
                        End If
                    End If
                End If
                
            End If
        End If

    Next

    If InStr(xmlErros.xml, "Erro") > 0 Then
        'Mensagens de erro
        LiquidarMultilateralFinanceira = xmlErros.xml
    Else
        LiquidarMultilateralFinanceira = ""
    End If
    
    Set xmlErros = Nothing
    Set xmlFiltro = Nothing
    Set objOperacao = Nothing
    Set objMsg = Nothing
    Set objMensagem = Nothing
    Set xmlRepeticao = Nothing
    Set xmlMensagens = Nothing
    Set objHistMensagem = Nothing
    Set objMensagemSLCC = Nothing
    Set objProcessoOperacao = Nothing
    
Exit Function
ErrorHandler:

    Set xmlErros = Nothing
    Set xmlFiltro = Nothing
    Set objOperacao = Nothing
    Set objMsg = Nothing
    Set objMensagem = Nothing
    Set xmlRepeticao = Nothing
    Set xmlMensagens = Nothing
    Set objHistMensagem = Nothing
    Set objMensagemSLCC = Nothing
    Set objProcessoOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarMultilateralFinanceira Function (vntPasso=" & vntPasso & ")", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    Exit Function
    Resume

End Function

'--------------------------------------------------------------------------------------------------
'Monstagem da mensagem de corretoras ( STR0007 )
'--------------------------------------------------------------------------------------------------

Private Function flMontarMsgEnvioLiquidacaoCorretoras(ByVal objDomNodeOrigem As MSXML2.IXMLDOMNode, _
                                                      Optional ByRef objOperacao As A8LQS.clsOperacao) As String

Dim xmlMsg                                  As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40

    
On Error GoTo ErrorHandler
    If objOperacao Is Nothing Then
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
    End If
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    Set xmlMsg = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlOperacao = objOperacao.LerXMLOperacao(objDomNodeOrigem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    If Not xmlOperacao.documentElement.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val("0" & xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
            Call fgAppendNode(xmlMsg, "", "SISMSG", "")
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_MESG", xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_OPER", xmlOperacao.documentElement.selectSingleNode("//TP_OPER").Text)
            
            Call fgAppendNode(xmlMsg, "SISMSG", "SG_SIST_ORIG", xmlOperacao.documentElement.selectSingleNode("//SG_SIST_ORIG").Text)
            
            Call fgAppendNode(xmlMsg, "SISMSG", "SG_SIST_DEST", xmlOperacao.documentElement.selectSingleNode("//SG_SIST_DEST").Text)
            
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_SOLI", xmlOperacao.documentElement.selectSingleNode("//TP_SOLI").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_MESG", xmlOperacao.documentElement.selectSingleNode("//DT_MESG").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "HO_MESG", xmlOperacao.documentElement.selectSingleNode("//HO_MESG").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_USUA_CADR_OPER", xmlOperacao.documentElement.selectSingleNode("//CO_USUA_CADR_OPER").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_LIQU_OPER_ATIV", xmlOperacao.documentElement.selectSingleNode("//DT_LIQU_OPER_ATIV").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_VEIC_LEGA", xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_PROD", xmlOperacao.documentElement.selectSingleNode("//CO_PROD").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CONT", xmlOperacao.documentElement.selectSingleNode("//TP_CONT").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_SEGM", xmlOperacao.documentElement.selectSingleNode("//CO_SEGM").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_EVEN_FINC", xmlOperacao.documentElement.selectSingleNode("//CO_EVEN_FINC").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_INDX", xmlOperacao.documentElement.selectSingleNode("//CO_INDX").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ORIG_RECU", xmlOperacao.documentElement.selectSingleNode("//CO_ORIG_RECU").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_FORM_LIQU", xmlOperacao.documentElement.selectSingleNode("//CO_FORM_LIQU").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_UNI_ORG", xmlOperacao.documentElement.selectSingleNode("//CO_UNI_ORG").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_BANC_DEBT_TRIB", xmlOperacao.documentElement.selectSingleNode("//CO_BANC_DEBT_TRIB").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_AGEN_DEBT_TRIB", xmlOperacao.documentElement.selectSingleNode("//CO_AGEN_DEBT_TRIB").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CONT_DEBT_TRIB", xmlOperacao.documentElement.selectSingleNode("//NU_CONT_DEBT_TRIB").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CONT_DEBT_TRIB", xmlOperacao.documentElement.selectSingleNode("//TP_CONT_DEBT_TRIB").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_DEBT_TRIB", xmlOperacao.documentElement.selectSingleNode("//TP_PESS_DEBT_TRIB").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_CRED", String(8 - Len(xmlOperacao.documentElement.selectSingleNode("//CO_ISPB_IF_CRED").Text), "0") & xmlOperacao.documentElement.selectSingleNode("//CO_ISPB_IF_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_BANC_CRED", xmlOperacao.documentElement.selectSingleNode("//CO_BANC_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_AGEN_CRED", xmlOperacao.documentElement.selectSingleNode("//CO_AGEN_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CONT_CRED", xmlOperacao.documentElement.selectSingleNode("//NU_CONT_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CONT_CRED", xmlOperacao.documentElement.selectSingleNode("//TP_CONT_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_CRED", xmlOperacao.documentElement.selectSingleNode("//TP_PESS_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_1", xmlOperacao.documentElement.selectSingleNode("//CO_CNPJ_CPF_CRED_1").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_CRED_1", xmlOperacao.documentElement.selectSingleNode("//NO_TITU_CRED_1").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_2", "")
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_CRED_2", "")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_BANC_DEBT_ISEN", xmlOperacao.documentElement.selectSingleNode("//CO_BANC_DEBT_ISEN").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_AGEN_DEBT_ISEN", xmlOperacao.documentElement.selectSingleNode("//CO_AGEN_DEBT_ISEN").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CONT_DEBT_ISEN", xmlOperacao.documentElement.selectSingleNode("//NU_CONT_DEBT_ISEN").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CONT_DEBT_ISEN", xmlOperacao.documentElement.selectSingleNode("//TP_CONT_DEBT_ISEN").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_DEBT_ISEN", xmlOperacao.documentElement.selectSingleNode("//TP_PESS_DEBT_ISEN").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_FLND_TED", xmlOperacao.documentElement.selectSingleNode("//CO_FLND_TED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_OPER_ATIV", xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text)


            Call fgAppendNode(xmlMsg, "SISMSG", "CO_MESG", "STR0007")
            Call fgAppendNode(xmlMsg, "SISMSG", "CORRETORAS", vbNullString)
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CTRL_IF", String$(23, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_EMPR", fgCompletaString(objDomNodeOrigem.selectSingleNode("CO_EMPR").Text, "0", 3, True))
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_DOCT", Format$(Now, "HHMMSS"))
            'RATS 1104 - Complementar com zeros a esquerda o ISPB
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_DEBT", String$(8, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_REME", "J")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_REME", String$(15, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_REME", String$(80, " "))
            'RATS 1104 - Complementar com zeros a esquerda o ISPB
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_CRED", String(8 - Len(xmlOperacao.documentElement.selectSingleNode("//CO_ISPB_IF_CRED").Text), "0") & xmlOperacao.documentElement.selectSingleNode("//CO_ISPB_IF_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_AGEN_CRED", xmlOperacao.documentElement.selectSingleNode("//CO_AGEN_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CNTA_CRED", "CC")
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CNTA_CRED", xmlOperacao.documentElement.selectSingleNode("//NU_CONT_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_CRED", xmlOperacao.documentElement.selectSingleNode("//TP_PESS_CRED").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_1", xmlOperacao.documentElement.selectSingleNode("//CO_CNPJ_CPF_CRED_1").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_1", objDomNodeOrigem.selectSingleNode("NO_CNPT").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_2", String$(15, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_2", String$(80, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "VA_OPER_ATIV", objDomNodeOrigem.selectSingleNode("VA_LANC").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_FIND_COBE", "21")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_IDEN_TRANS", String$(25, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "IN_TEA", "N")
            Call fgAppendNode(xmlMsg, "SISMSG", "DE_HIST_MESG", String$(200, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "NIVE_PREF", " ")
            
            If objDomNodeOrigem.selectSingleNode("CO_LOCA_LIQU") Is Nothing Then
                Call fgAppendNode(xmlMsg, "SISMSG", "CO_LOCA_LIQU", objDomNodeOrigem.selectSingleNode("//CO_LOCA_LIQU").Text)
            Else
                Call fgAppendNode(xmlMsg, "SISMSG", "CO_LOCA_LIQU", objDomNodeOrigem.selectSingleNode("CO_LOCA_LIQU").Text)
            End If
            
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_SIST", Format$(fgDataHoraServidor(DataAux), "YYYYMMDD"))
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_AGND", String$(8, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "TX_FILLER_STR0007", String$(667, " "))
           
               
  
    
           flMontarMsgEnvioLiquidacaoCorretoras = xmlMsg.xml
           Set xmlOperacao = Nothing
           Set xmlMsg = Nothing
           Exit Function
        End If
    End If

    
    Select Case Val(objDomNodeOrigem.selectSingleNode("TP_IF_CRED_DEBT").Text)
        Case 1
            Call fgAppendNode(xmlMsg, "", "SISMSG", "")

            Call fgAppendNode(xmlMsg, "SISMSG", "CO_MESG", "STR0007")
            Call fgAppendNode(xmlMsg, "SISMSG", "CORRETORAS", vbNullString)
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CTRL_IF", String$(23, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_EMPR", fgCompletaString(objDomNodeOrigem.selectSingleNode("CO_EMPR").Text, "0", 3, True))
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_DOCT", Format$(Now, "HHMMSS"))
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_DEBT", String$(8, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_REME", "J")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_REME", String$(15, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_REME", String$(80, " "))
            'RATS 1104 - Complementar com zeros a esquerda o ISPB
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_CRED", String(8 - Len(objDomNodeOrigem.selectSingleNode("CO_ISPB_CNPT").Text), "0") & objDomNodeOrigem.selectSingleNode("CO_ISPB_CNPT").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_AGEN_CRED", objDomNodeOrigem.selectSingleNode("CO_AGEN_COTR").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_CNTA_CRED", "CC")
            Call fgAppendNode(xmlMsg, "SISMSG", "NU_CNTA_CRED", objDomNodeOrigem.selectSingleNode("NU_CC_COTR").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "TP_PESS_CRED", "J")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_1", objDomNodeOrigem.selectSingleNode("CO_CNPJ_CNPT").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_1", objDomNodeOrigem.selectSingleNode("NO_CNPT").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_CNPJ_CPF_CRED_2", String$(15, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "NO_TITU_2", String$(80, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "VA_OPER_ATIV", objDomNodeOrigem.selectSingleNode("VA_LANC").Text)
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_FIND_COBE", "21")
            Call fgAppendNode(xmlMsg, "SISMSG", "CO_IDEN_TRANS", String$(25, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "IN_TEA", "N")
            Call fgAppendNode(xmlMsg, "SISMSG", "DE_HIST_MESG", String$(200, " "))
            Call fgAppendNode(xmlMsg, "SISMSG", "NIVE_PREF", " ")
            
            If objDomNodeOrigem.selectSingleNode("CO_LOCA_LIQU") Is Nothing Then
                Call fgAppendNode(xmlMsg, "SISMSG", "CO_LOCA_LIQU", objDomNodeOrigem.selectSingleNode("//CO_LOCA_LIQU").Text)
            Else
                Call fgAppendNode(xmlMsg, "SISMSG", "CO_LOCA_LIQU", objDomNodeOrigem.selectSingleNode("CO_LOCA_LIQU").Text)
            End If
            
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_SIST", Format$(fgDataHoraServidor(DataAux), "YYYYMMDD"))
            Call fgAppendNode(xmlMsg, "SISMSG", "DT_AGND", String$(8, "0"))
            Call fgAppendNode(xmlMsg, "SISMSG", "TX_FILLER_STR0007", String$(667, " "))

    End Select

    flMontarMsgEnvioLiquidacaoCorretoras = xmlMsg.xml
    Set xmlMsg = Nothing

    Exit Function

ErrorHandler:
    Set xmlMsg = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMsgEnvioLiquidacaoCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Monstagem da mensagem para envio ( STR0004 )
'--------------------------------------------------------------------------------------------------

Private Function flMontarMsgEnvioSTR0004(ByVal objDomNodeOrigem As MSXML2.IXMLDOMNode) As String

Dim xmlMsg                                  As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
        
    Set xmlMsg = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlMsg, "", "SISMSG", "")

    Call fgAppendNode(xmlMsg, "SISMSG", "CO_MESG", "STR0004")
    Call fgAppendNode(xmlMsg, "SISMSG", "CORRETORAS", vbNullString)
    Call fgAppendNode(xmlMsg, "SISMSG", "NU_CTRL_IF", String$(23, " "))
    Call fgAppendNode(xmlMsg, "SISMSG", "CO_EMPR", fgCompletaString(objDomNodeOrigem.selectSingleNode("CO_EMPR").Text, "0", 3, True))
    Call fgAppendNode(xmlMsg, "SISMSG", "NU_DOCT", Format$(Now, "HHMMSS"))
    Call fgAppendNode(xmlMsg, "SISMSG", "DT_SIST", Format$(fgDataHoraServidor(DataAux), "YYYYMMDD"))
    Call fgAppendNode(xmlMsg, "SISMSG", "CO_ISPB_IF_CRED", vbNullString)
    Call fgAppendNode(xmlMsg, "SISMSG", "VA_OPER_ATIV", objDomNodeOrigem.selectSingleNode("VA_OPER_ATIV").Text)
    Call fgAppendNode(xmlMsg, "SISMSG", "CO_FIND_COBE", "21")
    Call fgAppendNode(xmlMsg, "SISMSG", "DE_HIST_MESG", Space$(200))
    Call fgAppendNode(xmlMsg, "SISMSG", "DT_AGND", String$(8, 0))
    Call fgAppendNode(xmlMsg, "SISMSG", "TX_FILLER_STR0004", String$(969, " "))
    Call fgAppendNode(xmlMsg, "SISMSG", "CO_IDEF_TRAF", String$(25, " "))
    Call fgAppendNode(xmlMsg, "SISMSG", "CO_LOCA_LIQU", enumLocalLiquidacao.BMD)
            
    flMontarMsgEnvioSTR0004 = xmlMsg.xml
    Set xmlMsg = Nothing

    Exit Function

ErrorHandler:
    Set xmlMsg = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMsgEnvioSTR0004 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Montagem das repetições da LDL0003
'--------------------------------------------------------------------------------------------------

Public Function flMontaRepeticaoLDL0003(ByVal pintAcao As enumAcaoConciliacao, _
                                      ByVal plngNaturezaMovimento As Long, _
                                      ByRef pxml As MSXML2.DOMDocument40, _
                                      ByRef pxmlRepeticao As MSXML2.DOMDocument40, _
                                      ByVal plngLocalLiquidacao As Long)

'Monta a estrutura de repetição da LDL0003, de concordancia ou discordancia.
    
Dim objNodeMsg                  As IXMLDOMNode
Dim vntValor                    As Variant
Dim vntConfDivg                 As Variant
        
On Error GoTo ErrorHandler
        
    pxmlRepeticao.loadXML ""
    
    fgAppendNode pxmlRepeticao, "", "Repet_LDL0003_ResultLiqd", "", ""
    
    'Para cada item de Repetição
    For Each objNodeMsg In pxml.selectNodes("//Grupo_Mensagem[NU_SEQU_CNTR_REPE!='1']")
        
        'Não envia valores negativos. Se for negativo envia Zero
        
        If plngLocalLiquidacao = enumLocalLiquidacao.BMC Then
            vntValor = Val(Replace(objNodeMsg.selectSingleNode("./ValorCamara").Text, ",", "."))
            
            vntConfDivg = "C"
        
        Else
            vntValor = Val(Replace(objNodeMsg.selectSingleNode("./ValorOperacoes").Text, ",", "."))
            
            vntConfDivg = IIf(objNodeMsg.selectSingleNode("./ValorOperacoes").Text = objNodeMsg.selectSingleNode("./ValorCamara").Text, _
                                "C", _
                                "D")
        End If
                            
        If plngNaturezaMovimento = enumTipoDebitoCredito.Debito Then
            vntValor = vntValor * -1
        End If
        vntValor = fgMax(vntValor, 0)
    
        fgAppendNode pxmlRepeticao, "Repet_LDL0003_ResultLiqd", "Grupo_LDL0003_ResultLiqd", "", ""
        
        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_CNPJ_LIQU", _
                                                        objNodeMsg.selectSingleNode("./CNPJ").Text, "Repet_LDL0003_ResultLiqd"
        
        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_PARP_CAMR", _
                                                        "", "Repet_LDL0003_ResultLiqd"
        
        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "VA_RESL_LIQU", _
                                                        fgVlr_To_Xml(vntValor), "Repet_LDL0003_ResultLiqd"
        
        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "TP_CONF_DIVG", _
                                                        vntConfDivg, "Repet_LDL0003_ResultLiqd"
    Next
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaRepeticaoLDL0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    Exit Function
    Resume
    
End Function

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Public Function IncluirConciliacaoSEL1023(ByVal pstrNumeroControleIFConciliado As String, _
                                          ByVal pstrNumeroControleIF As String) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_MESG_RECB_SPB_CNCL (   " & vbNewLine & _
             "       NU_SEQU_CNTR_REPE,             " & vbNewLine & _
             "       NU_CTRL_IF,                    " & vbNewLine & _
             "       DT_OPER,                       " & vbNewLine & _
             "       CO_OPER_SELIC,             " & vbNewLine & _
             "       NU_COMD_OPER_ORIG,             " & vbNewLine & _
             "       VA_FINC,                       " & vbNewLine & _
             "       DH_REGT_MESG_SPB,              " & vbNewLine & _
             "       NU_ATIV_MERC,                  " & vbNewLine & _
             "       PU_ATIV_MERC,                  " & vbNewLine & _
             "       QT_ATIV_MERC,                  " & vbNewLine & _
             "       DT_LIQU,                       " & vbNewLine & _
             "       IN_OPER_DEBT_CRED,             " & vbNewLine & _
             "       DT_VENC,                       " & vbNewLine & _
             "       CO_CNPJ_CNPT,                  " & vbNewLine & _
             "       NO_CNPT,                       " & vbNewLine & _
             "       CO_CNTA_CUTD_SELIC_VEIC_LEGA,  " & vbNewLine & _
             "       TP_INFO_LDL,                   " & vbNewLine & _
             "       PE_TAXA_NEGO,                  " & vbNewLine & _
             "       CO_TITL_CUTD,                  " & vbNewLine & _
             "       TP_NEGO_BMA)                   " & vbNewLine
    strSQL = strSQL & _
             "         SELECT                           " & vbNewLine & _
             "          NU_SEQU_CNTR_REPE,              " & vbNewLine & _
                        "'" & pstrNumeroControleIF & "'," & vbNewLine & _
             "          DT_OPER,                        " & vbNewLine & _
             "          CO_OPER_SELIC,              " & vbNewLine & _
             "          NU_COMD_OPER_ORIG,              " & vbNewLine & _
             "          VA_FINC,                        " & vbNewLine & _
             "          DH_REGT_MESG_SPB,               " & vbNewLine & _
             "          NU_ATIV_MERC,                   " & vbNewLine & _
             "          PU_ATIV_MERC,                   " & vbNewLine & _
             "          QT_ATIV_MERC,                   " & vbNewLine & _
             "          DT_LIQU,                        " & vbNewLine & _
             "          IN_OPER_DEBT_CRED,              " & vbNewLine & _
             "          DT_VENC,                        " & vbNewLine & _
             "          CO_CNPJ_CNPT,                   " & vbNewLine & _
             "          NO_CNPT,                        " & vbNewLine & _
             "          CO_CNTA_CUTD_SELIC_VEIC_LEGA,   " & vbNewLine & _
             "          TP_INFO_LDL,                    " & vbNewLine & _
             "          PE_TAXA_NEGO,                   " & vbNewLine & _
             "          CO_TITL_CUTD,                   " & vbNewLine & _
             "          TP_NEGO_BMA                     " & vbNewLine & _
             "         FROM A8.TB_MESG_RECB_SPB_CNCL " & vbNewLine & _
             "         WHERE    NU_CTRL_IF  =   '" & pstrNumeroControleIFConciliado & "'"

    fgExecuteSQL strSQL

    strSQL = " INSERT INTO A8.TB_CNCL_OPER_ATIV (   " & vbNewLine & _
             "         NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbNewLine & _
             "         NU_SEQU_OPER_ATIV,           " & vbNewLine & _
             "         NU_CTRL_IF,                  " & vbNewLine & _
             "         DH_REGT_MESG_SPB,            " & vbNewLine & _
             "         QT_ATIV_MERC_CNCL,           " & vbNewLine & _
             "         NU_SEQU_CNTR_REPE)           " & vbNewLine & _
             "         SELECT                      " & vbNewLine & _
             "         NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbNewLine & _
             "         NU_SEQU_OPER_ATIV,           " & vbNewLine & _
                    "'" & pstrNumeroControleIF & "'   ," & vbNewLine & _
             "         DH_REGT_MESG_SPB,            " & vbNewLine & _
             "         QT_ATIV_MERC_CNCL,           " & vbNewLine & _
             "         NU_SEQU_CNTR_REPE            " & vbNewLine & _
             "         FROM A8.TB_CNCL_OPER_ATIV    " & vbNewLine & _
             "         WHERE    NU_CTRL_IF  =   '" & pstrNumeroControleIFConciliado & "'"

    fgExecuteSQL strSQL

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirConciliacaoSEL1023 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'
'--------------------------------------------------------------------------------------------------

Function flObterSequenciaConciliacaoMensagem(ByVal pstrNumeroControleIF As String, _
                                             ByVal pstrDataRegistroMensagemSPB As String, _
                                             ByVal pintNumeroSequencialContrRepeticao As Integer, _
                                    Optional ByVal pblnSequencialDaMensagemMae As Boolean = False, _
                                    Optional ByVal pblnCriarSequencia As Boolean = True, _
                                    Optional ByVal pstrCodigoMensagem As String = vbNullString, _
                                    Optional ByVal blnFluxoAutomatico As Boolean = False, _
                                    Optional ByVal pblnSegregarTipoBackOffice As Boolean = True)

'Obtém a sequencia de conciliação da mensagem, ou cria se pblnCriarSequencia=TRUE
'Se blnSequencialDaMensagemMae = True, o sequencial devolvido/criado é o da mensagem mãe da mensagem passada (nu_sequ_cntr_repe=1)
    
Dim xmlFiltro                           As MSXML2.DOMDocument40
Dim xmlMensagem                         As MSXML2.DOMDocument40
Dim objMensagemSLCC                     As A8LQS.clsMensagemSLCC
Dim objMensagem                         As A8LQS.clsMensagemEnviadaRecebida
Dim vntSequencia                        As Variant
Dim strRet                              As String
Dim strCampoSeq                         As String
Dim objDomNode                          As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    flObterSequenciaConciliacaoMensagem = Null
    vntSequencia = Null
    
    If pblnSequencialDaMensagemMae Then
        strCampoSeq = "NU_SEQU_CNCL_OPER_ATIV_MESG_MA"
    Else
        strCampoSeq = "NU_SEQU_CNCL_OPER_ATIV_MESG"
    End If

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    If pstrCodigoMensagem <> "LTR0007" Then
        xmlFiltro.Load ""
        fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
        
        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Chave", ""
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_CTRL_IF", pstrNumeroControleIF
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "DH_REGT_MESG_SPB", fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB)
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_SEQU_CNTR_REPE", pintNumeroSequencialContrRepeticao
        
        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_SegregaBackOffice", ""
        fgAppendNode xmlFiltro, "Grupo_SegregaBackOffice", "SegregaBackOffice", pblnSegregarTipoBackOffice
    
        strRet = objMensagemSLCC.ObterDetalheMensagemCamara(xmlFiltro, blnFluxoAutomatico)
        If xmlMensagem.loadXML(strRet) Then
            vntSequencia = xmlMensagem.selectSingleNode("//" & strCampoSeq).Text
            If vntSequencia = "0" And pblnCriarSequencia Then
                vntSequencia = IncluirJustificativaConciliacaoOperacao(0)
        
                Call objMensagem.AlterarSequenciaConciliacao(pstrNumeroControleIF, _
                                               fgDtHrStr_To_DateTime(pstrDataRegistroMensagemSPB), _
                                               vntSequencia, _
                                               IIf(pblnSequencialDaMensagemMae, 1, pintNumeroSequencialContrRepeticao))
            
            End If
        End If
        
    Else
        xmlFiltro.Load ""
        fgAppendNode xmlFiltro, "", "Repeat_Filtros", ""
        fgAppendNode xmlFiltro, "Repeat_Filtros", "Grupo_Chave", ""
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_CTRL_IF", pstrNumeroControleIF
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "DH_REGT_MESG_SPB", vbNullString
        fgAppendAttribute xmlFiltro, "Grupo_Chave", "NU_SEQU_CNTR_REPE", pintNumeroSequencialContrRepeticao
    
        strRet = objMensagemSLCC.ObterDetalheMensagemCamara(xmlFiltro, True)
    
        If xmlMensagem.loadXML(strRet) Then
            vntSequencia = "0"
            
            For Each objDomNode In xmlMensagem.selectNodes("Repeat_DetalheMensagemCamara/*")
                vntSequencia = objDomNode.selectSingleNode(strCampoSeq).Text
                If vntSequencia <> "0" Then Exit For
            Next
    
            If vntSequencia = "0" Then
                vntSequencia = IncluirJustificativaConciliacaoOperacao(0)
            End If
        
            Call objMensagem.AlterarSequenciaConciliacao(pstrNumeroControleIF, _
                                           fgDtHrStr_To_DateTime(pstrDataRegistroMensagemSPB), _
                                           vntSequencia, _
                                           IIf(pblnSequencialDaMensagemMae, 1, pintNumeroSequencialContrRepeticao))
            
        End If
    
    End If

    flObterSequenciaConciliacaoMensagem = vntSequencia

    Set xmlFiltro = Nothing
    Set xmlMensagem = Nothing
    Set objMensagem = Nothing
    Set objMensagemSLCC = Nothing

    Exit Function
ErrorHandler:

    Set xmlFiltro = Nothing
    Set xmlMensagem = Nothing
    Set objMensagem = Nothing
    Set objMensagemSLCC = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterSequenciaConciliacaoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a exclusão de um registro na tabela:
'    TB_MESG_RECB_SPB_CNCL
Public Function ExcluirConciliacaoOperacao(Optional ByVal pvntSequencialConciliacaoOperacao As Variant = vbNullString, _
                                  Optional ByVal pvntSequenciaOperacao As Variant = vbNullString, _
                                  Optional ByVal pvntNumeroControleIF As Variant = vbNullString, _
                                  Optional ByVal pvntDataRegistroMensagemSPB As Variant = vbNullString, _
                                  Optional ByVal pvntNumeroSequencialContrRepeticao As Variant = vbNullString) As Boolean

Dim strSQL                                  As String
Dim strWhere                                As String

On Error GoTo ErrorHandler
    
    strWhere = ""
    
    If Trim(pvntSequencialConciliacaoOperacao) <> vbNullString Then
        strWhere = strWhere & IIf(strWhere <> "", " AND ", " ") & _
                    "NU_SEQU_CNCL_OPER_ATIV_MESG = " & fgVlrToDBServer(pvntSequencialConciliacaoOperacao) & vbCrLf
    End If
    If Trim(pvntSequenciaOperacao) <> vbNullString Then
        strWhere = strWhere & IIf(strWhere <> "", " AND ", " ") & _
                    "NU_SEQU_OPER_ATIV = " & fgVlrToDBServer(pvntSequenciaOperacao) & vbCrLf
    End If
    If Trim(pvntNumeroControleIF) <> vbNullString Then
        strWhere = strWhere & IIf(strWhere <> "", " AND ", " ") & _
                    "NU_CTRL_IF = '" & pvntNumeroControleIF & "'" & vbCrLf
    End If
    If Trim(pvntDataRegistroMensagemSPB) <> vbNullString Then
        strWhere = strWhere & IIf(strWhere <> "", " AND ", " ") & _
                    "DH_REGT_MESG_SPB = " & fgDtHrXML_To_Oracle(pvntDataRegistroMensagemSPB) & vbCrLf
    End If
    If Trim(pvntNumeroSequencialContrRepeticao) <> vbNullString Then
        strWhere = strWhere & IIf(strWhere <> "", " AND ", " ") & _
                    "NU_SEQU_CNTR_REPE = " & fgVlrToDBServer(pvntNumeroSequencialContrRepeticao) & vbCrLf
    End If

    strSQL = " DELETE FROM A8.TB_CNCL_OPER_ATIV         " & vbNewLine & _
             " WHERE   " & strWhere
    
    fgExecuteSQL strSQL
    
    Exit Function
    
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExcluirConciliacaoOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a alteração do CNPJ da contraparte na tabela:
'    TB_MESG_RECB_SPB_CNCL
Public Function AlterarCNPJContraParte(ByVal pstrNumeroControleIF As String, _
                                       ByVal pstrDataRegistroMensagem As String, _
                                       ByVal pvntNumeroSequenciaRepeticao As Variant, _
                                       ByVal pvntCNPJContraparte As Variant) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_SPB_CNCL " & vbNewLine & _
             " Set      CO_CNPJ_CNPT            =   " & pvntCNPJContraparte & vbNewLine & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'" & vbNewLine & _
             " AND      DH_REGT_MESG_SPB        =   " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbNewLine & _
             " AND      NU_SEQU_CNTR_REPE       =   " & pvntNumeroSequenciaRepeticao
    
    Call fgExecuteSQL(strSQL)

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarCNPJContraParte Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Retorno de tipo de Operacao CETIP que fazem conciliacao
Public Function ObterTiposConciliacaoCETIP() As String
 
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT     DISTINCT TP_OPER                          " & vbNewLine & _
             " FROM       A8.TB_TIPO_OPER_CNCL_CETIP                " & vbNewLine & _
             " ORDER BY   TP_OPER                                   "

    ObterTiposConciliacaoCETIP = fgQueryXMLLerTodos("TiposConciliacaoCETIP", strSQL, App.EXEName & "." & TypeName(Me))

    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not ObjectContext Is Nothing Then
       'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterTiposConciliacaoCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Efetua a liquidação de operações, em conjunto com o sistem TQ (provisoriamente)
Public Function LiquidarOperacaoTQ(ByVal pvntSequencialOperacao As Long)

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim objPM                                   As A8LQS.clsProcessoMensagemBMA
    
On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objPM = CreateObject("A8LQS.clsProcessoMensagemBMA")
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    ''Alterar status para Liquidada
    'objOperacao.AlterarStatusAcao pvntSequencialOperacao, enumStatusOperacao.Liquidada, "", , , False
    
    'Chamar GerenciarChamada, para envio de A6, CC Disponivel e Resposta para Legado
    Set xmlOperacao = objOperacao.LerXMLOperacao(pvntSequencialOperacao)
    
    objPM.GerenciarChamada xmlOperacao, xmlMensagem, enumStatusOperacao.Liquidada, 0, "", enumIndicadorSimNao.Nao, "", False, True
        
        
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    Set objPM = Nothing
    
Exit Function
ErrorHandler:
    
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing
    Set objPM = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarOperacaoTQ Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Verifica se uma mensagem enviada (Ex: LTR0004) já está liquidada, para autorizar a regularização por contingência

Private Function flAutorizaRegularizacao(ByVal pstrCodigoMensagemSPB As String, _
                                         ByVal pvntSequenciaConciliacaoMensagemOperacao As Variant) As Boolean
                                         
Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim objMensagem                             As A8LQS.clsMensagemSLCC
Dim strCodMesgEnviada                       As String
Dim intStatusMensagem                       As Integer

    On Error GoTo ErrorHandler
    
    Select Case pstrCodigoMensagemSPB
        Case "BMA0015"
            strCodMesgEnviada = "SEL1023"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "BMC0101"
            strCodMesgEnviada = "BMC0102"
            intStatusMensagem = enumStatusMensagem.Confirmada
        Case "LDL0001"
            strCodMesgEnviada = "LDL0004"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LDL0007"
            strCodMesgEnviada = "LDL0008"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LTR0001"
            strCodMesgEnviada = "LTR0004"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
        Case "LTR0007"
            strCodMesgEnviada = "LTR0008"
            intStatusMensagem = enumStatusMensagem.MensagemLiquidada
    End Select
    
    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = CreateObject("A8LQS.clsMensagemSLCC")
    
    Call fgAppendNode(xmlFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltros, "Grupo_Status", "Status", intStatusMensagem)
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
    Call fgAppendNode(xmlFiltros, "Grupo_CodigoMensagem", "CodigoMensagem", strCodMesgEnviada)
    
    Call fgAppendNode(xmlFiltros, "Repeat_Filtros", "Grupo_SequencialConciliacao", "")
    Call fgAppendNode(xmlFiltros, "Grupo_SequencialConciliacao", "SequencialConciliacao", pvntSequenciaConciliacaoMensagemOperacao)
    
    Call xmlLeitura.loadXML(objMensagem.ObterDetalheMensagem(xmlFiltros))
    flAutorizaRegularizacao = IIf(xmlLeitura.xml = vbNullString, False, True)
    
    Set xmlFiltros = Nothing
    Set xmlLeitura = Nothing
    Set objMensagem = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlFiltros = Nothing
    Set xmlLeitura = Nothing
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAutorizaRegularizacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'KIDA - CBLC
Private Function flExisteLDL0005R2(ByVal objDomNodeProc As MSXML2.IXMLDOMNode) As Boolean

Dim strSQL                                  As String
Dim strNumCtrlCamara                        As String
Dim rsMESG                                  As ADODB.Recordset

On Error GoTo ErrorHandler
        
    strSQL = "  SELECT  A.NU_CTRL_CAMR             " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A " & vbCrLf & _
             "  Where   A.NU_CTRL_IF            = '" & objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text & "'" & vbCrLf & _
             "  AND     A.DH_REGT_MESG_SPB      =  " & fgDtHrXML_To_Oracle(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text)
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    If Not rsMESG.EOF Then
        strNumCtrlCamara = rsMESG("NU_CTRL_CAMR")
    Else
        Exit Function
    End If
    
    Set rsMESG = Nothing
    
    strSQL = "  SELECT  A.NU_CTRL_IF                 " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A   " & vbCrLf & _
             "  Where   A.NU_CTRL_CAMR            = '" & strNumCtrlCamara & "'" & vbCrLf & _
             "  AND     A.CO_MESG_SPB             = 'LDL0005R2'"
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    flExisteLDL0005R2 = Not rsMESG.EOF
    
    Set rsMESG = Nothing
    
    Exit Function

ErrorHandler:
    Set rsMESG = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteLDL0005R2 Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'KIDA - CBLC
Private Function flRegularizarRecebimentoMultiCBLC(ByVal objDomNodeProc As MSXML2.IXMLDOMNode)


Dim objMesgRecbEnviSPB                      As A8LQS.clsMensagemEnviadaRecebida
Dim objMensagemBMA                          As A8LQS.clsProcessoMensagemBMA
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlCociliacao                           As MSXML2.DOMDocument40
Dim xmlMesgRecbEnviSPB                      As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim vntSequenciaConciliacaoOperacao         As Variant
Dim vntSequenciaOperacao                    As Variant

Dim strSQL                                  As String
Dim strNumCtrlCamara                        As String
Dim strNumCtrlIF                            As String
Dim rsMESG                                  As ADODB.Recordset

On Error GoTo ErrorHandler
        
    strSQL = "  SELECT  A.NU_CTRL_CAMR             " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A " & vbCrLf & _
             "  Where   A.NU_CTRL_IF            = '" & objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text & "'" & vbCrLf & _
             "  AND     A.DH_REGT_MESG_SPB      =  " & fgDtHrXML_To_Oracle(objDomNodeProc.selectSingleNode("DH_REGT_MESG_SPB").Text)
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    If Not rsMESG.EOF Then
        strNumCtrlCamara = rsMESG("NU_CTRL_CAMR")
    Else
        Exit Function
    End If
    
    Set rsMESG = Nothing
    
    
    strSQL = "  SELECT  A.NU_CTRL_IF             " & vbCrLf & _
             "  From    A8.TB_MESG_RECB_ENVI_SPB A " & vbCrLf & _
             "  Where   A.NU_CTRL_CAMR           = '" & strNumCtrlCamara & "'" & vbCrLf & _
             "  AND     A.CO_MESG_SPB            = 'LDL0005R2'"
    
    Set rsMESG = fgQuerySQL(strSQL)
    
    If Not rsMESG.EOF Then
        strNumCtrlIF = rsMESG("NU_CTRL_IF")
    Else
        Exit Function
    End If
    
    Set rsMESG = Nothing
    
    
    strSQL = "  UPDATE  A8.TB_MESG_RECB_ENVI_SPB A   " & vbCrLf & _
             "  SET     A.CO_ULTI_SITU_PROC      = " & enumStatusMensagem.MensagemLiquidada & vbCrLf & _
             "  Where   A.NU_CTRL_IF           = '" & strNumCtrlIF & "'" & vbCrLf & _
             "  AND     A.CO_MESG_SPB            = 'LDL0005R2'"
    
    Call fgExecuteSQL(strSQL)
    
    Set objMesgRecbEnviSPB = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set xmlMesgRecbEnviSPB = CreateObject("MSXML2.DOMDocument.4.0")
    
    strSQL = objMesgRecbEnviSPB.LerXMLMensagem(objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text, "1").xml
    
    If Not xmlMesgRecbEnviSPB.loadXML(strSQL) Then
        'Número de Controle IF da Mensagem R1 inválido.
        lngCodigoErroNegocio = 3022
        GoTo ErrorHandler
    End If
                    
    vntSequenciaConciliacaoOperacao = IIf(xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text = "", 0, xmlMesgRecbEnviSPB.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
    
    If vntSequenciaConciliacaoOperacao = 0 Then
        Set xmlMesgRecbEnviSPB = Nothing
        Set objMesgRecbEnviSPB = Nothing
        Exit Function
    End If
    
    '-------------------------------------------------------------------------------------------------------------
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    Set xmlCociliacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
    
    If xmlCociliacao.loadXML(objConciliacao.LerTodos(vntSequenciaConciliacaoOperacao, 0, objDomNodeProc.selectSingleNode("NU_CTRL_IF").Text)) Then
    
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        
        
        For Each xmlNode In xmlCociliacao.selectNodes("Repeat_Conciliacao/*")
            
            Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            vntSequenciaOperacao = xmlNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
            Set xmlOperacao = objOperacao.LerXMLOperacao(vntSequenciaOperacao)
                
            objMensagemBMA.GerenciarChamada xmlOperacao, _
                                            xmlMensagem, _
                                            enumStatusOperacao.Liquidada, _
                                            0, _
                                            0, _
                                            enumIndicadorSimNao.Nao, _
                                            ""
            Set xmlOperacao = Nothing
        Next
        
        Set objOperacao = Nothing
    End If
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing

    Exit Function
ErrorHandler:
    
    Set objConciliacao = Nothing
    Set xmlCociliacao = Nothing
    Set xmlOperacao = Nothing
    Set xmlMesgRecbEnviSPB = Nothing
    Set objMesgRecbEnviSPB = Nothing
    Set objOperacao = Nothing
       
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flProcessaMensagemR1Pagamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Public Function AlterarStatusLote(ByVal plngIdLote As Long, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv As Date, ByVal pintSituacaoLote As Integer)
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    On Error GoTo ErrorHandler
    
        Set rsOperacoes = CreateObject("ADODB.Recordset")
        strSQL = "Update a8.tb_lote_oper_ativ l " & vbCrLf & _
                 " set l.in_liqu_lote = " & pintSituacaoLote & "," & vbCrLf & _
                 "     l.dh_ulti_atlz = sysdate " & vbCrLf & _
                 " Where L.CD_LOTE = " & plngIdLote & vbCrLf & _
                 " And   L.SG_SIST         =   '" & pstrSiglaSistema & "' " & vbCrLf & _
                 " And   L.DT_OPER_ATIV    =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv))
                                  
        Set rsOperacoes = fgQuerySQL(strSQL)
        
        Set AlterarStatusLote = rsOperacoes
        
        Set rsOperacoes = Nothing
        
        If Not ObjectContext Is Nothing Then
            ObjectContext.SetComplete
        End If
    
    Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function



