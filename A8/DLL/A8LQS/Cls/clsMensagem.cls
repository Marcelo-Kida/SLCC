VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF769BD02A7"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Manutenções e Consultas em Mensagens do sistema

Option Explicit

Private xmlDOMDesc                          As MSXML2.DOMDocument40

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strCodErroBACEN                     As String


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

'Método genérico para a execução de outros métodos desta classe
Public Function Executar(ByRef pxmlExecucao As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlExecucao.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodosTiposMensagens"
            Executar = flLerTodosTiposMensagens(pxmlExecucao)
            
        Case "LerTodos"
            Executar = LerTodosMensagemSPBGrupoUsuario()
        
        Case "Reenviar"
            Call ReenviarMensagem(pxmlExecucao.selectSingleNode("//NU_CTRL_IF").Text, _
                                  pxmlExecucao.selectSingleNode("//DH_ULTI_ATLZ").Text)
        
        Case "RejeitarTED"
            Call RejeitarTED(pxmlExecucao.selectSingleNode("//NU_CTRL_IF").Text, _
                             pxmlExecucao.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                             pxmlExecucao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                             pxmlExecucao.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
                        
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    
    End Select

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Obter as mensagens associadas pelo numero de controle IF.
Public Function ObterMensagensAssociadas(ByVal pstrNumeroControleIF As String, _
                                         ByVal pstrDataRegistro As String, _
                                         Optional ByVal pstrNumeroSequenciaControleRepeticao As String = 1) As String

Dim strSQL                                  As String
Dim datDataCorte                            As Date

On Error GoTo ErrorHandler

    strSQL = " SELECT     A.NU_CTRL_IF,                    " & vbCrLf & _
             "            A.DH_REGT_MESG_SPB,              " & vbCrLf & _
             "            A.CO_LOCA_LIQU,                  " & vbCrLf & _
             "            A.CO_VEIC_LEGA,                  " & vbCrLf & _
             "            A.SG_SIST,                       " & vbCrLf & _
             "            A.NU_SEQU_OPER_ATIV,             " & vbCrLf & _
             "            A.TP_BKOF,                       " & vbCrLf & _
             "            A.CO_EMPR,                       " & vbCrLf & _
             "            A.DH_RECB_ENVI_MESG_SPB,         " & vbCrLf & _
             "            A.CO_MESG_SPB,                   " & vbCrLf & _
             "            A.NU_COMD_OPER,                  " & vbCrLf & _
             "            A.CO_SITU_MESG_SPB,              " & vbCrLf & _
             "            B.DE_SITU_PROC,                  " & vbCrLf & _
             "            A.CO_TEXT_XML,                   " & vbCrLf & _
             "            A.HO_ENVI_MESG_SPB,              " & vbCrLf & _
             "            A.CO_ULTI_SITU_PROC,             " & vbCrLf & _
             "            A.IN_ENTR_MANU,                  " & vbCrLf & _
             "            A.CO_USUA_ULTI_ATLZ,             " & vbCrLf & _
             "            A.CO_ETCA_TRAB_ULTI_ATLZ,        " & vbCrLf & _
             "            A.DH_ULTI_ATLZ                   " & vbCrLf & _
             " FROM A8.TB_MESG_RECB_ENVI_SPB        A,     " & vbCrLf & _
             "      A8.TB_SITU_PROC                 B      " & vbCrLf
             
    strSQL = strSQL & _
             " WHERE      A.CO_ULTI_SITU_PROC = B.CO_SITU_PROC " & vbCrLf & _
             " AND        A.NU_CTRL_IF = '" & pstrNumeroControleIF & "'" & vbCrLf & _
             " AND        A.NU_SEQU_CNTR_REPE = " & pstrNumeroSequenciaControleRepeticao & ""
             
             '----------------------------------------------------------------------------------------
             '>>> Não filtrar por DATA de REGISTRO, todas as mensagens, inclusive a própria,
             '    deverão ser retornadas.
             '----------------------------------------------------------------------------------------
             '" AND        A.DH_REGT_MESG_SPB = " & fgDtHrXML_To_Oracle(pstrDataRegistro)
             '----------------------------------------------------------------------------------------


    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
    
    If fgDtXML_To_Date(Mid$(pstrDataRegistro, 1, 8)) < datDataCorte Then
        strSQL = strSQL & " UNION ALL " & Replace$(strSQL, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
    End If


    strSQL = strSQL & _
            " ORDER BY DH_REGT_MESG_SPB DESC "
    
    ObterMensagensAssociadas = fgQueryXMLLerTodos("Mensagem", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensAssociadas Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'' Ler a tabela TB_INSTITUICAO_SPB com as condições local de liquidação e empresa
'' fusinada.
Public Function flObterISPBLDL(ByVal plngCodigoEmpresaFusionada As Long, _
                               ByVal plngLocalLiquidacao As Long, _
                      Optional ByRef plngSequenciaLocalLiquidacao As Long) As Variant


Dim strSQL                                  As String
Dim rsISPB                                  As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT CO_ISPB, SQ_ISPB  " & vbCrLf & _
             " FROM A8.TB_INSTITUICAO_SPB   " & vbCrLf & _
             " WHERE SQ_ISPB = (SELECT SQ_ISPB FROM A8.TB_LOCAL_LIQUIDACAO " & vbCrLf & _
                              " WHERE  CO_LOCA_LIQU = " & plngLocalLiquidacao & "  " & vbCrLf & _
                              " AND    CO_EMPR_FUSI = " & plngCodigoEmpresaFusionada & ")"

    Set rsISPB = fgQuerySQL(strSQL)
    
    If Not rsISPB.EOF Then
        plngSequenciaLocalLiquidacao = rsISPB!SQ_ISPB
        flObterISPBLDL = rsISPB!CO_ISPB
    Else
        'Código ISPB LDL não encontrado
        lngCodigoErroNegocio = 3035
        GoTo ErrorHandler
    End If


Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterISPBLDL Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Montar mensagem de envio para a mensageria.
''
'' Executar o método clsVeiculoLegal.ObterTipoBackOffice;
''
'' Verificar o código da mensagem SPB se as 3 primeiras posições é "STR", executar
'' o método flMontarMensagemPZ, senão executar o método flMontarMensagemNZ;
''
'' Executar o método clsMensagemEnviadaRecebida.Incluir;

Public Function MontarMensagem(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                               ByRef pblnEnviarMensagem As Boolean, _
                      Optional ByVal pblnAlterarDebitoCredito As Boolean = True) As String

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim strNumeroControleIF                     As String
Dim lngStatusMensagem                       As Long
Dim lngTipoBackOffice                       As Long
Dim strRemessa                              As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim strSequencialConcilOper                 As String
Dim strValorConciliacao                     As String
Dim strDH_REGT_MESG_SPB                     As String
Dim lngProtocoloLG                          As Long
Dim intCanalVenda                           As Integer
Dim strAux                                  As String
Dim strNumCtrlCamara                        As String
Dim strDataOperacaoCambioSisbacen           As String

On Error GoTo ErrorHandler

    strSequencialConcilOper = "0"
    
    If Not xmlRemessa.documentElement.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
        lngTipoBackOffice = fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                  xmlRemessa.documentElement.selectSingleNode("//SG_SIST_ORIG").Text)
    Else
        lngTipoBackOffice = fgObterTipoBackOffice(xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
    End If
    
    If Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Or _
       Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "PAG" Then
       
        MontarMensagem = MontarMensagemPZ(xmlRemessa, lngTipoBackOffice)
        strNumeroControleIF = xmlRemessa.documentElement.selectSingleNode("NU_CTRL_IF").Text
        
        If Not xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
            
            Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(xmlRemessa.selectSingleNode("//NU_SEQU_OPER_ATIV").Text))
            If xmlConciliacao.xml <> vbNullString Then
                strSequencialConcilOper = xmlConciliacao.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text
            End If
        
            Set objConciliacao = Nothing
            Set xmlConciliacao = Nothing
        
        End If
    Else
        strNumeroControleIF = fgObterNumeroControleIF
        MontarMensagem = flMontarMensagemNZ(xmlRemessa, strNumeroControleIF, lngTipoBackOffice)
    End If

    'Gravar Mensagem
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    If xmlRemessa.documentElement.selectSingleNode("HO_AGND") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "HO_AGND", ""
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
        pblnEnviarMensagem = True
    Else
        If xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text = "" Or _
            xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text <= Format(Time, "HHMM") Then
            lngStatusMensagem = enumStatusMensagem.EnviadaBUS
            pblnEnviarMensagem = True
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemAgendada
            pblnEnviarMensagem = False
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_COMD_OPER", ""
    End If
    
    If Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Or _
       Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "PAG" Then
        If Not xmlRemessa.documentElement.selectSingleNode("NU_DOCT") Is Nothing Then
            xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text = xmlRemessa.documentElement.selectSingleNode("NU_DOCT").Text
        End If
    End If
       
    'RATS 413/0003
    If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.DespesasBMC Then
        If xmlRemessa.documentElement.selectSingleNode("CABEC") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "CABEC", MontarMensagem
        Else
            xmlRemessa.documentElement.selectSingleNode("CABEC").Text = MontarMensagem
        End If
    End If
    
    If pblnAlterarDebitoCredito = False Then
        lngStatusMensagem = enumStatusMensagem.ConcordanciaAutomatica
        fgAppendNode xmlRemessa, "MESG", "COM_REGRA", "S"
        strRemessa = MontarMensagem & xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ").Text & xmlRemessa.xml
    Else
        strRemessa = xmlRemessa.xml
    End If
    
    'Quando mensagem = BMC0002 ela so sera enviada na liberacao
    If xmlRemessa.selectSingleNode("//CO_MESG").Text = "BMC0002" Then
        lngStatusMensagem = enumStatusMensagem.ConcordanciaAutomatica
        pblnEnviarMensagem = False
        fgAppendNode xmlRemessa, "MESG", "COM_REGRA", "S"
        strRemessa = MontarMensagem & xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ").Text & xmlRemessa.xml
    End If
        
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
            
            lngStatusMensagem = enumStatusMensagem.PendenteLibAlcadaAdmArea
            pblnEnviarMensagem = False
            Call fgAppendNode(xmlRemessa, "MESG", "DT_SIST", xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text)
            strRemessa = xmlRemessa.xml
            strRemessa = Replace$(strRemessa, "<MESG>", "<SISMSG>")
            strRemessa = Replace$(strRemessa, "</MESG>", "</SISMSG>")
            
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then

            lngStatusMensagem = enumStatusMensagem.PendenteLibAlcadaAdmArea
            pblnEnviarMensagem = False
            Call fgAppendNode(xmlRemessa, "MESG", "DT_SIST", xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV").Text)
            strRemessa = xmlRemessa.xml
            strRemessa = Replace$(strRemessa, "<MESG>", "<SISMSG>")
            strRemessa = Replace$(strRemessa, "</MESG>", "</SISMSG>")
        
        'CCR
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.EmissaoOperacaoCCR Then
        
            strNumCtrlCamara = xmlRemessa.selectSingleNode("//CodReemb").Text
        
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.NegociacaoOperacaoCCR Then
        
            strNumCtrlCamara = xmlRemessa.selectSingleNode("//CodReemb").Text
            
           
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.DevolucaoRecolhimentoEstornoReembolsoCCR Then
        
            strNumCtrlCamara = xmlRemessa.selectSingleNode("//CodReemb").Text
            
        End If
    
    End If
    
    lngProtocoloLG = 0
    If Not xmlRemessa.documentElement.selectSingleNode("NU_PRTC_OPER_LG") Is Nothing Then
        lngProtocoloLG = Val(xmlRemessa.documentElement.selectSingleNode("NU_PRTC_OPER_LG").Text)
    End If
    
    intCanalVenda = 0
    If Not xmlRemessa.documentElement.selectSingleNode("TP_CNAL_VEND") Is Nothing Then
        intCanalVenda = Val(xmlRemessa.documentElement.selectSingleNode("TP_CNAL_VEND").Text)
    End If
    
    'quando mensagem = CAM0043, então grava conteudo da tag DT_EVEN_CAMB no campo A8.TB_MESG_RECB_ENVI_SPB.DT_OPER_CAMB_SISBACEN
    strDataOperacaoCambioSisbacen = vbNullString
    If xmlRemessa.selectSingleNode("//CO_MESG").Text = "CAM0043" Then
        If Not xmlRemessa.documentElement.selectSingleNode("DT_EVEN_CAMB") Is Nothing Then
            strDataOperacaoCambioSisbacen = fgDtXML_To_Oracle(xmlRemessa.documentElement.selectSingleNode("DT_EVEN_CAMB").Text)
        End If
    End If
    
    Call objMensagemEnviadaRecebida.Incluir(strRemessa, _
                                            strNumeroControleIF, _
                                            xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            strSequencialConcilOper, _
                                            xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                            lngTipoBackOffice, _
                                            xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                            xmlRemessa.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                            "", _
                                            lngStatusMensagem, _
                                            enumIndicadorSimNao.Nao, _
                                            xmlRemessa.documentElement.selectSingleNode("HO_AGND").Text, _
                                            xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                            xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                            xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                            1, _
                                            strDH_REGT_MESG_SPB, _
                                            lngProtocoloLG, _
                                            "", _
                                            strNumCtrlCamara, _
                                            intCanalVenda, _
                                            strDataOperacaoCambioSisbacen)
                                       
    Set objMensagemEnviadaRecebida = Nothing
    
    strValorConciliacao = fgObterValorMensagemPelaTag(xmlRemessa)
    
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            strSequencialConcilOper = objConciliacao.flObterSequenciaConciliacaoMensagem(strNumeroControleIF, _
                                                                                         strDH_REGT_MESG_SPB, _
                                                                                         1)
            Set objConciliacao = Nothing
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            strSequencialConcilOper = objConciliacao.flObterSequenciaConciliacaoMensagem(strNumeroControleIF, strDH_REGT_MESG_SPB, 1)
            Set objConciliacao = Nothing
           'Pagamento de Despesas
        
        'CCR
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.EmissaoOperacaoCCR Then
        
            strAux = xmlRemessa.selectSingleNode("//TpComerc").Text & _
                     xmlRemessa.selectSingleNode("//TpInstntoCCR").Text
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                        strDH_REGT_MESG_SPB, _
                                        xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                        0, 0, _
                                        strAux, _
                                        0, 0, _
                                        xmlRemessa.selectSingleNode("//VlrEms").Text, _
                                        1, , _
                                        enumTipoDebitoCredito.Debito, _
                                        xmlRemessa.selectSingleNode("//DtExprc").Text, _
                                        0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                        xmlRemessa.selectSingleNode("//CNPJBaseEntRespons").Text)
            Set objConciliacao = Nothing
            strValorConciliacao = ""
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.NegociacaoOperacaoCCR Then
        
            strAux = xmlRemessa.selectSingleNode("//TpComerc").Text & _
                     xmlRemessa.selectSingleNode("//TpInstntoCCR").Text
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                        strDH_REGT_MESG_SPB, _
                                        xmlRemessa.selectSingleNode("//DtNegc").Text, _
                                        0, 0, _
                                        strAux, _
                                        0, 0, _
                                        xmlRemessa.selectSingleNode("//VlrNegc").Text, _
                                        1, , _
                                        enumTipoDebitoCredito.Debito, _
                                        "", _
                                        0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                        xmlRemessa.selectSingleNode("//CNPJBaseEntRespons").Text)
            Set objConciliacao = Nothing
            strValorConciliacao = ""
        ElseIf xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemBUS.DevolucaoRecolhimentoEstornoReembolsoCCR Then
        
            strAux = " " & _
                     " " & _
                     xmlRemessa.selectSingleNode("//TpInstntoCCR").Text
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            Call objConciliacao.Incluir(Trim(strNumeroControleIF), _
                                        strDH_REGT_MESG_SPB, _
                                        xmlRemessa.selectSingleNode("//DT_OPER_ATIV").Text, _
                                        0, 0, _
                                        strAux, _
                                        0, 0, _
                                        xmlRemessa.selectSingleNode("//VlrDevRecolht_EstReemb").Text, _
                                        1, , _
                                        enumTipoDebitoCredito.Debito, _
                                        "", _
                                        0, "", 0, "", 0, 0, "", "", 0, "", "", 0, 0, "", 0, 0, 0, "", "", "", "", "", , _
                                        xmlRemessa.selectSingleNode("//CNPJBaseEntRespons").Text)
            Set objConciliacao = Nothing
            strValorConciliacao = ""
        End If
       
    End If
    
'    strValorConciliacao = fgObterValorMensagemPelaTag(xmlRemessa)
   
    If strValorConciliacao <> vbNullString And xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text <> "SEL1611" Then
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        
        If Not xmlRemessa.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
            Call objConciliacao.Incluir(strNumeroControleIF, _
                                        strDH_REGT_MESG_SPB, _
                                        xmlRemessa.documentElement.selectSingleNode("DT_SIST").Text, _
                                        0, 0, "", 0, 0, _
                                        strValorConciliacao)
        
            If Val(strSequencialConcilOper) <> 0 Then
                Call objConciliacao.IncluirConciliacaoOperacao( _
                                            strSequencialConcilOper, _
                                            xmlRemessa.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                            strNumeroControleIF, _
                                            strDH_REGT_MESG_SPB, _
                                            0, _
                                            1, _
                                            True)
            End If
        End If
                               
        Set objConciliacao = Nothing
    End If
    
    If pblnAlterarDebitoCredito Then
        If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
            'Transformar o Tipo de Débito/Crédito para Mensageria
            
            'RATS 399 - Ocor 01
            For Each xmlNode In xmlRemessa.selectNodes("//IN_OPER_DEBT_CRED")
                xmlNode.Text = IIf(CLng(xmlNode.Text) = enumTipoDebitoCredito.Debito, "D", "C")
            Next
        End If
    End If

    'Alteração de envio de mensagem conforme planilha de mensageria cetip (ManualdeMensageriaCTPv9.xls)
    If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
        Call MontarMensagemCETIP(xmlRemessa)
    End If
    
    'RATS 413/003
    If xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text <> enumTipoOperacaoLQS.DespesasBMC Then
        MontarMensagem = MontarMensagem & xmlRemessa.xml
    End If

    If pblnAlterarDebitoCredito Then
        If xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
            'Transformar o Tipo de Débito/Crédito para Mensageria
            'RATS 399 - Ocor 01
            For Each xmlNode In xmlRemessa.selectNodes("//IN_OPER_DEBT_CRED")
                xmlNode.Text = IIf(xmlNode.Text = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
            Next
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Monta uma mensagem para interface com o BUS, a partir dos dados de uma mensagem no padrão SPB

Public Function MontarMensagemPorMensagemSPB(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                             ByRef pblnEnviarMensagem As Boolean, _
                                    Optional ByVal pstrNumControleIF As String = vbNullString, _
                                    Optional ByVal plngNuOperAtiv As Long, _
                                    Optional ByVal pblnLiquidacaoPorLote As Boolean = False, _
                                    Optional ByVal plngTipoBackOfficeLote As Long) As String

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlConciliacao                          As MSXML2.DOMDocument40

Dim lngStatusMensagem                       As Long
Dim lngTipoBackOffice                       As Long
Dim strNumeroControleIFOld                  As String
Dim strNumeroControleIF                     As String
Dim strDataInclusao                         As String
Dim vntProtocoloAlcada                      As Variant
Dim strValorConciliacao                     As String
Dim blnIncluirConciliacao                   As Boolean
Dim blnGravaMensagem                        As Boolean
Dim strCodigoPartCamara                     As String
Dim strTipoInfoLDL                          As String
Dim strDataLiquidacao                       As String
Dim strNumeroControleCamara                 As String
Dim strCodigoPraca                          As String
Dim strNumeroDocumento                      As String
Dim strChACAM                               As String
Dim lngNrSequOperAtiv                       As Long
Dim strCodigoVeicLegal                      As String
Dim strSiglaSistema                         As String

Dim strNuCtrlIf_Original                    As String

    On Error GoTo ErrorHandler

    strNuCtrlIf_Original = vbNullString
    If Not xmlMensagem.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
        strNuCtrlIf_Original = xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text
    End If
    
    If Trim$(pstrNumControleIF) = vbNullString Then
        strNumeroControleIF = fgObterNumeroControleIF
    Else
        strNumeroControleIF = pstrNumControleIF
    End If

    If Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Then
        
        xmlMensagem.documentElement.selectSingleNode("NU_CTRL_IF").Text = strNumeroControleIF
        If pblnLiquidacaoPorLote Then
            MontarMensagemPorMensagemSPB = flMontarMensagemNZParaMensagem(xmlMensagem, _
                                                                          strNumeroControleIF, _
                                                                          plngTipoBackOfficeLote, _
                                                                          strNumeroControleIFOld)
        Else
            MontarMensagemPorMensagemSPB = flMontarMensagemNZParaMensagem(xmlMensagem, _
                                                                          strNumeroControleIF, _
                                                                          fgObterTipoBackOfficeUsuario, _
                                                                          strNumeroControleIFOld)
        End If
        
    Else

        MontarMensagemPorMensagemSPB = flMontarMensagemNZParaMensagem(xmlMensagem, _
                                                                      strNumeroControleIF, _
                                                                      CLng("0" & xmlMensagem.documentElement.selectSingleNode("TP_BKOF").Text), _
                                                                      strNumeroControleIFOld, _
                                                                      plngNuOperAtiv)
    End If

    'Gravar Mensagem
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    If xmlMensagem.documentElement.selectSingleNode("HO_AGND") Is Nothing Then
        fgAppendNode xmlMensagem, "SISMSG", "HO_AGND", ""
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
        pblnEnviarMensagem = True
    Else
        If xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text = "" Or _
            xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text <= Format(Time, "HHMM") Then
            lngStatusMensagem = enumStatusMensagem.EnviadaBUS
            pblnEnviarMensagem = True
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemAgendada
            pblnEnviarMensagem = False
        End If
    End If

    blnGravaMensagem = True
    blnIncluirConciliacao = False
    strDataInclusao = vbNullString
    vntProtocoloAlcada = 0
    
    If Not xmlMensagem.documentElement.selectSingleNode("NU_PRTC_MESG_LG") Is Nothing Then
    
        If Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Or _
           Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "LTR" Or _
           xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = "BMC0012" Then
            lngStatusMensagem = enumStatusMensagem.PendenteLibAlcadaAdmArea
        Else
            lngStatusMensagem = enumStatusMensagem.PendenteLibAlcadaAdmGeral
        End If
        
        vntProtocoloAlcada = xmlMensagem.documentElement.selectSingleNode("NU_PRTC_MESG_LG").Text
        
        If Not xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
            If xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusMensagem.PendenteLibAlcadaAdmArea Or _
               xmlMensagem.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusMensagem.PendenteLibAlcadaAdmGeral Then
                pblnEnviarMensagem = True
                blnIncluirConciliacao = False
                blnGravaMensagem = False
                strDataInclusao = xmlMensagem.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text
            Else
                vntProtocoloAlcada = xmlMensagem.documentElement.selectSingleNode("NU_PRTC_MESG_LG").Text
                pblnEnviarMensagem = False
                blnIncluirConciliacao = True
            End If
        Else
            vntProtocoloAlcada = xmlMensagem.documentElement.selectSingleNode("NU_PRTC_MESG_LG").Text
            pblnEnviarMensagem = False
            blnIncluirConciliacao = True
        End If
    
    End If

    If Not xmlMensagem.documentElement.selectSingleNode("CO_PARP_CAMR") Is Nothing Then
        strCodigoPartCamara = xmlMensagem.documentElement.selectSingleNode("CO_PARP_CAMR").Text
    End If
    
    If Not xmlMensagem.documentElement.selectSingleNode("TP_INFO_LDL") Is Nothing Then
        strTipoInfoLDL = xmlMensagem.documentElement.selectSingleNode("TP_INFO_LDL").Text
    End If
    
    If Not xmlMensagem.documentElement.selectSingleNode("NU_CTRL_CAMR") Is Nothing Then
        strNumeroControleCamara = xmlMensagem.documentElement.selectSingleNode("NU_CTRL_CAMR").Text
    End If
    
    If Not xmlMensagem.documentElement.selectSingleNode("CO_PRAC") Is Nothing Then
        strCodigoPraca = xmlMensagem.documentElement.selectSingleNode("CO_PRAC").Text
    End If
    
    If Val(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text) = 140 Then
        blnGravaMensagem = False
    End If
    
    strNumeroDocumento = vbNullString
    If Not xmlMensagem.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        strNumeroDocumento = xmlMensagem.documentElement.selectSingleNode("NU_COMD_OPER").Text
    End If
    
    If Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Or _
       Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "PAG" Then
        If Not xmlMensagem.selectSingleNode("//NU_DOCT") Is Nothing Then
            strNumeroDocumento = xmlMensagem.selectSingleNode("//NU_DOCT").Text
        End If
    End If
 
    If xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = "CAM0002" Then
        
        Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao(, strNuCtrlIf_Original))
        Set objConciliacao = Nothing
        
        If xmlConciliacao.xml <> vbNullString Then
        
            lngNrSequOperAtiv = Val(xmlConciliacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
            strChACAM = xmlMensagem.selectSingleNode("//ChACAM").Text
            
        End If
    ElseIf xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = "CAM0054" Then
        lngNrSequOperAtiv = plngNuOperAtiv
        strChACAM = xmlMensagem.selectSingleNode("//ChACAM").Text
    Else
        If xmlMensagem.documentElement.selectSingleNode("CORRETORAS") Is Nothing Then
            lngNrSequOperAtiv = Val(xmlMensagem.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
        Else
            lngNrSequOperAtiv = 0
        End If
        strChACAM = vbNullString
    End If

    If xmlMensagem.documentElement.selectSingleNode("CORRETORAS") Is Nothing Then
        If blnGravaMensagem Then
            objMensagemEnviadaRecebida.Incluir xmlMensagem.xml, _
                                               strNumeroControleIF, _
                                               lngNrSequOperAtiv, _
                                               xmlMensagem.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("TP_BKOF").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text, _
                                               strNumeroDocumento, _
                                               "", _
                                               lngStatusMensagem, _
                                               enumIndicadorSimNao.Nao, _
                                               xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                               xmlMensagem.documentElement.selectSingleNode("SG_SIST").Text, _
                                               1, _
                                               strDataInclusao, _
                                               vntProtocoloAlcada, _
                                               strCodigoPartCamara, _
                                               strNumeroControleCamara
        End If
    
    Else
        strCodigoVeicLegal = ""
        strSiglaSistema = "A8"
        If Not xmlMensagem.documentElement.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val("0" & xmlMensagem.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
                If Not xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then
                    strCodigoVeicLegal = xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text
                End If
                If Not xmlMensagem.documentElement.selectSingleNode("//SG_SIST_ORIG") Is Nothing Then
                    strSiglaSistema = xmlMensagem.documentElement.selectSingleNode("//SG_SIST_ORIG").Text
                End If
            End If
        End If
        objMensagemEnviadaRecebida.Incluir xmlMensagem.xml, _
                                           strNumeroControleIF, _
                                           0, _
                                           0, _
                                           xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, _
                                           xmlMensagem.documentElement.selectSingleNode("TP_BKOF").Text, _
                                           xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text, _
                                           strNumeroDocumento, _
                                           "", _
                                           lngStatusMensagem, _
                                           enumIndicadorSimNao.Nao, _
                                           xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text, _
                                           xmlMensagem.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                           strCodigoVeicLegal, _
                                           strSiglaSistema, _
                                           1, _
                                           strDataInclusao, _
                                           vntProtocoloAlcada, _
                                           strCodigoPartCamara
                                           
    
    End If

    strValorConciliacao = fgObterValorMensagemPelaTag(xmlMensagem)
    strDataLiquidacao = fgObterDataLiquidacaoPelaTag(xmlMensagem)
   
    If blnIncluirConciliacao Or strValorConciliacao <> vbNullString Then
        If blnGravaMensagem Then
            fgAppendNode xmlMensagem, "SISMSG", "DH_INCLUSAO", strDataInclusao
            
            Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
            objConciliacao.Incluir strNumeroControleIF, _
                                   strDataInclusao, _
                                   xmlMensagem.documentElement.selectSingleNode("DT_SIST").Text, _
                                   0, 0, "", 0, 0, _
                                   strValorConciliacao, _
                                   , _
                                   strDataLiquidacao, , , , , , _
                                   strTipoInfoLDL, , , , , , , , , , , , , , , , , _
                                   strCodigoPraca, , , , , , , _
                                   strChACAM
            Set objConciliacao = Nothing
        End If
    
    End If
    
    Set objMensagemEnviadaRecebida = Nothing

    If Mid(xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "STR" Then
        MontarMensagemPorMensagemSPB = Mid$(MontarMensagemPorMensagemSPB, 1, 20) & xmlMensagem.xml
    Else
        MontarMensagemPorMensagemSPB = MontarMensagemPorMensagemSPB & xmlMensagem.xml
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objMensagemEnviadaRecebida = Nothing
    Set objConciliacao = Nothing
    Set xmlConciliacao = Nothing
    Exit Function

ErrorHandler:
    Set objMensagemEnviadaRecebida = Nothing
    Set objConciliacao = Nothing
    Set xmlConciliacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemPorMensagemSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Montar mensagem a ser envia para a mensageria via PZ.
''
'' Montar o header do sistema A7.
''
'' Executar o método fgMontaHeaderMensagePZ;
''
'' Adionar as tags TX_HEAD_NZ, NIVE_PREF e TP_BKOF no xml de envio;
Public Function MontarMensagemPZ(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                 ByVal plngTipoBackOffice As Long) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strHeaderPZ                             As String
Dim strNumeroControleIF                     As String
Dim lngStatusMensagem                       As Long

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PZ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, " ", 9, False)
    LSet udtProtocoloAux = udtProtocolo
    
    If Not xmlRemessa.selectSingleNode("//TP_OPER") Is Nothing Then
        If xmlRemessa.selectSingleNode("//TP_OPER").Text = enumTipoOperacaoLQS.DespesasBMC Then
            Call MontarMsgEnvioBMC(xmlRemessa)
        End If
    End If
    
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        
        If xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioTEDClientes Or _
         xmlRemessa.selectSingleNode("//TP_MESG").Text = enumTipoMensagemLQS.EnvioPagDespesas Then
            strHeaderPZ = fgMontarHeaderPZTED(xmlRemessa)
        Else
            strHeaderPZ = fgMontarHeaderPZ(xmlRemessa)
        End If
        
    Else
        strHeaderPZ = fgMontarHeaderPZ(xmlRemessa)
    End If
    
    'Incluir Header NZ na mensagem para enviar BUS
    If xmlRemessa.documentElement.selectSingleNode("TX_HEAD_PZ") Is Nothing Then
        fgAppendNode xmlRemessa, xmlRemessa.documentElement.nodeName, "TX_HEAD_PZ", strHeaderPZ
    Else
        xmlRemessa.documentElement.selectSingleNode("TX_HEAD_PZ").Text = strHeaderPZ
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
        fgAppendNode xmlRemessa, xmlRemessa.documentElement.nodeName, "NIVE_PREF", ""
    Else
        xmlRemessa.documentElement.selectSingleNode("NIVE_PREF").Text = ""
    End If
    If xmlRemessa.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode xmlRemessa, xmlRemessa.documentElement.nodeName, "TP_BKOF", plngTipoBackOffice
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_BKOF").Text = plngTipoBackOffice
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//CO_EMPR") Is Nothing Then
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT").Text = fgObterISPBIF(xmlRemessa.selectSingleNode("//CO_EMPR").Text)
        If xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CPF_REME") Is Nothing Then
            Call fgAppendNode(xmlRemessa, xmlRemessa.documentElement.nodeName, "CO_CNPJ_CPF_REME", fgObterCNPJEmpresa(xmlRemessa.selectSingleNode("//CO_EMPR").Text))
        Else
            xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CPF_REME").Text = fgObterCNPJEmpresa(xmlRemessa.selectSingleNode("//CO_EMPR").Text)
        End If
    End If

    MontarMensagemPZ = udtProtocoloAux.String & xmlRemessa.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemPZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Montar mensagem a ser envia para a mensageria via NZ.
''
'' Montar o header do sistema A7.
''
'' Executar o método fgMontaHeaderMensageNZ;
''
'' Executar o método clsEmpresa.Ler;
''
'' Executar o método clsInstituicaoISPB.LerTodos;
''
'' Adionar as tags TX_HEAD_NZ, NU_CTRL_IF, NIVE_PREF, CO_ISPB_IF, TP_BKOF e
'' DT_SIST no xml de envio;
''
'' Verificar o valor da tag IN_OPER_DEBT_CRED igual a Débito;
''   1. Adicionar a tag CO_CNTA_CEDE com o valor da tag
'' CO_CNTA_CUTD_SELIC_VEIC_LEGA
''   2. Adicionar a tag CO_CNTA_CESS com o valor da tag CO_CNTA_CUTD_SELIC_CNPT
'' senão:
''   1. Adicionar a tag CO_CNTA_CEDE com o valor da tag CO_CNTA_CUTD_SELIC_CNPT
''   2. Adicionar a tag CO_CNTA_CESS com o valor da tag
'' CO_CNTA_CUTD_SELIC_VEIC_LEGA
''
'' Verificar o código da mensagem for igual a "RDC0009", adiocionar a tag
'' CO_ISPB_LDL executando o método flObterISPBLDL;

Private Function flMontarMensagemNZ(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                    ByRef pstrNumeroControleIF As String, _
                                    ByVal plngTipoBackOffice As Long) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strHeaderNZ                             As String
Dim lngStatusMensagem                       As Long
Dim strVeiculoLegal                         As String

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim objProcMensgemBMA                       As A8LQS.clsProcessoMensagemBMA

Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim strNumeroControleReqIF                  As String
Dim strNumeroControleEspc                   As String
Dim objEmpresa                              As A6A7A8.clsEmpresa
Dim xmlEmpresa                              As MSXML2.DOMDocument40

Dim intLayoutEntrada                        As Integer
    
'Bruno Oliveira - 30/08/2010 - Novo Cadastro Selic
Dim blnIndicadorPontaCliente1               As Boolean
    
    On Error GoTo ErrorHandler

    intLayoutEntrada = 0
    If Not xmlRemessa.selectSingleNode("//TP_MESG") Is Nothing Then
        intLayoutEntrada = Val(xmlRemessa.selectSingleNode("//TP_MESG").Text)
    End If

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("NZ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, " ", 9, False)
    LSet udtProtocoloAux = udtProtocolo

    strHeaderNZ = fgMontaHeaderMensageNZ(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                         xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                         pstrNumeroControleIF)

        
    'Obter o sequencial ISPB
    Set xmlEmpresa = CreateObject("MSXML2.DOMDocument.4.0")
    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    xmlEmpresa.loadXML objEmpresa.Ler(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    Set objEmpresa = Nothing

    If xmlRemessa.documentElement.selectSingleNode("DT_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("DT_MESG").Text
    End If

    'Incluir Header NZ na mensagem para enviar BUS
    If xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "TX_HEAD_NZ", strHeaderNZ
    Else
        xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ").Text = strHeaderNZ
    End If
    If xmlRemessa.documentElement.selectSingleNode("//NU_CTRL_IF") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_CTRL_IF", pstrNumeroControleIF
    Else
        xmlRemessa.documentElement.selectSingleNode("//NU_CTRL_IF").Text = pstrNumeroControleIF
    End If
    If xmlRemessa.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NIVE_PREF", ""
    Else
        xmlRemessa.documentElement.selectSingleNode("NIVE_PREF").Text = ""
    End If
    If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ISPB_IF", fgObterISPBIF(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text))
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF").Text = fgObterISPBIF(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text))
    End If
    If xmlRemessa.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "TP_BKOF", plngTipoBackOffice
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_BKOF").Text = plngTipoBackOffice
    End If
    If xmlRemessa.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "DT_SIST", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    Else
        xmlRemessa.documentElement.selectSingleNode("DT_SIST").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    End If
    If xmlRemessa.documentElement.selectSingleNode("DH_SIST") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "DH_SIST", fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    Else
        xmlRemessa.documentElement.selectSingleNode("DH_SIST").Text = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    End If

    strNumeroControleReqIF = fgObterNumeroControleIF
       
    If xmlRemessa.documentElement.selectSingleNode("NU_CTRL_REQU_IF") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "NU_CTRL_REQU_IF", strNumeroControleReqIF
    Else
        xmlRemessa.documentElement.selectSingleNode("NU_CTRL_REQU_IF").Text = strNumeroControleReqIF
    End If

    'KIDA 22/10/2010
    'RATS 973
    If Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT") Is Nothing Then
        If Len(Trim(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT").Text)) > 14 Then
            xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT").Text = Right(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT").Text, 14)
        End If
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_CNPT") Is Nothing Then
        If Len(Trim(xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_CNPT").Text)) > 14 Then
            xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_CNPT").Text = Right(xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_CNPT").Text, 14)
        End If
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_PART") Is Nothing Then
        If Len(Trim(xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_PART").Text)) > 14 Then
            xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_PART").Text = Right(xmlRemessa.documentElement.selectSingleNode("//CD_CNPJ_CPF_PART").Text, 14)
        End If
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT_CETIP21") Is Nothing Then
        If Len(Trim(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT_CETIP21").Text)) > 14 Then
            xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT_CETIP21").Text = Right(xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CNPT_CETIP21").Text, 14)
        End If
    End If
    
    
    'Identificar conta cedente e cessionária
    Select Case CLng(xmlRemessa.documentElement.selectSingleNode("TP_OPER").Text)
        
        Case enumTipoOperacaoLQS.Vinculo
            
            If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
            End If
            If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
            End If
        
        Case enumTipoOperacaoLQS.Desvinculo
            
            If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
            End If
            
            If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
            End If

        Case enumTipoOperacaoLQS.DepositoBMA
                  
             xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text = Format(xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text, "00")
                  
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
                
        Case enumTipoOperacaoLQS.RetiradaBMA
         
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
        
        Case enumTipoOperacaoLQS.TransferenciasBMCDeposito
                  
             xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text = Format(xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text, "00")
                  
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
                
        Case enumTipoOperacaoLQS.TransferenciasBMCRetirada
         
             xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text = Format(xmlRemessa.documentElement.selectSingleNode("TP_TRAF_LDL").Text, "00")
                  
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
             End If
             
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
        
        Case enumTipoOperacaoLQS.TransferenciaBMA, _
             enumTipoOperacaoLQS.MovimentacaoEntreCamarasBMA
        
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CEDE").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU").Text
             End If
             
             If xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text
             Else
                 xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CESS").Text = xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_LOCA_LIQU_DEST").Text
             End If
        
        Case enumTipoOperacaoLQS.EspecCompromissadaCobertura, _
             enumTipoOperacaoLQS.EspecCompromissadaIntermediacao, _
             enumTipoOperacaoLQS.EspecDefinitivaCobertura, _
             enumTipoOperacaoLQS.EspecDefinitivaIntermediacao, _
             enumTipoOperacaoLQS.EspecTermoCobertura, _
             enumTipoOperacaoLQS.EspecTermoIntermediacao
                
             Set objProcMensgemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
             strNumeroControleEspc = objProcMensgemBMA.ObterNumeroControleEspecificacaoOriginal(xmlRemessa.documentElement.selectSingleNode("//NU_COMD_OPER").Text, _
                                                                                                True, _
                                                                                                xmlRemessa.documentElement.selectSingleNode("//CO_EMPR").Text, _
                                                                                                xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text)
             Set objProcMensgemBMA = Nothing
        
             If xmlRemessa.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG") Is Nothing Then
                 fgAppendNode xmlRemessa, "MESG", "NU_CTRL_MESG_SPB_ORIG", strNumeroControleEspc
             Else
                 xmlRemessa.documentElement.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text = strNumeroControleEspc
             End If
        
        Case Else
                
            If Not xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
                
                'descobre se é ponta de Cliente1
                If Not xmlRemessa.documentElement.selectSingleNode("TP_CNPT") Is Nothing Then
                    If xmlRemessa.documentElement.selectSingleNode("TP_CNPT").Text = enumTipoContraparte.Cliente1 Then 'operação com Cliente1
                        If Not xmlRemessa.selectSingleNode("//IN_PONTA_CLI1") Is Nothing Then
                            If UCase$(xmlRemessa.selectSingleNode("//IN_PONTA_CLI1").Text) = "S" Then
                                blnIndicadorPontaCliente1 = True
                            End If
                        End If
                    End If
                End If
                
                If blnIndicadorPontaCliente1 = True Then 'ponta Cliente1
                    'regra de Identificação conta cedente e cessionária válida para a ponta do Cliente1 (regra inversa em relação à do Banco/Fundo)
                    If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                        End If
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                        End If
                    ElseIf xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                        End If
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                        End If
                    End If
                Else 'ponta Banco/Fundo
                    'regra de Identificação conta cedente e cessionária válida para a ponta do Banco/Fundo
                    If xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                        End If
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                        End If
                    ElseIf xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_CNPT").Text
                        End If
                        If Not xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA") Is Nothing Then
                            fgAppendNode xmlRemessa, "MESG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("CO_CNTA_CUTD_SELIC_VEIC_LEGA").Text
                        End If
                    End If
                End If
                
            End If
    
    End Select
    

    If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_ISPB_IF_DEBT", ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT").Text = ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    End If

    'Adicionar as tags CodSISBACENIFCompra e CodSISBACENIFVenda à mensagem BMC0001
    If Trim(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text) = "BMC0001" Then
        'Debito
        If Val(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = 1 Then
            fgAppendNode xmlRemessa, "MESG", "CodSISBACENIFCompra", xmlRemessa.documentElement.selectSingleNode("CO_SISB_IF_CP").Text
            fgAppendNode xmlRemessa, "MESG", "CodSISBACENIFVenda", xmlRemessa.documentElement.selectSingleNode("CO_SISB_IF_PT").Text
        'Credito
        ElseIf Val(xmlRemessa.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text) = 2 Then
            fgAppendNode xmlRemessa, "MESG", "CodSISBACENIFCompra", xmlRemessa.documentElement.selectSingleNode("CO_SISB_IF_PT").Text
            fgAppendNode xmlRemessa, "MESG", "CodSISBACENIFVenda", xmlRemessa.documentElement.selectSingleNode("CO_SISB_IF_CP").Text
        End If
    End If
    
    
    'Adicionar a tag ISPBLDL à mensagem RDC0009 e para o grupo LDL
    If Trim$(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text) = "RDC0009" Or _
        Mid(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, 1, 3) = "LDL" Then
        If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "CO_ISPB_LDL", flObterISPBLDL(xmlEmpresa.documentElement.selectSingleNode("CO_EMPR_FUSI").Text, _
                                                                           xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text)
        Else
            xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = flObterISPBLDL(xmlEmpresa.documentElement.selectSingleNode("CO_EMPR_FUSI").Text, _
                                                                                             xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text)
        End If
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = "RDC0014" Then
        If xmlRemessa.documentElement.selectSingleNode("//NumCtrlRDC") Is Nothing Then
            If Not xmlRemessa.documentElement.selectSingleNode("NU_CTRL_RDSC_ORIG") Is Nothing Then
                fgAppendNode xmlRemessa, "MESG", "NumCtrlRDC", xmlRemessa.documentElement.selectSingleNode("NU_CTRL_RDSC_ORIG").Text
            End If
        End If
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then
        Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
        Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
        xmlVeiculoLegal.loadXML objVeiculoLegal.Ler(xmlRemessa.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, xmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text)

        If xmlVeiculoLegal.documentElement.selectSingleNode("//TP_TITL_BMA").Text <> "PNA" And _
            xmlVeiculoLegal.documentElement.selectSingleNode("//TP_TITL_BMA").Text <> "PLC" Then
            strVeiculoLegal = fgObterCNPJEmpresa(xmlRemessa.documentElement.selectSingleNode("//CO_EMPR").Text)
        Else
            strVeiculoLegal = xmlVeiculoLegal.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text
        End If

        If xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA") Is Nothing Then
            fgAppendNode xmlRemessa, "MESG", "CO_CNPJ_VEIC_LEGA", strVeiculoLegal
        Else
            xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_VEIC_LEGA").Text = strVeiculoLegal
        End If
        Set objVeiculoLegal = Nothing
        Set xmlVeiculoLegal = Nothing
    End If

    If Not xmlRemessa.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If Val(xmlRemessa.selectSingleNode("//CO_LOCA_LIQU").Text) = enumLocalLiquidacao.CETIP Then
            
            If intLayoutEntrada <> enumTipoMensagemLQS.RegistroContratoSWAPCetip21 Then
                'Carregar no xml o Identificador do Banco Liquidante para mensagem SPB
                Call flIdentBancoLiquidanteCamara(xmlRemessa)
            End If
            
            If Not xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP") Is Nothing Then
                If Not xmlRemessa.selectSingleNode("//CO_MOED") Is Nothing Then
                    If Val(xmlRemessa.selectSingleNode("//CO_MOED").Text) = 220 Then
                        If InStr(1, xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text, ",") <> 0 Then
                            
                            xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text = _
                                Left$(xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text, _
                                    InStr(1, xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text, ",")) & _
                                Mid$(xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text, _
                                    InStr(1, xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text, ",") + 1, _
                                    6)
                                    
                        End If
                    End If
                End If
                
                If Not xmlRemessa.selectSingleNode("//PE_TAXA_NEGO") Is Nothing Then
                    xmlRemessa.selectSingleNode("//PE_TAXA_NEGO").Text = xmlRemessa.selectSingleNode("//PE_TAXA_NEGO_CETIP").Text
                End If
                    
            End If
            
        End If
    End If

    flMontarMensagemNZ = udtProtocoloAux.String

    Set xmlEmpresa = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:


    Set xmlEmpresa = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMensagemNZ Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Monta uma mensagem para envio ao sistema NZ, a partir dos dados de uma mensagem no padrão SPB
Private Function flMontarMensagemNZParaMensagem(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                                ByRef pstrNumeroControleIF As String, _
                                                ByVal plngTipoBackOffice As Long, _
                                                ByRef pstrNumeroControleIFConciliado As String, _
                                       Optional ByVal plngNuOperAtiv As Long) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strHeaderNZ                             As String
Dim lngStatusMensagem                       As Long

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("NZ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, " ", 9, False)
    LSet udtProtocoloAux = udtProtocolo

    strHeaderNZ = fgMontaHeaderMensageNZ(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, _
                                         xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text, _
                                         pstrNumeroControleIF)

    'Incluir Header NZ na mensagem para enviar BUS
    If xmlRemessa.documentElement.selectSingleNode("//CodMsg") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CodMsg", xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text
    Else
        xmlRemessa.documentElement.selectSingleNode("//CodMsg").Text = xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "TX_HEAD_NZ", strHeaderNZ
    Else
        xmlRemessa.documentElement.selectSingleNode("TX_HEAD_NZ").Text = strHeaderNZ
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "NU_CTRL_IF", pstrNumeroControleIF
    Else
        pstrNumeroControleIFConciliado = xmlRemessa.documentElement.selectSingleNode("NU_CTRL_IF").Text
        xmlRemessa.documentElement.selectSingleNode("NU_CTRL_IF").Text = pstrNumeroControleIF
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "NIVE_PREF", ""
    Else
        xmlRemessa.documentElement.selectSingleNode("NIVE_PREF").Text = ""
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_ISPB_IF", fgObterISPBIF(xmlRemessa.documentElement.selectSingleNode("//CO_EMPR").Text)
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF").Text = fgObterISPBIF(xmlRemessa.documentElement.selectSingleNode("//CO_EMPR").Text)
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "TP_BKOF", plngTipoBackOffice
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_BKOF").Text = plngTipoBackOffice
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("DT_SIST") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "DT_SIST", fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    Else
        xmlRemessa.documentElement.selectSingleNode("DT_SIST").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("DH_SIST") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "DH_SIST", fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    Else
        xmlRemessa.documentElement.selectSingleNode("DH_SIST").Text = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = "SEL1023" Then
        flMontarSEL1023 xmlRemessa
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = "LTR0008" Then
        Call fgAppendNode(xmlRemessa, "SISMSG", "TP_CONF_DIVG", "C")
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = "CAM0002" Then
        fgAppendNode xmlRemessa, "SISMSG", "CodSISBACENCamr", fgObterCodSISBACENCamr
        
        If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL") Is Nothing Then
            fgAppendNode xmlRemessa, "SISMSG", "CO_ISPB_LDL", ""
        End If
        
        'Quando a CAM0002 e enviada em contingencia o CO_PRAC deve ser 5885 (SP)
        If Not xmlRemessa.documentElement.selectSingleNode("CO_PRAC") Is Nothing Then
            If xmlRemessa.documentElement.selectSingleNode("CO_PRAC").Text = "" Then
                xmlRemessa.documentElement.selectSingleNode("CO_PRAC").Text = "5885"
            End If
        End If

        'ISPBLDL
        Select Case xmlRemessa.documentElement.selectSingleNode("CO_LOCA_LIQU").Text
            Case enumLocalLiquidacao.BMA
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMA, "00000000")

            Case enumLocalLiquidacao.CETIP
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbCETIP, "00000000")

            Case enumLocalLiquidacao.BMC
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMC, "00000000")

            Case enumLocalLiquidacao.BMD
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBMD, "00000000")

            Case enumLocalLiquidacao.SELIC
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbSELIC, "00000000")

            Case enumLocalLiquidacao.CLBCAcoes, _
                 enumLocalLiquidacao.CLBCTpPriv, _
                 enumLocalLiquidacao.CLBCTPub
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbCBLC, "00000000")

            Case enumLocalLiquidacao.SSTR
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(enumISPB.IspbBACEN, "00000000")

            Case enumLocalLiquidacao.CIP
                xmlRemessa.documentElement.selectSingleNode("CO_ISPB_LDL").Text = Format(4391007, "00000000")
    
        End Select
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text = "CAM0054" Then
        'Inclui campos da operacao na mensagem
        Set objOperacao = CreateObject("A8LQS.clsOperacao")
        Set xmlOperacao = objOperacao.LerXMLOperacao(plngNuOperAtiv)
        
        If xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_BASE_IF") Is Nothing Then
            fgAppendNode xmlRemessa, "SISMSG", "CO_CNPJ_BASE_IF", xmlOperacao.documentElement.selectSingleNode("CO_CNPJ_BASE_IF").Text
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_BASE_CAMR") Is Nothing Then
            fgAppendNode xmlRemessa, "SISMSG", "CO_CNPJ_BASE_CAMR", xmlOperacao.documentElement.selectSingleNode("CO_CNPJ_BASE_CAMR").Text
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_IF") Is Nothing Then
            fgAppendNode xmlRemessa, "SISMSG", "CO_CNPJ_IF", xmlOperacao.documentElement.selectSingleNode("CO_CNPJ_IF").Text
        End If
        
        If xmlRemessa.documentElement.selectSingleNode("CO_MOED_ISO") Is Nothing Then
            fgAppendNode xmlRemessa, "SISMSG", "CO_MOED_ISO", xmlOperacao.documentElement.selectSingleNode("CO_MOED_ISO").Text
        End If
        
    End If

    If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_ISPB_IF_DEBT", ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT").Text = ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    End If

    flMontarMensagemNZParaMensagem = udtProtocoloAux.String

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing

    Exit Function
ErrorHandler:
    Set xmlOperacao = Nothing
    Set objOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMensagemNZParaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Monta uma mensagem para envio ao sistema PZ, a partir dos dados de uma mensagem no padrão SPB
Private Function flMontarMensagemPZParaMensagem(ByRef xmlRemessa As MSXML2.DOMDocument40, _
                                                ByVal plngTipoBackOffice As Long) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strHeaderPZ                             As String
Dim strNumeroControleIF                     As String
Dim lngStatusMensagem                       As Long

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("PZ", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = fgCompletaString(xmlRemessa.documentElement.selectSingleNode("CO_MESG").Text, " ", 9, False)
    LSet udtProtocoloAux = udtProtocolo

    strHeaderPZ = fgMontarHeaderPZ(xmlRemessa)

    'Incluir Header NZ na mensagem para enviar BUS
    If xmlRemessa.documentElement.selectSingleNode("TX_HEAD_PZ") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "TX_HEAD_PZ", strHeaderPZ
    Else
        xmlRemessa.documentElement.selectSingleNode("TX_HEAD_PZ").Text = strHeaderPZ
    End If

    If xmlRemessa.documentElement.selectSingleNode("NIVE_PREF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "NIVE_PREF", ""
    Else
        xmlRemessa.documentElement.selectSingleNode("NIVE_PREF").Text = ""
    End If
    If xmlRemessa.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "TP_BKOF", plngTipoBackOffice
    Else
        xmlRemessa.documentElement.selectSingleNode("TP_BKOF").Text = plngTipoBackOffice
    End If
    
    If xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_ISPB_IF_DEBT", ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_ISPB_IF_DEBT").Text = ObterISPBIF(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    End If

    If Not xmlRemessa.documentElement.selectSingleNode("//CO_CNPJ_CPF_REME") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_CNPJ_CPF_REME", fgObterCNPJEmpresa(xmlRemessa.documentElement.selectSingleNode("CO_EMPR").Text)
    Else
        xmlRemessa.documentElement.selectSingleNode("CO_CNPJ_CPF_REME").Text = fgObterCNPJEmpresa(xmlRemessa.selectSingleNode("//CO_EMPR").Text)
    End If

    flMontarMensagemPZParaMensagem = udtProtocoloAux.String & xmlRemessa.xml
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:


    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarMensagemPZParaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Um resultado vbNullString significa que a operação foi efetuada com sucesso
'Senão, retornará o XML com o erro
Public Function Salvar(ByRef xmlMensagem As MSXML2.DOMDocument40) As String

Dim strErro                                 As String
Dim objProcessaOperacao                     As A8LQS.clsProcessoOperacao

On Error GoTo ErrorHandler

    Set objProcessaOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Call objProcessaOperacao.GerenciarRecebimentoOperacao(xmlMensagem, strErro)
    Salvar = strErro
    Set objProcessaOperacao = Nothing

    Exit Function
ErrorHandler:
    
    Set objProcessaOperacao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler mensagens SPB das tabelas A8.TB_MESG_RECB_ENVI_SPB e A8.TB_SITU_PROC pelo
'' número sequencial da operação.
Public Function LerMensagensPorOperacao(ByVal vntSequenciaOperacao As Variant) As String
    
Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT   A.NU_CTRL_IF,                               " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                         " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                        " & vbNewLine & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,              " & vbNewLine & _
             "          A.CO_MESG_SPB,                              " & vbNewLine & _
             "          A.CO_TEXT_XML,                              " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                        " & vbNewLine & _
             "          B.DE_SITU_PROC,                             " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ,                        " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,                   " & vbNewLine & _
             "          A.TP_BKOF                                   " & vbNewLine & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB        A,          " & vbNewLine & _
             "          A8.TB_SITU_PROC                 B           " & vbNewLine & _
             " WHERE    A.CO_ULTI_SITU_PROC     = B.CO_SITU_PROC    " & vbNewLine
             
    If vntSequenciaOperacao < 0 Then
        strSQL = Replace$(strSQL, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
        vntSequenciaOperacao = vntSequenciaOperacao * -1
    End If
             
    strSQL = strSQL & _
             " AND     (A.NU_SEQU_OPER_ATIV           =  " & vntSequenciaOperacao & "          " & vbNewLine & _
             " OR       A.NU_SEQU_CNCL_OPER_ATIV_MESG IN (SELECT   NU_SEQU_CNCL_OPER_ATIV_MESG " & vbNewLine & _
             "                                            FROM     A8.TB_CNCL_OPER_ATIV        " & vbNewLine & _
             "                                            WHERE    NU_SEQU_OPER_ATIV =  " & vntSequenciaOperacao & "))" & vbNewLine
             
    strSQL = strSQL & _
             " ORDER BY A.DH_REGT_MESG_SPB DESC "
    
    LerMensagensPorOperacao = fgQueryXMLLerTodos("OperacaoMensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagensPorOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
'' Ler o grupo da mensagem SPB das tabelas TB_GRUPO, TB_SERVICO, TB_EVENTO e
'' TB_MENSAGEM pelo código da mensagem.
Public Function LerGrupoMensagem(ByVal pstrCodigoMensagem As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   A.CO_GRUP " & vbCrLf & _
             "   FROM   A8.TB_GRUPO     A, " & vbCrLf & _
             "          A8.TB_SERVICO   B, " & vbCrLf & _
             "          A8.TB_EVENTO    C, " & vbCrLf & _
             "          A8.TB_MENSAGEM  D  " & vbCrLf & _
             "  WHERE   A.SQ_GRUP =     B.SQ_GRUP " & vbCrLf & _
             "    AND   B.SQ_SERV =     C.SQ_SERV " & vbCrLf & _
             "    AND   C.SQ_EVEN =     D.SQ_EVEN " & vbCrLf & _
             "    AND   D.CO_MESG =     '" & pstrCodigoMensagem & "'" & vbCrLf & _
             "    AND   C.SQ_TIPO_FLUX NOT IN(5,7)"

    LerGrupoMensagem = fgQueryXMLLer("Grupo", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerGrupoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler todos os grupos de mensagens SPB utilizando as tabelas TB_GRUPO, TB_SERVICO,
'' TB_EVENTO e TB_MENSAGEM.
Public Function LerTodosGrupoMensagem(Optional ByVal pblnMensagemTratadaSLCC As Boolean = True) As String

Dim strSQL                                  As String
Dim strNot                                  As String
Dim strSegragaDados                         As String

On Error GoTo ErrorHandler
    
    
    strSQL = " SELECT   DISTINCT(A.CO_GRUP), " & vbCrLf & _
             "          A.NO_GRUP,       " & vbCrLf & _
             "          A.SQ_GRUP        " & vbCrLf & _
             "   FROM   A8.TB_GRUPO   A, " & vbCrLf & _
             "          A8.TB_SERVICO B, " & vbCrLf & _
             "          A8.TB_EVENTO  C, " & vbCrLf & _
             "          A8.TB_MENSAGEM D," & vbCrLf & _
             "          A8.TB_GRUP_USUA_MESG_SPB E " & vbCrLf & _
             "  WHERE   A.SQ_GRUP = B.SQ_GRUP      " & vbCrLf & _
             "    AND   B.SQ_SERV = C.SQ_SERV      " & vbCrLf & _
             "    AND   C.SQ_EVEN = D.SQ_EVEN      " & vbCrLf & _
             "    AND   D.CO_MESG = E.CO_MESG_SPB  " & vbCrLf & _
             "    AND   E.TP_ACES = " & enumTipoAcessoMensagemSPB.Enviar & _
             "    AND   C.SQ_TIPO_FLUX NOT IN(5,7) " & vbCrLf & _
             "    AND   D.CO_MESG <> 'CAM0005' "

    strSegragaDados = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                     False, _
                                     "E", _
                                     "E", _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False, _
                                     True)
    
    If Trim(strSegragaDados) <> vbNullString Then
        strSQL = strSQL & strSegragaDados & vbCrLf
    End If

    
    If pblnMensagemTratadaSLCC Then
        strNot = "IN"
    Else
        strNot = "NOT IN"
    End If
    
    strSQL = strSQL & _
             "    AND   D.CO_MESG " & strNot & " (SELECT   CO_MESG_SPB_REGT_OPER " & vbCrLf & _
                                                 "  FROM   A8.TB_TIPO_OPER " & vbCrLf & _
                                                 " WHERE   CO_MESG_SPB_REGT_OPER IS NOT NULL)"
   
    strSQL = strSQL & " ORDER BY NO_GRUP "
    
    LerTodosGrupoMensagem = fgQueryXMLLerTodos("Grupo", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosGrupoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

Public Function LerTodosMensagemSPBGrupoUsuario() As String

Dim strSQL                                  As String
Dim strSegregaDados                         As String

On Error GoTo ErrorHandler
    
    
    strSQL = " SELECT   DISTINCT(B.CO_MESG_SPB),   " & vbCrLf & _
             "          A.NO_MESG                  " & vbCrLf & _
             "   FROM   A8.TB_MENSAGEM A,          " & vbCrLf & _
             "          A8.TB_GRUP_USUA_MESG_SPB B " & vbCrLf & _
             "  WHERE   A.CO_MESG = B.CO_MESG_SPB  "

    strSegregaDados = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                     False, _
                                     "B", _
                                     "B", _
                                     False, _
                                     False, _
                                     False, _
                                     False, _
                                     False, _
                                     True)
    
    If Trim(strSegregaDados) <> vbNullString Then
        strSQL = strSQL & strSegregaDados & vbCrLf
    End If
   
    strSQL = strSQL & " ORDER BY B.CO_MESG_SPB "
    
    LerTodosMensagemSPBGrupoUsuario = fgQueryXMLLerTodos("MensagemSPB", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosMensagemSPBGrupoUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler todos os serviços de mensagens SPB utilizando as tabelas TB_GRUPO,
'' TB_SERVICO, TB_EVENTO, TB_MENSAGEM e A8.TB_GRUP_USUA_MESG_SPB pelo grupo da
'' mensagem do parâmetro.
Public Function LerTodosServicos(ByVal plngSqGrupo As Long, _
                        Optional ByVal pblnMensagemTratadaSLCC As Boolean = True) As String

Dim strSQL                                  As String
Dim strNot                                  As String
Dim strSegragaDados                         As String

On Error GoTo ErrorHandler
    
    
    strSQL = " SELECT   DISTINCT(B.NO_SERV), " & _
             "          B.SQ_SERV  " & _
             "   FROM   A8.TB_GRUPO   A, " & _
             "          A8.TB_SERVICO B, " & _
             "          A8.TB_EVENTO  C, " & _
             "          A8.TB_MENSAGEM D, " & _
             "          A8.TB_GRUP_USUA_MESG_SPB E " & _
             "  WHERE   A.SQ_GRUP = B.SQ_GRUP     " & _
             "    AND   A.SQ_GRUP = " & plngSqGrupo & _
             "    AND   B.SQ_SERV = C.SQ_SERV     " & _
             "    AND   C.SQ_EVEN = D.SQ_EVEN     " & _
             "    AND   D.CO_MESG = E.CO_MESG_SPB " & _
             "    AND   E.TP_ACES = " & enumTipoAcessoMensagemSPB.Enviar & vbCrLf & _
             "    AND   C.SQ_TIPO_FLUX NOT IN(5,7) " & vbCrLf & _
             "    AND   D.CO_MESG <> 'CAM0005' "

    strSegragaDados = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                     False, _
                                     "E", _
                                     "E", _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False, _
                                     True)
    
    If Trim(strSegragaDados) <> vbNullString Then
        strSQL = strSQL & strSegragaDados
    End If

    
    If pblnMensagemTratadaSLCC Then
        strNot = "IN"
    Else
        strNot = "NOT IN"
    End If
    
    strSQL = strSQL & _
             "    AND   D.CO_MESG " & strNot & " (SELECT   CO_MESG_SPB_REGT_OPER " & _
                                                 "  FROM   A8.TB_TIPO_OPER " & _
                                                 " WHERE   CO_MESG_SPB_REGT_OPER IS NOT NULL)"
    
    strSQL = strSQL & " ORDER BY NO_SERV "
    
    LerTodosServicos = fgQueryXMLLerTodos("Servico", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosServicos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler todos os eventos de mensagens SPB utilizando as tabelas TB_GRUPO,
'' TB_SERVICO, TB_EVENTO, TB_MENSAGEM e A8.TB_GRUP_USUA_MESG_SPB pelo serviço do
'' parâmetro.
Public Function LerTodosEventos(ByVal plngSqServico As Long, _
                       Optional ByVal pblnMensagemTratadaSLCC As Boolean = True) As String

Dim strSQL                                  As String
Dim strNot                                  As String
Dim strSegragaDados                         As String

On Error GoTo ErrorHandler
    
    
    strSQL = " SELECT   DISTINCT(C.SQ_EVEN),  " & _
             "          C.CO_EVEN,  " & _
             "          C.NO_EVEN   " & _
             "   FROM   A8.TB_GRUPO   A, " & _
             "          A8.TB_SERVICO B, " & _
             "          A8.TB_EVENTO  C, " & _
             "          A8.TB_MENSAGEM D, " & _
             "          A8.TB_GRUP_USUA_MESG_SPB E " & _
             "  WHERE   A.SQ_GRUP = B.SQ_GRUP     " & _
             "    AND   B.SQ_SERV = C.SQ_SERV     " & _
             "    AND   C.SQ_EVEN = D.SQ_EVEN     " & _
             "    AND   C.SQ_SERV = " & plngSqServico & _
             "    AND   D.CO_MESG = E.CO_MESG_SPB " & _
             "    AND   E.TP_ACES = " & enumTipoAcessoMensagemSPB.Enviar & vbCrLf & _
             "    AND   C.SQ_TIPO_FLUX NOT IN (5,7) " & vbCrLf & _
             "    AND   D.CO_MESG <> 'CAM0005' "
             
             
    strSegragaDados = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                     False, _
                                     "E", _
                                     "E", _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False, _
                                     True)
    
    If Trim(strSegragaDados) <> vbNullString Then
        strSQL = strSQL & strSegragaDados
    End If

             
             
    If pblnMensagemTratadaSLCC Then
        strNot = "IN"
    Else
        strNot = "NOT IN"
    End If
    
    strSQL = strSQL & _
             "    AND   D.CO_MESG " & strNot & " (SELECT   CO_MESG_SPB_REGT_OPER " & _
                                                 "  FROM   A8.TB_TIPO_OPER " & _
                                                 " WHERE   CO_MESG_SPB_REGT_OPER IS NOT NULL)"
    
    strSQL = strSQL & " ORDER BY CO_EVEN  "
    
    
    LerTodosEventos = fgQueryXMLLerTodos("Evento", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If


    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosEventos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler todas as mensagens tratadas ou não pelo SLCC conforme a sequencia de grupo
'' do parâmetro.
Public Function LerTodosMensagem(ByVal plngSequenciaGrupo As Long, _
                        Optional ByVal plngSequenciaServico As Long = 0, _
                        Optional ByVal plngSequenciaEvento As Long = 0, _
                        Optional ByVal pblnMensagemTratadaSLCC As Boolean = True, _
                        Optional ByVal pstrCdGrupoMesg As String = "") As String


Dim strSQL                                  As String
Dim strNot                                  As String
Dim strSegragaDados                         As String

On Error GoTo ErrorHandler
    
    If pblnMensagemTratadaSLCC Then
        
        strSQL = "  SELECT  A.CO_GRUP AS CO_GRUP,  " & vbCrLf & _
                 "          A.NO_GRUP AS NO_GRUP,  " & vbCrLf & _
                 "          A.SQ_GRUP AS SQ_GRUP,  " & vbCrLf & _
                 "          B.SQ_SERV AS SQ_SERV,  " & vbCrLf & _
                 "          B.NO_SERV AS NO_SERV,  " & vbCrLf & _
                 "          C.SQ_EVEN AS SQ_EVEN,  " & vbCrLf & _
                 "          C.CO_EVEN AS CO_EVEN,  " & vbCrLf & _
                 "          C.NO_EVEN AS NO_EVEN,  " & vbCrLf & _
                 "          D.CO_MESG AS CO_MESG,  " & vbCrLf & _
                 "          D.SQ_MESG AS SQ_MESG,  " & vbCrLf & _
                 "          F.TP_MESG_RECB_INTE AS TP_MESG_RECB_INTE, " & vbCrLf & _
                 "          F.TP_OPER AS TP_OPER, " & vbCrLf & _
                 "          F.NO_TIPO_OPER AS NO_TIPO_OPER "

        strSQL = strSQL & _
                 "   FROM   A8.TB_GRUPO    A,           " & vbCrLf & _
                 "          A8.TB_SERVICO  B,           " & vbCrLf & _
                 "          A8.TB_EVENTO   C,           " & vbCrLf & _
                 "          A8.TB_MENSAGEM D,           " & vbCrLf & _
                 "          A8.TB_GRUP_USUA_MESG_SPB E, " & vbCrLf & _
                 "          A8.TB_TIPO_OPER F           " & vbCrLf & _
                 "  WHERE   A.SQ_GRUP = B.SQ_GRUP       " & vbCrLf & _
                 "    AND   A.SQ_GRUP = " & plngSequenciaGrupo & vbCrLf & _
                 "    AND   B.SQ_SERV = C.SQ_SERV     " & vbCrLf & _
                 "    AND   C.SQ_EVEN = D.SQ_EVEN     " & vbCrLf & _
                 "    AND   D.CO_MESG = E.CO_MESG_SPB " & vbCrLf & _
                 "    AND   D.CO_MESG = F.CO_MESG_SPB_REGT_OPER " & vbCrLf & _
                 "    AND   E.TP_ACES = " & enumTipoAcessoMensagemSPB.Enviar & vbCrLf & _
                 "    AND   C.SQ_TIPO_FLUX NOT IN(5,7)"

    Else
    
        strSQL = "  SELECT  A.CO_GRUP,  " & _
                 "          A.NO_GRUP,  " & _
                 "          A.SQ_GRUP,  " & _
                 "          B.SQ_SERV,  " & _
                 "          B.NO_SERV,  " & _
                 "          C.SQ_EVEN,  " & _
                 "          C.CO_EVEN,  " & _
                 "          C.NO_EVEN,  " & _
                 "          D.CO_MESG,  " & _
                 "          D.SQ_MESG,  " & _
                 "          D.NO_TAG_PRIN_MESG, " & _
                 "          D.NO_MESG " & _
                 "   FROM   A8.TB_GRUPO   A, " & _
                 "          A8.TB_SERVICO B, " & _
                 "          A8.TB_EVENTO  C, " & _
                 "          A8.TB_MENSAGEM D, " & _
                 "          A8.TB_GRUP_USUA_MESG_SPB E " & _
                 "  WHERE   A.SQ_GRUP = B.SQ_GRUP     " & _
                 "    AND   A.SQ_GRUP = " & plngSequenciaGrupo & _
                 "    AND   B.SQ_SERV = C.SQ_SERV     " & _
                 "    AND   C.SQ_EVEN = D.SQ_EVEN     " & _
                 "    AND   D.CO_MESG = E.CO_MESG_SPB " & _
                 "    AND   E.TP_ACES = " & enumTipoAcessoMensagemSPB.Enviar & vbCrLf & _
                 "    AND   C.SQ_TIPO_FLUX NOT IN (5,7) " & vbCrLf & _
                 "    AND   D.CO_MESG <> 'CAM0005' "
    
    End If
    
    strSegragaDados = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                     False, _
                                     "E", _
                                     "E", _
                                     False, _
                                     False, _
                                     True, _
                                     False, _
                                     False, _
                                     True)
    
    If Trim(strSegragaDados) <> vbNullString Then
        strSQL = strSQL & strSegragaDados
    End If
    
    If plngSequenciaServico <> 0 Then
        strSQL = strSQL & " AND   C.SQ_SERV = " & plngSequenciaServico
    End If
             
    If plngSequenciaEvento <> 0 Then
        strSQL = strSQL & " AND   D.SQ_EVEN = " & plngSequenciaEvento
    End If
    
    If pblnMensagemTratadaSLCC Then
        strNot = "IN"
    Else
        strNot = "NOT IN"
    End If
    
    strSQL = strSQL & _
             "    AND   D.CO_MESG " & strNot & " (SELECT   CO_MESG_SPB_REGT_OPER " & _
                                                 "  FROM   A8.TB_TIPO_OPER " & _
                                                 " WHERE   CO_MESG_SPB_REGT_OPER IS NOT NULL " & _
                                                 "   AND   TP_MESG_RECB_INTE IS NOT NULL)"
                 
'-------------------
    If pblnMensagemTratadaSLCC Then
    
        If pstrCdGrupoMesg = "SEL" Then
            strSQL = strSQL & _
                    "Union " & vbCrLf & _
                    "   SELECT  'SEL' AS CO_GRUP, " & vbCrLf & _
                    "   'Sistema Especial de Liquidação e Custódia' AS NO_GRUP, " & vbCrLf & _
                    "   12 AS SQ_GRUP, " & vbCrLf & _
                    "   65 AS SQ_SERV, " & vbCrLf & _
                    "   'Controle de Operações SELIC' AS NO_SERV, " & vbCrLf & _
                    "   0 AS SQ_EVEN, " & vbCrLf & _
                    "   'SEL' AS CO_EVEN, " & vbCrLf & _
                    "   A.NO_TIPO_OPER as NO_EVEN, " & vbCrLf & _
                    "   'SEL' as CO_MESG, " & vbCrLf & _
                    "   0 as SQ_MESG, " & vbCrLf & _
                    "   A.TP_MESG_RECB_INTE, " & vbCrLf & _
                    "   A.TP_OPER, " & vbCrLf & _
                    "   A.NO_TIPO_OPER " & vbCrLf & _
                    " FROM    A8.TB_TIPO_OPER  A " & vbCrLf & _
                    " Where a.CO_MESG_SPB_REGT_OPER Is Null " & vbCrLf & _
                    " and     A.TP_OPER IN (6,12,13,14,15) " & vbCrLf
        
        ElseIf pstrCdGrupoMesg = "CTP" Then
            strSQL = strSQL & _
                    "Union " & vbCrLf & _
                    "   SELECT  'CTP' AS CO_GRUP, " & vbCrLf & _
                    "   'Câmara CTP' AS NO_GRUP, " & vbCrLf & _
                    "   5 AS SQ_GRUP, " & vbCrLf & _
                    "   30 AS SQ_SERV, " & vbCrLf & _
                    "   'Registro de Operações - Títulos e Valores Mobiliários' AS NO_SERV, " & vbCrLf & _
                    "   0 AS SQ_EVEN, " & vbCrLf & _
                    "   'CTP' AS CO_EVEN, " & vbCrLf & _
                    "   A.NO_TIPO_OPER as NO_EVEN, " & vbCrLf & _
                    "   'CTP' as CO_MESG, " & vbCrLf & _
                    "   0 as SQ_MESG, " & vbCrLf & _
                    "   A.TP_MESG_RECB_INTE, " & vbCrLf & _
                    "   A.TP_OPER, " & vbCrLf & _
                    "   A.NO_TIPO_OPER " & vbCrLf & _
                    " FROM    A8.TB_TIPO_OPER  A " & vbCrLf & _
                    " Where a.CO_MESG_SPB_REGT_OPER Is Null " & vbCrLf & _
                    " and     A.TP_OPER IN (67,68,105,102) " & vbCrLf
        
        ElseIf pstrCdGrupoMesg = "BMC" Then
            strSQL = strSQL & _
                    "Union " & vbCrLf & _
                    "   SELECT  'BMC' AS CO_GRUP, " & vbCrLf & _
                    "   'Câmara de Registro, Compensação e Liquidação de Operações de Câmbio da BM&F' AS NO_GRUP, " & vbCrLf & _
                    "   1 AS SQ_GRUP, " & vbCrLf & _
                    "   63 AS SQ_SERV, " & vbCrLf & _
                    "   'Operações Interbancárias de Câmbio BMC' AS NO_SERV, " & vbCrLf & _
                    "   0 AS SQ_EVEN, " & vbCrLf & _
                    "   'BMC' AS CO_EVEN, " & vbCrLf & _
                    "   A.NO_TIPO_OPER as NO_EVEN, " & vbCrLf & _
                    "   'BMC' as CO_MESG, " & vbCrLf & _
                    "   0 as SQ_MESG, " & vbCrLf & _
                    "   A.TP_MESG_RECB_INTE, " & vbCrLf & _
                    "   A.TP_OPER, " & vbCrLf & _
                    "   A.NO_TIPO_OPER " & vbCrLf & _
                    " FROM    A8.TB_TIPO_OPER  A " & vbCrLf & _
                    " Where a.CO_MESG_SPB_REGT_OPER Is Null " & vbCrLf & _
                    " and     A.TP_OPER IN (" & enumTipoOperacaoLQS.RegistroOperacoesBMC & "," & _
                                                enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao & "," & _
                                                enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica & ") " & vbCrLf
        
        End If
    
    End If

'-------------------
                 
    If pblnMensagemTratadaSLCC Then
        strSQL = strSQL & " ORDER BY TP_OPER "
    End If
                 
    LerTodosMensagem = fgQueryXMLLerTodos("Mensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
       objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Ler tipos de mensagem conforme filtro especificado.
Private Function flLerTodosTiposMensagens(ByVal pxmlFiltro As MSXML2.DOMDocument40) As String

Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereFiltro                          As String
Dim strValores                              As String

    On Error GoTo ErrorHandler
    
    strValores = vbNullString
    For Each objXMLNode In pxmlFiltro.selectNodes("//Grupo_CodigoMensagem/*")
        strValores = strValores & "'" & objXMLNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND A.CO_MESG IN (" & strValores & ")" & vbNewLine
    End If
    
    strValores = vbNullString
    For Each objXMLNode In pxmlFiltro.selectNodes("//Grupo_GrupoMensagem/*")
        strValores = strValores & "'" & objXMLNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND SUBSTR(A.CO_MESG, 1, 3) IN (" & strValores & ")" & vbNewLine
        strWhereFiltro = strWhereFiltro & " AND SUBSTR(A.CO_MESG, 8, 1) = ' '                  " & vbNewLine
    End If
    
    strSQL = " SELECT  A.CO_MESG,       " & vbNewLine & _
             "         A.NO_MESG        " & vbNewLine & _
             " FROM    A8.TB_MENSAGEM A " & vbNewLine

    If strWhereFiltro <> vbNullString Then
        strWhereFiltro = " WHERE" & strWhereFiltro
        strWhereFiltro = Replace$(strWhereFiltro, "WHERE AND", "WHERE")
        strSQL = strSQL & strWhereFiltro
    End If
    
    strSQL = strSQL & " ORDER BY A.CO_MESG "
                 
    flLerTodosTiposMensagens = fgQueryXMLLerTodos("Mensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
       objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLerTodosTiposMensagens Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Retorna o conteúdo de uma mensagem, a partir de seu código
Public Function LerMensagem(ByVal pstrCodigoMensagem As String, _
                   Optional ByVal plngTipoMensagem As Long = 0, _
                   Optional ByVal plngCodigoEmpresa As Long, _
                   Optional ByVal plngTipoOperacao As Long = 0) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim objMensagemSPB                          As A8LQS.clsMensagemSPB

On Error GoTo ErrorHandler

    'Se a mensagem for tratada pelo SLCC
    'Obter mensagem no A7
    If VerificaMensagemTratadaSLCC(Trim$(pstrCodigoMensagem)) Or _
       Trim$(pstrCodigoMensagem) = "BMC0012" Then

        Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
        
        If plngTipoMensagem <> 0 Then
            LerMensagem = objMensagemSLCC.LerMensagem(plngTipoMensagem, plngTipoOperacao)
        Else
            LerMensagem = objMensagemSLCC.LerMensagem(flObterTipoMensagem(pstrCodigoMensagem), plngTipoOperacao)
        End If
        
        Set objMensagemSLCC = Nothing

    Else
        'Se mensagem não tratada pelo SLCC
        'Obter Mensagem PK
        Set objMensagemSPB = CreateObject("A8LQS.clsMensagemSPB")
        LerMensagem = objMensagemSPB.LerMensagem(pstrCodigoMensagem)
        Set objMensagemSPB = Nothing
    End If

    Exit Function
ErrorHandler:

    Set objMensagemSLCC = Nothing
    Set objMensagemSPB = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Selecionar na tabela A8.TB_TIPO_OPER a quantidade de mensagens iguais ao
'' parâmetro;
''
'' Caso a quantidade de mensagens for maior que 0 retornar Verdadeiro senão
'' retornar falso.
Public Function VerificaMensagemTratadaSLCC(ByVal pstrCodigoMensagem As String) As Boolean

Dim strSQL                                  As String
Dim rs                                      As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT COUNT(CO_MESG_SPB_REGT_OPER) " & _
             "   FROM A8.TB_TIPO_OPER " & _
             " WHERE  CO_MESG_SPB_REGT_OPER = '" & pstrCodigoMensagem & "'"

    Set rs = fgQuerySQL(strSQL)

    If rs.fields(0) = 0 Then
        VerificaMensagemTratadaSLCC = False
        
        If pstrCodigoMensagem = "SEL" Or pstrCodigoMensagem = "CTP" Or pstrCodigoMensagem = "BMC" Then
            VerificaMensagemTratadaSLCC = True
        Else
            VerificaMensagemTratadaSLCC = False
        End If
        
    Else
        If pstrCodigoMensagem = "BMC0012" Then
            VerificaMensagemTratadaSLCC = False
        Else
            VerificaMensagemTratadaSLCC = True
        End If
    
    End If

    rs.Close

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificaMensagemTratadaSLCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Selecionar a coluna TP_MESG_RECB_INTE da tabela A8.TB_TIPO_OPER conforme o
'' parâmetro recebido.
Private Function flObterTipoMensagem(ByVal pstrCodigoMensagem As String) As Long

Dim strSQL                                  As String
Dim rs                                      As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT TP_MESG_RECB_INTE " & _
             "   FROM A8.TB_TIPO_OPER " & _
             "  WHERE CO_MESG_SPB_REGT_OPER = '" & Trim(pstrCodigoMensagem) & "'"
    
    Set rs = fgQuerySQL(strSQL)
    
    flObterTipoMensagem = rs.fields(0)
        
    rs.Close

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Formata uma mensagem no estilo HTML
Public Function ObterMensagemHTML(ByVal plngSequencial As Long) As String

Dim objMensagemSLCC                         As A8LQS.clsMensagemSLCC
Dim objMensagemSPB                          As A8LQS.clsMensagemSPB
Dim xmlDOMDocument                          As MSXML2.DOMDocument40

Dim strXML                                  As String
Dim strHTML                                 As String
Dim strCodigoMensagem                       As String
Dim lngCodigoTipoMensagem                   As Long
Dim blnXMLValido                            As Boolean
Dim blnOperacao                             As Boolean

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloNZ                          As udtProtocoloNZ

    On Error GoTo ErrorHandler

    'Obtem mensagem XML a partir do código identificador informado
    If plngSequencial <> 0 Then
        strXML = fgSelectVarchar4000(plngSequencial)
    End If
        
    Set xmlDOMDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlDOMDocument.loadXML(strXML) Then
        blnXMLValido = True
    Else
        'Se não conseguir ler a STRING retornada pela função << fgSelectVarchar4000 >>,
        'tenta ler removendo os HEADERs (SLCC e NZ)
        If xmlDOMDocument.loadXML(Mid(strXML, Len(udtProtocolo) + Len(udtProtocoloNZ) + 1)) Then
            blnXMLValido = True
        Else
            blnXMLValido = False
        End If
    End If
    
    If blnXMLValido Then
        
        flFiltrarConteudoExibicao xmlDOMDocument
        
        If Not xmlDOMDocument.selectSingleNode("//TP_MESG") Is Nothing Then
            lngCodigoTipoMensagem = xmlDOMDocument.selectSingleNode("//TP_MESG").Text
        End If
        
        Set xmlDOMDesc = CreateObject("MSXML2.DOMDocument.4.0")
        
        'Verifica se a mensagem é proveniente do SPB
        If xmlDOMDocument.documentElement.nodeName = "SISMSG" And _
           (lngCodigoTipoMensagem <> enumTipoMensagemBUS.EnvioTEDClientes) And _
           (lngCodigoTipoMensagem <> enumTipoMensagemBUS.EnvioPagDespesas) Then
            
            Set objMensagemSPB = CreateObject("A8LQS.clsMensagemSPB")
            
            strCodigoMensagem = xmlDOMDocument.selectSingleNode("//CodMsg").Text
            
            If UCase(Mid(strCodigoMensagem, 8, 1)) = "E" Then
                'Caso seja mensagem de erro, pega o nome da mensagem original para descrição
                strCodigoMensagem = Left$(strCodigoMensagem, 7)
            End If
            
            
            If Left$(strCodigoMensagem, 3) = "STR" Or _
               Left$(strCodigoMensagem, 3) = "PAG" Then
            
                ObterMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                                    "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                    "Mensagem Interna LQS para o BUS" & _
                                    "</Font></TD></TR>" & _
                                    flMontaHTMLLayoutInterno(strXML) & _
                                    "</Table></Center></Body></Html>"
            
            Else
            
                If strCodigoMensagem = "BMC0012" Then
                    If Not xmlDOMDocument.selectSingleNode("//CO_SITU_ACTE") Is Nothing Then
                        Call fgAppendNode(xmlDOMDocument, "BMCReqConfcOpInterbancCamIntmddCortr", "SitActe", xmlDOMDocument.selectSingleNode("//CO_SITU_ACTE").Text)
                    End If
                End If
                
                If xmlDOMDesc.loadXML(objMensagemSPB.LerMensagem(strCodigoMensagem)) Then
                    
                    strHTML = "<HTML><Body><Center>"
                    
                    If xmlDOMDocument.selectSingleNode("SISMSG").childNodes(0).attributes.length = 0 Then
                        
                        strHTML = strHTML & "<Table border = 0>" & _
                                        "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                        xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem/NO_MESG").Text & _
                                        "</Font></TD></TR>" & _
                                        flMontaHTML(xmlDOMDocument.selectSingleNode("SISMSG").childNodes(0).xml)
                    Else
                    
                        strHTML = strHTML & "<Table border = 0>" & _
                                        "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=Red Size=2 Face=Verdana>" & _
                                        "*" & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem/NO_MESG").Text & _
                                        "</Font></TD></TR>" & _
                                        flMontaHTML(xmlDOMDocument.selectSingleNode("SISMSG").childNodes(0).xml)
                    
                    
                    End If
                        
                        
                    If Not xmlDOMDocument.selectSingleNode("//NumCtrlIF") Is Nothing Then
                        'Exibe o No. Controle se a mensagem não conter o mesmo
                        strHTML = strHTML & "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                          "No. Controle IF" & _
                                          "</TD><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                                          Replace(xmlDOMDocument.selectSingleNode("//NumCtrlIF").Text, " ", "&nbsp") & _
                                          "</Font></TD></TR>"
                    End If
                        
                    If xmlDOMDocument.selectSingleNode("SISMSG").childNodes(0).attributes.length > 0 Then
                        strCodErroBACEN = xmlDOMDocument.selectSingleNode("SISMSG").childNodes(0).attributes.Item(0).Text
                    End If
                        
                        
                    Dim arrCodErroBACEN() As String
                    Dim liAux             As Integer
                    Dim strDescricaoErro  As String
                    Dim objLogErro                              As A6A7A8CA.clsLogErro
                    
                    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")
                    arrCodErroBACEN = Split(strCodErroBACEN, "|")
                        
                    If strCodErroBACEN <> "" Then
                        
                      For liAux = LBound(arrCodErroBACEN) To UBound(arrCodErroBACEN)
                        
                         If UCase(Mid(xmlDOMDocument.selectSingleNode("//CodMsg").Text, 8, 1)) = "E" Then
                            strDescricaoErro = fgObterDescricaoErro(arrCodErroBACEN(liAux))
                         Else
                            strDescricaoErro = Mid(objLogErro.ObterDescricaoErroNZPKPJ(arrCodErroBACEN(liAux)), 4)
                         End If
                        
                        strHTML = strHTML & "<TR><TD BGColor=""#BBBBBB""><Font Color=red Size=2 Face=Verdana>" & _
                                          "* Código do Erro" & _
                                          "</TD><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana Color=Red>" & _
                                          arrCodErroBACEN(liAux) & " - " & strDescricaoErro & _
                                          "</Font></TD></TR>"
                                          
                      Next
                    
                    End If
                    
                    strHTML = strHTML & "</Table>"
        
                    strHTML = strHTML & "</Body></HTML>"
                    
                    ObterMensagemHTML = strHTML
                Else
                    ObterMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                                        "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                        "Dicionário Inexistente para a Mensagem" & _
                                        "</Font></TD></TR>" & _
                                        "<TR><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                                        strXML & _
                                        "</Font></TD></TR>" & _
                                        "</Table></Center></Body></Html>"
                End If
            End If
        '...então, é uma mensagem interna BUS -> LQS
        Else
            lngCodigoTipoMensagem = xmlDOMDocument.selectSingleNode("//TP_MESG").Text
            
            Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
                     
            strCodErroBACEN = ""
                     
            If xmlDOMDesc.loadXML(objMensagemSLCC.LerMensagem(lngCodigoTipoMensagem)) Then
            
                blnOperacao = IIf(Not xmlDOMDocument.selectSingleNode("//MESG") Is Nothing, True, False)
                
                strHTML = "<HTML><Body><Center>"
                strHTML = strHTML & "<Table border = 0>" & _
                                "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem/NO_TIPO_MESG").Text & _
                                "</Font></TD></TR>" & _
                                flMontaHTML(xmlDOMDocument.documentElement.xml, , blnOperacao)
                strHTML = strHTML & "</Table>"
                strHTML = strHTML & "</Body></HTML>"
                
                ObterMensagemHTML = strHTML
            Else
                ObterMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                                    "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                    "Mensagem Interna LQS para o BUS" & _
                                    "</Font></TD></TR>" & _
                                    flMontaHTMLLayoutInterno(strXML) & _
                                    "</Table></Center></Body></Html>"
            End If
            
            Set objMensagemSLCC = Nothing
        End If
    Else
        
        If Trim$(strXML) <> vbNullString Then
            'o XML da mensagens não pode ser carregado
            'a mensagem esta errada (Copy Fixo ou XML Mal Formado)
            ObterMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                                "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                "XML Mal Formado" & _
                                "</Font></TD></TR>" & _
                                "<TR><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                                strXML & _
                                "</Font></TD></TR>" & _
                                "</Table></Center></Body></Html>"
        Else
            'o XML da mensagens não pode ser carregado
            'a mensagem esta errada (Copy Fixo ou XML Mal Formado)
            ObterMensagemHTML = "<Html><Body><Center><Table border = 0>" & _
                                "<TR><TD BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                                "Conciliação sem mensagem" & _
                                "</Font></TD></TR>" & _
                                "<TR><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                                strXML & _
                                "</Font></TD></TR>" & _
                                "</Table></Center></Body></Html>"
        End If
    End If
    
    Set xmlDOMDesc = Nothing
    Set xmlDOMDocument = Nothing
    Set objMensagemSPB = Nothing
    
    Exit Function
    
ErrorHandler:
    Set xmlDOMDesc = Nothing
    Set xmlDOMDocument = Nothing
    Set objMensagemSLCC = Nothing
    Set objMensagemSPB = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagemHTML Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function



'Filtra o conteúdo de uma mensagem de acordo com o perfil do usuário
Private Sub flFiltrarConteudoExibicao(ByRef pxmlMensagem As DOMDocument40)

Dim xmlNode                                 As IXMLDOMNode
Dim strXpath                                As String
Dim strSQL                                  As String
Dim strCodMsg                               As String
Dim rs                                      As Recordset
Dim strTagFiltro                            As String
Dim lngTipoFiltro                           As Long
Dim strValorBusca                           As String

On Error GoTo ErrorHandler
    
    If Not pxmlMensagem.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = Trim$(pxmlMensagem.selectSingleNode("//CodMsg").Text)
        If strCodMsg = "LTR0007" Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    'Verificar se a mensagem está configurada para a aplicação do filtro
    strSQL = "SELECT NO_TAG_FILT, TP_FILT " & vbCrLf & _
               "FROM A8.TB_CTRL_RETN_MESG " & vbCrLf & _
              "WHERE CO_MESG_RECB = '" & strCodMsg & "' " & vbCrLf & _
                "AND NU_NIVE_REPE = 1 " & vbCrLf & _
                "AND NOT NO_TAG_FILT IS NULL "
              
    Set rs = fgQuerySQL(strSQL)
    
    If rs.EOF Then
        Set rs = Nothing
        Exit Sub
    End If
    
    strTagFiltro = Trim$(rs!NO_TAG_FILT)
    lngTipoFiltro = CLng(rs!TP_FILT)
    
    Set rs = Nothing
    
    'Obtem valores possiveis
    Select Case lngTipoFiltro
        Case enumTIpoFiltroExibicaoMsg.FiltroConta
            'strColuna = "SUBSTR(LPAD(CO_CNTA_CUTD_PADR_SELIC,9,0),1,4) as CO_CNTA_CUTD_PADR_SELIC"
            strSQL = "SELECT SUBSTR(LPAD(CO_CNTA_CUTD_PADR_SELIC,9,0),1,4) as CO_CNTA_CUTD_PADR_SELIC " & _
                      " FROM A8.TB_VEIC_LEGA " & _
                     " WHERE TP_BKOF = " & fgObterTipoBackOfficeUsuario() & _
                       " AND CO_CNTA_CUTD_PADR_SELIC IS NOT NULL"
        Case enumTIpoFiltroExibicaoMsg.FiltroCPNJ
            'strColuna = "CO_CNPJ_VEIC_LEGA"
            strSQL = "SELECT CO_CNPJ_VEIC_LEGA " & _
                      " FROM A8.TB_VEIC_LEGA " & _
                     " WHERE TP_BKOF = " & fgObterTipoBackOfficeUsuario() & _
                       " AND CO_CNPJ_VEIC_LEGA IS NOT NULL"
            
        Case Else
            
            Exit Sub
    End Select
    
    Set rs = fgQuerySQL(strSQL)
    
    strXpath = rs.GetString(, , , "||")
    
    Set rs = Nothing
    
    If strXpath = vbNullString Then Exit Sub
    
    For Each xmlNode In pxmlMensagem.selectNodes("//" & strTagFiltro)
    
        Select Case lngTipoFiltro
            Case enumTIpoFiltroExibicaoMsg.FiltroConta
                strValorBusca = Left$(Right$(String$(9, "0") & xmlNode.Text, 9), 4)
            Case enumTIpoFiltroExibicaoMsg.FiltroCPNJ
                strValorBusca = xmlNode.Text
        End Select
            
        If InStr(1, strXpath, strValorBusca) = 0 Then
            'Remover tag
'            xmlNode.parentNode.parentNode.removeChild xmlNode.parentNode
        End If
        
    Next
    
    Exit Sub
ErrorHandler:
    Set rs = Nothing
    Set xmlNode = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flFiltrarConteudoExibicao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub


'' Obter a cor do nível que deverá ser exibido no HTML.
Private Function flCorNivel(ByVal pintNivel As Integer) As String
    Select Case pintNivel
        Case 0
            flCorNivel = "#EEEEEE"
        Case 1
            flCorNivel = "#CEDAEA"
        Case 2
            flCorNivel = "#E6EFD8"
        Case Else
            flCorNivel = "#F3ECD4"
    End Select

End Function

'' Montar um HTML sobre um xml que contém a operação ou mensagem.

Private Function flMontaHTML(ByVal pstrDOMDoc As String, _
                    Optional ByVal pintNivel As Integer, _
                    Optional ByRef pblnOperacao As Boolean) As String

Dim xmlDOMDoc                               As MSXML2.DOMDocument40
Dim intX                                    As Integer
Dim intNivel                                As Integer
Dim strFontColor                            As String
Dim strAux                                  As String
Dim strData                                 As String
Dim strNovaLinhaHTML                        As String
Dim blnExibirTag                            As Boolean


On Error GoTo ErrorHandler

    Set xmlDOMDoc = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlDOMDoc.loadXML(pstrDOMDoc) Then
        
        For intX = 0 To xmlDOMDoc.childNodes(0).childNodes.length - 1
            
            If Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 6) = "Repet_" Or _
               Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 4) = "REPE" Then
                
                flMontaHTML = flMontaHTML & flMontaHTML(xmlDOMDoc.childNodes(0).childNodes(intX).xml, pintNivel, pblnOperacao)
            
            ElseIf Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 1) <> "/" Then
               
                If Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 6) = "Grupo_" Or _
                   Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 5) = "GRUP_" Then
                    intNivel = pintNivel + 1
                Else
                    intNivel = pintNivel
                End If
                
                If UCase(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName) <> "ERROGEN" Then
                    strFontColor = "Black"
                Else
                    strFontColor = "Red"
                    strCodErroBACEN = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                End If
                
                strNovaLinhaHTML = vbCr & _
                            "<TR BGColor=" & flCorNivel(intNivel) & ">" & vbCr & _
                            "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"

                blnExibirTag = True
                
                'Verifica se a mensagem recebida é proveniente do BUS -> LQS...
                'If Not xmlDOMDoc.selectSingleNode("//MESG") Is Nothing Then
                If pblnOperacao Then
                    If xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_ATRB_MESG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/NO_TRAP_ATRB") Is Nothing Then
                        
                        Select Case UCase(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName)
                            
                            Case "QT_ATIV_MERC"
                                
                                If Not xmlDOMDoc.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
                                    If xmlDOMDoc.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
                                        blnExibirTag = False
                                    End If
                                End If
                            
                            Case "TP_MESG", "DT_SIST", "SG_SIST_ORIG", "SG_SIST_DEST", "CO_EMPR", "IN_ENTR_MANU", _
                                 "VA_LANC_CC", "TP_OPER", "NU_SEQU_OPER_ATIV", "CO_ULTI_SITU_PROC", "TP_ACAO_OPER_ATIV_EXEC", _
                                 "NU_COMD_ACAO_EXEC", "DH_ULTI_ATLZ", "CO_MESG", "TP_MESG_RETN_INTE", "CO_SITU_MOVI_CAIX_SUB_RESE", _
                                 "DE_TIPO_LIQU", "TX_HEAD_NZ", "NU_CTRL_IF", "NIVE_PREF", "CO_ISPB_IF", _
                                 "TP_BKOF", "CO_CNTA_CEDE", "CO_CNTA_CESS", "NU_PRTC_OPER_LG", "IN_SIST_SITU_CNTG", _
                                 "IN_ENVI_PREV_SIST_PJ", "IN_ENVI_RELZ_SIST_PJ", "IN_ENVI_PREV_SIST_A6", "IN_ENVI_RELZ_SIST_A6", _
                                 "TP_LIQU_OPER_ATIV", "CABEC", "CO_CNPJ_VEIC_LEGA", "NU_CTRL_REQU_IF", "DH_SIST", "NU_ATIV_MERC", "NU_PROC", "NU_CTRL_MESG_SPB_ORIG", "CO_ISPB_LDL"

                                'Não faz nada, pois estas são apenas TAGs de apoio para o BUS

                                blnExibirTag = False

                            Case Else
                                flMontaHTML = flMontaHTML & strNovaLinhaHTML & _
                                    String(intNivel * 3, Chr(1)) & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName
                        End Select
                    Else
                        flMontaHTML = flMontaHTML & strNovaLinhaHTML & _
                            String(intNivel * 3, Chr(1)) & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_ATRB_MESG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/NO_TRAP_ATRB").Text
                    End If

                '...se não, é uma mensagem SPB
                Else
                    If xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_TAG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/DE_TAG") Is Nothing Then
                        flMontaHTML = flMontaHTML & strNovaLinhaHTML & _
                            String(intNivel * 3, Chr(1)) & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName
                    Else

                        If xmlDOMDoc.childNodes(0).childNodes(intX).attributes.length > 0 Then
                            
                            strNovaLinhaHTML = Replace(strNovaLinhaHTML, "Color=Black", "Color=red")
                            
                            flMontaHTML = flMontaHTML & strNovaLinhaHTML & _
                                            String(intNivel * 3, Chr(1)) & "* " & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_TAG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/DE_TAG").Text
                            
                            strCodErroBACEN = strCodErroBACEN & IIf(strCodErroBACEN = "", "", "|") & xmlDOMDoc.childNodes(0).childNodes(intX).attributes.Item(0).Text
                        Else
                        
                            flMontaHTML = flMontaHTML & strNovaLinhaHTML & _
                                            String(intNivel * 3, Chr(1)) & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_TAG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/DE_TAG").Text

                        End If
                    End If
                End If
                
                If blnExibirTag Then
                    flMontaHTML = flMontaHTML & "</Font></TD>" & vbCr & _
                                "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"
 
                    If Not xmlDOMDoc.childNodes(0).childNodes(intX).childNodes(0) Is Nothing Then
                        
                        If xmlDOMDoc.childNodes(0).childNodes(intX).childNodes(0).hasChildNodes Then
                            
                            flMontaHTML = flMontaHTML & flMontaHTML(xmlDOMDoc.childNodes(0).childNodes(intX).xml, pintNivel + 1, pblnOperacao)
                            
                        Else
                            If UCase$(Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "DTH" Then
                                strData = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                                'YYYYMMDDHHMMSS
                                If IsDate(Mid$(strData, 7, 2) & "/" & Mid$(strData, 5, 2) & "/" & Left$(strData, 4)) Then
                                    flMontaHTML = flMontaHTML & "|DH|" & strData & "|"
                                Else
                                    flMontaHTML = flMontaHTML & strData
                                End If
                            
                            ElseIf UCase$(Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 2)) = "DT" Then
                                
                                strData = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                                'YYYYMMDD
                                If IsDate(Mid$(strData, 7, 2) & "/" & Mid$(strData, 5, 2) & "/" & Left$(strData, 4)) Then
                                    flMontaHTML = flMontaHTML & "|DT|" & strData & "|" 'FormatDateTime(fgDtHrStr_To_DateTime(Left$(strData, 8) & "000000"), vbShortDate)
                                Else
                                    flMontaHTML = flMontaHTML & strData
                                End If
                            
                            ElseIf UCase$(Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "VLR" Or _
                                   UCase$(Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "SLD" Then
                                
                                'Utiliza a função << fgVlrXml_To_Decimal >> para formatar corretamente as
                                'casas decimais a serem exibidas, de acordo com a formatação do SERVIDOR
                                
                                flMontaHTML = flMontaHTML & fgVlrXml_To_Decimal(xmlDOMDoc.childNodes(0).childNodes(intX).Text)
                            
                            ElseIf UCase$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName) = "QTDCTP" Then
                                
                                If InStr(1, xmlDOMDoc.childNodes(0).childNodes(intX).Text, ",", vbBinaryCompare) > 0 Then
                                    flMontaHTML = flMontaHTML & fgVlrXml_To_InterfaceDecimais(xmlDOMDoc.childNodes(0).childNodes(intX).Text, 8)
                                Else
                                    flMontaHTML = flMontaHTML & fgVlrXml_To_Interface(xmlDOMDoc.childNodes(0).childNodes(intX).Text, False)
                                End If
                                
                            ElseIf UCase$(Left$(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "QTD" Then
                                'Utiliza a função << fgVlrXml_To_Interface >> para formatar Ex.: 1000, como 1.000
                                
                                flMontaHTML = flMontaHTML & fgVlrXml_To_Interface(xmlDOMDoc.childNodes(0).childNodes(intX).Text, False)
                                
                                
                            Else
                                strAux = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                                strAux = Replace(strAux, "&", "&amp;")
                                strAux = Replace(strAux, "<", "&lt;")
                                strAux = Replace(strAux, ">", "&gt;")
                                flMontaHTML = flMontaHTML & strAux
                            End If
 
                            flMontaHTML = flMontaHTML & "</Font></TD>" & vbCr & _
                                                "</TR>"
                        End If
                    End If
                End If
            End If
        Next
    End If

    flMontaHTML = Replace(flMontaHTML, Chr(1), "&nbsp")

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaHTML Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Compõe NET de mensagens para ser utilizado nas conciliações
Public Function ObterNetMensagemConciliacao(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim intTipoLiquidacao                       As enumTipoLiquidacao
Dim intTipoContraparte                      As enumTipoContraparte
Dim intPagamentoRecebimento                 As enumTipoDebitoCredito
Dim strPreviaDefinitiva                     As String
Dim blnNegacao                              As Boolean
                                                                
    On Error GoTo ErrorHandler
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusMensagem/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por Ação da Mensagem
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_AcaoMensagem/*")
        If InStr(1, objDomNode.Text, "CONDIÇÃO") = 0 Then
            strValores = strValores & "'" & objDomNode.Text & "', "
        Else
            strValores = strValores & objDomNode.Text
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If InStr(1, strValores, "CONDIÇÃO") = 0 Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.TP_ACAO_MESG_SPB_EXEC  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    
    'Filtro por Canal Operação Interna
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalOperacaoInterna/*")
        If InStr(1, objDomNode.Text, "CONDIÇÃO") = 0 Then
            strValores = strValores & "'" & objDomNode.Text & "', "
        Else
            strValores = strValores & objDomNode.Text
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If InStr(1, strValores, "CONDIÇÃO") = 0 Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      C.CO_CNAL_OPER_INTE  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro Tipo de Liquidação
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        intTipoLiquidacao = Val(objDomNode.Text)
    Next
    
    If intTipoLiquidacao = enumTipoLiquidacao.Bruta Then
       
        strWhereFiltro = strWhereFiltro & " AND     (A.NU_COMD_OPER IS NOT NULL " & vbNewLine & _
                                          " OR       A.NU_COMD_OPER != '')      " & vbNewLine
                                          
    ElseIf intTipoLiquidacao = enumTipoLiquidacao.Bilateral Then
           
        strWhereFiltro = strWhereFiltro & " AND     (A.NU_COMD_OPER IS     NULL " & vbNewLine & _
                                          " OR       A.NU_COMD_OPER  = '')      " & vbNewLine
                                          
    End If
    
    'Filtro Tipo Transferência / Natureza Movimento / Tipo Informação
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoContraparte/*")
        intTipoContraparte = Val(objDomNode.Text)
    Next
    
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NaturezaMovimento/*")
        intPagamentoRecebimento = Val(objDomNode.Text)
    Next
    
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoInformacao/*")
        strPreviaDefinitiva = objDomNode.Text
    Next
    
    'Mercado
    If intTipoContraparte = enumTipoContraparte.Externo Then
        
        'Pagamento
        If intPagamentoRecebimento = enumTipoDebitoCredito.Debito Then
            strWhereFiltro = strWhereFiltro & " AND      A.CO_MESG_SPB = 'LTR0001'" & vbNewLine
            
        'Recebimento
        Else
            strWhereFiltro = strWhereFiltro & " AND      A.CO_MESG_SPB IN ('LTR0003R2', 'LTR0005R2')" & vbNewLine
            
        End If
        
        strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_CNTR_REPE  = 1        "
        
    'Book Transfer
    Else
        
        strWhereFiltro = strWhereFiltro & " AND      A.CO_MESG_SPB        = 'LTR0007'" & vbNewLine
        
        'Pagamento
        If intPagamentoRecebimento = enumTipoDebitoCredito.Debito Then
            strWhereFiltro = strWhereFiltro & " AND      C.IN_OPER_DEBT_CRED = " & enumTipoDebitoCredito.Debito & vbNewLine
            
        'Recebimento
        Else
            strWhereFiltro = strWhereFiltro & " AND      C.IN_OPER_DEBT_CRED = " & enumTipoDebitoCredito.Credito & vbNewLine
            
        End If
        
        strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_CNTR_REPE  = 1        "
        
    End If
    
    'Filtro Data
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & "AND      C.DT_OPER between " & objDomNode.selectSingleNode("DataIni").Text & " AND " & _
                                                                          objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU     IN (" & strValores & ")"
    End If
    '---------------------------------------------------------------------------------------------------------
    
    strWhereSegregacao = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", False, "A", "B", False, False, True, True, True, False)
    
    strSQL = " SELECT   A.CO_EMPR,                                                         " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          A.SG_SIST,                                                         " & vbNewLine & _
             "          A.NU_COMD_OPER,                                                    " & vbNewLine & _
             "          A.NU_CTRL_IF,                                                      " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                                                " & vbNewLine & _
             "          A.NU_SEQU_CNTR_REPE,                                               " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                               " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                    " & vbNewLine & _
             "          A.CO_MESG_SPB,                                                     " & vbNewLine & _
             "          A.TP_BKOF,                                                         " & vbNewLine & _
             "          DECODE(A.IN_CONF_MESG_LTR, " & enumIndicadorSimNao.Sim & ",        " & vbNewLine & _
             "          'Sim', 'Não') AS IN_CONF_MESG_LTR,                                 " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                               " & vbNewLine & _
             "          A.NU_CTRL_CAMR,                                                    " & vbNewLine & _
             "          A.CO_PARP_CAMR,                                                    " & vbNewLine

    strSQL = strSQL & _
             "          B.NO_VEIC_LEGA,                                                    " & vbNewLine & _
             "          B.CO_CNPJ_VEIC_LEGA,                                               " & vbNewLine & _
             "          B.ID_PART_CAMR_CETIP,                                              " & vbNewLine & _
             "          C.IN_OPER_DEBT_CRED,                                               " & vbNewLine & _
             "          C.VA_FINC,                                                         " & vbNewLine & _
             "          C.NU_ATIV_MERC,                                                    " & vbNewLine & _
             "          C.QT_ATIV_MERC,                                                    " & vbNewLine & _
             "          C.CO_ISPB_BANC_LIQU_CNPT,                                          " & vbNewLine & _
             "          C.CO_CNPJ_CNPT,                                                    " & vbNewLine & _
             "          C.CO_CNPT_CAMR                                                     " & vbNewLine
             
    strSQL = strSQL & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB        A,                                 " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                 " & vbNewLine & _
             "          A8.TB_MESG_RECB_SPB_CNCL        C                                  " & vbNewLine & _
             " WHERE    A.CO_VEIC_LEGA                = B.CO_VEIC_LEGA                     " & vbNewLine & _
             " AND      A.SG_SIST                     = B.SG_SIST                          " & vbNewLine & _
             " AND      A.NU_CTRL_IF                  = C.NU_CTRL_IF                       " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB            = C.DH_REGT_MESG_SPB                 " & vbNewLine & _
             " AND      A.NU_SEQU_CNTR_REPE           = C.NU_SEQU_CNTR_REPE                " & vbNewLine

    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados
    If Trim(strWhereSegregacao) <> "" Then
       strSQL = strSQL & vbNewLine & strWhereSegregacao
    End If
    
    strSQL = strSQL & vbNewLine & _
             " ORDER BY B.NO_VEIC_LEGA, A.SG_SIST, A.NU_COMD_OPER"
    
    ObterNetMensagemConciliacao = fgQueryXMLLerTodos("NetMensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetMensagemConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Compõe NET de Mensagens e Operações para utilização na Liquidação Física (Títulos)
Public Function ObterNetLiquidacaoFisica(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltroMesgOper                  As String
Dim strWhereFiltroMesg                      As String
Dim strWhereFiltroOper                      As String
Dim strWhereFiltroOperMesg                  As String
Dim strValores                              As String
Dim intDebitoCredito                        As enumTipoDebitoCredito
                                                                
    On Error GoTo ErrorHandler
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    'Filtro por Status da Operacao
    strValores = vbNullString
    strWhereFiltroOper = vbNullString
    strWhereFiltroMesg = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
                
                strWhereFiltroOper = strWhereFiltroOper & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")" & vbCrLf
        strWhereFiltroOperMesg = strWhereFiltroOperMesg & " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Status da Mensagem
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusMensagem/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltroMesg = strWhereFiltroMesg & " AND         A.CO_ULTI_SITU_PROC  IN (" & strValores & ") " & vbCrLf
        strWhereFiltroMesgOper = strWhereFiltroMesgOper & " AND       f.CO_ULTI_SITU_PROC  IN (" & strValores & ") " & vbCrLf
    End If
    
    'Filtro por Tipo de Operacao
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltroOper = strWhereFiltroOper & " AND      A.TP_OPER  IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltroOper = strWhereFiltroOper & " AND      V.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbNewLine
        strWhereFiltroMesg = strWhereFiltroMesg & " AND      V.CO_GRUP_VEIC_LEGA      IN (" & strValores & ")" & vbNewLine
    End If
    
    'Filtro Tipo de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltroOper = strWhereFiltroOper & " AND      A.TP_LIQU_OPER_ATIV     IN (" & strValores & ")" & vbNewLine
    End If
    
    'Data da operação/mensagem
    For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_DataOperacaoMensagem")
        'Data da operacao
        strWhereFiltroOper = strWhereFiltroOper & "AND A.DT_OPER_ATIV between " & fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataIni").Text) & " AND" & vbNewLine & _
                         fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataFim").Text) & vbCrLf
        
        'Data da mensagem vs. operacao
        'strWhereFiltroOper = strWhereFiltroOper & "AND TRUNC(f.DH_REGT_MESG_SPB (+)) = TRUNC(a.dt_oper_ativ) " & vbNewLine
        
        'Data da mensagem
        strWhereFiltroMesg = strWhereFiltroMesg & "AND A.DH_REGT_MESG_SPB between " & fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataIni").Text) & " AND " & vbNewLine & _
                         fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataFim").Text) & vbCrLf
        
        'Data da mensagem
        strWhereFiltroMesgOper = strWhereFiltroMesgOper & "AND f.DH_REGT_MESG_SPB between " & fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataIni").Text) & " AND " & vbNewLine & _
                         fgDtXML_To_Oracle(objDomNode.selectSingleNode("DataFim").Text) & vbCrLf
    Next objDomNode
    
    
    '---------------------------------------------------------------------------------------------------------
    strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "V", False, False, True, True, True, False)

    strSQL = "SELECT /*+ ORDERED */ 'MESG' as ORIGEM, c.de_bkof, " & vbCrLf & _
            "       B.nu_ativ_merc, " & vbCrLf & _
            "       to_char(B.nu_comd_oper_ORIG) as nu_comd_oper, " & vbCrLf & _
            "       B.co_cnta_cutd_selic_veic_lega, " & vbCrLf & _
            "       B.dt_venc as dt_venc_ativ, " & vbCrLf & _
            "       D.DE_SITU_PROC, " & vbCrLf & _
            "       DECODE(b.in_oper_debt_cred, 1, b.qt_ativ_merc * -1, b.qt_ativ_merc) AS qt_camara, " & vbCrLf & _
            "       0 AS qt_operacao, " & vbCrLf & _
            "       A.NU_CTRL_IF, " & vbCrLf & _
            "       A.DH_REGT_MESG_SPB, " & vbCrLf & _
            "       A.NU_SEQU_CNTR_REPE, " & vbCrLf & _
            "       A.DH_ULTI_ATLZ, " & vbCrLf & _
            "       A.CO_MESG_SPB " & vbCrLf & _
            "FROM   a8.tb_mesg_recb_ENVI_spb a, " & vbCrLf & _
            "       a8.tb_mesg_recb_spb_cncl b, " & vbCrLf & _
            "       a8.tb_SITU_PROC D, " & vbCrLf & _
            "       a8.tb_tipo_bkof c, " & vbCrLf & _
            "       a8.tb_veic_lega v " & vbCrLf
    strSQL = strSQL & _
            "WHERE  B.nu_sequ_cntr_repe = A.nu_sequ_cntr_repe " & vbCrLf & _
            "       AND A.nu_sequ_cntr_repe > 1 " & vbCrLf & _
            "       AND B.nu_ctrl_if =  A.nu_ctrl_if " & vbCrLf & _
            "       AND B.dh_regt_mesg_spb = A.dh_regt_mesg_spb " & vbCrLf & _
            "       AND D.CO_SITU_PROC =  A.CO_ULTI_SITU_PROC " & vbCrLf & _
            "       AND c.tp_bkof =   v.tp_bkof " & vbCrLf & _
            "       AND a.co_veic_lega =   v.co_veic_lega " & vbCrLf & _
            "       AND V.SG_SIST = A.SG_SIST " & vbCrLf & _
            "       AND A.CO_MESG_SPB = 'BMA0015' " & vbCrLf & _
            strWhereFiltroMesg & vbCrLf & _
            strWhereSegregacao & vbCrLf & _
            "       AND not exists  " & vbCrLf
    strSQL = strSQL & _
            "               (SELECT              'x'                               " & vbCrLf & _
            "               FROM                 a8.tb_oper_ativ A1,               " & vbCrLf & _
            "                                    a8.tb_VEIC_LEGA V1                " & vbCrLf & _
            "               WHERE                a1.nu_ativ_merc                 = B.nu_ativ_merc                 " & vbCrLf & _
            "               AND                  a1.co_cnta_cutd_selic_veic_lega = B.co_cnta_cutd_selic_veic_lega " & vbCrLf & _
            "               AND                  a1.dt_venc_ativ                 = B.dt_venc                      " & vbCrLf & _
            "               AND                  a1.co_veic_lega                 = V1.co_veic_lega                " & vbCrLf & _
            "               AND                  a1.sg_sist                      = V1.sg_sist                     " & vbCrLf & _
            "               " & Replace(Replace(strWhereFiltroOper, "A.", "A1."), "V.", "V1.") & _
            "               ) " & vbCrLf
            
    '"AND a1.nu_comd_oper = B.nu_comd_oper_ORIG " & vbCrLf & _

    strSQL = strSQL & " UNION ALL " & vbCrLf

    strSQL = strSQL & _
            "SELECT /*+ ORDERED */'OPER' as ORIGEM, c.de_bkof,                             " & vbCrLf & _
            "       a.nu_ativ_merc,                                                        " & vbCrLf & _
            "       a.nu_comd_oper,                                                        " & vbCrLf & _
            "       a.co_cnta_cutd_selic_veic_lega,                                        " & vbCrLf & _
            "       a.dt_venc_ativ,                                                        " & vbCrLf & _
            "       D.DE_SITU_PROC,                                                        " & vbCrLf & _
            "       DECODE(f.in_oper_debt_cred, 1, f.qt_ativ_merc * -1, f.qt_ativ_merc) AS qt_camara, " & _
            "       NVL(SUM(DECODE(a.in_oper_debt_cred, 2, a.qt_ativ_merc, 0)), 0) - NVL(SUM(DECODE(a.in_oper_debt_cred, 1, a.qt_ativ_merc, 0)), 0) AS qt_operacao, " & vbCrLf & _
            "       F.NU_CTRL_IF,                                                          " & vbCrLf & _
            "       F.dH_REGT_MESG_SPB,                                                    " & vbCrLf & _
            "       F.NU_SEQU_CNTR_REPE,                                                   " & vbCrLf & _
            "       F.DH_ULTI_ATLZ,                                                        " & vbCrLf & _
            "       F.CO_MESG_SPB                                                          " & vbCrLf
    strSQL = strSQL & _
            "FROM   a8.tb_oper_ativ a,                                                     " & vbCrLf & _
            "       a8.tb_veic_lega v,                                                      " & vbCrLf & _
            "       a8.tb_SITU_PROC D,                                                     " & vbCrLf & _
            "       a8.tb_tipo_bkof c,                                                     " & vbCrLf & _
            "       (SELECT     f.*                             ,                          " & vbCrLf & _
            "                   c.nu_ativ_merc                  ,                          " & vbCrLf & _
            "                   c.co_cnta_cutd_selic_veic_lega  ,                          " & vbCrLf & _
            "                   c.dt_venc                       ,                          " & vbCrLf & _
            "                   c.in_oper_debt_cred             ,                          " & vbCrLf & _
            "                   c.qt_ativ_merc                                             " & vbCrLf & _
            "        FROM       a8.tb_mesg_recb_envi_spb f,                                " & vbCrLf & _
            "                   a8.tb_mesg_recb_spb_cncl c                                 " & vbCrLf & _
            "        WHERE      f.co_mesg_spb           = 'BMA0015'                        " & vbCrLf & _
            "        AND        f.nu_sequ_cntr_repe     > 1                                " & vbCrLf & _
            "        AND    c.nu_sequ_cntr_repe             = F.nu_sequ_cntr_repe          " & vbCrLf & _
            "        AND    c.nu_ctrl_if                    = F.nu_ctrl_if                 " & vbCrLf & _
            "        AND    c.dh_regt_mesg_spb              = F.dh_regt_mesg_spb           " & vbCrLf & _
            "        " & strWhereFiltroMesgOper & _
            "        ) f                                                                    " & vbCrLf
    
    strSQL = strSQL & _
            "Where  a.co_veic_lega                  = v.co_veic_lega                        " & vbCrLf & _
            "AND    a.sg_sist                       = v.sg_sist                             " & vbCrLf & _
            "AND    c.tp_bkof                       = v.tp_bkof                             " & vbCrLf & _
            "AND    D.CO_SITU_PROC                  = A.CO_ULTI_SITU_PROC                   " & vbCrLf & _
            "AND    a.nu_ativ_merc                  = f.nu_ativ_merc                    (+) " & vbCrLf & _
            "AND    a.co_cnta_cutd_selic_veic_lega  = f.co_cnta_cutd_selic_veic_lega    (+) " & vbCrLf & _
            "AND    a.dt_venc_ativ                  = f.dt_venc                         (+) " & vbCrLf & _
            "AND    TRUNC(a.dt_oper_ativ)           = TRUNC(f.DH_REGT_MESG_SPB (+))         " & vbCrLf & _
            strWhereFiltroOper & vbCrLf & strWhereSegregacao & vbCrLf & _
            "GROUP BY c.de_bkof, a.nu_ativ_merc, a.nu_comd_oper, a.co_cnta_cutd_selic_veic_lega, a.dt_venc_ativ, d.de_situ_proc, f.in_oper_debt_cred, f.qt_ativ_merc, f.nu_ctrl_if, f.Dh_regt_mesg_spb, f.nu_sequ_cntr_repe, f.dh_ulti_atlz, f.co_mesg_spb"

       '"AND a.nu_comd_oper =          b.nu_comd_oper_orig (+) " & vbCrLf & _

    ObterNetLiquidacaoFisica = fgQueryXMLLerTodos("NetLiquidacaoFisica", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNetLiquidacaoFisica Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Executar a chamada do método fgObterNumeroControleIF para a obtenção do número
'' de controle IF.
Public Function ObterNumeroControleIF() As String

On Error GoTo ErrorHandler
    
    ObterNumeroControleIF = fgObterNumeroControleIF

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterNumeroControleIF Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'' Obter o sequencial de ISPB de Destino para a mensagem SPB.
Public Function ObterSequencialISPBDestino(ByVal pstrCodigoMensagem As String, _
                                  Optional ByVal plngCodigoMensagemXML As Long = 0, _
                                  Optional ByVal plngLocaLiquidacao As Long = 0, _
                                  Optional ByVal pstrMensagemXML As String = vbNullString) As Long

Dim objLocaLiquidacao                       As A6A7A8.clsLocalLiquidacao
Dim objInstituicaoSPB                       As A8LQS.clsInstituicaoISPB
Dim objMensagemSPB                          As A8LQS.clsMensagemSPB
Dim xmlDomAux                               As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim strGrupoMensagem                        As String
Dim lngSequencialMensagem                   As Long
Dim strCodigoISPB                           As String * 8
Dim lngISPBDestino                          As Long
Dim blnBuscaEspecifica                      As Boolean      '<-- Utilizado para os Grupos: LDL e LTR ou
                                                            '    para as mensagens: GEN0001 e GEN0003
    On Error GoTo ErrorHandler
    
    strGrupoMensagem = UCase(Mid(Trim(pstrCodigoMensagem), 1, 3))
    blnBuscaEspecifica = (strGrupoMensagem = "LDL" Or strGrupoMensagem = "LTR" Or _
                          UCase(Trim(pstrCodigoMensagem)) = "GEN0001" Or _
                          UCase(Trim(pstrCodigoMensagem)) = "GEN0003")
    
    'Verifica se a busca do ISPB Destino será pelo conteúdo da mensagem...
    If blnBuscaEspecifica Then
        Set xmlDomAux = CreateObject("MSXML2.DOMDocument.4.0")
        
        If plngCodigoMensagemXML <> 0 Then
            'Obtem mensagem XML a partir do código identificador informado
            Call xmlDomAux.loadXML(fgSelectVarchar4000(plngCodigoMensagemXML))
        Else
            Call xmlDomAux.loadXML(pstrMensagemXML)
        End If
        
        'Verifica se a mensagem é uma LDL
        If Not xmlDomAux.selectSingleNode("//ISPBLDL") Is Nothing Then
            strCodigoISPB = Format(xmlDomAux.getElementsByTagName("ISPBLDL").Item(0).Text, "00000000")
        
        'Verifica se a mensagem é uma LTR
        ElseIf Not xmlDomAux.selectSingleNode("//ISPBLTR") Is Nothing Then
            strCodigoISPB = Format(xmlDomAux.getElementsByTagName("ISPBLTR").Item(0).Text, "00000000")
        
        'Verifica se a mensagem é uma GEN (1 ou 3)
        ElseIf Not xmlDomAux.selectSingleNode("//ISPBDestinatario") Is Nothing Then
            strCodigoISPB = Format(xmlDomAux.getElementsByTagName("ISPBDestinatario").Item(0).Text, "00000000")
        
        End If
    
        Set objInstituicaoSPB = CreateObject("A8LQS.clsInstituicaoISPB")
        Call xmlDomAux.loadXML(objInstituicaoSPB.LerTodos(, strCodigoISPB))
        Set objInstituicaoSPB = Nothing
        
        lngISPBDestino = _
            xmlDomAux.selectSingleNode("Repeat_Instituicao_ISPB/Grupo_Instituicao_ISPB/SQ_ISPB").Text
        
        Set xmlDomAux = Nothing
    Else
        
        Set objMensagemSPB = CreateObject("A8LQS.clsMensagemSPB")
        lngSequencialMensagem = objMensagemSPB.ObterSequenciaMensagem(pstrCodigoMensagem)
        Set objMensagemSPB = Nothing
        
        Set xmlDomAux = CreateObject("MSXML2.DOMDocument.4.0")
        
        strSQL = " SELECT   DECODE(D.CO_ISPB, NULL, C.SQ_ISPB, D.SQ_ISPB)  SQ_ISPB      " & vbNewLine & _
                 " FROM     A8.TB_MENSAGEM              A,                              " & vbNewLine & _
                 "          A8.TB_EVENTO                B,                              " & vbNewLine & _
                 "          A8.TB_INSTITUICAO_SPB       C,                              " & vbNewLine & _
                 "          A8.TB_INSTITUICAO_SPB       D                               " & vbNewLine & _
                 " WHERE    A.SQ_EVEN                   = B.SQ_EVEN                     " & vbNewLine & _
                 " AND      B.SQ_ISPB                   = C.SQ_ISPB                     " & vbNewLine & _
                 " AND      C.SQ_ISPB_SPER              = D.SQ_ISPB     (+)             " & vbNewLine & _
                 " AND      A.SQ_MESG                   = " & lngSequencialMensagem & " " & vbNewLine & _
                 " AND      ROWNUM                      = 1"
        
        Call xmlDomAux.loadXML(fgQueryXMLLerTodos("LocalLiquidacao", strSQL, App.EXEName & "." & TypeName(Me)))
        
        If Not xmlDomAux.selectSingleNode("Repeat_LocalLiquidacao/Grupo_LocalLiquidacao/SQ_ISPB") Is Nothing Then
            lngISPBDestino = xmlDomAux.selectSingleNode("Repeat_LocalLiquidacao/Grupo_LocalLiquidacao/SQ_ISPB").Text
        End If
        
        Set xmlDomAux = Nothing
    End If
    
    ObterSequencialISPBDestino = lngISPBDestino
    
    Exit Function

ErrorHandler:
    Set objLocaLiquidacao = Nothing
    Set objInstituicaoSPB = Nothing
    Set objMensagemSPB = Nothing
    Set xmlDomAux = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSequencialISPBDestino Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Reenviar a mensagem para a mensageria gerando um novo header para o sistema A7.
Public Function Reenviar(ByVal pvntNumeroSequencia As Variant) As String

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim strMensagemSPB                          As String
Dim strXMLMensagem                          As String
Dim xmlMensagemSPB                          As MSXML2.DOMDocument40
Dim xmlDOMMensagem                          As MSXML2.DOMDocument40
Dim strNumeroCtrlIf                         As String
Dim strHeaderNZ                             As String

On Error GoTo ErrorHandler
    
    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    
    'Obtem a Primeira mensagem para a Operação
    strMensagemSPB = objMensagem.LerPrimeiraMensagem(pvntNumeroSequencia)
    Set xmlMensagemSPB = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlMensagemSPB.loadXML(strMensagemSPB) Then
        'Mensagem para reenvio inexistente
        lngCodigoErroNegocio = 3021
        GoTo ErrorHandler
    End If
    
    'Obtem o texto da mensagem e gera um novo cabeçalho
    With xmlMensagemSPB
        Set xmlDOMMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlDOMMensagem.loadXML fgSelectVarchar4000(.selectSingleNode("//CO_TEXT_XML").Text)

        If Not xmlDOMMensagem.selectSingleNode("//DT_MESG") Is Nothing Then
            If fgDtXML_To_Date(xmlDOMMensagem.selectSingleNode("//DT_MESG").Text) <> fgDataHoraServidor(enumFormatoDataHora.Data) Then
                If Not xmlDOMMensagem.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
                    If fgDtXML_To_Date(xmlDOMMensagem.selectSingleNode("//DT_OPER_ATIV").Text) <> fgDataHoraServidor(enumFormatoDataHora.Data) Then
                        ' 3101 - Data da Mensagem Inválida para o Reenvio.
                        lngCodigoErroNegocio = 3101
                        GoTo ErrorHandler
                    End If
                Else
                    ' 3101 - Data da Mensagem Inválida para o Reenvio.
                    lngCodigoErroNegocio = 3101
                    GoTo ErrorHandler
                End If
            End If
        End If

        udtProtocolo.CodigoEmpresa = Format(CLng(.selectSingleNode("//CO_EMPR").Text), "00000")
        udtProtocolo.SiglaSistemaDestino = "NZ "
        udtProtocolo.SiglaSistemaOrigem = "A8 "
        udtProtocolo.TipoMensagem = .selectSingleNode("//CO_MESG_SPB").Text & "  "
        LSet udtProtocoloAux = udtProtocolo
        
        strNumeroCtrlIf = fgObterNumeroControleIF
        strHeaderNZ = fgMontaHeaderMensageNZ(.selectSingleNode("//CO_MESG_SPB").Text, _
                                             .selectSingleNode("//CO_EMPR").Text, _
                                             strNumeroCtrlIf)
        
        xmlDOMMensagem.selectSingleNode("//NU_CTRL_IF").Text = strNumeroCtrlIf
        xmlDOMMensagem.selectSingleNode("//TX_HEAD_NZ").Text = strHeaderNZ
        
        Call objMensagem.Incluir(xmlDOMMensagem.xml, _
                                 strNumeroCtrlIf, _
                                 pvntNumeroSequencia, _
                                 0, .selectSingleNode("//CO_MESG_SPB").Text, _
                                 CLng("0" & .selectSingleNode("//TP_BKOF").Text), _
                                 CLng("0") & .selectSingleNode("//CO_EMPR").Text, _
                                 .selectSingleNode("//NU_COMD_OPER").Text, _
                                 "", _
                                 enumStatusMensagem.EnviadaBUS, _
                                 enumIndicadorSimNao.Sim, _
                                 vbNullString, _
                                 CLng("0" & .selectSingleNode("//CO_LOCA_LIQU").Text), _
                                 .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                 .selectSingleNode("//SG_SIST").Text)

    End With
    
    Set objMensagem = Nothing
    Set xmlMensagemSPB = Nothing

    Reenviar = udtProtocoloAux.String & xmlDOMMensagem.xml

    Set xmlDOMMensagem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set xmlMensagemSPB = Nothing
    Set xmlDOMMensagem = Nothing
    Set objMensagem = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reenviar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Executar a chamada do método fgObterISPBIF.
Public Function ObterISPBIF(ByVal plngCodigoEmpresa As Long)

On Error GoTo ErrorHandler

    ObterISPBIF = fgObterISPBIF(plngCodigoEmpresa)

Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterISPBIF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Montar um HTML de mensagens internas.
Private Function flMontaHTMLLayoutInterno(ByVal pxmlLayoutInterno As String) As String

Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRetorno                              As String

On Error GoTo ErrorHandler

    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    If Mid(pxmlLayoutInterno, 1, 1) = "<" Then
        Call xmlDomLeitura.loadXML(pxmlLayoutInterno)
    Else
        Call xmlDomLeitura.loadXML(Mid(pxmlLayoutInterno, 221))
    End If
    
    For Each objDomNode In xmlDomLeitura.documentElement.childNodes
        strRetorno = strRetorno & _
            "<TR><TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                    objDomNode.nodeName & _
                "</Font></TD>" & _
                "<TD BGColor=""#EEEEEE""><Font Size=2 Face=Verdana>" & _
                         objDomNode.selectSingleNode("//" & objDomNode.nodeName).Text & _
                "</Font></TD></TR>" & vbNewLine
    Next
    
    flMontaHTMLLayoutInterno = strRetorno

    Set xmlDomLeitura = Nothing
    
Exit Function
ErrorHandler:
    Set xmlDomLeitura = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaHTMLLayoutInterno Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Formata a mensagem de uma operação no estilo HTML
Public Function ObterMensagemHTMLPorOperacao(ByVal pvntSequencialOper As Variant) As String

Dim strSQL                                  As String
Dim rsOperacao                              As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT CO_TEXT_XML   " & vbNewLine & _
             " FROM   A8.TB_OPER_ATIV_MESG_INTE  " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & Abs(pvntSequencialOper) & "  " & vbNewLine & _
             " AND    DH_MESG_INTE   = ( " & vbNewLine & _
                                         " SELECT MAX(DH_MESG_INTE) " & vbNewLine & _
                                         "   FROM A8.TB_OPER_ATIV_MESG_INTE " & vbNewLine & _
                                         "  WHERE NU_SEQU_OPER_ATIV = " & Abs(pvntSequencialOper) & vbNewLine & _
                                         "    AND TP_SOLI_MESG_INTE IN (1,2)" & ")"

    If pvntSequencialOper < 0 Then
        strSQL = Replace(strSQL, "A8.TB_", "A8HIST.TB_")
    End If

    Set rsOperacao = fgQuerySQL(strSQL)
    
    If pvntSequencialOper < 0 Then
        ObterMensagemHTMLPorOperacao = ObterMensagemHTML(rsOperacao!CO_TEXT_XML * -1)
    Else
        ObterMensagemHTMLPorOperacao = ObterMensagemHTML(rsOperacao!CO_TEXT_XML)
    End If

    Set rsOperacao = Nothing

Exit Function
ErrorHandler:

    Set rsOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagemHTMLPorOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Processa o reenvio da mensagem identificada
Public Function ReenviarMensagem(ByVal pstrNumeroControleIF As String, _
                                 ByVal pstrDHUltimaAtualizacao As String) As Boolean

Dim udtHeaderMensagem                       As udtHeaderMensagem

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux

Dim objMensagem                             As A8LQS.clsMensagemEnviadaRecebida
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objProcessoMensagemBMA                  As A8LQS.clsProcessoMensagemBMA
Dim objOperacao                             As A8LQS.clsOperacao
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlMensagemSPB                          As MSXML2.DOMDocument40
Dim xmlDOMMensagem                          As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim strNumeroCtrlIf                         As String
Dim strHeaderNZ                             As String
Dim strMensagemSPB                          As String
Dim strXMLMensagem                          As String
Dim strDataRegistroMensagem                 As String
Dim lngSequenciaConciliacaoMensagem         As Long

On Error GoTo ErrorHandler

    Set objMensagem = CreateObject("A8LQS.clsMensagemEnviadaRecebida")

    strMensagemSPB = objMensagem.LerMensagemIda(pstrNumeroControleIF)

    Set xmlMensagemSPB = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMensagemSPB.loadXML(strMensagemSPB) Then
        'Mensagem para reenvio inexistente
        lngCodigoErroNegocio = 3021
        GoTo ErrorHandler
    End If

    objMensagem.AlterarStatus pstrNumeroControleIF, _
                              xmlMensagemSPB.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                              enumStatusMensagem.Encerrada, _
                              pstrDHUltimaAtualizacao

    'Alterar o status da operação relacionada a mensagem que está sendo reenviada
    If Not xmlMensagemSPB.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
        If Val("0" & xmlMensagemSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) <> 0 Then

            Set objOperacao = CreateObject("A8LQS.clsOperacao")

            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMensagemSPB.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

            If xmlOperacao.xml = vbNullString Then
                Exit Function
            End If
            
            If xmlMensagemSPB.selectSingleNode("//CO_MESG_SPB").Text = "CTP0100" Then
                objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                              0, _
                                              enumTipoAcao.CancelamentoEnviado
            Else
                objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                              enumStatusOperacao.Liberada, _
                                              xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text
            End If
            
            Set objOperacao = Nothing
        End If
    End If

    'Obtem o texto da mensagem e gera um novo cabeçalho
    With xmlMensagemSPB
        
        Set xmlDOMMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        
        If xmlMensagemSPB.selectSingleNode("//CO_MESG_SPB").Text = "CTP0100" Then
            
            strXMLMensagem = Mid(fgSelectVarchar4000(.selectSingleNode("//CO_TEXT_XML").Text), Len(udtHeaderMensagem) + Len(udtProtocolo) + 1)
            
            If Not xmlDOMMensagem.loadXML(strXMLMensagem) Then
                xmlDOMMensagem.loadXML fgSelectVarchar4000(.selectSingleNode("//CO_TEXT_XML").Text)
            End If
            
        Else
            xmlDOMMensagem.loadXML fgSelectVarchar4000(.selectSingleNode("//CO_TEXT_XML").Text)
        End If

        udtProtocolo.CodigoEmpresa = Format(CLng(.selectSingleNode("//CO_EMPR").Text), "00000")
        udtProtocolo.SiglaSistemaDestino = "NZ "
        udtProtocolo.SiglaSistemaOrigem = "A8 "
        udtProtocolo.TipoMensagem = .selectSingleNode("//CO_MESG_SPB").Text & "  "
        LSet udtProtocoloAux = udtProtocolo

        strNumeroCtrlIf = fgObterNumeroControleIF
        strHeaderNZ = fgMontaHeaderMensageNZ(.selectSingleNode("//CO_MESG_SPB").Text, _
                                             .selectSingleNode("//CO_EMPR").Text, _
                                             strNumeroCtrlIf)

        If Not xmlDOMMensagem.selectSingleNode("//TX_HEAD_NZ") Is Nothing Then
            xmlDOMMensagem.selectSingleNode("//NU_CTRL_IF").Text = strNumeroCtrlIf
        Else
            Call fgAppendNode(xmlDOMMensagem, xmlDOMMensagem.documentElement.nodeName, "TX_HEAD_NZ", strNumeroCtrlIf)
        End If
       
        If Not xmlDOMMensagem.selectSingleNode("//TX_HEAD_NZ") Is Nothing Then
            xmlDOMMensagem.selectSingleNode("//TX_HEAD_NZ").Text = strHeaderNZ
        Else
            Call fgAppendNode(xmlDOMMensagem, xmlDOMMensagem.documentElement.nodeName, "TX_HEAD_NZ", strHeaderNZ)
        End If
        
        strDataRegistroMensagem = vbNullString
        lngSequenciaConciliacaoMensagem = 0
        
        Select Case xmlMensagemSPB.selectSingleNode("//CO_MESG_SPB").Text
            Case "LTR0003", "LTR0004", "LDL0004", "STR0004"
                If Not .selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG") Is Nothing Then
                    lngSequenciaConciliacaoMensagem = Val(.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text)
                End If
        End Select
        
        Call objMensagem.Incluir(xmlDOMMensagem.xml, _
                                 strNumeroCtrlIf, _
                                 .selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                 lngSequenciaConciliacaoMensagem, _
                                 .selectSingleNode("//CO_MESG_SPB").Text, _
                                 CLng("0" & .selectSingleNode("//TP_BKOF").Text), _
                                 CLng("0") & .selectSingleNode("//CO_EMPR").Text, _
                                 .selectSingleNode("//NU_COMD_OPER").Text, _
                                 "", _
                                 enumStatusMensagem.EnviadaBUS, _
                                 enumIndicadorSimNao.Sim, _
                                 vbNullString, _
                                 CLng("0" & .selectSingleNode("//CO_LOCA_LIQU").Text), _
                                 .selectSingleNode("//CO_VEIC_LEGA").Text, _
                                 .selectSingleNode("//SG_SIST").Text, , _
                                 strDataRegistroMensagem)

    End With

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Select Case xmlMensagemSPB.selectSingleNode("//CO_MESG_SPB").Text
        Case "LTR0003", "LTR0004"
            Call objProcessoOperacao.EnviarMensagemMQ(udtProtocoloAux.String & xmlDOMMensagem.xml, enumIdentificadorFila.BUS)
            
            Set objProcessoMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
            Call fgAppendNode(xmlDOMMensagem, "SISMSG", "DataRegistroNovaMensagem", strDataRegistroMensagem)
            Call objProcessoMensagemBMA.GerenciarMensagemAvisoInformacao(xmlDOMMensagem, vbNullString)
            Set objProcessoMensagemBMA = Nothing
        
        Case "STR0004"
            strMensagemSPB = MontarMensagemPZ(xmlDOMMensagem, xmlMensagemSPB.selectSingleNode("//TP_BKOF").Text)
            Call objProcessoOperacao.EnviarMensagemMQ(strMensagemSPB, enumIdentificadorFila.BUS)
            
        Case Else
            
            If xmlDOMMensagem.selectSingleNode("//TP_OPER") Is Nothing Then
                If lngSequenciaConciliacaoMensagem <> 0 Then
                    
                    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
                    Set xmlConciliacao = CreateObject("MSXML2.DOMDocument.4.0")
                    
                    Call xmlConciliacao.loadXML(objConciliacao.ConsultaConciliacao("> 0", , , , , , , _
                                                                                   lngSequenciaConciliacaoMensagem))
                    
                    If xmlConciliacao.xml <> vbNullString Then
                        If Not xmlDOMMensagem.selectSingleNode("SISMSG") Is Nothing Then
                            
                            Set objOperacao = CreateObject("A8LQS.clsOperacao")
                            Set xmlConciliacao = objOperacao.LerXMLOperacao(xmlConciliacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
                            Set objOperacao = Nothing
                            
                            Call fgAppendXML(xmlDOMMensagem, "SISMSG", xmlConciliacao.xml)
                        
                        End If
                    End If
                    
                    Set objConciliacao = Nothing
                    Set xmlConciliacao = Nothing
                End If
            End If
            
            Call objProcessoOperacao.GerenciarReenvio(udtProtocoloAux.String & xmlDOMMensagem.xml)
    End Select
    Set objProcessoOperacao = Nothing

    ReenviarMensagem = True

    Set xmlDOMMensagem = Nothing
    Set objMensagem = Nothing
    Set xmlMensagemSPB = Nothing

    Exit Function

ErrorHandler:
    Set xmlMensagemSPB = Nothing
    Set xmlDOMMensagem = Nothing
    Set objMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ReenviarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function



'Monta uma mensagem para envio ao STR
Public Function MontarMensagemSTR(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                  ByRef pblnEnviarMensagem As Boolean) As String

Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim lngStatusMensagem                       As Long
Dim lngTipoBackOffice                       As Long
Dim strNumeroControleIFOld                  As String
Dim strNumeroControleIF                     As String

    On Error GoTo ErrorHandler

    'MontarMensagemSTR = flMontarMensagemPZParaMensagem(xmlMensagem, CLng(xmlMensagem.documentElement.selectSingleNode("TP_BKOF").Text))
    strNumeroControleIF = fgObterNumeroControleIF()

    'Gravar Mensagem
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    If xmlMensagem.documentElement.selectSingleNode("HO_AGND") Is Nothing Then
        fgAppendNode xmlMensagem, "SISMSG", "HO_AGND", ""
        lngStatusMensagem = enumStatusMensagem.EnviadaBUS
        pblnEnviarMensagem = True
    Else
        If xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text = "" Or _
            xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text <= Format(Time, "HHMM") Then
            lngStatusMensagem = enumStatusMensagem.EnviadaBUS
            pblnEnviarMensagem = True
        Else
            lngStatusMensagem = enumStatusMensagem.MensagemAgendada
            pblnEnviarMensagem = False
        End If
    End If

    objMensagemEnviadaRecebida.Incluir xmlMensagem.xml, _
                                       strNumeroControleIF, _
                                       xmlMensagem.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text, _
                                       lngTipoBackOffice, _
                                       xmlMensagem.documentElement.selectSingleNode("CO_EMPR").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("NU_COMD_OPER").Text, _
                                       "", lngStatusMensagem, enumIndicadorSimNao.Nao, _
                                       xmlMensagem.documentElement.selectSingleNode("HO_AGND").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("CO_LOCA_LIQU").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                       xmlMensagem.documentElement.selectSingleNode("SG_SIST").Text
    Set objMensagemEnviadaRecebida = Nothing

    MontarMensagemSTR = MontarMensagemSTR & xmlMensagem.xml

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemSTR Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Monta uma mensagem SEL1023
Private Function flMontarSEL1023(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    fgAppendNode xmlRemessa, "SISMSG", "TP_TRAF_LDL", xmlRemessa.documentElement.selectSingleNode("//TpTransfLDL").Text
    
    If Not xmlRemessa.documentElement.selectSingleNode("//NumOpSEL") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "NU_COMD_OPER", xmlRemessa.documentElement.selectSingleNode("//NumOpSEL").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "NU_COMD_OPER", ""
    End If

    fgAppendNode xmlRemessa, "SISMSG", "IN_OPER_DEBT_CRED", xmlRemessa.documentElement.selectSingleNode("//TpDeb_Cred").Text

    If UCase(xmlRemessa.documentElement.selectSingleNode("//TpDeb_Cred").Text) = "D" Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("//CtCstdia").Text
        fgAppendNode xmlRemessa, "SISMSG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("//CtCstdiaBMA").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "CO_CNTA_CEDE", xmlRemessa.documentElement.selectSingleNode("//CtCstdiaBMA").Text
        fgAppendNode xmlRemessa, "SISMSG", "CO_CNTA_CESS", xmlRemessa.documentElement.selectSingleNode("//CtCstdia").Text
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//IdentdTitSEL") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "NU_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("//IdentdTitSEL").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "NU_ATIV_MERC", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//DtVenc") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "DT_VENC_ATIV", xmlRemessa.documentElement.selectSingleNode("//DtVenc").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "DT_VENC_ATIV", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//PU") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "PU_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("//PU").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "PU_ATIV_MERC", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//Qtd") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "QT_ATIV_MERC", xmlRemessa.documentElement.selectSingleNode("//Qtd").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "QT_ATIV_MERC", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//VlrFinanc") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "VA_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("//VlrFinanc").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "VA_OPER_ATIV", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//ChASEL") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "CO_CHAV_ASSO_SELIC", xmlRemessa.documentElement.selectSingleNode("//ChASEL").Text
    Else
        fgAppendNode xmlRemessa, "SISMSG", "CO_CHAV_ASSO_SELIC", ""
    End If
    
    If Not xmlRemessa.documentElement.selectSingleNode("//DtMovto") Is Nothing Then
        fgAppendNode xmlRemessa, "SISMSG", "DT_OPER_ATIV", xmlRemessa.documentElement.selectSingleNode("//DtMovto").Text
    End If

    'RATS 399 - Ocor 01
    For Each xmlNode In xmlRemessa.selectNodes("//IN_OPER_DEBT_CRED")
        If IsNumeric(xmlNode.Text) Then
            xmlNode.Text = IIf(xmlNode.Text = enumTipoDebitoCredito.Debito, "D", "C")
        End If
    Next
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarSEL1023 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Monta mensagem CETIP de acordo com a planilha de mensageria CETIP - ManualdeMensageriaCTPv9.xls

Public Function MontarMensagemCETIP(ByRef pxmlremessa As MSXML2.DOMDocument40) As Boolean

Dim xmlMensageriaCETIP                      As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlNodeTag                              As MSXML2.IXMLDOMNode
Dim xmlNodeRemove                           As MSXML2.IXMLDOMNode
Dim strCodigoMensagem                       As String
Dim strXpathQuery                           As String
Dim strParentContext                        As String
Dim strCodFormaPagtoCTP                     As String
Dim strTaxaJurosCTP                         As String
Dim blnPreencheTAG                          As Boolean

Dim vntFiltro                               As Variant
Dim lngContFiltro                           As Long
Dim strSiglaSistemaCETIP                    As String
Dim strCodigoOperacaoCETIP                  As String
Dim strCodigoOperacao                       As String
Dim intIndicadorDebCred                     As Integer

Dim xmlNodeRemoveList                       As MSXML2.IXMLDOMNodeList

Dim strSubTipoAtivo                         As String
Dim intTipoContraparte                      As Integer
Dim blnAtivosImobiliarios                   As Boolean

    On Error GoTo ErrorHandler
    
    strCodigoMensagem = pxmlremessa.selectSingleNode("//CO_MESG").Text
    
    'Regra de Negócio CETIP
    'RATS 180 - PM 1362040
    If Trim$(strCodigoMensagem) = "CTP0001" Then
        
        If Not pxmlremessa.selectSingleNode("//SG_SIST_CETIP") Is Nothing Then
            
            If pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text = "CETIP" Then
                If pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                   pxmlremessa.selectSingleNode("//CO_PARP_CAMR").Text = pxmlremessa.selectSingleNode("//CO_CNPT_CAMR").Text
                
                ElseIf pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                   If Not pxmlremessa.selectSingleNode("//CO_OPER_CETIP") Is Nothing Then
                        If pxmlremessa.selectSingleNode("//CO_OPER_CETIP").Text = enumOperacaoCETIP_CTP0011.DepositoEmissaoSemFinaceiro Then
                            If Not pxmlremessa.selectSingleNode("//TP_BKOF") Is Nothing Then
                                If pxmlremessa.selectSingleNode("//TP_BKOF").Text <> enumTipoBackOffice.FundosProprios And _
                                   pxmlremessa.selectSingleNode("//TP_BKOF").Text <> enumTipoBackOffice.FundosTerceiros Then
                                   pxmlremessa.selectSingleNode("//DT_AQUI_ATIV_MERC").Text = vbNullString
                                End If
                            End If
                        End If
                   End If
                End If
                
                If Not pxmlremessa.selectSingleNode("//TP_RENT") Is Nothing Then
                    If pxmlremessa.selectSingleNode("//TP_RENT").Text = enumTipoRentabilidade.Pre Then
                        If Not pxmlremessa.selectSingleNode("//VA_FINC_BASE_RESG") Is Nothing Then
                            pxmlremessa.selectSingleNode("//QT_TITU_CETIP").Text = pxmlremessa.selectSingleNode("//VA_FINC_BASE_RESG").Text
                            If Not pxmlremessa.selectSingleNode("//PE_TAXA_CETIP") Is Nothing Then
                                pxmlremessa.selectSingleNode("//PE_TAXA_CETIP").Text = pxmlremessa.selectSingleNode("//PE_TAXA_CETIP").Text
                            End If
                        
                        End If
                    ElseIf pxmlremessa.selectSingleNode("//TP_RENT").Text = enumTipoRentabilidade.pos Then
                        pxmlremessa.selectSingleNode("//QT_TITU_CETIP").Text = pxmlremessa.selectSingleNode("//VA_OPER_ATIV").Text
                        If Not pxmlremessa.selectSingleNode("//PE_TAXA_CETIP") Is Nothing Then
                            pxmlremessa.selectSingleNode("//PE_TAXA_CETIP").Text = ""
                        End If
                    
                    End If
                End If
            
                If Not pxmlremessa.selectSingleNode("//TP_TABE_RESG") Is Nothing Then
                    Select Case pxmlremessa.selectSingleNode("//TP_TABE_RESG").Text
                        Case "S"
                            pxmlremessa.selectSingleNode("//TP_TABE_RESG").Text = "N"
                        Case "M"
                            pxmlremessa.selectSingleNode("//TP_TABE_RESG").Text = "M"
                        Case "N"
                            pxmlremessa.selectSingleNode("//TP_TABE_RESG").Text = vbNullString
                    End Select
                End If
            
            End If
        
        End If
    
    ElseIf Trim$(strCodigoMensagem) = "CTP0052" Then
        
        If Not pxmlremessa.selectSingleNode("//SG_SIST_CETIP") Is Nothing Then
            If pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text = "CETIP" Then
                
                If Not pxmlremessa.selectSingleNode("//TP_BKOF") Is Nothing Then
                    If pxmlremessa.selectSingleNode("//TP_BKOF").Text = enumTipoBackOffice.Tesouraria Then
                        If pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Debito Then
                           pxmlremessa.selectSingleNode("//DT_AQUI_ATIV_MERC").Text = vbNullString
                           pxmlremessa.selectSingleNode("//VA_PU_AQUI").Text = vbNullString
                        End If
                    Else
                        If pxmlremessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text = enumTipoDebitoCredito.Credito Then
                           pxmlremessa.selectSingleNode("//DT_AQUI_ATIV_MERC").Text = vbNullString
                           pxmlremessa.selectSingleNode("//VA_PU_AQUI").Text = vbNullString
                        End If
                    End If
                End If
            
            ElseIf pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text <> "SND" Then
                
                'RATS 229 - PM 1487457
                'RATS 241 - PM 1508014
                If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
                    If pxmlremessa.selectSingleNode("//TP_CNPT").Text = enumTipoContraparte.Cliente1 Then
                        
                        If Not pxmlremessa.selectSingleNode("//VA_PU_NEGO") Is Nothing Then
                            pxmlremessa.selectSingleNode("//VA_PU_NEGO").Text = vbNullString
                        End If
                        
                        If Not pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR") Is Nothing Then
                            pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR").Text = vbNullString
                        End If
                        
                    End If
                End If
        
            ElseIf pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text = "SND" Then
        
                If Not pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR") Is Nothing Then
                    pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR").Text = vbNullString
                End If
        
            End If
        End If
            
    ElseIf Trim$(strCodigoMensagem) = "CTP9001" Then
    
        If Not pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP") Is Nothing Then
            
            If Trim$(pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text) <> vbNullString Then
                If fgVlrXml_To_Decimal(pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text) <> 0 And _
                   pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text <> "15" And _
                   pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text <> "65" And _
                   pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text <> "59" Then
    
                    If Not pxmlremessa.selectSingleNode("//PU_TERM") Is Nothing Then
                        pxmlremessa.selectSingleNode("//PU_TERM").Text = vbNullString
                    End If
    
                End If
                
                If fgVlrXml_To_Decimal(pxmlremessa.selectSingleNode("//CO_INDX_TERM_CETIP").Text) <> 0 Then
                
                    If Not pxmlremessa.selectSingleNode("//TP_INDX_TERM_CETIP") Is Nothing Then
                        pxmlremessa.selectSingleNode("//TP_INDX_TERM_CETIP").Text = vbNullString
                    End If
                    
                    If Not pxmlremessa.selectSingleNode("//DE_INDX_TERM_CETIP") Is Nothing Then
                        pxmlremessa.selectSingleNode("//DE_INDX_TERM_CETIP").Text = vbNullString
                    End If
                
                End If
            End If
        
        End If
        
        blnPreencheTAG = False
                                                       
        If pxmlremessa.selectSingleNode("//TP_CNPT").Text = enumTipoContraparte.Cliente1 Then
            For Each xmlNode In pxmlremessa.selectNodes("//REPE_INFO_CNTR/*")
                If Not xmlNode.selectSingleNode("CO_INDX_CETIP") Is Nothing Then
                    If fgVlrXml_To_Decimal(xmlNode.selectSingleNode("CO_INDX_CETIP").Text) <> 0 Then
                        If Not xmlNode.selectSingleNode("TP_INDX_ESPC") Is Nothing Then
                            xmlNode.selectSingleNode("TP_INDX_ESPC").Text = vbNullString
                        End If
                    Else
                        blnPreencheTAG = True
                    End If
                End If
            Next
        Else
            For Each xmlNode In pxmlremessa.selectNodes("//REPE_INFO_CNTR/*")
                If Not xmlNode.selectSingleNode("CO_INDX_CETIP") Is Nothing Then
                    If fgVlrXml_To_Decimal(xmlNode.selectSingleNode("CO_INDX_CETIP").Text) <> 0 Then
                        If Not xmlNode.selectSingleNode("TP_INDX_ESPC") Is Nothing Then
                            xmlNode.selectSingleNode("TP_INDX_ESPC").Text = vbNullString
                        End If
                    Else
                        If xmlNode.selectSingleNode("IN_PARP_CNPT").Text = "P" Then
                            blnPreencheTAG = True
                        Else
                            blnPreencheTAG = False
                        End If
                    End If
                End If
            Next
        End If
                
        If Not blnPreencheTAG Then
            If Not pxmlremessa.selectSingleNode("//TP_INDX_TERM_CETIP") Is Nothing Then
                If Val(pxmlremessa.selectSingleNode("//TP_INDX_TERM_CETIP").Text) <> 0 Then
                    If Not pxmlremessa.selectSingleNode("//TX_OBSE_QULF_PARM") Is Nothing Then
                        pxmlremessa.selectSingleNode("//TX_OBSE_QULF_PARM").Text = vbNullString
                    End If
                End If
            End If
        End If
    
    End If
    
    'Verifica a existencia do XML de apoio MensageriaCetip.xml
    If Dir(App.Path & "\MensageriaCetip.xml") = "" Then Exit Function

    Set xmlMensageriaCETIP = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlMensageriaCETIP.Load(App.Path & "\MensageriaCetip.xml") Then
        lngCodigoErroNegocio = 9999
        GoTo ErrorHandler
    End If
           
    'Monta a Query Xpath com de acordo com os Filtros para cada mensagem
    strXpathQuery = "//*/MESG[@CO_MESG='" & strCodigoMensagem & "' and "
    
    If Not xmlMensageriaCETIP.selectSingleNode("//Repet_Filtro/Grupo_Filtro[CO_MESG='" & strCodigoMensagem & "']/FILTRO") Is Nothing Then
    
        vntFiltro = Split(xmlMensageriaCETIP.selectSingleNode("//Repet_Filtro/Grupo_Filtro[CO_MESG='" & strCodigoMensagem & "']/FILTRO").Text, ";", , vbBinaryCompare)
    
        For lngContFiltro = 0 To UBound(vntFiltro)
            If Not pxmlremessa.selectSingleNode("//" & vntFiltro(lngContFiltro)) Is Nothing Then
                If lngContFiltro <> UBound(vntFiltro) Then
                    strXpathQuery = strXpathQuery & "@" & vntFiltro(lngContFiltro) & "='" & pxmlremessa.selectSingleNode("//" & vntFiltro(lngContFiltro)).Text & "' and "
                Else
                    strXpathQuery = strXpathQuery & "@" & vntFiltro(lngContFiltro) & "='" & pxmlremessa.selectSingleNode("//" & vntFiltro(lngContFiltro)).Text & "']"
                End If
                If vntFiltro(lngContFiltro) = "SG_SIST_CETIP" Then
                    strSiglaSistemaCETIP = pxmlremessa.selectSingleNode("//" & vntFiltro(lngContFiltro)).Text
                ElseIf vntFiltro(lngContFiltro) = "CO_OPER_CETIP" Then
                    strCodigoOperacaoCETIP = pxmlremessa.selectSingleNode("//" & vntFiltro(lngContFiltro)).Text
                End If
            Else
                strXpathQuery = vbNullString
                Exit For
            End If
        Next
                       
        If strXpathQuery <> vbNullString Then
            For Each xmlNode In xmlMensageriaCETIP.selectNodes(strXpathQuery)
                For Each xmlNodeTag In xmlNode.childNodes
                    'Se conteudo da tag N remover a tag da mensagem de envio para A7
                    If xmlNodeTag.Text = "N" Then
                        
                        'Verifica todos os grupos de repeticao
                        Set xmlNodeRemoveList = pxmlremessa.selectNodes("//" & xmlNodeTag.nodeName)
                        
                        'Para cada tag encontrada remover do xmlRemessa
                        For Each xmlNodeRemove In xmlNodeRemoveList
                            If Not xmlNodeRemove Is Nothing Then
                                
                                Set xmlNode = pxmlremessa.selectSingleNode("//*[" & xmlNodeRemove.nodeName & "='" & xmlNodeRemove.Text & "']")
                                If Not xmlNode Is Nothing Then
                                    xmlNode.removeChild xmlNodeRemove
                                End If
                                
                            End If
                        Next
                        
                    End If
                Next
            Next
        End If
    End If
    
    If Not pxmlremessa.selectSingleNode("//CO_SUB_TIPO_ATIV") Is Nothing Then
        strSubTipoAtivo = pxmlremessa.selectSingleNode("//CO_SUB_TIPO_ATIV").Text
    End If

    If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
        intTipoContraparte = pxmlremessa.selectSingleNode("//TP_CNPT").Text
    End If

    blnAtivosImobiliarios = False
    
    If (strSubTipoAtivo = "CRI" Or _
        strSubTipoAtivo = "CCI" Or _
        strSubTipoAtivo = "LCI" Or _
        strSubTipoAtivo = "LH") And _
        intTipoContraparte = enumTipoContraparte.Cliente1 Then
            
        blnAtivosImobiliarios = True
        
    End If
    
    If strCodigoMensagem = "CTP0052" Then
        
        'Ivan 31/01/2011 - Codigo removido, pois estas regras serao incluidas no sistema de Origem
'        'Tratamento especial para novo fluxo de mensagens com Operacao 64
'        'Ivan 25/05/2010
'        If Not pxmlremessa.selectSingleNode("//TP_MESG") Is Nothing Then
'            strCodigoOperacao = pxmlremessa.selectSingleNode("//TP_MESG").Text
'        End If
'
'        If Right(strCodigoOperacao, 2) = "64" Then
'            If strSubTipoAtivo = "DI" _
'            Or strSubTipoAtivo = "DII" _
'            Or strSubTipoAtivo = "DIM" _
'            Or strSubTipoAtivo = "DIR" _
'            Or strSubTipoAtivo = "DIRP" _
'            Or strSubTipoAtivo = "DIRR" _
'            Or strSubTipoAtivo = "DIRS" _
'            Or strSubTipoAtivo = "DIRG" Then
'                If strCodigoOperacaoCETIP = "14" Then
'                    Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
'                    Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
'                    Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
'                End If
'            End If
'        End If
'        'Fim
        
        If strSubTipoAtivo <> "CCB" Or strSiglaSistemaCETIP <> "SNA" Then
            If Not pxmlremessa.selectSingleNode("//IN_TIT_CTP_INDP") Is Nothing Then
                Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
            End If
        End If
            
        If blnAtivosImobiliarios Then
                
            If strCodigoOperacaoCETIP = "52" Then
                Call fgRemoveNode(pxmlremessa, "PE_TAXA_INTM")
                Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "DT_AGND")
            End If
        
            If strCodigoOperacaoCETIP = "14" Then
                Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
                Call fgRemoveNode(pxmlremessa, "PE_TAXA_INTM")
                Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "DT_AGND")
                Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
            End If
            
            If strCodigoOperacaoCETIP = "352" Then
                Call fgRemoveNode(pxmlremessa, "PE_TAXA_INTM")
                Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                Call fgRemoveNode(pxmlremessa, "DT_AGND")
                Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
                Call fgRemoveNode(pxmlremessa, "TP_OPER_RESE_TECN")
                Call fgRemoveNode(pxmlremessa, "CO_CNPJ_CNPT")
                Call fgRemoveNode(pxmlremessa, "TP_PESS")
            End If
            
            If strCodigoOperacaoCETIP = "53" Then
                Call fgRemoveNode(pxmlremessa, "IN_OPER_DEBT_CRED")
                Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
                Call fgRemoveNode(pxmlremessa, "PE_TAXA_INTM")
                Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
                Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                Call fgRemoveNode(pxmlremessa, "DT_AGND")
                Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
                Call fgRemoveNode(pxmlremessa, "TP_OPER_RESE_TECN")
                Call fgRemoveNode(pxmlremessa, "CO_CNPJ_CNPT")
                Call fgRemoveNode(pxmlremessa, "TP_PESS")
            End If
            
        Else
        
            Call fgRemoveNode(pxmlremessa, "TP_OPER_RESE_TECN")
  
        End If
            
    End If

    'Retirado no Book 3.03 - A pedido da Aline
'    'Para CTP0012, remover tag "IN_VINC_RES_TEC"
'    If strCodigoMensagem = "CTP0012" Then
'        If Not pxmlremessa.selectSingleNode("//IN_VINC_RES_TEC") Is Nothing Then
'            Call fgRemoveNode(pxmlremessa, "IN_VINC_RES_TEC")
'        End If
'    End If

    'Tratamentos especiais para CTP0054
    If strCodigoMensagem = "CTP0054" Then
        If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
            If Val(pxmlremessa.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Cliente1 Then
                
                Select Case strSiglaSistemaCETIP
                    Case "SNA", "CETIP"
                        Select Case strCodigoOperacaoCETIP
                            Case "54", "57"
                                If Not pxmlremessa.selectSingleNode("//VA_PU_NEGO") Is Nothing Then
                                    Call fgRemoveNode(pxmlremessa, "VA_PU_NEGO")
                                End If
                        End Select
                End Select
                
                Select Case strSiglaSistemaCETIP
                    Case "SNA", "CETIP"
                        Select Case strCodigoOperacaoCETIP
                            Case "54"
                                If Not pxmlremessa.selectSingleNode("//VA_PU_RETN_CETIP") Is Nothing Then
                                    Call fgRemoveNode(pxmlremessa, "VA_PU_RETN_CETIP")
                                End If
                        End Select
                End Select
                
            End If
        End If
    
        If Not pxmlremessa.selectSingleNode("//TP_LIQU_OPER_ATIV") Is Nothing Then
            If Val(pxmlremessa.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.SemModalidade Then
                If Not pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR") Is Nothing Then
                    Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                End If
            End If
        End If
        
        If blnAtivosImobiliarios Then
            
                           
                    If strCodigoOperacaoCETIP = "57" Then
                        Call fgRemoveNode(pxmlremessa, "VA_PU_RETN_CETIP")
                        Call fgRemoveNode(pxmlremessa, "DT_AGND")
                        Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
                        Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_NEGO")
                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP_RETR")
                    End If
                    
                    If strCodigoOperacaoCETIP = "54" Then
                        Call fgRemoveNode(pxmlremessa, "DT_AGND")
                        Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
                        Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP")
                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_NEGO")
                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP_RETR")
                        Call fgRemoveNode(pxmlremessa, "DS_INDC_VALR_CNTR_PART")
                        Call fgRemoveNode(pxmlremessa, "TP_INDX_INDC_VALR_CNTR_PART")
                        Call fgRemoveNode(pxmlremessa, "PE_PARM_JURO")
                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_JURO_CETIP")
                        Call fgRemoveNode(pxmlremessa, "CR_CALC_JURO")
                    End If
                    
                    If strCodigoOperacaoCETIP = "554" Then
                        Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
                        Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_RETN")
                        Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                        Call fgRemoveNode(pxmlremessa, "DT_AGND")
                        Call fgRemoveNode(pxmlremessa, "CO_ASSO_CETIP")
                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP")
                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_NEGO")
                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP_RETR")
                        Call fgRemoveNode(pxmlremessa, "DS_INDC_VALR_CNTR_PART")
                        Call fgRemoveNode(pxmlremessa, "TP_INDX_INDC_VALR_CNTR_PART")
                        Call fgRemoveNode(pxmlremessa, "PE_PARM_JURO")
                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_JURO_CETIP")
                        Call fgRemoveNode(pxmlremessa, "CR_CALC_JURO")
                    End If
      
        
        Else
        
              '========================================================================
              'CETIP21 - DEMANDA LEGAL - CTP0054 ENVIADA PELO SIGOM
              'Dia 05/02/2013 - Inclusão do item Sub Tipo Ativo: Debênture
              'O SLCC verifica se o tipo ativo é Debênture e não retira as TAGs
              '========================================================================
              
              If strSubTipoAtivo <> "DEB" Then

                Call fgRemoveNode(pxmlremessa, "DS_INDC_VALR_CNTR_PART")
                Call fgRemoveNode(pxmlremessa, "TP_INDX_INDC_VALR_CNTR_PART")
                Call fgRemoveNode(pxmlremessa, "PE_PARM_JURO")
                Call fgRemoveNode(pxmlremessa, "PE_TAXA_JURO_CETIP")
                Call fgRemoveNode(pxmlremessa, "CR_CALC_JURO")
        
              End If
              
              '========================================================================
              'Fim Inclusão - Cléber Santos - 05/02/2013
              '========================================================================
        
        End If
        
    End If

    'Tratamentos especiais para CTP0055
    If strCodigoMensagem = "CTP0055" Then
        If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
            If Val(pxmlremessa.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Cliente1 Then
                
              '========================================================================
              'CETIP21 - DEMANDA LEGAL - CTP0055 ENVIADA PELO SIGOM - TAG OPICIONAL
              'Dia 14/02/2013 - Permitir o envio da TAG "VA_PU_NEGO" na MSG CTP0055
              'O SLCC não irá mais controlar o envio dessa tag, passará ao legado.
              '========================================================================
                
'                Select Case strSiglaSistemaCETIP
'                    Case "CETIP", "SNA"
'                        Select Case strCodigoOperacaoCETIP
'                            Case "55", "82", "355"
'                                If Not pxmlremessa.selectSingleNode("//VA_PU_NEGO") Is Nothing Then
'                                    Call fgRemoveNode(pxmlremessa, "VA_PU_NEGO")
'                                End If
'                        End Select
'                End Select
'
'                Select Case strSiglaSistemaCETIP
'                    Case "SND"
'                        Select Case strCodigoOperacaoCETIP
'                            Case "355"
'                                If Not pxmlremessa.selectSingleNode("//VA_PU_NEGO") Is Nothing Then
'                                    Call fgRemoveNode(pxmlremessa, "VA_PU_NEGO")
'                                End If
'                        End Select
'                End Select
                
              '========================================================================
              'Fim Inclusão - Cléber Santos - 05/02/2013
              '========================================================================
              
              
                Select Case strSiglaSistemaCETIP
                    Case "SNA"
                        Select Case strCodigoOperacaoCETIP
                            Case "55"
                                If Not pxmlremessa.selectSingleNode("//VA_PU_CETIP_ORIG") Is Nothing Then
                                    pxmlremessa.selectSingleNode("//VA_PU_CETIP_ORIG").Text = "0"
                                End If
                        End Select
                End Select
                
                If Not pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR") Is Nothing Then
                    Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                End If
            
            End If
        End If
    
        Select Case strSiglaSistemaCETIP
            Case "CETIP"
                Select Case strCodigoOperacaoCETIP
                    Case "55", "355"
                        If Not pxmlremessa.selectSingleNode("//VA_PU_CETIP_ORIG") Is Nothing Then
                            pxmlremessa.selectSingleNode("//VA_PU_CETIP_ORIG").Text = "0"
                        End If
                End Select
        End Select
        
    End If

    'Tratamentos especiais para CTP0056
    If strCodigoMensagem = "CTP0056" Then
        If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
            If Val(pxmlremessa.selectSingleNode("//TP_CNPT").Text) = enumTipoContraparte.Cliente1 Then
                
                Select Case strSiglaSistemaCETIP
                    Case "CETIP", "SNA"
                        Select Case strCodigoOperacaoCETIP
                            Case "58"
                                If Not pxmlremessa.selectSingleNode("//VA_PU_NEGO") Is Nothing Then
                                    Call fgRemoveNode(pxmlremessa, "VA_PU_NEGO")
                                End If
                        End Select
                End Select
                
                If Not pxmlremessa.selectSingleNode("//CO_BANC_LIQU_CAMR") Is Nothing Then
                    Call fgRemoveNode(pxmlremessa, "CO_BANC_LIQU_CAMR")
                End If
            
            End If
        End If
    End If

    'Outros tratamentos especiais para CTP0001
    If strCodigoMensagem = "CTP0001" Then
        
        'Ivan 31/01/2011 - Codigo removido, pois estas regras serao incluidas no sistema de Origem
'        'Tratamento especial para novo fluxo de mensagens com CTP4001
'        'Ivan 05/05/2010
'        If Not pxmlremessa.selectSingleNode("//TP_MESG") Is Nothing Then
'            strCodigoOperacao = pxmlremessa.selectSingleNode("//TP_MESG").Text
'        End If
'
'        If Not pxmlremessa.selectSingleNode("//SG_SIST_CETIP") Is Nothing Then
'            strSiglaSistemaCETIP = pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text
'        End If
'
'        If Right(strCodigoOperacao, 2) = "52" Then
'            If strSubTipoAtivo = "DI" _
'            Or strSubTipoAtivo = "DII" _
'            Or strSubTipoAtivo = "DIM" _
'            Or strSubTipoAtivo = "DIR" _
'            Or strSubTipoAtivo = "DIRP" _
'            Or strSubTipoAtivo = "DIRR" _
'            Or strSubTipoAtivo = "DIRS" _
'            Or strSubTipoAtivo = "DIRG" Then
'                If strSiglaSistemaCETIP = "CETIP21" Then
'                    Call fgRemoveNode(pxmlremessa, "CO_IDEF_EMIS")
'                End If
'            End If
'        End If
'        'Fim
        
        If strSubTipoAtivo = "DIR" Then
            If Not pxmlremessa.selectSingleNode("//TP_TABE_RESG") Is Nothing Then
                pxmlremessa.selectSingleNode("//TP_TABE_RESG").Text = vbNullString
            End If
        End If
            
        If blnAtivosImobiliarios Then
            Call fgRemoveNode(pxmlremessa, "CO_IDEF_EMIS")
            Call fgRemoveNode(pxmlremessa, "PE_TAXA_CETIP")
            Call fgRemoveNode(pxmlremessa, "DT_EMIS_ATIV_MERC")
            Call fgRemoveNode(pxmlremessa, "DT_VENC_ATIV")
            Call fgRemoveNode(pxmlremessa, "QT_COTA")
            Call fgRemoveNode(pxmlremessa, "IN_VINC_RES_TEC")
            Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
            Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
            Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
            Call fgRemoveNode(pxmlremessa, "TP_TABE_RESG")
            Call fgRemoveNode(pxmlremessa, "DT_INIC_RESG")
            Call fgRemoveNode(pxmlremessa, "DT_AGND")
            Call fgRemoveNode(pxmlremessa, "DE_HIST_LANC")
        End If
        
    End If
    
    'Ivan 31/01/2011 - Codigo removido, pois estas regras serao incluidas no sistema de Origem
'    'Outros tratamentos especiais para CTP4001
'    'Ivan 05/05/2010
'    If strCodigoMensagem = "CTP4001" Then
'
'        If Not pxmlremessa.selectSingleNode("//TP_MESG") Is Nothing Then
'            strCodigoOperacao = pxmlremessa.selectSingleNode("//TP_MESG").Text
'        End If
'
'        If Not pxmlremessa.selectSingleNode("//SG_SIST_CETIP") Is Nothing Then
'            strSiglaSistemaCETIP = pxmlremessa.selectSingleNode("//SG_SIST_CETIP").Text
'        End If
'
'        If Right(strCodigoOperacao, 2) = "52" Then
'            If strSubTipoAtivo = "DI" _
'            Or strSubTipoAtivo = "DII" _
'            Or strSubTipoAtivo = "DIM" _
'            Or strSubTipoAtivo = "DIR" _
'            Or strSubTipoAtivo = "DIRP" _
'            Or strSubTipoAtivo = "DIRR" _
'            Or strSubTipoAtivo = "DIRS" _
'            Or strSubTipoAtivo = "DIRG" Then
'                If strSiglaSistemaCETIP = "CETIP21" Then
'
'                    Call fgRemoveNode(pxmlremessa, "CO_IDEF_EMIS")
'
'                    If Not pxmlremessa.selectSingleNode("//CD_FORM_PAGTO") Is Nothing Then
'                        strCodFormaPagtoCTP = pxmlremessa.selectSingleNode("//CD_FORM_PAGTO").Text
'                    End If
'
'                    'Codigo Forma Pagamento CETIP = Pre
'                    If strCodFormaPagtoCTP = "12" Then
'                        Call fgRemoveNode(pxmlremessa, "QT_TITU_CETIP")
'                        Call fgRemoveNode(pxmlremessa, "VA_NOML")
'                        Call fgRemoveNode(pxmlremessa, "CO_INDX_CETIP")
'                        Call fgRemoveNode(pxmlremessa, "PE_TAXA_CETIP")
'                        Call fgRemoveNode(pxmlremessa, "CO_CRIT_CALC_JUROS")
'                        Call fgRemoveNode(pxmlremessa, "VA_PU_NEGO")
'                        Call fgRemoveNode(pxmlremessa, "PE_FLUT_JURO_CETIP")
'
'                    'Codigo Forma Pagamento CETIP = Pos
'                    Else
'
'                        Call fgRemoveNode(pxmlremessa, "VA_FINC_BASE_RESG")
'
'                        If Not pxmlremessa.selectSingleNode("//PE_TAXA_CETIP") Is Nothing Then
'                            strTaxaJurosCTP = pxmlremessa.selectSingleNode("//PE_TAXA_CETIP").Text
'                            If strTaxaJurosCTP = vbNullString _
'                            Or strTaxaJurosCTP = 0 Then
'                                Call fgRemoveNode(pxmlremessa, "PE_TAXA_CETIP")
'                                Call fgRemoveNode(pxmlremessa, "CO_CRIT_CALC_JUROS")
'                            End If
'                        Else
'                            Call fgRemoveNode(pxmlremessa, "PE_TAXA_CETIP")
'                            Call fgRemoveNode(pxmlremessa, "CO_CRIT_CALC_JUROS")
'                        End If
'                    End If
'                End If
'            End If
'        End If
'    End If
'    'Fim

    'Outros tratamentos especiais para CTP0011
    If strCodigoMensagem = "CTP0011" Then
        If Val(strCodigoOperacaoCETIP) = enumOperacaoCETIP_CTP0011.RetiradaCustodia Then
            If blnAtivosImobiliarios Then
                Call fgRemoveNode(pxmlremessa, "VA_OPER_ATIV")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "DT_AGND")
            Else
                Call fgRemoveNode(pxmlremessa, "TP_OPER_RESE_TECN")
                Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
            End If
        End If
    End If
    
    'Outros tratamentos especiais para CTP0013
    If strCodigoMensagem = "CTP0013" Then
        If blnAtivosImobiliarios Then
                
            If strCodigoOperacaoCETIP = "13" Then
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_VENC_ATIV")
            End If
            
            If strCodigoOperacaoCETIP = "23" Then
                Call fgRemoveNode(pxmlremessa, "IN_DIRE_CAUC")
                Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_VENC_ATIV")
                Call fgRemoveNode(pxmlremessa, "TP_GRTA")
            End If
            
            If (strCodigoOperacaoCETIP = "35") Or (strCodigoOperacaoCETIP = "3") Then
                Call fgRemoveNode(pxmlremessa, "IN_DIRE_CAUC")
                Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "DT_VENC_ATIV")
                Call fgRemoveNode(pxmlremessa, "TP_GRTA")
            End If
            
            If strCodigoOperacaoCETIP = "71" Then
                Call fgRemoveNode(pxmlremessa, "IN_DIRE_CAUC")
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
                Call fgRemoveNode(pxmlremessa, "NU_COMD_OPER_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_OPER_ATIV_ORIG")
                Call fgRemoveNode(pxmlremessa, "DT_VENC_ATIV")
                Call fgRemoveNode(pxmlremessa, "TP_GRTA")
            End If
            
        Else
        
            Call fgRemoveNode(pxmlremessa, "IN_TIT_CTP_INDP")
            Call fgRemoveNode(pxmlremessa, "TP_GRTA")
        
        End If
    End If
    
    'Outros tratamentos especiais para CTP0021
    If strCodigoMensagem = "CTP0021" Then
        If Val(strCodigoOperacaoCETIP) = enumOperacaoCETIP_CTP0011.RetiradaCustodia Then
        
            If blnAtivosImobiliarios Then
                Call fgRemoveNode(pxmlremessa, "DT_AQUI_ATIV_MERC")
                Call fgRemoveNode(pxmlremessa, "VA_PU_AQUI")
            Else
                Call fgRemoveNode(pxmlremessa, "CO_INVE_CAMR")
            End If
            
        End If
    End If
        
    'Kida 10/03/2010
    'Book 3.01
    If (strSiglaSistemaCETIP = "CETIP21" _
    Or strSiglaSistemaCETIP = "SNA") Then
        If strSubTipoAtivo = "DI" _
        Or strSubTipoAtivo = "DII" _
        Or strSubTipoAtivo = "DIM" _
        Or strSubTipoAtivo = "DIR" _
        Or strSubTipoAtivo = "DIRP" _
        Or strSubTipoAtivo = "DIRR" _
        Or strSubTipoAtivo = "DIRS" _
        Or strSubTipoAtivo = "DIRG" Then
            If Not pxmlremessa.selectSingleNode("//TP_CNPT") Is Nothing Then
                If Val(pxmlremessa.selectSingleNode("//TP_CNPT").Text) <> enumTipoContraparte.Cliente1 Then
                    Call fgRemoveNode(pxmlremessa, "CO_CNPJ_CNPT")
                    Call fgRemoveNode(pxmlremessa, "TP_PESS")
                End If
            End If
        End If
    End If
    
    Set xmlMensageriaCETIP = Nothing

    Exit Function

ErrorHandler:
    Set xmlMensageriaCETIP = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMensagemCETIP Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flIdentBancoLiquidanteCamara(ByRef xmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim objEmpresa                              As A6A7A8.clsEmpresa
Dim strXML                                  As String
Dim xmlEmpresa                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    strXML = objEmpresa.Ler(CLng(xmlRemessa.selectSingleNode("//CO_EMPR").Text))
    Set objEmpresa = Nothing
    
    If xmlRemessa.documentElement.selectSingleNode("CO_BANC_LIQU_CAMR") Is Nothing Then
        fgAppendNode xmlRemessa, "MESG", "CO_BANC_LIQU_CAMR", vbNullString
    End If
    
    If strXML <> vbNullString Then
        Set xmlEmpresa = CreateObject("MSXML2.DOMDocument.4.0")
        xmlEmpresa.loadXML strXML
        xmlRemessa.documentElement.selectSingleNode("CO_BANC_LIQU_CAMR").Text = xmlEmpresa.selectSingleNode("//ID_PART_CAMR_CETIP").Text
        Set xmlEmpresa = Nothing
    End If
    
    Exit Function
ErrorHandler:
    
    Set objEmpresa = Nothing
    Set xmlEmpresa = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIdentBancoLiquidanteCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Montagem das repetições da LDL0003 CBLC
'--------------------------------------------------------------------------------------------------
Public Function MontarRepeticaoLDL0003(ByRef pxmlLDL0001 As MSXML2.DOMDocument40, _
                                       ByRef pxmlRepeticao As MSXML2.DOMDocument40)

'Monta a estrutura de repetição da LDL0003, de concordancia ou discordancia.

Dim ojDomNode                               As IXMLDOMNode

On Error GoTo ErrorHandler

    pxmlRepeticao.loadXML ""

    fgAppendNode pxmlRepeticao, "", "Repet_LDL0003_ResultLiqd", "", ""

    'Para cada item de Repetição
    For Each ojDomNode In pxmlLDL0001.selectNodes("//Repet_LDL0001_ResultLiqd/Grupo_LDL0001_ResultLiqd")

        fgAppendNode pxmlRepeticao, "Repet_LDL0003_ResultLiqd", "Grupo_LDL0003_ResultLiqd", "", ""

        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_CNPJ_LIQU", _
                                                        ojDomNode.selectSingleNode("CNPJNLiqdant").Text, "Repet_LDL0003_ResultLiqd"

        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_PARP_CAMR", _
                                                        "", "Repet_LDL0003_ResultLiqd"

        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "VA_RESL_LIQU", _
                                                        ojDomNode.selectSingleNode("VlrResultLiqdNLiqdant").Text, "Repet_LDL0003_ResultLiqd"

        fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "TP_CONF_DIVG", _
                                                        "C", "Repet_LDL0003_ResultLiqd"
    Next

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarRepeticaoLDL0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function



'KIDA - BMF
Public Function MontarRepeticaoLDL0003_BMF(ByRef pxmlLDL0001 As MSXML2.DOMDocument40, _
                                           ByRef pxmlRepeticao As MSXML2.DOMDocument40)

Dim xmlAux                                  As MSXML2.DOMDocument40
Dim ojDomNode                               As MSXML2.IXMLDOMNode
Dim objNodeOperacao                         As MSXML2.IXMLDOMNode
Dim strXpath                                As String
Dim dblValorLanc                            As Double
Dim strIN_DEBT_CRED                         As String

On Error GoTo ErrorHandler
    
    Set xmlAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlAux.loadXML pxmlRepeticao.xml
    
    pxmlRepeticao.loadXML ""

    fgAppendNode pxmlRepeticao, "", "Repet_LDL0003_ResultLiqd", "", ""

    strIN_DEBT_CRED = pxmlLDL0001.selectSingleNode("//TpDeb_Cred").Text

    'Para cada item de Repetição
    For Each ojDomNode In pxmlLDL0001.selectNodes("//Repet_LDL0001_ResultLiqd/Grupo_LDL0001_ResultLiqd")
        
        strXpath = "Repeat_Processamento/*[IN_DEBT_CRED='" & strIN_DEBT_CRED & "' and CO_PARP_CAMR='" & ojDomNode.selectSingleNode("IdentdPartCamr").Text & "']"
        
        Set objNodeOperacao = xmlAux.selectSingleNode(strXpath)
        
        If Not objNodeOperacao Is Nothing Then
            
            fgAppendNode pxmlRepeticao, "Repet_LDL0003_ResultLiqd", "Grupo_LDL0003_ResultLiqd", "", ""
            fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_CNPJ_LIQU", ojDomNode.selectSingleNode("CNPJNLiqdant").Text, "Repet_LDL0003_ResultLiqd"
            fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "CO_PARP_CAMR", ojDomNode.selectSingleNode("IdentdPartCamr").Text, "Repet_LDL0003_ResultLiqd"
            fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "VA_RESL_LIQU", Replace(ojDomNode.selectSingleNode("VlrResultLiqdNLiqdant").Text, "-", ""), "Repet_LDL0003_ResultLiqd"
            fgAppendNode pxmlRepeticao, "Grupo_LDL0003_ResultLiqd", "TP_CONF_DIVG", objNodeOperacao.selectSingleNode("TP_CONF_DIVG").Text, "Repet_LDL0003_ResultLiqd"
            
            
        End If

        
    Next
    

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarRepeticaoLDL0003 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


'--------------------------------------------------------------------------------------------------
'Monstagem da mensagem STR0007 - Pagamento BMA0112
'--------------------------------------------------------------------------------------------------

Private Function MontarMsgEnvioBMC(ByRef pxmlOperacao As MSXML2.DOMDocument40) As String


On Error GoTo ErrorHandler
        
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NU_CTRL_IF", fgObterNumeroControleIF())
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NU_DOCT", Format$(Now, "HHMMSS"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_ISPB_IF_DEBT", String$(8, "0"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "TP_PESS_REME", pxmlOperacao.selectSingleNode("//TP_PESS_CLIE").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_CNPJ_CPF_REME", String$(15, "0"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NO_REME", String$(80, " "))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_ISPB_IF_CRED", enumISPB.IspbBMC)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_AGEN_CRED", pxmlOperacao.selectSingleNode("//CO_AGEN").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "TP_CNTA_CRED", pxmlOperacao.selectSingleNode("//TP_CONT_PGTO").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NU_CNTA_CRED", pxmlOperacao.selectSingleNode("//NU_CC").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "TP_PESS_CRED", pxmlOperacao.selectSingleNode("//TP_PESS_CLIE").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_CNPJ_CPF_CRED_1", pxmlOperacao.selectSingleNode("//CO_CNPJ_TTLR").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NO_TITU_1", pxmlOperacao.selectSingleNode("//NO_TTLR").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_CNPJ_CPF_CRED_2", String$(15, "0"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NO_TITU_2", String$(80, " "))
    
    If Not pxmlOperacao.selectSingleNode("//CO_FIND_COBE") Is Nothing Then
        Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_FIND_COBE", pxmlOperacao.selectSingleNode("//CO_FIND_COBE").Text)
    ElseIf Not pxmlOperacao.selectSingleNode("//CO_FIND_IF") Is Nothing Then
        Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_FIND_COBE", pxmlOperacao.selectSingleNode("//CO_FIND_IF").Text)
    Else
        Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_FIND_COBE", "")
    End If
    
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_IDEN_TRANS", pxmlOperacao.selectSingleNode("//CO_IDEF_TRAF").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "IN_TEA", "N")
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "DE_HIST_MESG", String$(200, " "))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "NIVE_PREF", " ")
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "CO_LOCA_LIQU", pxmlOperacao.selectSingleNode("//CO_LOCA_LIQU").Text)
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "DT_SIST", Format$(fgDataHoraServidor(DataAux), "YYYYMMDD"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "DT_AGND", String$(8, "0"))
    Call fgAppendNode(pxmlOperacao, pxmlOperacao.documentElement.nodeName, "TX_FILLER_STR0007", String$(667, " "))

    Exit Function

ErrorHandler:
    Set pxmlOperacao = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarMsgEnvioBMC Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'-------------------------------------------------------------------------------------------------
'Controlar o rejeição de TEDs enviadas a Clientes
'-------------------------------------------------------------------------------------------------

Public Function RejeitarTED(ByVal pstrNU_CTRL_IF As String, _
                            ByVal pstrDH_REGT_MESG_SPB As String, _
                            ByVal pstrNU_SEQU_OPER_ATIV As String, _
                            ByVal pstrNU_SEQU_CNCL_OPER_ATIV_MESG As String) As Boolean

Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objOperacao                             As A8LQS.clsOperacao
Dim objMensagemEnviadaRecebida              As A8LQS.clsMensagemEnviadaRecebida
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objConciliacao                          As A8LQS.clsConciliacao

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlConciliacao                          As MSXML2.DOMDocument40
Dim objNode                                 As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set objMensagemEnviadaRecebida = CreateObject("A8LQS.clsMensagemEnviadaRecebida")
    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")

    'Cancelamento do lote de operações pela conciliação, pois para o caso de
    'liquidação de corretora é uma mensagem STR0007 para várias operações
    If pstrNU_SEQU_CNCL_OPER_ATIV_MESG <> vbNullString Then
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        Set xmlConciliacao = New MSXML2.DOMDocument
            
        'Busca as operações com base no código de conciliação da mensagem STR0007
         xmlConciliacao.loadXML (objConciliacao.LerTodos(pstrNU_SEQU_CNCL_OPER_ATIV_MESG))
         
         'Para cada operação do lote sistema executa a mudança de status de cancelamento da STR0007
         For Each objNode In xmlConciliacao.selectNodes("//Grupo_Conciliacao/NU_SEQU_OPER_ATIV")
         
            Set xmlOperacao = objOperacao.LerXMLOperacao(objNode.Text)
            
            Call objProcessoOperacao.GerenciarCancelamento(xmlOperacao, False)
            
            Call objOperacao.AlterarStatusAcao(objNode.Text, _
                                               enumStatusOperacao.EmSer, _
                                               vbNullString, _
                                               , , _
                                               False)
        
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
            
         Next
        
    Else
    
        Set xmlOperacao = objOperacao.LerXMLOperacao(pstrNU_SEQU_OPER_ATIV)
        
        Call objProcessoOperacao.GerenciarCancelamento(xmlOperacao, False)
        
        Call objOperacao.AlterarStatusAcao(pstrNU_SEQU_OPER_ATIV, _
                                           enumStatusOperacao.EmSer, _
                                           vbNullString, _
                                           , , _
                                           False)
    
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao)
    
        
    End If
    
    Call objMensagemEnviadaRecebida.AlterarStatus(pstrNU_CTRL_IF, _
                                                  pstrDH_REGT_MESG_SPB, _
                                                  enumStatusMensagem.MensagemCancelada, _
                                                  vbNullString, _
                                                  , _
                                                  False)
    
    
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing
    Set objConciliacao = Nothing

    RejeitarTED = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objProcessoOperacao = Nothing
    Set objOperacao = Nothing
    Set objMensagemEnviadaRecebida = Nothing
    Set objFluxoProcessamento = Nothing
    Set xmlOperacao = Nothing
    Set objConciliacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "RejeitarTED Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterXMLMensagem(ByVal plngSequencial As Long) As String


On Error GoTo ErrorHandler

    'Obtem mensagem XML a partir do código identificador informado
    If plngSequencial <> 0 Then
        ObterXMLMensagem = fgSelectVarchar4000(plngSequencial)
    End If
   
    Exit Function
    
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterXMLMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'KIDA - BMF
Public Function MontarRepeticaoLDL0004Previa_BMF(ByRef pxmlLDL0001 As MSXML2.DOMDocument40, _
                                                 ByVal pxmlRepeticao As MSXML2.DOMDocument40)


Dim ojDomNode                               As MSXML2.IXMLDOMNode
Dim objNodeEnvioMensagem                    As MSXML2.IXMLDOMNode
Dim objListEnvioMensagem                    As MSXML2.IXMLDOMNodeList
Dim objNodeOperacao                         As MSXML2.IXMLDOMNode
Dim strXpath                                As String
Dim dblValorLanc                            As Double

On Error GoTo ErrorHandler

    'Para cada item de Repetição
    For Each ojDomNode In pxmlLDL0001.selectNodes("//Repet_LDL0001_ResultLiqd/Grupo_LDL0001_ResultLiqd")

        strXpath = "//Repeat_Processamento/*[IN_DEBT_CRED='D']"
        
        Set objListEnvioMensagem = pxmlRepeticao.selectNodes(strXpath)
        
        For Each objNodeEnvioMensagem In objListEnvioMensagem
        
            strXpath = "Repeat_Operacao/*[CO_PARP_CAMR='" & ojDomNode.selectSingleNode("IdentdPartCamr").Text & "']"
            Set objNodeOperacao = objNodeEnvioMensagem.selectSingleNode(strXpath)
            
            If Not objNodeOperacao Is Nothing Then
                ojDomNode.selectSingleNode("VlrResultLiqdNLiqdant").Text = Replace(objNodeEnvioMensagem.selectSingleNode("VA_OPER_ATIV").Text, "-", "")
                dblValorLanc = dblValorLanc + fgVlrXml_To_Decimal(ojDomNode.selectSingleNode("VlrResultLiqdNLiqdant").Text)
            End If
        Next

    Next
    
    pxmlLDL0001.selectSingleNode("//VlrLanc").Text = fgVlr_To_Xml(dblValorLanc)
    
    If Not pxmlLDL0001.selectSingleNode("//VA_FINC") Is Nothing Then
        pxmlLDL0001.selectSingleNode("//VA_FINC").Text = fgVlr_To_Xml(dblValorLanc)
    End If
    

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MontarRepeticaoLDL0004Previa_BMF Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


