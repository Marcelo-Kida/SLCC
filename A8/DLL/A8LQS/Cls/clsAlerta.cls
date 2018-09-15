VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F2945BC01DB"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsavel pelo fornecimento de informações relativas aos alertas, bem
'' como a manipulação dos dados cadastrados e envio de e-mails padrão para
'' destinatários configurados.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                          As COMSVCSLib.ObjectContext
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

'' Consultar todos os alertas cadastrados para o processo de verificação..
Public Function LerTodos(Optional ByVal pstrDataEmissaoAlerta As String, _
                         Optional plngCodigoFatoGeradorAlerta As Long = 0) As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

On Error GoTo ErrorHandler
             
    strWhereSegregacao = fgSegregaDados("A8.TB_FATO_GERA_ALER_GRUP_USUA", _
                                        False, _
                                        "C", _
                                        "C", _
                                        False, _
                                        False, _
                                        True, _
                                        False, _
                                        False, _
                                        True)
             
    strSQL = " SELECT DISTINCT A.NU_SEQU_ALER_EMIT,                 " & vbCrLf & _
             "        A.NU_SEQU_OPER_ATIV,                 " & vbCrLf & _
             "        A.CO_OPER_ATIV,                      " & vbCrLf & _
             "        A.CO_FATO_GERA_ALER,                 " & vbCrLf & _
             "        A.DH_EMIS_ALER,                      " & vbCrLf & _
             "        A.TP_BKOF,                           " & vbCrLf & _
             "        A.NO_VEIC_LEGA,                      " & vbCrLf & _
             "        A.NO_CNPT,                           " & vbCrLf & _
             "        A.NO_TIPO_OPER,                      " & vbCrLf & _
             "        A.NU_COMD_OPER,                      " & vbCrLf & _
             "        A.VA_OPER_ATIV,                      " & vbCrLf & _
             "        A.DH_ULTI_ATLZ,                      " & vbCrLf & _
             "        B.DE_FATO_GERA_ALER                  " & vbCrLf
             
    strSQL = strSQL & _
             " FROM   A8.TB_ALER_EMIT                 A,   " & vbCrLf & _
             "        A8.TB_FATO_GERA_ALER            B,   " & vbCrLf & _
             "        A8.TB_FATO_GERA_ALER_GRUP_USUA  C    " & vbCrLf & _
             " WHERE  A.CO_FATO_GERA_ALER = C.CO_FATO_GERA_ALER " & vbCrLf & _
             " AND    A.TP_BKOF = C.TP_BKOF                " & vbCrLf & _
             " AND    A.CO_FATO_GERA_ALER = B.CO_FATO_GERA_ALER " & vbCrLf & _
             " AND    A.DH_DESA_ALER IS NULL                    " & vbCrLf
             
    If pstrDataEmissaoAlerta <> vbNullString Then
        strSQL = strSQL & _
            " AND    A.DH_EMIS_ALER > " & fgDtXML_To_Oracle(pstrDataEmissaoAlerta) & vbCrLf
    End If
             
    If plngCodigoFatoGeradorAlerta <> 0 Then
        strSQL = strSQL & _
            " AND    A.CO_FATO_GERA_ALER = " & plngCodigoFatoGeradorAlerta
    End If
             
    strSQL = strSQL & strWhereSegregacao

    LerTodos = fgQueryXMLLerTodos("Alerta", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function


'' Enviar e-mail a um destinatário pré-cadastrado, através do recebimento de
'' parâmetros via XML, de acordo com a estrutura abaixo:
'' '<Email>
'' '   <EmailTo></EmailTo>                 'Lista de Email separados por ;
'' '   <EmailSender></EmailSender>         'Nome do Computador (Servidor)
'' '   <EmailSubject></EmailSubject>       'Asssunto do Email
'' '   <EmailTextBody></EmailTextBody>     'Texto do e-mail
'' '</Email>
''
'' A rotina enviará a mensagem no seguinte formato:
'' *********************************************************
'' Notificação de Ocorrência - Sistema A8 - Liquidação
'' Ocorrência             : ...
'' Fato Gerador do Alerta : ...
'' Número da Operação     : ...
'' Veículo Legal          : ...
'' Contraparte            : ...
'' Tipo da Operação       : ...
'' Número do comando      : ...
'' Valor da operação      : ...
'' *********************************************************
'' Não responder este e-mail
'' *********************************************************
Private Sub flEnviaEmail(pxmlDomAlerta As MSXML2.DOMDocument40)

'<Email>
'   <EmailTo></EmailTo>                 'Lista de Email separados por ;
'   <EmailSender></EmailSender>         'Nome do Computador (Servidor)
'   <EmailSubject></EmailSubject>       'Asssunto do Email
'   <EmailTextBody></EmailTextBody>     'Texto do e-mail
'</Email>

Dim objEmail                                As A6A7A8CA.clsEmail
Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice
Dim xmlDomFatoAlerta                        As MSXML2.DOMDocument40
Dim xmlEmail                                As MSXML2.DOMDocument40

Dim strFatoAlerta                           As String
Dim strAssunto                              As String
Dim strCorpo                                As String
Dim strPara                                 As String
Dim strNomeContraparte                      As String

On Error GoTo ErrorHandler

    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    With pxmlDomAlerta.documentElement
    strFatoAlerta = objFatoGeraAlerBkOffice.LerDetalhe(.selectSingleNode("CO_FATO_GERA_ALER").Text, _
                                                       .selectSingleNode("TP_BKOF").Text)
    End With
    Set objFatoGeraAlerBkOffice = Nothing

    Set xmlDomFatoAlerta = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFatoAlerta.loadXML(strFatoAlerta) Then
       'Registro inexistente ou excluído por outro usuário
       lngCodigoErroNegocio = 4002
       GoTo ErrorHandler
    End If
    
    If CLng("0" & xmlDomFatoAlerta.documentElement.selectSingleNode("IN_ENVI_EMAIL").Text) <> enumIndicadorSimNao.Sim Then
        Set xmlDomFatoAlerta = Nothing
        Exit Sub
    End If
    
    strPara = xmlDomFatoAlerta.documentElement.selectSingleNode("DE_EMAIL_ENVI_ALER").Text
    
    strAssunto = xmlDomFatoAlerta.documentElement.selectSingleNode("TX_ASSU_EMAIL_ENVI_ALER").Text
    
    With pxmlDomAlerta.documentElement
        If .selectSingleNode("NO_CNPT") Is Nothing Then
            strNomeContraparte = vbNullString
        Else
            strNomeContraparte = .selectSingleNode("NO_CNPT").Text
        End If

        strCorpo = _
        " *********************************************************" & vbCrLf & _
        " Notificação de Ocorrência - Sistema A8 - Liquidação" & vbCrLf & _
        " Ocorrência             : " & xmlDomFatoAlerta.documentElement.selectSingleNode("TX_EMAIL_ENVI_ALER").Text & vbCrLf & _
        " Fato Gerador do Alerta : " & xmlDomFatoAlerta.documentElement.selectSingleNode("DE_FATO_GERA_ALER").Text & vbCrLf & _
        " Número da Operação     : " & .selectSingleNode("NU_SEQU_OPER_ATIV").Text & vbCrLf & _
        " Veículo Legal          : " & .selectSingleNode("NO_VEIC_LEGA").Text & vbCrLf & _
        " Contraparte            : " & strNomeContraparte & vbCrLf & _
        " Tipo da Operação       : " & .selectSingleNode("NO_TIPO_OPER").Text & vbCrLf & _
        " Número do comando      : " & .selectSingleNode("NU_COMD_OPER").Text & vbCrLf & _
        " Valor da operação      : " & .selectSingleNode("//VA_OPER_ATIV").Text & vbCrLf & _
        " *********************************************************" & vbCrLf
        
        If Not .selectSingleNode("TX_ANEX") Is Nothing Then
        
        If xmlDomFatoAlerta.documentElement.selectSingleNode("IN_ENVI_EMAIL_ANEX").Text = enumIndicadorSimNao.Sim Then
            
            strCorpo = strCorpo & _
            " Mensagem: " & vbCrLf
            
            strCorpo = strCorpo & _
            .selectSingleNode("TX_ANEX").Text & vbCrLf
        
        End If
        End If
        strCorpo = strCorpo & _
        " *********************************************************" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
        " Não responder este e-mail                                "
        
        
    End With
    
    Set xmlDomFatoAlerta = Nothing

    Set xmlEmail = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlEmail, "", "Email", "")
    Call fgAppendNode(xmlEmail, "Email", "EmailTo", strPara)
    Call fgAppendNode(xmlEmail, "Email", "EmailSender", "")
    Call fgAppendNode(xmlEmail, "Email", "EmailSubject", strAssunto)
    Call fgAppendNode(xmlEmail, "Email", "EmailTextBody", strCorpo)
    
    Set objEmail = CreateObject("A6A7A8CA.clsEmail")
    
    objEmail.EnviaEmail xmlEmail.xml
    
    Set objEmail = Nothing
    Set xmlEmail = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Sub
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    Set objFatoGeraAlerBkOffice = Nothing
    Set xmlDomFatoAlerta = Nothing
    Set xmlEmail = Nothing
    Set objEmail = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    fgRaiseError App.EXEName, TypeName(Me), "flEnviaEmail", 0
End Sub

'' Promover a geração de alerta (inclusão do registro e envio de e-mail aos
'' dstinatátios pré-configurados) de acordo com o tipo de mensagem recebida (XML
'' da operação ou mensagem)
Public Function GerarAlerta(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                   Optional ByRef xmlMensagem As MSXML2.DOMDocument40, _
                   Optional ByVal pstrSituacaoMensagem As String, _
                   Optional ByVal penumFatorGerador As enumFatorGeradorAlerta, _
                   Optional ByVal plngStatusOperacao As Long = 0) As Boolean

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim lngTipoOperacao                         As String
Dim strCofatoGeraAler                       As String

On Error GoTo ErrorHandler
    
    'Pika - Carlos 21/10/2003
    'Verificar o que fazer quando não existe CO_VEIC_LEGA
    'Exemplo: quando o tP_mesg = 21 e não encontrou o NUmero de Controle RDC Original
    If xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then Exit Function

    'Tratar se a geração for no recebimento da mensagem
    If Not xmlMensagem Is Nothing Then
        Select Case Mid(xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text, 1, 3)
            Case "SEL"
                Select Case pstrSituacaoMensagem
                    Case "EXP"
                        penumFatorGerador = enumFatorGeradorAlerta.SelicExpiradoPorTempo
                    Case "PEN"
                        penumFatorGerador = enumFatorGeradorAlerta.SelicPendenteTitulo
                    Case "PEO"
                        penumFatorGerador = enumFatorGeradorAlerta.SelicPendenteOperacao
                    Case "RST"
                        penumFatorGerador = enumFatorGeradorAlerta.SelicRejeitadaPeloSTR
                    Case "ERR"
                        penumFatorGerador = enumFatorGeradorAlerta.Erro
                End Select
            Case "RDC"
                Select Case pstrSituacaoMensagem
                    Case "10", "11", "12"
                        penumFatorGerador = enumFatorGeradorAlerta.RDCCanceladaMensagem
                    Case "4"
                        penumFatorGerador = enumFatorGeradorAlerta.SelicExpiradoPorTempo
                    Case "18", "19", "20", "21", "26", "27", "28", "92"
                        penumFatorGerador = enumFatorGeradorAlerta.RDCPendente
                    Case "5", "6", "9", "24", "41"
                        penumFatorGerador = enumFatorGeradorAlerta.RDCRejeitada
                    Case "93"
                        penumFatorGerador = enumFatorGeradorAlerta.RDCConversaoMensagem
                    Case "51"
                        penumFatorGerador = enumFatorGeradorAlerta.RDCRetornoAutomatico
                End Select
            Case "LDL"
                Select Case plngStatusOperacao
                    Case enumStatusOperacao.Rejeitada
                        penumFatorGerador = enumFatorGeradorAlerta.LDLRejeitada
                End Select
            Case "BMA"
                Select Case plngStatusOperacao
                    Case enumStatusOperacao.Rejeitada
                        penumFatorGerador = enumFatorGeradorAlerta.BMARejeitada
                End Select
            Case "LTR"
                Select Case plngStatusOperacao
                    Case enumStatusOperacao.Rejeitada
                        penumFatorGerador = enumFatorGeradorAlerta.LTRRejeitada
                    Case enumStatusOperacao.Pendencia
                        penumFatorGerador = enumFatorGeradorAlerta.LTRPendente
                End Select
            Case "CTP"
                Select Case plngStatusOperacao
                    Case enumStatusOperacao.Rejeitada
                        penumFatorGerador = enumFatorGeradorAlerta.CTPRejeitada
                    Case enumStatusOperacao.Pendencia
                        penumFatorGerador = enumFatorGeradorAlerta.CTPPendente
                End Select
            Case "STR"
                penumFatorGerador = enumFatorGeradorAlerta.MensagemSTRE
        
        End Select
        
    End If
    
    fgAppendNode xmlOperacao, "MESG", "CO_FATO_GERA_ALER", penumFatorGerador

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    strXML = objVeiculoLegal.Ler(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                 xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    xmlVeiculoLegal.loadXML strXML

    fgAppendNode xmlOperacao, "MESG", "TP_BKOF", xmlVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text
    fgAppendNode xmlOperacao, "MESG", "NO_VEIC_LEGA", xmlVeiculoLegal.documentElement.selectSingleNode("NO_VEIC_LEGA").Text

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            strXML = objTipoOperacao.LerPorTipoMensagemRecebimento(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text)
        End If
    Else
        strXML = objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
    End If
    
    If strXML <> vbNullString Then
    '    ' 3024 - Tipo de Operação Inválido.
    '    lngCodigoErroNegocio = 3024
    '    GoTo ErrorHandler
        xmlTipoOperacao.loadXML strXML
    Else
        Exit Function
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "TP_OPER", xmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text
    Else
    
    End If
    
    If xmlTipoOperacao Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", ""
    Else
        fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", xmlTipoOperacao.documentElement.selectSingleNode("NO_TIPO_OPER").Text
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER", vbNullString
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_SEQU_OPER_ATIV", vbNullString
    End If


    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> 0 Then
        'Gerar Alerta
        If flExisteFatoGeraAlerBkOffice(xmlOperacao) Then
            flIncluir xmlOperacao
            flEnviaEmail xmlOperacao
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarAlerta Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Verificar a existência do fato gerador de alerta para o tipo de BackOffice
'' especificado
Private Function flExisteFatoGeraAlerBkOffice(ByRef pxmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objFatoGeraAlerBkOffice                 As A8LQS.clsFatoGeraAlerBkOffice

Dim lngFatoGeraAlert                        As Long
Dim lngTipoBkOffice                         As Long

On Error GoTo ErrorHandler

    lngFatoGeraAlert = CLng("0" & pxmlOperacao.documentElement.selectSingleNode("CO_FATO_GERA_ALER").Text)
    lngTipoBkOffice = CLng("0" & pxmlOperacao.documentElement.selectSingleNode("TP_BKOF").Text)

    Set objFatoGeraAlerBkOffice = CreateObject("A8LQS.clsFatoGeraAlerBkOffice")
    
    flExisteFatoGeraAlerBkOffice = objFatoGeraAlerBkOffice.Existe(lngFatoGeraAlert, lngTipoBkOffice)
    
    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objFatoGeraAlerBkOffice = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExisteFatoGeraAlerBkOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'' Promover a geração de alerta (inclusão do registro e envio de e-mail aos
'' dstinatátios pré-configurados) de acordo com o tipo de mensagem recebida (XML da mensagem sem operação)
Public Function GerarAlertaSemOperacao(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                                       ByVal penumFatorGerador As enumFatorGeradorAlerta) As Boolean

Dim strXML                                  As String
Dim xmlAlerta                               As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    strXML = ObterPropriedades
    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")

    Call xmlAlerta.loadXML(strXML)
    
    'na propriedas alterar:
    xmlAlerta.documentElement.selectSingleNode("CO_FATO_GERA_ALER").Text = penumFatorGerador
    If xmlMensagem.documentElement.selectSingleNode("//CodMsg").Text = "SEL1099" Then
        xmlAlerta.documentElement.selectSingleNode("//VA_OPER_ATIV").Text = xmlMensagem.documentElement.selectSingleNode("//VlrFinanc").Text
    End If

    'Gerar Alerta
    If flExisteFatoGeraAlerBkOffice(xmlAlerta) Then
        flIncluir xmlAlerta
        flEnviaEmail xmlAlerta
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarAlertaSemOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Promover a inclusão de um novo alerta
Private Function flIncluir(ByRef pxmlAlerta As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim vntSequenciaAlerta                      As Variant
Dim strNomeContraparte                      As String

On Error GoTo ErrorHandler

    vntSequenciaAlerta = fgExecuteSequence("A8.SQ_A8_NU_SEQU_ALER_EMIT")

    strSQL = " INSERT INTO A8.TB_ALER_EMIT (" & vbCrLf & _
             "        NU_SEQU_ALER_EMIT,           " & vbCrLf & _
             "        NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "        DH_EMIS_ALER,                " & vbCrLf & _
             "        CO_FATO_GERA_ALER,           " & vbCrLf & _
             "        TP_BKOF,                     " & vbCrLf & _
             "        NO_VEIC_LEGA,                " & vbCrLf & _
             "        NO_CNPT,                     " & vbCrLf & _
             "        NO_TIPO_OPER,                " & vbCrLf & _
             "        NU_COMD_OPER,                " & vbCrLf & _
             "        VA_OPER_ATIV,                " & vbCrLf & _
             "        DH_DESA_ALER,                " & vbCrLf & _
             "        CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
             "        CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
             "        DH_ULTI_ATLZ                 " & vbCrLf & _
             " )"

    With pxmlAlerta.documentElement
        
        If .selectSingleNode("NO_CNPT") Is Nothing Then
            strNomeContraparte = vbNullString
        Else
            strNomeContraparte = .selectSingleNode("NO_CNPT").Text
        End If

        If .selectSingleNode("//VA_OPER_ATIV") Is Nothing Then
            Call fgAppendNode(pxmlAlerta, pxmlAlerta.documentElement.nodeName, "VA_OPER_ATIV", "")
        End If

        strSQL = strSQL & _
             " VALUES ( " & vbCrLf & _
             vntSequenciaAlerta & ", " & vbCrLf & _
             IIf(.selectSingleNode("NU_SEQU_OPER_ATIV").Text = vbNullString, " NULL ", .selectSingleNode("NU_SEQU_OPER_ATIV").Text) & ", " & vbCrLf & _
             fgDataHoraServidor_To_Oracle & ", " & vbCrLf & _
             IIf(.selectSingleNode("CO_FATO_GERA_ALER").Text = vbNullString, " NULL ", .selectSingleNode("CO_FATO_GERA_ALER").Text) & ", " & vbCrLf & _
             IIf(.selectSingleNode("TP_BKOF").Text = vbNullString, " NULL ", .selectSingleNode("TP_BKOF").Text) & ", " & vbCrLf & _
             IIf(.selectSingleNode("NO_VEIC_LEGA").Text = vbNullString, " NULL ", "'" & .selectSingleNode("NO_VEIC_LEGA").Text & "'") & ", " & vbCrLf & _
             IIf(strNomeContraparte = vbNullString, " NULL ", ("'" & strNomeContraparte & "'")) & ", " & vbCrLf & _
             IIf(.selectSingleNode("NO_TIPO_OPER").Text = vbNullString, " NULL ", "'" & .selectSingleNode("NO_TIPO_OPER").Text & "'") & ", " & vbCrLf & _
             IIf(.selectSingleNode("NU_COMD_OPER").Text = vbNullString, " NULL ", "'" & .selectSingleNode("NU_COMD_OPER").Text & "'") & ", " & vbCrLf & _
             IIf(.selectSingleNode("//VA_OPER_ATIV").Text = vbNullString, " 0", fgVlrToDBServer(.selectSingleNode("//VA_OPER_ATIV").Text)) & ", " & vbCrLf & _
             " Null, " & vbCrLf & _
             "'" & fgUsuarioRede & "'," & vbCrLf & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & vbCrLf & _
             "SYSDATE)"

    End With

    fgExecuteSQL strSQL

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'' Controlar o roteamento e execução dos métodos privados de manipulação de dados,
'' tais como:
''      Incluir e
''      Alterar
Public Function Salvar(ByVal pxmlAlerta As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    If pxmlAlerta.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlAlerta
    ElseIf pxmlAlerta.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlAlerta
    ElseIf pxmlAlerta.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        'flExcluir pxmlAlerta
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados, tais como:
''      Incluir
''      Alterar
''      Ler     e
''      LerTodos
Public Function Executar(ByVal pxmlAlerta As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlAlerta.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlAlerta)
        Case "Ler"
        Case "LerTodos"
            Executar = LerTodos(pxmlAlerta.documentElement.selectSingleNode("//DH_EMIS_ALER").Text, _
                                CLng("0" & pxmlAlerta.documentElement.selectSingleNode("//CO_FATO_GERA_ALER").Text))
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

'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
''      TB_ALER_EMIT
Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT  NU_SEQU_ALER_EMIT,          " & vbCrLf & _
             "         NU_SEQU_OPER_ATIV,          " & vbCrLf & _
             "         CO_OPER_ATIV,               " & vbCrLf & _
             "         CO_FATO_GERA_ALER,          " & vbCrLf & _
             "         DH_EMIS_ALER,               " & vbCrLf & _
             "         TP_BKOF,                    " & vbCrLf & _
             "         NO_VEIC_LEGA,               " & vbCrLf & _
             "         NO_CNPT,                    " & vbCrLf & _
             "         NO_TIPO_OPER,               " & vbCrLf & _
             "         NU_COMD_OPER,               " & vbCrLf & _
             "         VA_OPER_ATIV,               " & vbCrLf & _
             "         DH_DESA_ALER,               " & vbCrLf & _
             "         DH_ULTI_ATLZ,               " & vbCrLf & _
             "         '' AS TX_ANEX               " & vbCrLf & _
             " FROM A8.TB_ALER_EMIT                " & vbCrLf

    ObterPropriedades = fgPropriedades("Grupo_Alerta", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a alteração de um alerta existente
Private Function flAlterar(ByRef pxmlAlerta As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_ALER_EMIT " & _
            "  SET    CO_FATO_GERA_ALER       = " & pxmlAlerta.documentElement.selectSingleNode("CO_FATO_GERA_ALER").Text & "," & _
            "         TP_BKOF                 = " & pxmlAlerta.documentElement.selectSingleNode("TP_BKOF").Text & "," & _
            "         NO_VEIC_LEGA            = '" & pxmlAlerta.documentElement.selectSingleNode("NO_VEIC_LEGA").Text & "'," & _
            "         NO_CNPT                 = '" & pxmlAlerta.documentElement.selectSingleNode("NO_CNPT").Text & "'," & _
            "         NO_TIPO                 = '" & pxmlAlerta.documentElement.selectSingleNode("NO_TIPO").Text & "'," & _
            "         NU_COMD_OPER            = '" & pxmlAlerta.documentElement.selectSingleNode("NU_COMD_OPER").Text & "'," & _
            "         VA_OPER_ATIV            = " & fgVlrToDBServer(pxmlAlerta.documentElement.selectSingleNode("//VA_OPER_ATIV").Text) & "," & _
            "         DH_DESA_ALER            = " & fgDtXML_To_Oracle(pxmlAlerta.documentElement.selectSingleNode("DH_DESA_ALER").Text) & "," & _
            "         CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & _
            "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & _
            "         DH_ULTI_ATLZ            =  SYSDATE " & _
            " WHERE   NU_SEQU_OPER_ATIV       = " & pxmlAlerta.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text & _
            " AND     DH_EMIS_ALER            = " & fgDtHrXML_To_Oracle(pxmlAlerta.documentElement.selectSingleNode("DH_EMIS_ALER").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a geração de alerta (inclusão do registro e envio de e-mail aos
'' dstinatátios pré-configurados) para Câmara
Public Function GerarAlertaCamara(ByRef pxmlAlerta As MSXML2.DOMDocument40) As Boolean

Dim strXML                                  As String
Dim xmlAlerta                               As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    'Gerar Alerta
    If flExisteFatoGeraAlerBkOffice(pxmlAlerta) Then
        flIncluir pxmlAlerta
        flEnviaEmail pxmlAlerta
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarAlertaCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter Intervalo de parada do verifica Server
Public Function ObterIntervaloVerificaServer() As String
    
Dim xmlFilas                                As MSXML2.DOMDocument40
    
On Error GoTo ErrorHandler
    
    Set xmlFilas = CreateObject("MSXML2.DOMDocument.4.0")
        
    If xmlFilas.Load(App.Path & "\FilasNET.xml") Then
        ObterIntervaloVerificaServer = xmlFilas.selectSingleNode("//Janela_ParadaVerificacao").xml
    End If
    
    Set xmlFilas = Nothing

    Exit Function

ErrorHandler:
    
    Set xmlFilas = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterIntervaloVerificaServer Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



'' Promover a geração de alerta A6(inclusão do registro e envio de e-mail aos
'' dstinatátios pré-configurados) de acordo com o tipo de mensagem recebida (XML
'' da operação ou mensagem)

Public Function GerarAlertaA6(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                     Optional ByVal penumFatorGerador As enumFatorGeradorAlerta) As Boolean

Dim objVeiculoLegal                         As A6A7A8.clsVeiculoLegal
Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim strXML                                  As String
Dim lngTipoOperacao                         As String
Dim strCofatoGeraAler                       As String

On Error GoTo ErrorHandler
    
    If xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA") Is Nothing Then Exit Function
    
    fgAppendNode xmlOperacao, "MESG", "CO_FATO_GERA_ALER", penumFatorGerador

    Set objVeiculoLegal = CreateObject("A6A7A8.clsVeiculoLegal")
    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    strXML = objVeiculoLegal.Ler(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                 xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    xmlVeiculoLegal.loadXML strXML

    fgAppendNode xmlOperacao, "MESG", "TP_BKOF", xmlVeiculoLegal.documentElement.selectSingleNode("TP_BKOF").Text
    fgAppendNode xmlOperacao, "MESG", "NO_VEIC_LEGA", xmlVeiculoLegal.documentElement.selectSingleNode("NO_VEIC_LEGA").Text

    Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
    Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
            strXML = objTipoOperacao.LerPorTipoMensagemRecebimento(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text)
        End If
    Else
        strXML = objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
    End If
    
    If strXML <> vbNullString Then
        xmlTipoOperacao.loadXML strXML
    Else
        Exit Function
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "TP_OPER", xmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text
    Else
    
    End If
    
    If xmlTipoOperacao Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", ""
    Else
        fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", xmlTipoOperacao.documentElement.selectSingleNode("NO_TIPO_OPER").Text
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER", vbNullString
    End If

    If xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_SEQU_OPER_ATIV", vbNullString
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> 0 Then
        'Gerar Alerta
        If flExisteFatoGeraAlerBkOffice(xmlOperacao) Then
            flEnviaEmail xmlOperacao
        End If
    End If
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarAlertaA6 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function GerarAlertaVeiculoLegalInvalido(ByVal xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim lngTipoBackOffice                       As Long
Dim objTipoOperacao                         As A6A7A8.clsTipoOperacao
Dim xmlTipoOperacao                         As MSXML2.DOMDocument40
Dim objRS                                   As ADODB.Recordset
Dim strXML                                  As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler
        
    strSQL = " SELECT TP_BKOF FROM A8.TB_TIPO_BKOF "
            
    Set objRS = fgQuerySQL(strSQL)
                
    While Not objRS.EOF
        
        lngTipoBackOffice = objRS!TP_BKOF
    
        fgAppendNode xmlOperacao, "MESG", "CO_FATO_GERA_ALER", enumFatorGeradorAlerta.RemessaLegadoDadosInvalidos
        fgAppendNode xmlOperacao, "MESG", "TP_BKOF", lngTipoBackOffice
        fgAppendNode xmlOperacao, "MESG", "NO_VEIC_LEGA", "Não Existe"
        
        Set objTipoOperacao = CreateObject("A6A7A8.clsTipoOperacao")
        Set xmlTipoOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        
        
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
            If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
                strXML = objTipoOperacao.LerPorTipoMensagemRecebimento(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text)
            End If
        Else
            strXML = objTipoOperacao.Ler(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        End If
        
        If strXML <> vbNullString Then
            xmlTipoOperacao.loadXML strXML
        Else
            Exit Function
        End If
        
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "TP_OPER", xmlTipoOperacao.documentElement.selectSingleNode("TP_OPER").Text
        Else
        
        End If
        
        If xmlTipoOperacao Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", ""
        Else
            fgAppendNode xmlOperacao, "MESG", "NO_TIPO_OPER", xmlTipoOperacao.documentElement.selectSingleNode("NO_TIPO_OPER").Text
        End If
    
        If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER", vbNullString
        End If
    
        If xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "NU_SEQU_OPER_ATIV", vbNullString
        End If
    
        If xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text <> 0 Then
            'Gerar Alerta
            If flExisteFatoGeraAlerBkOffice(xmlOperacao) Then
                flIncluir xmlOperacao
                flEnviaEmail xmlOperacao
            End If
        End If
        
        objRS.MoveNext
    Wend
    
    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set objRS = Nothing

    Exit Function
ErrorHandler:


    Set objTipoOperacao = Nothing
    Set xmlTipoOperacao = Nothing
    Set objRS = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerarAlertaVeiculoLegalInvalido Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


