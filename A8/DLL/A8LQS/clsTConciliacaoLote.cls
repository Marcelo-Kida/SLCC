VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTConciliacaoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Variável utilizada para tratamento de erros
Private xmlRetornoErro                      As MSXML2.DOMDocument40
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
Public Function LiberarLoteParaAlcada() As String
Dim rsLotesALiberar             As ADODB.Recordset
Dim rsSomaLote                  As ADODB.Recordset
Dim rsOperacao                  As ADODB.Recordset
Dim blnLoteValido               As Boolean
Dim strLogErro                  As String


On Error GoTo ErrorHandler

    Set rsLotesALiberar = CreateObject("ADODB.Recordset")
    Set rsLotesALiberar = SelecionarLotesALiberar()
    
    
    If Not rsLotesALiberar.EOF Then
        rsLotesALiberar.MoveFirst
        Do Until rsLotesALiberar.EOF

            'Validar as Operações do Lote
            blnLoteValido = ValidarOperacoesPorLote(rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV)
            If blnLoteValido Then
                ProcessarLote rsLotesALiberar, enumAcaoConciliacao.BOConcordar
                ProcessarLote rsLotesALiberar, enumAcaoConciliacao.AdmAreaLiberar
                'Marcar Lote como Liberado
                AlterarStatusLote rsLotesALiberar!CD_LOTE, rsLotesALiberar!SG_SIST, rsLotesALiberar!DT_OPER_ATIV, enumSituacaoLote.Liberado
            End If
            rsLotesALiberar.MoveNext
        Loop
    End If
    
    Exit Function
ErrorHandler:
    Set rsLotesALiberar = Nothing
    Set rsOperacao = Nothing
    Set rsSomaLote = Nothing
    strLogErro = Err.Description
    
    
    
    fgGravaArquivo App.EXEName & "." & "LiberarLoteParaAlcada", strLogErro
    
End Function
Public Function AlterarStatusLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv As Date, ByVal pintSituacaoLote As Integer)
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    Dim objConciliacao                          As A8LQS.clsConciliacao
    On Error GoTo ErrorHandler
    
    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    AlterarStatusLote = objConciliacao.AlterarStatusLote(pintIdLote, pstrSiglaSistema, pdtDataOperAtiv, pintSituacaoLote)
    Set objConciliacao = Nothing
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
ErrorHandler:
    Set objConciliacao = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
Private Function ProcessarLote(prsLotesALiberar As ADODB.Recordset, ByVal pintAcaoProcessamento As Integer)
Dim rsSomaLote                  As ADODB.Recordset
Dim rsOperacao                  As ADODB.Recordset
Dim strXMLProc                  As String
Dim strResultadoOperacao        As String

    On Error GoTo ErrorHandler


    Set rsSomaLote = CreateObject("ADODB.Recordset")
    Set rsSomaLote = ObterSomatoriaOperacaoPorLote(prsLotesALiberar!CD_LOTE, prsLotesALiberar!SG_SIST, prsLotesALiberar!DT_OPER_ATIV)
    
    'Obter as operações do Lote
    Set rsOperacao = CreateObject("ADODB.Recordset")
    Set rsOperacao = ObterUltimaOperacaoPorLote(prsLotesALiberar!CD_LOTE, prsLotesALiberar!SG_SIST, prsLotesALiberar!DT_OPER_ATIV)
    strXMLProc = flMontarXMLProcessamento(rsOperacao, rsSomaLote)
                
    Set rsOperacao = Nothing
    Set rsSomaLote = Nothing
                
    'Concordar o Lote
    flProcessar strXMLProc, pintAcaoProcessamento
    Exit Function
ErrorHandler:
    Set rsSomaLote = Nothing
    Set rsOperacao = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ProcessarLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'Enviar itens de mensagem e operações para liquidação
Private Function flProcessar(ByVal pstrXMLProc As String, ByVal pintAcaoProcessamento As Integer)

Dim strXMLRetorno                           As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

        
    If pstrXMLProc <> vbNullString Then
                
        'Set objOperacaoMensagem = fgCriarObjetoMIU("A8MIU.clsOperacaoMensagem")
        'strXMLRetorno = objOperacaoMensagem.LiquidarCorretoras(intAcaoProcessamento, _
        '                                                       strXMLProc, _
        '                                                       vntCodErro, _
        '                                                       vntMensagemErro)
        
        strXMLRetorno = LiquidarLoteCorretoras(pintAcaoProcessamento, _
                                                pstrXMLProc, _
                                                vntCodErro, _
                                                vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
    End If
    
    If strXMLRetorno <> vbNullString Then
        Set xmlRetornoErro = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetornoErro.loadXML(strXMLRetorno)
    Else
        Set xmlRetornoErro = Nothing
    End If
    
    flProcessar = strXMLRetorno
    
Exit Function
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
End Function

'Monta XML com as chaves das operações que serão processadas
Private Function flMontarXMLProcessamento(ByVal rsOperacao As ADODB.Recordset, ByVal rsSomaLote As ADODB.Recordset) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlProcessamento                        As MSXML2.DOMDocument40
Dim xmlItemEnvioMsg                         As MSXML2.DOMDocument40
Dim xmlItemOperacao                         As MSXML2.DOMDocument40
Dim intIgnoraGradeHorario                   As Integer
Dim strItemKey                              As String
Dim rsOperacoes                             As ADODB.Recordset
    
On Error GoTo ErrorHandler
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")
    
    Set xmlItemEnvioMsg = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlItemEnvioMsg, "", "Grupo_EnvioMensagem", "")
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_EMPR", _
                                       rsOperacao!CO_EMPR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_LOCA_LIQU", _
                                       rsOperacao!CO_LOCA_LIQU)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_ISPB_CNPT", _
                                       rsOperacao!CO_ISPB_BANC_LIQU_CNPT)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_CNPJ_CNPT", _
                                       rsOperacao!CO_CNPJ_CNPT)
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "NO_CNPT", _
                                       rsOperacao!NO_CNPT)
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "TP_IF_CRED_DEBT", _
                                       rsOperacao!TP_IF_CRED_DEBT)
    
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "CO_AGEN_COTR", _
                                       rsOperacao!CO_AGEN_COTR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "NU_CC_COTR", _
                                       rsOperacao!NU_CC_COTR)
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "ID_PGTO_RECB_GRUP", _
                                       IIf(rsSomaLote!VL_TOTL_OPER >= 0, "P", "R"))
                                       
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "VA_LANC", _
                                       fgVlr_To_Xml(rsOperacao!VL_TOTL_LOTE))
            
    Call fgAppendNode(xmlItemEnvioMsg, "Grupo_EnvioMensagem", _
                                       "Repeat_Operacao", _
                                       "")
                                       
    Set rsOperacoes = CreateObject("ADODB.Recordset")
    Set rsOperacoes = ObterOperacoesPorLote(rsOperacao!CD_LOTE, rsOperacao!SG_SIST, rsOperacao!DT_OPER_ATIV)
    
    rsOperacoes.MoveFirst
    Do Until rsOperacoes.EOF
        Set xmlItemOperacao = CreateObject("MSXML2.DOMDocument.4.0")
        Call fgAppendNode(xmlItemOperacao, "", "Grupo_Operacao", "")
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "NU_SEQU_OPER_ATIV", _
                                           rsOperacoes!NU_SEQU_OPER_ATIV)
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "DH_ULTI_ATLZ", _
                                           fgDtHr_To_Xml(rsOperacoes!dh_ulti_atlz))
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "CO_ULTI_SITU_PROC", _
                                            rsOperacoes!CO_ULTI_SITU_PROC)
        Call fgAppendNode(xmlItemOperacao, "Grupo_Operacao", _
                                           "ID_PGTO_RECB", _
                                            IIf(rsSomaLote!VL_TOTL_OPER >= 0, "P", "R"))
                    
        Call fgAppendXML(xmlItemEnvioMsg, "Repeat_Operacao", xmlItemOperacao.xml)
                    
        Set xmlItemOperacao = Nothing
        rsOperacoes.MoveNext
    Loop
    
    Call fgAppendXML(xmlProcessamento, "Repeat_Processamento", xmlItemEnvioMsg.xml)
    Set xmlItemEnvioMsg = Nothing

    
    If xmlProcessamento.selectNodes("Repeat_Processamento/*").length = 0 Then
        flMontarXMLProcessamento = vbNullString
    Else
        flMontarXMLProcessamento = xmlProcessamento.xml
    End If
    
    Set xmlProcessamento = Nothing
    
Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flMontarXMLProcessamento", 0

End Function
Private Function ObterOperacoesPorLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As ADODB.Recordset
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    On Error GoTo ErrorHandler
    
        Set rsOperacoes = CreateObject("ADODB.Recordset")
        strSQL = "Select O.*, " & vbCrLf & _
                 "L.TP_DEBT_CRED_LOTE, " & vbCrLf & _
                 "L.VL_TOTL_LOTE, " & vbCrLf & _
                 "L.QT_TOTL_LOTE, " & vbCrLf & _
                 "L.IN_LIQU_LOTE " & vbCrLf & _
                 " From  A8.TB_OPER_ATIV O, " & vbCrLf & _
                 "       A8.TB_LOTE_OPER_ATIV L  " & vbCrLf & _
                 " Where O.CD_LOTE = L.CD_LOTE " & vbCrLf & _
                 " And   O.SG_SIST = L.SG_SIST " & vbCrLf & _
                 " And   O.DT_OPER_ATIV = L.DT_OPER_ATIV " & vbCrLf & _
                 " And   L.CD_LOTE = " & pintIdLote & vbCrLf & _
                 " And   L.SG_SIST         =   '" & pstrSiglaSistema & "' " & vbCrLf & _
                 " And   L.DT_OPER_ATIV    =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv))
                 
                 
        Set rsOperacoes = fgQuerySQL(strSQL)
        
        Set ObterOperacoesPorLote = rsOperacoes
        
        Set rsOperacoes = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    
    Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
Private Function ObterUltimaOperacaoPorLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As ADODB.Recordset
    Dim strSQL                                  As String
    Dim rsOperacoes                             As ADODB.Recordset
    On Error GoTo ErrorHandler
    
        Set rsOperacoes = CreateObject("ADODB.Recordset")
        strSQL = "Select O.*, " & vbCrLf & _
                 "L.TP_DEBT_CRED_LOTE, " & vbCrLf & _
                 "L.VL_TOTL_LOTE, " & vbCrLf & _
                 "L.QT_TOTL_LOTE, " & vbCrLf & _
                 "L.IN_LIQU_LOTE " & vbCrLf & _
                 " From  A8.TB_OPER_ATIV O, " & vbCrLf & _
                 "       A8.TB_LOTE_OPER_ATIV L  " & vbCrLf & _
                 " Where O.CD_LOTE = L.CD_LOTE " & vbCrLf & _
                 " And   O.SG_SIST = L.SG_SIST " & vbCrLf & _
                 " And   O.DT_OPER_ATIV = L.DT_OPER_ATIV " & vbCrLf & _
                 " And   L.CD_LOTE = " & pintIdLote & vbCrLf & _
                 " And   L.SG_SIST         =   '" & pstrSiglaSistema & "' " & vbCrLf & _
                 " And   L.DT_OPER_ATIV    =   " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf & _
                 " And   rownum = 1 " & vbCrLf & _
                 " Order by O.dh_ulti_atlz desc "
                 
        Set rsOperacoes = fgQuerySQL(strSQL)
        
        Set ObterUltimaOperacaoPorLote = rsOperacoes
        
        Set rsOperacoes = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    
    Exit Function
ErrorHandler:
    Set rsOperacoes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Private Function ObterSomatoriaOperacaoPorLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As ADODB.Recordset
Dim rsLote                                  As ADODB.Recordset
Dim strSQL                                  As String
On Error GoTo ErrorHandler
    strSQL = "select decode(l.tp_debt_cred_lote,1,l.vl_totl_lote*-1,l.vl_totl_lote) vl_totl_lote, " & vbCrLf & _
             "  l.qt_totl_lote, " & vbCrLf & _
             "  sum (decode(o.in_oper_debt_cred,2,o.va_oper_ativ*-1,o.va_oper_ativ)) vl_totl_oper, " & vbCrLf & _
             "  count(o.co_oper_ativ) QT_OPER " & vbCrLf & _
             "from a8.tb_oper_ativ o, " & vbCrLf & _
             "     a8.tb_lote_oper_ativ l " & vbCrLf & _
             "Where o.cd_lote = l.cd_lote " & vbCrLf & _
             "and   o.dt_oper_ativ = l.dt_oper_ativ " & vbCrLf & _
             "and   o.sg_sist = l.sg_sist " & vbCrLf & _
             "and   l.cd_lote = " & pintIdLote & vbCrLf & _
             "and   l.sg_sist = '" & pstrSiglaSistema & "' " & vbCrLf & _
             "and   l.dt_oper_ativ = " & fgDtXML_To_Oracle(fgDt_To_Xml(pdtDataOperAtiv)) & vbCrLf & _
             "group by decode(l.tp_debt_cred_lote,1,l.vl_totl_lote*-1,l.vl_totl_lote), " & vbCrLf & _
             "         l.qt_totl_lote "
    
    Set rsLote = CreateObject("ADODB.Recordset")
    Set rsLote = fgQuerySQL(strSQL)
    Set ObterSomatoriaOperacaoPorLote = rsLote
        
    Set rsLote = Nothing
        
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
        
    Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Private Function ValidarOperacoesPorLote(ByVal pintIdLote As Integer, ByVal pstrSiglaSistema As String, ByVal pdtDataOperAtiv) As Boolean
Dim strSQL                                  As String
Dim rsLote                                  As ADODB.Recordset
Dim intDifQuantidade                        As Integer
Dim lngDifValor                             As Long

    On Error GoTo ErrorHandler
    
    Set rsLote = CreateObject("ADODB.Recordset")
    Set rsLote = ObterSomatoriaOperacaoPorLote(pintIdLote, pstrSiglaSistema, pdtDataOperAtiv)
    
    
    If Not rsLote.EOF Then
        intDifQuantidade = (rsLote!QT_TOTL_LOTE) - CInt(rsLote!QT_OPER)
        lngDifValor = CLng(rsLote!VL_TOTL_LOTE) - CLng(rsLote!VL_TOTL_OPER)
    Else
        ValidarOperacoesPorLote = False
        Exit Function
    End If
    
    'Se houver diferenças do valor e da quantidade no lote não liberar
    If intDifQuantidade <> 0 Or lngDifValor <> 0 Then
        ValidarOperacoesPorLote = False
        Exit Function
    End If
    ValidarOperacoesPorLote = True
    
    Set rsLote = Nothing
    
    Exit Function
ErrorHandler:
    Set rsLote = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "ValidarOperacoesPorLote Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function
Private Function SelecionarLotesALiberar() As ADODB.Recordset
Dim strSQL                                  As String
Dim rsLotes                                 As ADODB.Recordset
    On Error GoTo ErrorHandler

        Set rsLotes = CreateObject("ADODB.Recordset")
        strSQL = "Select * " & vbCrLf & _
                 " From  A8.TB_LOTE_OPER_ATIV " & vbCrLf & _
                 " Where IN_LIQU_LOTE = " & enumSituacaoLote.ALiberar & vbCrLf & _
                 " and   DT_OPER_ATIV  = TRUNC(SYSDATE)"

        Set rsLotes = fgQuerySQL(strSQL)
            
        Set SelecionarLotesALiberar = rsLotes
        
        Set rsLotes = Nothing
        If Not objContext Is Nothing Then
            objContext.SetComplete
        End If
    Exit Function
ErrorHandler:
    Set rsLotes = Nothing
    Call fgRaiseError(App.EXEName, TypeName(Me), "SelecionarLotesALiberar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

' Este método instancia a classe clsTConciliacao para a liquidação com corretoras, de operações e mensagens,
' na camada de negócios, e retorna para a camada de interface.
' Também devolve à camada de interface o resultado do processamento em lote.
Public Function LiquidarLoteCorretoras(ByVal pintAcao As Integer, _
                                   ByVal strXMLOperacoes As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As String

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlDocProcessar                         As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim xmlControleResultado                    As MSXML2.DOMDocument40
Dim xmlControleErro                         As MSXML2.DOMDocument40

Dim strxmlErro                              As String
Dim lngCodErro                              As Long
Dim intTipoErro                             As Integer
Dim strErro                                 As String

Dim blnIgnoraGradeHorario                   As Boolean

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    'Monta XML para controle de erros
    Set xmlControleErro = CreateObject("MSXML2.DOMDocument.4.0")

    'Monta XML para controle de resultados
    Set xmlControleResultado = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlControleResultado, "", "Repeat_ControleErro", "")

    Set xmlDocProcessar = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDocProcessar.loadXML strXMLOperacoes
    For Each xmlDomNode In xmlDocProcessar.documentElement.childNodes
        'blnIgnoraGradeHorario = Val(xmlDomNode.selectSingleNode("IgnoraGradeHorario").Text)

        'Incrementa o controle de erros com a MENSAGEM atual
        Call fgAppendNode(xmlControleResultado, "Repeat_ControleErro", "Grupo_ControleErro", "")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoConfirmacao", 0, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Operacao", vbNullString, "Repeat_ControleErro")

        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_LOCA_LIQU", xmlDomNode.selectSingleNode("CO_LOCA_LIQU").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_ISPB_CNPT", xmlDomNode.selectSingleNode("CO_ISPB_CNPT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_CNPJ_CNPT", xmlDomNode.selectSingleNode("CO_CNPJ_CNPT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TP_IF_CRED_DEBT", xmlDomNode.selectSingleNode("TP_IF_CRED_DEBT").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CO_AGEN_COTR", xmlDomNode.selectSingleNode("CO_AGEN_COTR").Text, "Repeat_ControleErro")
        Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "NU_CC_COTR", xmlDomNode.selectSingleNode("NU_CC_COTR").Text, "Repeat_ControleErro")
                        
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        strxmlErro = objConciliacao.LiquidarCorretoras(pintAcao, xmlDomNode, True, True)
        

        If strxmlErro <> vbNullString Then
            Call xmlControleErro.loadXML(strxmlErro)
            intTipoErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/ErrorType").Text
            lngCodErro = Val(xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Number").Text)

            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "ErroNegocioEspecifico", 0, "Repeat_ControleErro")

            strErro = xmlControleErro.selectSingleNode("Erro/Grupo_ErrorInfo/Description").Text

            'Incrementa o controle de erros com o STATUS do processo (ERRO)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "TipoErro", intTipoErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", strErro, "Repeat_ControleErro")
        Else
            lngCodErro = 0

            'Incrementa o controle de erros com o STATUS do processo (OK)
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "CodigoErro", lngCodErro, "Repeat_ControleErro")
            Call fgAppendNode(xmlControleResultado, "Grupo_ControleErro", "Status", "OK", "Repeat_ControleErro")
        End If
    Next

    LiquidarLoteCorretoras = xmlControleResultado.xml

    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing
    Set xmlDocProcessar = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objConciliacao = Nothing
    Set xmlDocProcessar = Nothing
    Set xmlControleResultado = Nothing
    Set xmlControleErro = Nothing

    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "LiquidarCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

