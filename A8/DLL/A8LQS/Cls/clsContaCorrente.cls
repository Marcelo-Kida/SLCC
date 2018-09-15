VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F1BE7DA0214"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'' Objeto responsável por promover as integrações com os sistemas de C/C:
''      DV - C/C Santander
''      BG - C/C Banespa

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private strConstruct                        As String
Private objContext                          As COMSVCSLib.ObjectContext

Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
    
On Error Resume Next
    
    strConstruct = pobjContruct.ConstructString

End Sub

'Gerenciar recebimento das remessas de resposta de Conta Corrente (DV e BG)
Public Function GerenciarRecebimento(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler

    If Val(xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.RetornoContaCorrenteDV Then
        flRetornoDV xmlMensagem
    ElseIf Val(xmlMensagem.documentElement.selectSingleNode("TP_MESG").Text) = enumTipoMensagemBUS.RetornoContaCorrenteBG Then
        flRetornoBG xmlMensagem
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GerenciarRecebimento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Processa o retorno da remessa de conta corrente com o sistema DV ( Santander )
Private Function flRetornoDV(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim lngStatusIntegracao                     As enumStatusIntegracao
Dim lngStatusIntegracaoEstorno              As enumStatusIntegracao

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim xmlLancamentoCCContabil                 As MSXML2.DOMDocument40
Dim strLancamentoCCContabil                 As String

Dim strLancamentoEstorno                    As String
Dim xmlLancamentoEstorno                    As MSXML2.DOMDocument40

Dim blnEstorno                              As Boolean

On Error GoTo ErrorHandler

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    strLancamentoCCContabil = objLancamentoCCContabil.LerTodos(xmlMensagem.documentElement.selectSingleNode("DV_NU_CTRL_ITGR").Text, _
                                                               enumTipoLancamentoIntegracao.Normal)

    If strLancamentoCCContabil = vbNullString Then
        ' 3041 - Número de controle da integração com sistema DV inválido.
        lngCodigoErroNegocio = 3041
        GoTo ErrorHandler
    End If

    Set xmlLancamentoCCContabil = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLancamentoCCContabil.loadXML strLancamentoCCContabil

    'Verifica se o lançamento é um estorno
    strLancamentoEstorno = objLancamentoCCContabil.LerTodos(xmlMensagem.documentElement.selectSingleNode("DV_NU_CTRL_ITGR").Text, _
                                                            enumTipoLancamentoIntegracao.Estorno)
                                                               
    If strLancamentoEstorno <> vbNullString Then
        Set xmlLancamentoEstorno = CreateObject("MSXML2.DOMDocument.4.0")
        xmlLancamentoEstorno.loadXML strLancamentoEstorno
        blnEstorno = True
    Else
        blnEstorno = False
    End If

    If UCase(xmlMensagem.documentElement.selectSingleNode("DV_CD_RETORNO").Text) = "S" Then
        lngStatusIntegracao = enumStatusIntegracao.IntegradoCC
        lngStatusIntegracaoEstorno = enumStatusIntegracao.Estornado
    Else
        If blnEstorno Then
            lngStatusIntegracao = enumStatusIntegracao.Integrado
            lngStatusIntegracaoEstorno = enumStatusIntegracao.ErroEstornoCC
        Else
            lngStatusIntegracao = enumStatusIntegracao.ErroCC
        End If
    End If

    If CLng(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusIntegracao.EnviadoCC And _
       CLng(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusIntegracao.EnviadoCCAutomatico And _
       CLng(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusIntegracao.EnviadoEstornoCC Then
        ' 3045 -  Situação do registro para integração inválido.
        lngCodigoErroNegocio = 3045
        GoTo ErrorHandler
    End If
    
    objLancamentoCCContabil.AlterarStatus xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                          enumTipoLancamentoIntegracao.Normal, _
                                          lngStatusIntegracao, _
                                          , _
                                          , _
                                          False, _
                                          CLng(xmlMensagem.documentElement.selectSingleNode("DV_CD_ERRO").Text), _
                                          xmlMensagem.documentElement.selectSingleNode("DV_MSG_ERRO").Text
                                          
    If blnEstorno Then
        objLancamentoCCContabil.AlterarStatus xmlLancamentoEstorno.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                              enumTipoLancamentoIntegracao.Estorno, _
                                              lngStatusIntegracaoEstorno, _
                                              , _
                                              , _
                                              False, _
                                              CLng(xmlMensagem.documentElement.selectSingleNode("DV_CD_ERRO").Text), _
                                              xmlMensagem.documentElement.selectSingleNode("DV_MSG_ERRO").Text
                                              
    Else
    
        If lngStatusIntegracao = enumStatusIntegracao.IntegradoCC Then
            'Se não for Estorno, Integra com contabil
            IntegrarContabil xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                             enumIntegrarCCContabilidade.Contabilidade, _
                             True
        End If
    
    
    End If

    IncluirCC "DV", _
              xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR").Text, _
              xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
              enumTipoLancamentoIntegracao.Normal, _
              enumTipoEnvioRetorno.Retorno, _
              xmlMensagem.xml

    Set objLancamentoCCContabil = Nothing

    Exit Function
ErrorHandler:

    Set objLancamentoCCContabil = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRetornoDV Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Processa o retorno da remessa de conta corrente com o sistema BG ( Banespa )
Private Function flRetornoBG(ByRef xmlMensagem As MSXML2.DOMDocument40) As Boolean

Dim lngStatusIntegracao                     As enumStatusIntegracao
Dim lngStatusLancamento                     As enumStatusIntegracao
Dim lngTipoIntegracao                       As enumTipoLancamentoIntegracao
Dim lngSequenciaLancamento                  As Long

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil

Dim xmlLancamentoCCContabil                 As MSXML2.DOMDocument40
Dim strLancamentoCCContabil                 As String

Dim strDescricaoErro                        As String
Dim xmlNodeValue                            As MSXML2.IXMLDOMNode
Dim strNumeroMovimento                      As String
Dim strComplemento                          As String

Dim strNetOperacoes                         As String
Dim intIndicadorDebCred                     As Integer
Dim dblValorLancCC                          As Double
Dim blnEnviaRetorno                         As Boolean
Dim datHoraUltiReproc                       As Date
Dim datHoraGradeHorario                     As Date

'KIDA
'RATS 941
Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objRegra                                As A8LQS.clsRegra
Dim objParamReprocCC                        As A8LQS.clsParamReprocCC
Dim xmlParamReproc                          As MSXML2.DOMDocument

On Error GoTo ErrorHandler

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    strLancamentoCCContabil = objLancamentoCCContabil.LerTodos(, _
                                                               , _
                                                               xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_ITGR").Text)

    Set xmlLancamentoCCContabil = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlLancamentoCCContabil.loadXML strLancamentoCCContabil

    lngStatusLancamento = 0
    If Not xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
        lngStatusLancamento = CLng(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text)
    End If
    
    If lngStatusLancamento <> enumStatusIntegracao.EnviadoCC And _
       lngStatusLancamento <> enumStatusIntegracao.EnviadoCCAutomatico Then
        ' 3045 -  Situação do registro para integração inválido.
        lngCodigoErroNegocio = 3045
        GoTo ErrorHandler
    End If
    
    lngTipoIntegracao = enumTipoLancamentoIntegracao.Normal
    If Not xmlLancamentoCCContabil.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
        lngTipoIntegracao = Val(xmlLancamentoCCContabil.selectSingleNode("//TP_LANC_ITGR").Text)
    End If
    
    If UCase(xmlMensagem.documentElement.selectSingleNode("//dse_status").Text) = "OK" Then
        lngStatusIntegracao = enumStatusIntegracao.IntegradoCC
    Else
        lngStatusIntegracao = enumStatusIntegracao.ErroCC
    End If
    
    strDescricaoErro = ""
    strComplemento = ""
    strNumeroMovimento = ""
    
    If UCase(xmlMensagem.documentElement.selectSingleNode("//dse_status").Text) = "OK" Then
        
        For Each xmlNodeValue In xmlMensagem.selectNodes("//kColl[@id='avisos']/*")
            If UCase(xmlNodeValue.selectSingleNode("@id").Text) = "@AVBGA0000" Then
                strNumeroMovimento = xmlNodeValue.selectSingleNode("@value").Text
            Else
                strComplemento = strComplemento & " | " & xmlNodeValue.selectSingleNode("@value").Text
            End If
        Next
        
        strDescricaoErro = strNumeroMovimento & "|" & strComplemento
        
    Else
        For Each xmlNodeValue In xmlMensagem.selectNodes("//kColl[@id='errores']/*")
            
            strComplemento = xmlNodeValue.selectSingleNode("@value").Text
            
            If strComplemento <> "" Then
                If Len(strComplemento) <= 80 Then
                    strDescricaoErro = strDescricaoErro & " | " & strComplemento
                End If
            End If
            
            'FREITAS 12/07/2010 - Implementação solicitada pelo projeto Teimosinha.
            'Quando o erro for de saldo, o sistema deverá atualizar o status da integração com 114 - Erro Saldo CC
            If Not xmlNodeValue.selectSingleNode("@id") Is Nothing Then
                If xmlNodeValue.selectSingleNode("@id").Text = "@ERBGE0515" Then
                    lngStatusIntegracao = enumStatusIntegracao.ErroSaldoCC
                End If
            End If
        
        Next
    End If
    
    If Len(strDescricaoErro) >= 80 Then
        strDescricaoErro = Mid(strDescricaoErro, 1, 80)
    Else
        strDescricaoErro = strDescricaoErro
    End If
        
    strNetOperacoes = vbNullString
    dblValorLancCC = 0
    intIndicadorDebCred = enumTipoDebitoCredito.Credito
    
    For Each xmlNodeValue In xmlLancamentoCCContabil.selectNodes("Repeat_LancamentoCCContabil/*")
        
        lngSequenciaLancamento = 0
        If Not xmlNodeValue.selectSingleNode("NR_SEQU_LANC") Is Nothing Then
            lngSequenciaLancamento = Val(xmlNodeValue.selectSingleNode("NR_SEQU_LANC").Text)
        End If
    
        Call objLancamentoCCContabil.AlterarStatus(xmlNodeValue.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                   xmlNodeValue.selectSingleNode("TP_LANC_ITGR").Text, _
                                                   lngStatusIntegracao, _
                                                   , _
                                                   , _
                                                   False, _
                                                   , _
                                                   strDescricaoErro, , _
                                                   lngSequenciaLancamento)
                                                   
        If Not xmlLancamentoCCContabil.selectSingleNode("//TP_MESG_RECB_INTE") Is Nothing Then
            If Val(xmlLancamentoCCContabil.selectSingleNode("//TP_MESG_RECB_INTE").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
                
                strNetOperacoes = strNetOperacoes & "|" & xmlNodeValue.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                
                If Val(xmlNodeValue.selectSingleNode("IN_LANC_DEBT_CRED").Text) = enumTipoDebitoCredito.Credito Then
                    dblValorLancCC = dblValorLancCC + _
                                     fgVlrXml_To_Decimal(xmlNodeValue.selectSingleNode("VA_LANC_CC").Text)
                Else
                    dblValorLancCC = dblValorLancCC - _
                                     fgVlrXml_To_Decimal(xmlNodeValue.selectSingleNode("VA_LANC_CC").Text)
                End If
            
            End If
        End If
        
        blnEnviaRetorno = False
        
        'RATS 1023
        'Verifica se existe parametrização para envio automática.
        'Se existir, sistema não envia retorno para o legado até que o horário limite seja excedido,
        'caso contrário o sistema retorna ao legado a informação.
        If lngStatusIntegracao = enumStatusIntegracao.ErroSaldoCC Then
             
            Set objParamReprocCC = CreateObject("A8LQS.clsParamReprocCC")
            
            If Not xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR") Is Nothing Then
                
                Set objOperacao = CreateObject("A8LQS.clsOperacao")
                Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNodeValue.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
                
                If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
                
                    Set xmlParamReproc = CreateObject("MSXML2.DOMDocument.4.0")
                
                    xmlParamReproc.loadXML objParamReprocCC.Ler(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR").Text, _
                                                             xmlLancamentoCCContabil.documentElement.selectSingleNode("//TP_BKOF").Text, _
                                                             xmlLancamentoCCContabil.documentElement.selectSingleNode("//TP_OPER").Text, _
                                                             xmlLancamentoCCContabil.documentElement.selectSingleNode("//IN_LANC_DEBT_CRED").Text, _
                                                             IIf(xmlOperacao.documentElement.selectSingleNode("//TP_CNAL_VEND").Text = "", 0, xmlOperacao.documentElement.selectSingleNode("//TP_CNAL_VEND").Text))
                    
                    If Not xmlParamReproc.selectSingleNode("//TP_OPER") Is Nothing Then
                        
                        If Not xmlParamReproc.documentElement.selectSingleNode("//HO_LIMI_ITGR") Is Nothing Then
                            If xmlParamReproc.documentElement.selectSingleNode("//HO_LIMI_ITGR").Text = gstrDataVazia Then
                                datHoraGradeHorario = BuscarGradeHorario(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR").Text, "BG")
                            Else
                                datHoraGradeHorario = CDate(Mid(xmlParamReproc.documentElement.selectSingleNode("//HO_LIMI_ITGR").Text, 9, 2) & ":" & Mid(xmlParamReproc.documentElement.selectSingleNode("//HO_LIMI_ITGR").Text, 11, 2))
                            End If
                        Else
                            datHoraGradeHorario = BuscarGradeHorario(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR").Text, "BG")
                        End If
                        
                        datHoraUltiReproc = TimeSerial(Hour(DateTime.Now), Minute(DateTime.Now), Second(DateTime.Now))
                        If datHoraUltiReproc > datHoraGradeHorario Then
                            blnEnviaRetorno = True
                        Else
                            blnEnviaRetorno = False
                        End If
                    
                    Else
                        blnEnviaRetorno = True
                    End If
                End If
            End If
        Else
            blnEnviaRetorno = True
        End If
        
        'KIDA
        'RATS 941
        If blnEnviaRetorno Then
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlNodeValue.selectSingleNode("NU_SEQU_OPER_ATIV").Text)
            
            If Not xmlOperacao.selectSingleNode("//TP_OPER") Is Nothing Then
            
                If fgVerificaOperacaoCC(xmlOperacao.selectSingleNode("//TP_OPER").Text) = True Then
                
                    Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
                    
                    If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
                        Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
                    End If
                    
                    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = lngStatusIntegracao
                    
                    Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
                    Set objFluxoProcessamento = Nothing
                End If
                
            End If
            
        End If
        
        Set objOperacao = Nothing
        Set xmlOperacao = Nothing
        
    Next
    
    If dblValorLancCC < 0 Then
        intIndicadorDebCred = IIf(intIndicadorDebCred = 2, 1, 2)
        dblValorLancCC = Abs(dblValorLancCC)
    End If
    
    If lngStatusIntegracao = enumStatusIntegracao.IntegradoCC Then
        'Integra com contabilidade
        
        If strNetOperacoes = vbNullString Then
            IntegrarContabil xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                             enumIntegrarCCContabilidade.Contabilidade, _
                             True, _
                             lngTipoIntegracao, , , , , _
                             lngSequenciaLancamento
        
        Else
            IntegrarContabil xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                             enumIntegrarCCContabilidade.Contabilidade, _
                             True, _
                             , , _
                             fgVlr_To_Xml(dblValorLancCC), _
                             Format$(intIndicadorDebCred, "0"), _
                             strNetOperacoes

        End If
        
        IncluirCC "BG", _
                  xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_EMPR").Text, _
                  xmlLancamentoCCContabil.documentElement.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                  lngTipoIntegracao, _
                  enumTipoEnvioRetorno.Retorno, _
                  xmlMensagem.xml, _
                  lngSequenciaLancamento
    
    End If
    
    Exit Function
    
    Set objLancamentoCCContabil = Nothing
    
ErrorHandler:

    Set objLancamentoCCContabil = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRetornoBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Verifica Operacoes Autorizadas a enviar retorno CC para legado
Public Function fgVerificaOperacaoCC(ByVal pstrTipoOperacao As String) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT DISTINCT(TP_OPER) " & _
             " FROM   A8.TB_FCAO_SIST_TIPO_OPER " & _
             " WHERE  CO_FCAO_SIST = 16 " & _
             " AND    TP_OPER = " & Trim(pstrTipoOperacao)
             
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

'Metodo de exclusão de registro de hitorico de conta corrente
Private Function flExcluirCC(ByVal pvntSequenciaOperacao As Variant, _
                             ByVal plngTipoLancamentoIntegracao As Long) As Boolean

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A8.TB_HIST_ENVI_INFO_CC " & vbNewLine & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & " " & vbNewLine & _
             " AND    TP_LANC_ITGR = " & plngTipoLancamentoIntegracao
             
    fgExecuteSQL strSQL

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluirCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Processo de disponibilização de remessas em conta corrente
Public Function Disponibilizar(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                      Optional ByVal blnEnviaRetornoLegado As Boolean = True) As Boolean

Dim xmlLancCCContabil                       As MSXML2.DOMDocument40
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim objA6A7A8Funcoes                        As A6A7A8.clsA6A7A8Funcoes
Dim datOperacao                             As Date
Dim datServidorD1                           As Date
Dim lngStatus                               As Long
Dim vntPasso                                As Variant
Dim lngNumeroSequencia                      As Long
Dim lngOrigemRecurso                        As Long
Dim blnOrigemRecurso                        As Boolean

    On Error GoTo ErrorHandler
    
    If xmlOperacao.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "TP_LANC_ITGR", enumTipoLancamentoIntegracao.Normal
    End If
        
    If Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) <> enumTipoOperacaoLQS.LancamentoContaCorrenteOperacoesManuais _
    And Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) <> enumTipoOperacaoLQS.LancamentoCCCashFlow _
    And Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) <> enumTipoOperacaoLQS.LancamentoCCCashFlowStrikeFixo _
    And Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) <> enumTipoOperacaoLQS.LancamentoCCSwapJuros Then
    
        'Pagamento de Despesas
        If Not xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU") Is Nothing Then
            If CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrente And _
            CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrenteComTributacao Then
                If Not xmlOperacao.documentElement.selectSingleNode("TP_MESG") Is Nothing Then
                    If CLng(xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text) <> enumTipoMensagemLQS.EnvioPagDespesas Then
                        Exit Function
                    End If
                Else
                    Exit Function
                End If
            End If
        Else
            Exit Function
        End If
                
        'KIDA - CBLC - 22/09/2008
        If Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralCBLC Then
            If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada _
               And Val(xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Normal Then
                    Exit Function
            End If
        'KIDA - CBLC - EVENTOS
        ElseIf Val(xmlOperacao.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.RegistroLiquidacaoEventoCBLC Then
            If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada _
               And Val(xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Normal Then
                    Exit Function
            End If
        Else
            If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada And _
               Val(xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Normal Then
                Exit Function
            End If
        End If
        
    End If
    
    If Not xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV") Is Nothing Then
        If fgVlrXml_To_Decimal(xmlOperacao.documentElement.selectSingleNode("VA_OPER_ATIV").Text) <= 0 Then
            Exit Function
        End If
    End If
    
    If Not xmlOperacao.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        
        If Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bruta Or _
           Val(xmlOperacao.selectSingleNode("//TP_LIQU_OPER_ATIV").Text) = enumTipoLiquidacao.Bilateral Then
            
            If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusOperacao.Liquidada Then
                Exit Function
            End If
            
         End If
         
    End If

    'Verificar se Definitiva ou Compromissada é menor que D-1 não disponibilizar lançamento.
    If CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.Definitiva Or _
       CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaIda Or _
       CLng(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text) = enumTipoOperacaoLQS.CompromissadaIda Then
        
        datOperacao = fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text)

        vntPasso = 10

        Set objA6A7A8Funcoes = CreateObject("A6A7A8.clsA6A7A8Funcoes")
        datServidorD1 = objA6A7A8Funcoes.AdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior)

        vntPasso = 20

        If datOperacao = fgDataHoraServidor(enumFormatoDataHora.Data) Then
            If datServidorD1 <> objA6A7A8Funcoes.AdicionarDiasUteis(datOperacao, 1, enumPaginacao.Anterior) Then
                Exit Function
            End If
        Else
            If datServidorD1 <> datOperacao Then
                Exit Function
            End If
        End If
    
    End If
    
    vntPasso = 30
    If xmlOperacao.documentElement.selectSingleNode("SG_SIST_CC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "SG_SIST_CC", "BG"
    Else
        xmlOperacao.documentElement.selectSingleNode("SG_SIST_CC").Text = "BG"
    End If

    vntPasso = 40
    lngStatus = enumStatusIntegracao.Disponível

    If xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CO_ULTI_SITU_PROC_CC", lngStatus
    Else
        xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC_CC").Text = lngStatus
    End If
    
    Set xmlLancCCContabil = CreateObject("MSXML2.DOMDocument.4.0")
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    If xmlLancCCContabil.loadXML(objLancamentoCCContabil.Ler(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                             xmlOperacao.selectSingleNode("//TP_LANC_ITGR").Text)) Then
        
        If xmlLancCCContabil.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusIntegracao.Antecipado Then
            If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada Then
            
                Call objLancamentoCCContabil.AlterarStatus(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                           enumTipoLancamentoIntegracao.Normal, _
                                                           enumStatusIntegracao.Disponível, , , False)
                Exit Function
            
            End If
        End If
        
    End If
    
    If Not xmlOperacao.documentElement.selectSingleNode("CO_ORIG_RECU") Is Nothing Then
        lngOrigemRecurso = CLng("0" & xmlOperacao.documentElement.selectSingleNode("CO_ORIG_RECU").Text)
    End If
    
    If lngOrigemRecurso = enumOrigemRecurso.TransfIsentaTributada Then
        blnOrigemRecurso = True
    Else
        blnOrigemRecurso = False
    End If
    
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    If blnOrigemRecurso Then
        
        lngNumeroSequencia = 2
        objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                                   enumTipoDebitoCredito.Debito, _
                                                   lngNumeroSequencia, _
                                                   enumTipoContaLancCC.ContaIsenta, _
                                                   True
        
        lngNumeroSequencia = 3
        objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                                   enumTipoDebitoCredito.Credito, _
                                                   lngNumeroSequencia, _
                                                   enumTipoContaLancCC.ContaTributada, _
                                                   True
    
    End If
        
    lngNumeroSequencia = 0
    
    Select Case Val(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
        
        Case enumTipoOperacaoLQS.EnvioPagDespesasTributos, _
             enumTipoOperacaoLQS.EnvioPagDespesasTributosIsenta, _
             enumTipoOperacaoLQS.EnvioPagDespesasTributosTrib
           
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Debito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaTributada, _
                                   False
           
        Case enumTipoOperacaoLQS.EnvioPagDespesasContaCorrente, _
             enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteTrib
          
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Debito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaTributada, _
                                   False
            
            lngNumeroSequencia = lngNumeroSequencia + 1
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Credito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaContraparte, _
                                   False
    
        Case enumTipoOperacaoLQS.EnvioPagDespesasContaCorrenteIsenta
          
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Debito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaIsenta, _
                                   False
            
            lngNumeroSequencia = lngNumeroSequencia + 1
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Credito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaContraparte, _
                                   False
    
        Case enumTipoOperacaoLQS.EnvioPagDespesasBoleto, _
             enumTipoOperacaoLQS.EnvioPagDespesasBoletoIsenta, _
             enumTipoOperacaoLQS.EnvioPagDespesasBoletoTrib
            
            objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                   enumTipoDebitoCredito.Debito, _
                                   lngNumeroSequencia, _
                                   enumTipoContaLancCC.ContaTributada, _
                                   False
        
        Case enumTipoOperacaoLQS.EnvioSTR0006PagDespesas, _
             enumTipoOperacaoLQS.EnvioSTR0007PagDespesas, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesas, _
             enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib, _
             enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib
             
             objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                                        enumTipoDebitoCredito.Debito, _
                                                        lngNumeroSequencia, _
                                                        enumTipoContaLancCC.ContaTributada, _
                                                        False
            
        Case enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta, _
             enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta, _
             enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta
             
             objLancamentoCCContabil.IncluirPagDespesas xmlOperacao, _
                                                        enumTipoDebitoCredito.Debito, _
                                                        lngNumeroSequencia, _
                                                        enumTipoContaLancCC.ContaIsenta, _
                                                        False
            
        Case Else
            
            If Not xmlLancCCContabil.selectSingleNode("//CO_ULTI_SITU_PROC") Is Nothing Then
                If xmlLancCCContabil.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusIntegracao.ErroSaldoCC Then
                    Call objLancamentoCCContabil.AlterarStatus(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                                                            enumTipoLancamentoIntegracao.Normal, _
                                                                                            enumStatusIntegracao.Disponível, , , False)
                End If
            Else
                Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
                objLancamentoCCContabil.Incluir xmlOperacao
            End If
            
    End Select

    '********************************************************************************************
    '  RATS 1008 - 02/08/2010 - Enviar o retorno de lancamento CC disponivel para o legado
    '********************************************************************************************
    If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
        Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
    End If

    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.Disponível
    
    If blnEnviaRetornoLegado Then
        Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
        Set objFluxoProcessamento = Nothing
    End If
    
    '********************************************************************************************
    
    Set objLancamentoCCContabil = Nothing
    Set xmlLancCCContabil = Nothing

    Exit Function
    
ErrorHandler:

    Set objFluxoProcessamento = Nothing
    Set objLancamentoCCContabil = Nothing
    Set xmlLancCCContabil = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Disponibilizar Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Linha =" & vntPasso)

End Function

'Processo de cancelamento de lançamento de conta corrente
Public Function Cancelar(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                         ByVal pblnCanceladoOrigem As Boolean, _
                Optional ByVal blnEnviaRetornoLegado As Boolean = True) As Boolean

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim xmlLancamentoCC                         As MSXML2.DOMDocument40
Dim strLancamentoCC                         As String

On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrente And _
        CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrenteComTributacao Then
        Exit Function
    End If

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    strLancamentoCC = objLancamentoCCContabil.Ler(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                  enumTipoLancamentoIntegracao.Normal)

    If strLancamentoCC = vbNullString Then
        Exit Function
    End If

    Set xmlLancamentoCC = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLancamentoCC.loadXML strLancamentoCC

    If pblnCanceladoOrigem = False Then
        If Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.Integrado Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.IntegradoCC Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.ErroEstornoCC Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.EnviadoCC Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.EnviadoCCAutomatico Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.EnviadoEstornoCC Or _
           Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) = enumStatusIntegracao.CanceladoOrigem Then
            ' 3054 - Operação com status de Conta Corrente inválido para cancelamento.
            lngCodigoErroNegocio = 3054
            GoTo ErrorHandler
        End If
    End If

    objLancamentoCCContabil.Estornar xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    '********************************************************************************************
    '  RATS 1008 - 02/08/2010 - Enviar o retorno de lancamento CC cancelado para o legado
    '********************************************************************************************
    If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
        Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
    End If

    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.Cancelado
    
    If blnEnviaRetornoLegado Then
        Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
        Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
        Set objFluxoProcessamento = Nothing
    End If
    
    '********************************************************************************************

    Set objLancamentoCCContabil = Nothing
    Set xmlLancamentoCC = Nothing

    Exit Function

ErrorHandler:
    
    Set objFluxoProcessamento = Nothing
    Set objLancamentoCCContabil = Nothing
    Set xmlLancamentoCC = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Cancelar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Inclusão de registro de historico conta corrente
Public Function IncluirCC(ByVal pstrSiglaSistema As String, _
                          ByVal plngCodigoEmpresa As Long, _
                          ByVal pvntSequenciaOperacao As Variant, _
                          ByVal plngTipoLancamentoIntegracao As Long, _
                          ByVal plngTipoEnvioRetorno As Long, _
                          ByVal pstrTextoXML As String, _
                 Optional ByVal pLngNrSequLancamento As Long = 0) As Boolean

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long

On Error GoTo ErrorHandler

    lngCodigoTxtXML = fgInsertVarchar4000(pstrTextoXML)

    strSQL = " Insert Into A8.TB_HIST_ENVI_INFO_CC (" & vbNewLine & _
             "  SG_SIST, " & vbNewLine & _
             "  CO_EMPR, " & vbNewLine & _
             "  NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "  TP_LANC_ITGR, " & vbNewLine & _
             "  TP_ENVI_RETN, " & vbNewLine & _
             "  DH_ENVI_CC, " & vbNewLine & _
             "  CO_TEXT_XML, " & vbNewLine & _
             "  NR_SEQU_LANC ) "

    strSQL = strSQL & vbNewLine & " Values (" & _
             "'" & pstrSiglaSistema & "', " & vbNewLine & _
             plngCodigoEmpresa & ", " & vbNewLine & _
             pvntSequenciaOperacao & ", " & vbNewLine & _
             plngTipoLancamentoIntegracao & ", " & vbNewLine & _
             plngTipoEnvioRetorno & ", " & vbNewLine & _
             fgDataHoraServidor_To_Oracle & ", " & vbNewLine & _
             lngCodigoTxtXML & ", " & vbNewLine & _
             pLngNrSequLancamento & ") "
    
    fgExecuteSQL strSQL

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IncluirCC Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Gerenciamento de integração com conta corrente ( DV e BG ) e contabil (HA)
Public Function Integrar(ByRef pvntSequenciaOperacao As Variant, _
                         ByVal plngIntegrarCCContabilidade As enumIntegrarCCContabilidade, _
                         ByVal pblnIgnoraGradeHorario As Boolean, _
                Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal, _
                Optional ByVal pstrJustificativa As String = vbNullString, _
                Optional ByVal pstrVaLancCC As String = vbNullString, _
                Optional ByVal pstrInLancDebCred As String = vbNullString, _
                Optional ByVal pstrNetOperacoes As String = vbNullString, _
                Optional ByVal pNumSeqLancamento As Long = 0, _
                Optional ByVal pblnIntegrarCCAutomatico As Boolean = False) As Boolean

Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim xmlLancamentoEstorno                    As MSXML2.DOMDocument40
Dim xmlLancamentoCCContabil                 As MSXML2.DOMDocument40
Dim objOperacao                             As A8LQS.clsOperacao
Dim objContabilidade                        As A8LQS.clsContabilidade
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objHoraLimIntSistLega                   As A8LQS.clsHoraLimIntSistLega
Dim objFluxoProcessamento                   As A8LQS.clsFluxoProcessamento
Dim intStatusIntegracao                     As enumStatusIntegracao

Dim strLancamentoCCContabil                 As String
Dim strLancamentoEstornoCC                  As String
Dim lngTipoBackOffice                       As Long
Dim lngDebitoCredito                        As Long
Dim vntPasso                                As Variant
Dim datDataOperacao                         As Date
Dim intCondicaoNet                          As Integer

    On Error GoTo ErrorHandler

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(pvntSequenciaOperacao)

    If xmlOperacao.xml = vbNullString Then
        ' 3005 - Operação inválida.
        lngCodigoErroNegocio = 3005
        GoTo ErrorHandler
    End If

    Set objContabilidade = CreateObject("A8LQS.clsContabilidade")
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
 
    strLancamentoCCContabil = objLancamentoCCContabil.Ler(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                          plngTipoLancamentoIntegracao, _
                                                          pNumSeqLancamento)
 
    If strLancamentoCCContabil = vbNullString Then
        ' 3005 - Operação inválida.
        lngCodigoErroNegocio = 3005
        GoTo ErrorHandler
    End If
    
    Set xmlLancamentoCCContabil = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLancamentoCCContabil.loadXML strLancamentoCCContabil
    
    If Not xmlLancamentoCCContabil.documentElement.selectSingleNode("//NR_SEQU_LANC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NR_SEQU_LANC", xmlLancamentoCCContabil.documentElement.selectSingleNode("NR_SEQU_LANC").Text
    Else
        fgAppendNode xmlOperacao, "MESG", "NR_SEQU_LANC", pNumSeqLancamento
    End If
    '************************************************************************************************************
    'Esta data é alterada pois qdo grava na tabela verifica o proximo dia util, e qdo envia utiliza a data do xmlOperacao
    'Carlos 05/07/04
    If Not xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
        xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV").Text = xmlLancamentoCCContabil.documentElement.selectSingleNode("//DT_OPER").Text
    Else
        If Not xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV_RETN") Is Nothing Then
            xmlOperacao.documentElement.selectSingleNode("DT_OPER_ATIV_RETN").Text = xmlLancamentoCCContabil.documentElement.selectSingleNode("//DT_OPER").Text
        Else
            xmlOperacao.documentElement.selectSingleNode("DT_MESG").Text = xmlLancamentoCCContabil.documentElement.selectSingleNode("//DT_OPER").Text
        End If
    End If
    '************************************************************************************************************

    If pblnIgnoraGradeHorario = False Then
        
        If ExcedeuHorarioLimiteTeimosinha(pvntSequenciaOperacao, plngTipoLancamentoIntegracao, pNumSeqLancamento) = True Then
            ' 3043 - Horário limite para Integração excedido.
            lngCodigoErroNegocio = 3043
            GoTo ErrorHandler
        End If
        
        Set objHoraLimIntSistLega = CreateObject("A8LQS.clsHoraLimIntSistLega")
        If Not objHoraLimIntSistLega.VerificarGradeHorario("BG", xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text) Then
            ' 3043 - Horário limite para Integração excedido.
            lngCodigoErroNegocio = 3043
            GoTo ErrorHandler
        End If
        Set objHoraLimIntSistLega = Nothing
    End If

    If Val(xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) = enumTipoAcao.EstornoEnviado Or _
       Val(xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text) = enumTipoAcao.EstornoSolicitado Then
        ' 3044 - Operação estornada, ou em processo de estorno.
        lngCodigoErroNegocio = 3044
        GoTo ErrorHandler
    End If
    
    Select Case plngIntegrarCCContabilidade
        Case enumIntegrarCCContabilidade.ContaCorrente
            
            If xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.Disponível And _
               xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.ErroCC And _
               xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.Antecipado And _
               xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.ErroEstornoCC And _
               xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.ErroSaldoCC And _
               plngTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal Then
                ' 3045 - Situação do registro para integração inválido.
                lngCodigoErroNegocio = 3045
                GoTo ErrorHandler
                
            ElseIf xmlLancamentoCCContabil.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.Disponível And _
                  plngTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Estorno Then
                  '3096 - Situação do registro para estorno inválido.
                  lngCodigoErroNegocio = 3096
                  GoTo ErrorHandler
            End If
            
            'Verifica o tipo de integracao de conta corrente
            If pblnIntegrarCCAutomatico = True Then
                intStatusIntegracao = enumStatusIntegracao.EnviadoCCAutomatico
            Else
                intStatusIntegracao = enumStatusIntegracao.EnviadoCC
            End If
            
            If pstrNetOperacoes <> vbNullString Then
                For intCondicaoNet = 1 To UBound(Split(pstrNetOperacoes, "|"))
                    Call objLancamentoCCContabil.AlterarStatus(Split(pstrNetOperacoes, "|")(intCondicaoNet), _
                                                               plngTipoLancamentoIntegracao, _
                                                               intStatusIntegracao, _
                                                               , , _
                                                               False, , , , _
                                                               pNumSeqLancamento)
                
                    If Not xmlOperacao.documentElement.selectSingleNode("//TP_MESG") Is Nothing Then
                        If Val(xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
                            Call objLancamentoCCContabil.IncluirConsolidacaoLanctoCC(pvntSequenciaOperacao, _
                                                                                     Split(pstrNetOperacoes, "|")(intCondicaoNet))
                        End If
                    End If
                Next
            Else
                
                objLancamentoCCContabil.AlterarStatus pvntSequenciaOperacao, _
                                                      plngTipoLancamentoIntegracao, _
                                                      intStatusIntegracao, _
                                                      , , _
                                                      False, , , , _
                                                      pNumSeqLancamento
            End If
                                                   
            'Cesar 07/05/2007 - Conta Corrente Corretoras
            If Val(xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
                
                xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = pstrVaLancCC
                If xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED") Is Nothing Then
                    Call fgAppendNode(xmlOperacao, "MESG", "IN_LANC_DEBT_CRED", pstrInLancDebCred)
                Else
                    xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED").Text = pstrInLancDebCred
                End If
           
            End If
            
            If pstrVaLancCC <> vbNullString Then
                xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = pstrVaLancCC
            End If
            
            'Busca tipo de backoffice
            lngTipoBackOffice = fgObterTipoBackOffice(xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                      xmlLancamentoCCContabil.documentElement.selectSingleNode("//SG_SIST").Text)
                            
            IntegrarBG xmlOperacao, plngTipoLancamentoIntegracao, pNumSeqLancamento, lngTipoBackOffice
            
        Case enumIntegrarCCContabilidade.Contabilidade
            
            vntPasso = 1
            If xmlLancamentoCCContabil.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text <> enumStatusIntegracao.IntegradoCC Then
                ' 3045 - Situação do registro para integração inválido.
                lngCodigoErroNegocio = 3045
                GoTo ErrorHandler
            End If
            
            vntPasso = 2
            
            If Val(xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
                
                xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = pstrVaLancCC
                If xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED") Is Nothing Then
                    Call fgAppendNode(xmlOperacao, "MESG", "IN_LANC_DEBT_CRED", pstrInLancDebCred)
                Else
                    xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED").Text = pstrInLancDebCred
                End If
           
            End If
                                        
            If Val(xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
                
                xmlOperacao.selectSingleNode("//CO_BANC").Text = xmlLancamentoCCContabil.selectSingleNode("//CO_BANC").Text
                xmlOperacao.selectSingleNode("//CO_AGEN").Text = xmlLancamentoCCContabil.selectSingleNode("//CO_AGEN").Text
                xmlOperacao.selectSingleNode("//NU_CC").Text = xmlLancamentoCCContabil.selectSingleNode("//NU_CC").Text
           
            End If
                                        
            objContabilidade.Integrar xmlOperacao, , pNumSeqLancamento
            
            If Val(xmlOperacao.documentElement.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.LancamentoContaCorrenteBG Then
                
                Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                   enumStatusOperacao.Liquidada, _
                                                   vbNullString, , , _
                                                   False)
                                                   
                xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liquidada
            
            End If
            
            '********************************************************************************************
            '  RATS 1008 - 02/08/2010 - Enviar o retorno de lancamento CC integrado para o legado
            '********************************************************************************************
            Set objFluxoProcessamento = CreateObject("A8LQS.clsFluxoProcessamento")
        
            If xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC") Is Nothing Then
                Call fgAppendNode(xmlOperacao, "MESG", "CO_ULTI_SITU_LANC_CC", vbNullString)
            End If
        
            xmlOperacao.selectSingleNode("//CO_ULTI_SITU_LANC_CC").Text = enumStatusIntegracao.Integrado
        
            Call objFluxoProcessamento.EnviarMensagemRetorno(xmlOperacao, False)
            '********************************************************************************************
            
            vntPasso = 3
            If pstrNetOperacoes <> vbNullString Then
                For intCondicaoNet = 1 To UBound(Split(pstrNetOperacoes, "|"))
                    objLancamentoCCContabil.AlterarStatus Split(pstrNetOperacoes, "|")(intCondicaoNet), _
                                                          plngTipoLancamentoIntegracao, _
                                                          enumStatusIntegracao.Integrado, _
                                                          , , _
                                                          False, , , , _
                                                          pNumSeqLancamento
                Next
            Else
                objLancamentoCCContabil.AlterarStatus pvntSequenciaOperacao, _
                                                      plngTipoLancamentoIntegracao, _
                                                      enumStatusIntegracao.Integrado, _
                                                      , , _
                                                      False, , , , _
                                                      pNumSeqLancamento
            End If
                                               
    End Select

    Set objLancamentoCCContabil = Nothing
    Set objFluxoProcessamento = Nothing
    Set objContabilidade = Nothing
    Set objOperacao = Nothing

    Exit Function
ErrorHandler:

    Set objContabilidade = Nothing
    Set objOperacao = Nothing
    Set objLancamentoCCContabil = Nothing
    Set objHoraLimIntSistLega = Nothing
    Set objFluxoProcessamento = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Integrar Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Passo: " & vntPasso)

End Function

'Monta e envia remessa DV
Public Function IntegrarDV(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                  Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal) As Boolean

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao

On Error GoTo ErrorHandler

    'Criar Tags para conversão no BUS
    'flMontarEnvioDV xmlOperacao, plngTipoLancamentoIntegracao
    
    '-------------------------------------------------
    'Pikachu - 01/07/2004
    'Para Liberar a versão com param gerais comentar a chamada 'flMontarEnvioDV xmlOperacao, plngTipoLancamentoIntegracao'
    'E descomentar a chamada abaixo 'Call flMontarEnvioDV_Param xmlOperacao, plngTipoLancamentoIntegracao'
    
    flMontarEnvioDV_Param xmlOperacao, plngTipoLancamentoIntegracao
    '-------------------------------------------------

    udtProtocolo.CodigoEmpresa = Format(CLng(xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("DV", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "1005"
    LSet udtProtocoloAux = udtProtocolo

    'Alterar o Indicador de débito/crédito no Lançamento em conta corrente
    

    'Incluir o histórico da mensagem enviada ao DV
    IncluirCC "DV", xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
              xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
              plngTipoLancamentoIntegracao, _
              enumTipoEnvioRetorno.Envio, _
              xmlOperacao.xml

    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.EnviarMensagemMQ udtProtocoloAux.String & xmlOperacao.xml, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing


    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IntegrarDV Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Monta e envia remessa BG
Public Function IntegrarBG(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                  Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal, _
                  Optional ByVal plngNumSequLancamento As Long = 0, _
                  Optional ByVal plngTipoBackOffice As Long) As Boolean

Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim lngDebitoCredito                        As Long
Dim strRemessaContaCorrente                 As String

On Error GoTo ErrorHandler

    'Criar Tags para conversão no BUS
    Call flMontarEnvioBG_Param(xmlOperacao, lngDebitoCredito, plngNumSequLancamento, plngTipoLancamentoIntegracao, plngTipoBackOffice)
    
    udtProtocolo.CodigoEmpresa = Format(CLng(xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text), "00000")
    udtProtocolo.SiglaSistemaDestino = fgCompletaString("BG", " ", 3, False)
    udtProtocolo.SiglaSistemaOrigem = fgCompletaString("A8", " ", 3, False)
    udtProtocolo.TipoMensagem = "000001007"
    
    LSet udtProtocoloAux = udtProtocolo

    'Incluir o histórico da mensagem enviada ao DV
    IncluirCC "BG", _
              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
              xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
              plngTipoLancamentoIntegracao, _
              0, _
              xmlOperacao.xml, plngNumSequLancamento
    
    strRemessaContaCorrente = udtProtocoloAux.String
    strRemessaContaCorrente = strRemessaContaCorrente & "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
    strRemessaContaCorrente = strRemessaContaCorrente & xmlOperacao.xml
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.EnviarMensagemMQ strRemessaContaCorrente, enumIdentificadorFila.BUS, True
    Set objProcessoOperacao = Nothing

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "IntegrarBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o numero de movimento (NUMMOV) original para estorno (BG)
Private Function flObterNumeroMovimentoBG(ByRef xmlOperacao As MSXML2.DOMDocument40)

Dim objHistSituacaoIntegracao               As A8LQS.clsHistSituacaoIntegracao
Dim strHistSituacaoIntegracao               As String
Dim xmlHistSituacaoIntegracao               As MSXML2.DOMDocument40
Dim vntNumMov                               As Variant
Dim strNumMov                               As String

On Error GoTo ErrorHandler:

    Set objHistSituacaoIntegracao = CreateObject("A8LQS.clsHistSituacaoIntegracao")
    
    strHistSituacaoIntegracao = objHistSituacaoIntegracao.LerTodos(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, 0, enumStatusIntegracao.IntegradoCC)
                                        
    If strHistSituacaoIntegracao <> "" Then
        
        Set xmlHistSituacaoIntegracao = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlHistSituacaoIntegracao.loadXML strHistSituacaoIntegracao
        
        strNumMov = xmlHistSituacaoIntegracao.selectSingleNode("//TX_MESG_ERRO").Text
        
        vntNumMov = Split(strNumMov, "|", , vbBinaryCompare)
        
        strNumMov = Trim(vntNumMov(0))
        strNumMov = Trim(Mid(strNumMov, 10, Len(strNumMov)))
       
        Set xmlHistSituacaoIntegracao = Nothing
    End If
    
    Set objHistSituacaoIntegracao = Nothing

    flObterNumeroMovimentoBG = strNumMov
        
    Exit Function
    
ErrorHandler:

    Set objHistSituacaoIntegracao = Nothing
    Set xmlHistSituacaoIntegracao = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterNumeroMovimentoBG Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Chamada para rotina de aprovação de alçada (LG)
Private Function flAprovarAlcada(ByVal pstrValor As String) As Boolean

Dim objPerfil                               As A6A7A8.clsPerfil
Dim vntProtovoloAlcada                      As Variant
    
On Error GoTo ErrorHandler

    'Digitação da Alçada
    Set objPerfil = CreateObject("A6A7A8.clsPerfil")
    
    Call objPerfil.AprovarAlcada(fgUsuarioRede, _
                                 pstrValor, _
                                 vntProtovoloAlcada)
                                     

    Set objPerfil = Nothing

    Exit Function
ErrorHandler:
    
    Set objPerfil = Nothing
    
    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAprovarAlcada", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'Disponibilizar registro para lançamento contabil (HA)
Public Function IntegrarContabil(ByRef pvntSequenciaOperacao As Variant, _
                                 ByVal plngIntegrarCCContabilidade As enumIntegrarCCContabilidade, _
                                 ByVal pblnIgnoraGradeHorario As Boolean, _
                        Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal, _
                        Optional ByVal pstrJustificativa As String = vbNullString, _
                        Optional ByVal pstrVaLancCC As String = vbNullString, _
                        Optional ByVal pstrInLancDebCred As String = vbNullString, _
                        Optional ByVal pstrNetOperacoes As String = vbNullString, _
                        Optional ByVal plngSequenciaLancamento As Long = 0) As Boolean
                

    On Error GoTo ErrorHandler
    
    Call Integrar(pvntSequenciaOperacao, _
                  plngIntegrarCCContabilidade, _
                  pblnIgnoraGradeHorario, _
                  plngTipoLancamentoIntegracao, _
                  pstrJustificativa, _
                  pstrVaLancCC, _
                  pstrInLancDebCred, _
                  pstrNetOperacoes, _
                  plngSequenciaLancamento)
    
    Exit Function
    
ErrorHandler:
    
    Call fgGravaArquivo("ERRO_INTEGRACAO_CONTABIL", Err.Description)
    
    Err.Clear

End Function

Private Function flObterHistoricoCC(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                    ByVal lngTipoBackOffice As Long, _
                                    ByVal plngDebitoCredito As Long) As String

Dim strParametroHistoricoCC                 As String
Dim strRemessaBG                            As String
Dim strFinalidadeTED                        As String
Dim strSubTipoAtivo                         As String
Dim strSubTipoAtivoOrig                     As String
Dim strProduto                              As String
Dim strProdutoOrig                          As String
Dim lngTipoCanalVenda                       As Long
Dim lngTipoCanalVendaOrig                   As Long
Dim lngFinalidadeTED                        As Long
Dim objParametroHistoricoCC                 As A8LQS.clsParametroHistoricoCC

On Error GoTo ErrorHandler
    
    '------------------------------------------------------------------------------------------------------
    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Or _
       xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CONTA_CORRENTE Then
       
        lngTipoCanalVenda = 0
        
        If Not xmlOperacao.documentElement.selectSingleNode("TP_CNAL_VEND") Is Nothing Then
            lngTipoCanalVenda = Val(xmlOperacao.documentElement.selectSingleNode("TP_CNAL_VEND").Text)
        Else
            lngTipoCanalVenda = 0
        End If
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
            strSubTipoAtivo = xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV").Text
        Else
            strSubTipoAtivo = "0"
        End If
        
        If Not xmlOperacao.documentElement.selectSingleNode("CD_SUB_PROD") Is Nothing Then
            strProduto = xmlOperacao.documentElement.selectSingleNode("CD_SUB_PROD").Text
        Else
            strProduto = "0"
        End If
        
        Set objParametroHistoricoCC = CreateObject("A8LQS.clsParametroHistoricoCC")
        
        strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                              plngDebitoCredito, _
                                                              strSubTipoAtivo, _
                                                              0, _
                                                              strProduto, _
                                                              lngTipoCanalVenda)
        
        'Salva os valores originais
        strSubTipoAtivoOrig = strSubTipoAtivo
        strProdutoOrig = strProduto
        lngTipoCanalVendaOrig = lngTipoCanalVenda
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = "0"
            strProduto = strProdutoOrig
            lngTipoCanalVenda = lngTipoCanalVendaOrig
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = strSubTipoAtivoOrig
            strProduto = "0"
            lngTipoCanalVenda = lngTipoCanalVendaOrig
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = strSubTipoAtivoOrig
            strProduto = strProdutoOrig
            lngTipoCanalVenda = 0
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = "0"
            strProduto = "0"
            lngTipoCanalVenda = lngTipoCanalVendaOrig
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = "0"
            strProduto = strProdutoOrig
            lngTipoCanalVenda = 0
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = strSubTipoAtivoOrig
            strProduto = "0"
            lngTipoCanalVenda = 0
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        
        End If
        
        If strParametroHistoricoCC = vbNullString Then
            strSubTipoAtivo = "0"
            strProduto = "0"
            lngTipoCanalVenda = 0
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  0, _
                                                                  strProduto, _
                                                                  lngTipoCanalVenda)
        
        End If
        
        Set objParametroHistoricoCC = Nothing
    
        If strParametroHistoricoCC = vbNullString Then
            ' 3042 - Código do histórico de Conta Corrente inválido.
            lngCodigoErroNegocio = 3042
            GoTo ErrorHandler
        End If
        
        GoTo Saida
    End If
    
    '------------------------------------------------------------------------------------------------------
    If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioTEDClientes Then
        
        lngFinalidadeTED = 0
        strSubTipoAtivo = "0"
        strProduto = "0"
        
        If Not xmlOperacao.documentElement.selectSingleNode("CO_FLND_TED") Is Nothing Then
            lngFinalidadeTED = Val("0" & xmlOperacao.documentElement.selectSingleNode("CO_FLND_TED").Text)
        Else
            lngFinalidadeTED = 0
        End If
        
        Set objParametroHistoricoCC = CreateObject("A8LQS.clsParametroHistoricoCC")
        strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                              plngDebitoCredito, _
                                                              strSubTipoAtivo, _
                                                              lngFinalidadeTED, _
                                                              strProduto)
        
        If strParametroHistoricoCC = vbNullString And _
           lngFinalidadeTED <> 0 Then
            
            lngFinalidadeTED = 0
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  plngDebitoCredito, _
                                                                  strSubTipoAtivo, _
                                                                  lngFinalidadeTED, _
                                                                  strProduto)
        
        End If
        Set objParametroHistoricoCC = Nothing
    
        If strParametroHistoricoCC = vbNullString Then
            ' 3042 - Código do histórico de Conta Corrente inválido.
            lngCodigoErroNegocio = 3042
            GoTo ErrorHandler
        End If
        
        GoTo Saida
        
    End If
    '------------------------------------------------------------------------------------------------------
    
    strSubTipoAtivo = "0"
    strProduto = "0"
    strFinalidadeTED = "0"

    Set objParametroHistoricoCC = CreateObject("A8LQS.clsParametroHistoricoCC")
    strParametroHistoricoCC = objParametroHistoricoCC.Ler("BG", _
                                                          xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                          lngTipoBackOffice, _
                                                          xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                          plngDebitoCredito, _
                                                          strSubTipoAtivo, _
                                                          strFinalidadeTED, _
                                                          strProduto)
    Set objParametroHistoricoCC = Nothing

    If strParametroHistoricoCC = vbNullString Then
        ' 3042 - Código do histórico de Conta Corrente inválido.
        lngCodigoErroNegocio = 3042
        GoTo ErrorHandler
    End If
    
Saida:
    flObterHistoricoCC = strParametroHistoricoCC

    Exit Function
    
ErrorHandler:
    
    Set objParametroHistoricoCC = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterHistoticoCC() Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montagem da remessa BG utilizando os parametros gerais
Private Function flMontarEnvioBG_Param(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                       ByRef plngDebitoCredito As Long, _
                              Optional ByVal plngNumSeqLancamento As Long = 0, _
                              Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal, _
                              Optional ByVal plngTipoBackOffice As Long) As String

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim objCadastroWorkflow                     As A8LQS.clsCadastroWorkflow
Dim xmlParametroIntegracaoCC                As MSXML2.DOMDocument

Dim xmlParametroHistoricoCC                 As MSXML2.DOMDocument40
Dim xmlRemessaBG                            As MSXML2.DOMDocument40
Dim xmlParametroGeral                       As MSXML2.DOMDocument40

Dim xmlLancCCCntb                           As MSXML2.DOMDocument40
Dim strLancCCCntb                           As String
Dim strParametroGeral                       As String
Dim strParametroHistoricoCC                 As String
Dim strRemessaBG                            As String
Dim strSubTipoAtivo                         As String
Dim strFinalidadeTED                        As String

Dim strImporte                              As String
Dim vntValorLancCC                          As Variant
Dim strCCC                                  As String
Dim strOperacao                             As String
Dim strCodigo                               As String
Dim strValorLancCC                          As String
Dim strDataOperacao                         As String
Dim strDataContabilizacao                   As String
Dim datDataOperacao                         As Date
Dim lngTipoBackOffice                       As Long

Dim strUsuario                              As String
Dim strCARTA                                As String
Dim strAUTORI                               As String
Dim strIMPERATIVA                           As String
Dim strDIVISA                               As String
Dim strCAJERO                               As String
Dim strAPLIC                                As String

Dim lngDebitoCreditoHist                    As Long

    On Error GoTo ErrorHandler

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    Set xmlLancCCCntb = CreateObject("MSXML2.DOMDocument.4.0")
    strLancCCCntb = objLancamentoCCContabil.Ler(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                                plngTipoLancamentoIntegracao, _
                                                plngNumSeqLancamento)
    
    xmlLancCCCntb.loadXML strLancCCCntb
    
    xmlOperacao.selectSingleNode("//CO_BANC").Text = xmlLancCCCntb.selectSingleNode("//CO_BANC").Text
    xmlOperacao.selectSingleNode("//CO_AGEN").Text = xmlLancCCCntb.selectSingleNode("//CO_AGEN").Text
    xmlOperacao.selectSingleNode("//NU_CC").Text = xmlLancCCCntb.selectSingleNode("//NU_CC").Text

    lngTipoBackOffice = fgObterTipoBackOffice(xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text)
    
    plngDebitoCredito = xmlLancCCCntb.selectSingleNode("//IN_LANC_DEBT_CRED").Text
    
    If Not xmlOperacao.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val(xmlOperacao.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.OperacoesCorretoras Then
            If Not xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED") Is Nothing Then
                plngDebitoCredito = Val(xmlOperacao.selectSingleNode("//IN_LANC_DEBT_CRED").Text)
            End If
        End If
    End If
    
    Set xmlLancCCCntb = Nothing
    
    If plngTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Estorno Then
        
        lngDebitoCreditoHist = IIf(plngDebitoCredito = enumTipoEntradaSaida.ENTRADA, _
                                                       enumTipoDebitoCreditoEstorno.EstornoCredito, _
                                                       enumTipoDebitoCreditoEstorno.EstornoDebito)
    
    Else
        lngDebitoCreditoHist = plngDebitoCredito
    End If
    
    strParametroHistoricoCC = flObterHistoricoCC(xmlOperacao, lngTipoBackOffice, lngDebitoCreditoHist)
    
    strParametroGeral = fgSelectVarchar4000(0, False)

    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")

    xmlParametroGeral.preserveWhiteSpace = True
    
    xmlParametroGeral.loadXML strParametroGeral

    strDataContabilizacao = vbNullString
    If Not xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO") Is Nothing Then
        If xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text <> vbNullString And _
           xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO/@OBRIG").Text = "S" Then
            strDataContabilizacao = Replace$(xmlParametroGeral.selectSingleNode("//DATA_CONTABILIZACAO").Text, "-", vbNullString)
            strDataContabilizacao = Replace$(strDataContabilizacao, "/", vbNullString)
            strDataContabilizacao = Right$(strDataContabilizacao, 4) & Mid$(strDataContabilizacao, 3, 2) & Left$(strDataContabilizacao, 2)
        End If
    End If
    
    If plngDebitoCredito = enumTipoDebitoCredito.Credito Then
        xmlParametroGeral.loadXML xmlParametroGeral.documentElement.selectSingleNode("//PARM_CC_BG_CRED").xml
    Else
        xmlParametroGeral.loadXML xmlParametroGeral.documentElement.selectSingleNode("//PARM_CC_BG_DEBT").xml
    End If
               
    '-----------------------------------------------------------------

    Set xmlRemessaBG = CreateObject("MSXML2.DOMDocument.4.0")
    
    strOperacao = xmlParametroGeral.selectSingleNode("//dse_operationName").Text
    
    Call fgAppendNode(xmlRemessaBG, "", "requestMsg", vbNullString)
    Call fgAppendNode(xmlRemessaBG, "requestMsg", "dse_operationName", strOperacao)
    Call fgAppendNode(xmlRemessaBG, "requestMsg", "dse_formattedData", vbNullString)
    Call fgAppendNode(xmlRemessaBG, "dse_formattedData", "kColl", "")
    
    Call fgAppendAttribute(xmlRemessaBG, "kColl", "id", "entryData")
    
    '-----------------------------------------------------------------------------------------------------
    
    strUsuario = xmlParametroGeral.selectSingleNode("//Usuario").Text
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "Usuario")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strUsuario)
    
    '-----------------------------------------------------------------------------------------------------
    strCCC = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_BANC").Text, "0", 4, True) & _
             fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) & _
             fgCompletaString(xmlOperacao.documentElement.selectSingleNode("NU_CC").Text, "0", 12, True)
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "CCC")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strCCC)
    
    '-----------------------------------------------------------------------------------------------------
    
    If Not xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC") Is Nothing Then
        xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text = Replace$(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text, "-", vbNullString)
        vntValorLancCC = Split(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text, ",", , vbBinaryCompare)
    Else
        vntValorLancCC = Split("0,0", ",", , vbBinaryCompare)
    End If
        
    If UBound(vntValorLancCC) = 0 Then
        strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 13, True) & "00"
    Else
        strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 13, True) & fgCompletaString(vntValorLancCC(1), "0", 2, False)
    End If
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "IMPORTE")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strValorLancCC)
    
    '-----------------------------------------------------------------------------------------------------
    
    Set xmlParametroHistoricoCC = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametroHistoricoCC.loadXML strParametroHistoricoCC
    strCodigo = fgCompletaString(xmlParametroHistoricoCC.documentElement.selectSingleNode("//CO_HIST_CC").Text, "0", 4, True)
    Set xmlParametroHistoricoCC = Nothing
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "CODIGO")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strCodigo)
    
    '-----------------------------------------------------------------------------------------------------
        
    If xmlParametroGeral.selectSingleNode("//CARTA/@OBRIG").Text = "S" Then
        
        strCARTA = IIf(Trim(xmlParametroGeral.selectSingleNode("//CARTA").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CARTA").Text)
        
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "CARTA")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strCARTA)
    
    End If
    
    '-----------------------------------------------------------------------------------------------------
    
    If plngDebitoCredito = enumTipoDebitoCredito.Debito Then
    
        'FREITAS - 08/07  /  SSS### - Incondicional ao Saldo - Branco é condicional
        If Not xmlOperacao.documentElement.selectSingleNode("TP_OPER") Is Nothing Then
            Set objCadastroWorkflow = CreateObject("A8LQS.clsCadastroWorkflow")
            Set xmlParametroIntegracaoCC = New MSXML2.DOMDocument
            xmlParametroIntegracaoCC.loadXML objCadastroWorkflow.ObterParamCondicionalSaldo(xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                                            enumFuncaoSistema.IntegracaoCC, _
                                                                                            xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                                            plngTipoBackOffice)
            
            If Not xmlParametroIntegracaoCC.documentElement.selectSingleNode("TP_COND_SALD") Is Nothing Then
                Select Case xmlParametroIntegracaoCC.documentElement.selectSingleNode("TP_COND_SALD").Text
                    Case enumTipoIncondicionalSaldoCC.Sim
                        strAUTORI = "SSS   "
                    Case enumTipoIncondicionalSaldoCC.Nao
                        strAUTORI = vbNullString
                        
                End Select
            End If
        End If

        If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
            If Val("0" & xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text) = enumCanalDeVenda.SGC Then
                strAUTORI = vbNullString
            End If
        End If

        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "AUTORI")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strAUTORI)
        
        Set objCadastroWorkflow = Nothing
        Set xmlParametroIntegracaoCC = Nothing
        
'        If xmlParametroGeral.selectSingleNode("//AUTORI/@OBRIG").Text = "S" Then
'
'            strAUTORI = IIf(Trim(xmlParametroGeral.selectSingleNode("//AUTORI").Text) = "", "SSS   ", xmlParametroGeral.selectSingleNode("//AUTORI").Text)
'
'            'KIDA - 18/02/2008
'            If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
'                If Val("0" & xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text) = enumCanalDeVenda.SGC Then
'                    strAUTORI = vbNullString
'                End If
'            End If
'
'            Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
'            Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "AUTORI")
'            Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strAUTORI)
'
'        End If
        
        '-----------------------------------------------------------------------------------------------------
        
        If xmlParametroGeral.selectSingleNode("//IMPERATIVA/@OBRIG").Text = "S" Then
            
            strIMPERATIVA = IIf(Trim(xmlParametroGeral.selectSingleNode("//IMPERATIVA").Text) = "", "N", xmlParametroGeral.selectSingleNode("//IMPERATIVA").Text)
            
            'KIDA - 18/02/2008
            If Not xmlOperacao.selectSingleNode("//TP_CNAL_VEND") Is Nothing Then
                If Val("0" & xmlOperacao.selectSingleNode("//TP_CNAL_VEND").Text) = enumCanalDeVenda.SGC Then
                    strIMPERATIVA = "S"
                End If
            End If
        
            Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
            Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "IMPERATIVA")
            Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strIMPERATIVA)
        
        End If
        
    End If
   
    '-----------------------------------------------------------------------------------------------------
    If xmlParametroGeral.selectSingleNode("//DIVISA/@OBRIG").Text = "S" Then
        
        strDIVISA = IIf(Trim(xmlParametroGeral.selectSingleNode("//DIVISA").Text) = "", "BRL", xmlParametroGeral.selectSingleNode("//DIVISA").Text)
    
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "DIVISA")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strDIVISA)
    
    End If
    
    '-----------------------------------------------------------------------------------------------------
    
    If Not xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
        datDataOperacao = fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text)
    Else
        datDataOperacao = fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV_RETN").Text)
    End If

    If strDataContabilizacao <> vbNullString Then
        datDataOperacao = fgDtXML_To_Date(strDataContabilizacao)
    End If

    If datDataOperacao < fgDataHoraServidor(DataAux) Then
        
        strDataOperacao = Format$(datDataOperacao, "yyyy-MM-dd")
        
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "FECCONT")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strDataOperacao)
        
    End If
    
    '-----------------------------------------------------------------------------------------------------
    
    If xmlParametroGeral.selectSingleNode("//CAJERO/@OBRIG").Text = "S" Then
        
        strCAJERO = IIf(Trim(xmlParametroGeral.selectSingleNode("//CAJERO").Text) = "", "A", xmlParametroGeral.selectSingleNode("//CAJERO").Text)
        
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "CAJERO")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strCAJERO)
    
    End If
    
    '-----------------------------------------------------------------------------------------------------
    
    If xmlParametroGeral.selectSingleNode("//APLIC/@OBRIG").Text = "S" Then
        
        strAPLIC = IIf(Trim(xmlParametroGeral.selectSingleNode("//APLIC").Text) = "", "A8", xmlParametroGeral.selectSingleNode("//APLIC").Text)
    
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "APLIC")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strAPLIC)
    
    End If
        
    '-----------------------------------------------------------------------------------------------------
    
    If Not xmlOperacao.selectSingleNode("//MESG") Is Nothing Then
        If xmlOperacao.selectSingleNode("//TP_LANC_ITGR") Is Nothing Then
            Call fgAppendNode(xmlOperacao, "MESG", "TP_LANC_ITGR", plngTipoLancamentoIntegracao)
        End If
    End If
    
    Call fgAppendXML(xmlOperacao, "MESG", xmlRemessaBG.xml)
    
    Set xmlRemessaBG = Nothing
    Set objLancamentoCCContabil = Nothing
    Set xmlParametroGeral = Nothing
    
    Exit Function
    
ErrorHandler:
        
    Set objLancamentoCCContabil = Nothing
    Set xmlParametroHistoricoCC = Nothing
    Set xmlParametroGeral = Nothing
    Set xmlLancCCCntb = Nothing
    Set xmlRemessaBG = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarEnvioBG_Param Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Montagem da remessa DV utilizando os parametros gerais
Private Function flMontarEnvioDV_Param(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                              Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal) As Boolean

Dim objParametroHistoricoCC                 As A8LQS.clsParametroHistoricoCC
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim xmlParametroHistoricoCC                 As MSXML2.DOMDocument40
Dim strParametroHistoricoCC                 As String
Dim lngTipoBackOffice                       As Long
Dim lngDebitoCredito                        As Long
Dim lngTipoMovimento                        As Long
Dim datDataOperacao                         As Date
Dim strSubTipoAtivo                         As String

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim strParametroGeral                       As String

Dim xmlLancCCCntb                           As MSXML2.DOMDocument40
Dim strLancCCCntb                           As String

On Error GoTo ErrorHandler

    'Converter o Débito/Crédito para o Conta Corrente
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
        
    'Pikachu - 26/10/2004
    'Obter dados da conta corrente da tabela A8.TB_LANC_CC_CNTB
    Set xmlLancCCCntb = CreateObject("MSXML2.DOMDocument.4.0")
    strLancCCCntb = objLancamentoCCContabil.Ler(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               plngTipoLancamentoIntegracao)
    
    xmlLancCCCntb.loadXML strLancCCCntb
    
    xmlOperacao.selectSingleNode("//CO_AGEN").Text = xmlLancCCCntb.selectSingleNode("//CO_AGEN").Text
    xmlOperacao.selectSingleNode("//NU_CC").Text = xmlLancCCCntb.selectSingleNode("//NU_CC").Text
    
    lngTipoBackOffice = 0
    lngDebitoCredito = objLancamentoCCContabil.ObterTipoDebitoCreditoCC(xmlOperacao.documentElement.selectSingleNode("//CO_VEIC_LEGA").Text, _
                                                                        xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                                        xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text, _
                                                                        lngTipoBackOffice, _
                                                                        xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
                                                                        
    lngDebitoCredito = xmlLancCCCntb.selectSingleNode("//IN_LANC_DEBT_CRED").Text
    
    Set xmlLancCCCntb = Nothing
    '----------------------------------------------------
                                                                        
    If plngTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal Then
        If lngDebitoCredito = enumTipoDebitoCredito.Credito Then
            If xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "IN_OPER_DEBT_CRED_CC", "C"
            Else
                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC").Text = "C"
            End If
            lngTipoMovimento = enumTipoDebitoCredito.Credito
        Else
            If xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "IN_OPER_DEBT_CRED_CC", "D"
            Else
                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC").Text = "D"
            End If
            lngTipoMovimento = enumTipoDebitoCredito.Debito
        End If
    Else
        If lngDebitoCredito = enumTipoDebitoCredito.Credito Then
            If xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "IN_OPER_DEBT_CRED_CC", "D"
            Else
                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC").Text = "D"
            End If
            lngTipoMovimento = enumTipoDebitoCreditoEstorno.EstornoCredito
        Else
            If xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC") Is Nothing Then
                fgAppendNode xmlOperacao, "MESG", "IN_OPER_DEBT_CRED_CC", "C"
            Else
                xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED_CC").Text = "C"
            End If
            lngTipoMovimento = enumTipoDebitoCreditoEstorno.EstornoDebito
        End If
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
    
        If Not xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV") Is Nothing Then
            strSubTipoAtivo = xmlOperacao.documentElement.selectSingleNode("CO_SUB_TIPO_ATIV").Text
        Else
            strSubTipoAtivo = "0"
        End If
        
        'sE NÃO EXISTIR CONTA CC COM SUB_TIPO_ATIVO , BUSCAR COM SUB TIPO ATIVO =0
        Set objParametroHistoricoCC = CreateObject("A8LQS.clsParametroHistoricoCC")
        strParametroHistoricoCC = objParametroHistoricoCC.Ler("DV", _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                              lngTipoMovimento, _
                                                              strSubTipoAtivo)
        
        If strParametroHistoricoCC = vbNullString And strSubTipoAtivo <> "0" Then
            strSubTipoAtivo = "0"
            strParametroHistoricoCC = objParametroHistoricoCC.Ler("DV", _
                                                                  xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                                  lngTipoBackOffice, _
                                                                  xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                                  lngTipoMovimento, _
                                                                  strSubTipoAtivo)
        
        End If
        Set objParametroHistoricoCC = Nothing
    
        If strParametroHistoricoCC = vbNullString Then
            ' 3042 - Código do histórico de Conta Corrente inválido.
            lngCodigoErroNegocio = 3042
            GoTo ErrorHandler
        End If
    
    
    Else
        strSubTipoAtivo = "0"
    
        Set objParametroHistoricoCC = CreateObject("A8LQS.clsParametroHistoricoCC")
        strParametroHistoricoCC = objParametroHistoricoCC.Ler("DV", _
                                                              xmlOperacao.documentElement.selectSingleNode("CO_EMPR").Text, _
                                                              lngTipoBackOffice, _
                                                              xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text, _
                                                              lngTipoMovimento, _
                                                              strSubTipoAtivo)
        Set objParametroHistoricoCC = Nothing
    
        If strParametroHistoricoCC = vbNullString Then
            ' 3042 - Código do histórico de Conta Corrente inválido.
            lngCodigoErroNegocio = 3042
            GoTo ErrorHandler
        End If
    
    End If
    
    '------------------------------------

    Set xmlParametroHistoricoCC = CreateObject("MSXML2.DOMDocument.4.0")
    xmlParametroHistoricoCC.loadXML strParametroHistoricoCC

    'Para atender a Despesas Selic que não tem a TAG - Carlos 09/06/2004
    If xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_COMD_OPER", "000001"
    Else
        If Not IsNumeric(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text) Then
            xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text = "000001"
        Else
            'PIKACHU - 20/07/2004
            'VERIFICAR - ACONTECEU EM PRODUCAO
            'PARA EVEVENTOS - PAGAMENTO DE JUROS O NUMERO DO COMANDO CHEGOU ZERADO '000000'
            If Val(xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text) = 0 Then
                xmlOperacao.documentElement.selectSingleNode("NU_COMD_OPER").Text = "000001"
            End If
        End If
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_HIST_DEST") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_HIST_DEST", xmlParametroHistoricoCC.documentElement.selectSingleNode("//CO_HIST_CC").Text
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_HIST_DEST").Text = xmlParametroHistoricoCC.documentElement.selectSingleNode("//CO_HIST_CC").Text
    End If

    Set xmlParametroHistoricoCC = Nothing

    If xmlOperacao.documentElement.selectSingleNode("NU_CTRL_ITGR") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NU_CTRL_ITGR", xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    Else
        xmlOperacao.documentElement.selectSingleNode("NU_CTRL_ITGR").Text = xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text
    End If
    
    'Verificar data
    If xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "DT_OPER_ATIV", xmlOperacao.documentElement.selectSingleNode("//DT_MESG").Text
        datDataOperacao = fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_MESG").Text)
    Else
        datDataOperacao = fgDtXML_To_Date(xmlOperacao.documentElement.selectSingleNode("//DT_OPER_ATIV").Text)
    End If

    '-----------------------------------------------------------------
    'Pikachu - 22/06/2004
    'Obter parametros gerais
    
    strParametroGeral = fgSelectVarchar4000(0, False)

    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlParametroGeral.preserveWhiteSpace = True

    xmlParametroGeral.loadXML strParametroGeral

    xmlParametroGeral.loadXML xmlParametroGeral.documentElement.selectSingleNode("//PARM_CC_DV").xml
    '-----------------------------------------------------------------
    
    'Identificador do programa de origem
    If xmlOperacao.documentElement.selectSingleNode("ID_PGM_ORIG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ID_PGM_ORIG", IIf(Trim(xmlParametroGeral.selectSingleNode("//ID_PGM_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//ID_PGM_ORIG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("ID_PGM_ORIG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//ID_PGM_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//ID_PGM_ORIG").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("ID_SIST_ORIG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "ID_SIST_ORIG", IIf(Trim(xmlParametroGeral.selectSingleNode("//ID_SIST_ORIG").Text) = "", "A8", xmlParametroGeral.selectSingleNode("//ID_SIST_ORIG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("ID_SIST_ORIG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//ID_SIST_ORIG").Text) = "", "A8", xmlParametroGeral.selectSingleNode("//ID_SIST_ORIG").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CO_USUA_CADR_OPER") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CO_USUA_CADR_OPER", fgUsuarioRede
    Else
        xmlOperacao.documentElement.selectSingleNode("CO_USUA_CADR_OPER").Text = fgUsuarioRede
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_FUNCAO") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_FUNCAO", enumTipoFuncaoDV.ConsisteAtualizaMovimento
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_FUNCAO").Text = enumTipoFuncaoDV.ConsisteAtualizaMovimento
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_PVENDA_ORIG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_PVENDA_ORIG", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_PVENDA_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_PVENDA_ORIG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_PVENDA_ORIG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_PVENDA_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_PVENDA_ORIG").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_IDENT_CCS_S") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_IDENT_CCS_S", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS_S").Text) = "", " ", xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS_S").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_IDENT_CCS_S").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS_S").Text) = "", " ", xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS_S").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_IDENT_CCS") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_IDENT_CCS", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_IDENT_CCS").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_IDENT_CCS").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("NSU_MOV_ORIG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NSU_MOV_ORIG", IIf(Trim(xmlParametroGeral.selectSingleNode("//NSU_MOV_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NSU_MOV_ORIG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("NSU_MOV_ORIG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//NSU_MOV_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NSU_MOV_ORIG").Text)
    End If
    
    'Se for Retroativo TP_PROC = 2
    
    If datDataOperacao < fgDataHoraServidor(DataAux) Then
        
        If xmlOperacao.documentElement.selectSingleNode("TP_PROC") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "TP_PROC", "2"
        Else
            xmlOperacao.documentElement.selectSingleNode("TP_PROC").Text = "2"
        End If
        
    Else
    
        If xmlOperacao.documentElement.selectSingleNode("TP_PROC") Is Nothing Then
            fgAppendNode xmlOperacao, "MESG", "TP_PROC", "1"
        Else
            xmlOperacao.documentElement.selectSingleNode("TP_PROC").Text = "1"
        End If
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_HIST_ORIG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_HIST_ORIG", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_HIST_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_HIST_ORIG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_HIST_ORIG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_HIST_ORIG").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_HIST_ORIG").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_IDENT_EVC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_IDENT_EVC", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_EVC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_IDENT_EVC").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_IDENT_EVC").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_IDENT_EVC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_IDENT_EVC").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CD_SUB_EVC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CD_SUB_EVC", IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_SUB_EVC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_SUB_EVC").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CD_SUB_EVC").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CD_SUB_EVC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CD_SUB_EVC").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ1") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ1", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ1").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ2") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ2", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ2").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ3") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ3", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ3").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ4") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ4", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ4").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ5") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ5", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ5").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ6") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ6", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ6").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ7") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ7", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ7").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ8") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_CHQ_BLOQ8", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_CHQ_BLOQ8").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_AGEN") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_AGEN", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_AGEN").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_AGEN").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_AGEN").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_AGEN").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_AGEN").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_CL") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_CTA_CL", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_CL").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_CL").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_CL").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_CL").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_CL").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_ENC") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_CTA_ENC", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_ENC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_ENC").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_ENC").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_ENC").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_ENC").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_SLD") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_CTA_SLD", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_SLD").Text) = "", "1", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_SLD").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_SLD").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_SLD").Text) = "", "1", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_SLD").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_SUST_RED") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_SUST_RED", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_SUST_RED").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_SUST_RED").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_SUST_RED").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_SUST_RED").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_SUST_RED").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUTA_SLD_AG") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUTA_SLD_AG", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUTA_SLD_AG").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_AUTA_SLD_AG").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUTA_SLD_AG").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUTA_SLD_AG").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_AUTA_SLD_AG").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUTA_BATCH") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUTA_BATCH", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUTA_BATCH").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_AUTA_BATCH").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUTA_BATCH").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUTA_BATCH").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_AUTA_BATCH").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_CHQ_PAGTO") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_CHQ_PAGTO", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CHQ_PAGTO").Text) = "", "N", xmlParametroGeral.selectSingleNode("//IN_CHQ_PAGTO").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_CHQ_PAGTO").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CHQ_PAGTO").Text) = "", "N", xmlParametroGeral.selectSingleNode("//IN_CHQ_PAGTO").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("NR_SEQ_NSU") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NR_SEQ_NSU", IIf(Trim(xmlParametroGeral.selectSingleNode("//NR_SEQ_NSU").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NR_SEQ_NSU").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("NR_SEQ_NSU").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//NR_SEQ_NSU").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NR_SEQ_NSU").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("DT_AGENDA") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "DT_AGENDA", IIf(Trim(xmlParametroGeral.selectSingleNode("//DT_AGENDA").Text) = "", "0", xmlParametroGeral.selectSingleNode("//DT_AGENDA").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("DT_AGENDA").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//DT_AGENDA").Text) = "", "0", xmlParametroGeral.selectSingleNode("//DT_AGENDA").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("DADOS_CARTAO") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "DADOS_CARTAO", IIf(Trim(xmlParametroGeral.selectSingleNode("//DADOS_CARTAO").Text) = "", "0", xmlParametroGeral.selectSingleNode("//DADOS_CARTAO").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("DADOS_CARTAO").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//DADOS_CARTAO").Text) = "", "0", xmlParametroGeral.selectSingleNode("//DADOS_CARTAO").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_TRAN_CT") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_TRAN_CT", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_TRAN_CT").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_TRAN_CT").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_TRAN_CT").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_TRAN_CT").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_TRAN_CT").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("VL_LIM_DIA") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "VL_LIM_DIA", "0"
    Else
        xmlOperacao.documentElement.selectSingleNode("VL_LIM_DIA").Text = "0"
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_CPMF_TRANSF") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_CPMF_TRANSF", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CPMF_TRANSF").Text) = "", "N", xmlParametroGeral.selectSingleNode("//IN_CPMF_TRANSF").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_CPMF_TRANSF").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CPMF_TRANSF").Text) = "", "N", xmlParametroGeral.selectSingleNode("//IN_CPMF_TRANSF").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_ORIG_MOV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_ORIG_MOV", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_ORIG_MOV").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_ORIG_MOV").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_ORIG_MOV").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_ORIG_MOV").Text) = "", "S", xmlParametroGeral.selectSingleNode("//IN_ORIG_MOV").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_TIP_MOV") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_TIP_MOV", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_TIP_MOV").Text) = "", "C", xmlParametroGeral.selectSingleNode("//IN_TIP_MOV").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_TIP_MOV").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_TIP_MOV").Text) = "", "C", xmlParametroGeral.selectSingleNode("//IN_TIP_MOV").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_CTA_INT") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_CTA_INT", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CTA_INT").Text) = "", " ", xmlParametroGeral.selectSingleNode("//IN_CTA_INT").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_CTA_INT").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_CTA_INT").Text) = "", " ", xmlParametroGeral.selectSingleNode("//IN_CTA_INT").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("MOSTRA_CI") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "MOSTRA_CI", IIf(Trim(xmlParametroGeral.selectSingleNode("//MOSTRA_CI").Text) = "", " ", xmlParametroGeral.selectSingleNode("//MOSTRA_CI").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("MOSTRA_CI").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//MOSTRA_CI").Text) = "", " ", xmlParametroGeral.selectSingleNode("//MOSTRA_CI").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("NSU_ESTORNO") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "NSU_ESTORNO", IIf(Trim(xmlParametroGeral.selectSingleNode("//NSU_ESTORNO").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NSU_ESTORNO").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("NSU_ESTORNO").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//NSU_ESTORNO").Text) = "", "0", xmlParametroGeral.selectSingleNode("//NSU_ESTORNO").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_BLQ") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_AUT_CTA_BLQ", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_BLQ").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_BLQ").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_AUT_CTA_BLQ").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_BLQ").Text) = "", "0", xmlParametroGeral.selectSingleNode("//IN_AUT_CTA_BLQ").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("IN_COMPL_HIST") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "IN_COMPL_HIST", IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_COMPL_HIST").Text) = "", " ", xmlParametroGeral.selectSingleNode("//IN_COMPL_HIST").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("IN_COMPL_HIST").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//IN_COMPL_HIST").Text) = "", " ", xmlParametroGeral.selectSingleNode("//IN_COMPL_HIST").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("CO_TP_COMPL_HIST") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "CO_TP_COMPL_HIST", IIf(Trim(xmlParametroGeral.selectSingleNode("//CO_TP_COMPL_HIST").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CO_TP_COMPL_HIST").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("CO_TP_COMPL_HIST").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//CO_TP_COMPL_HIST").Text) = "", "0", xmlParametroGeral.selectSingleNode("//CO_TP_COMPL_HIST").Text)
    End If

    If xmlOperacao.documentElement.selectSingleNode("TX_COMPL_HIST") Is Nothing Then
        fgAppendNode xmlOperacao, "MESG", "TX_COMPL_HIST", IIf(Trim(xmlParametroGeral.selectSingleNode("//TX_COMPL_HIST").Text) = "", " ", xmlParametroGeral.selectSingleNode("//TX_COMPL_HIST").Text)
    Else
        xmlOperacao.documentElement.selectSingleNode("TX_COMPL_HIST").Text = IIf(Trim(xmlParametroGeral.selectSingleNode("//TX_COMPL_HIST").Text) = "", " ", xmlParametroGeral.selectSingleNode("//TX_COMPL_HIST").Text)
    End If

    Set xmlParametroGeral = Nothing
    Set objLancamentoCCContabil = Nothing
 
    Exit Function
    
ErrorHandler:

    Set xmlParametroGeral = Nothing
    Set objLancamentoCCContabil = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarEnvioDV_Param Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Montagem da remessa de estorno BG utilizando os parametros gerais
Private Function flMontarEnvioEstornoBG_Parm(ByRef xmlOperacao As MSXML2.DOMDocument40, _
                                             ByRef plngDebitoCredito As Long, _
                                             Optional ByVal plngNumSequLancamento As Long = 0) As String

Dim objParametroHistoricoCC                 As A8LQS.clsParametroHistoricoCC
Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim xmlParametroHistoricoCC                 As MSXML2.DOMDocument40
Dim xmlRemessaBG                            As MSXML2.DOMDocument40

Dim lngTipoBackOffice                       As Long
Dim strParametroHistoricoCC                 As String
Dim strCCC                                  As String
Dim strOperacao                             As String
Dim strCodigo                               As String
Dim strRemessaBG                            As String
Dim strImporte                              As String
Dim vntValorLancCC                          As Variant
Dim strValorLancCC                          As String
Dim strNIO                                  As String
Dim strNumeroMovimento                      As String
Dim lngTipoMovimento                        As Long

Dim xmlParametroGeral                       As MSXML2.DOMDocument40
Dim strParametroGeral                       As String

Dim strUsuario                              As String

On Error GoTo ErrorHandler
   
    'Converter o Débito/Crédito para o Conta Corrente
    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    
    lngTipoBackOffice = 0
    plngDebitoCredito = objLancamentoCCContabil.ObterTipoDebitoCreditoCC(xmlOperacao.documentElement.selectSingleNode("CO_VEIC_LEGA").Text, _
                                                                         xmlOperacao.documentElement.selectSingleNode("SG_SIST_ORIG").Text, _
                                                                         xmlOperacao.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text, _
                                                                         lngTipoBackOffice, _
                                                                         xmlOperacao.documentElement.selectSingleNode("TP_OPER").Text)
   
    Set objLancamentoCCContabil = Nothing
   
   
    If Not xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC") Is Nothing Then
        vntValorLancCC = Split(xmlOperacao.documentElement.selectSingleNode("VA_LANC_CC").Text, ",", , vbBinaryCompare)
    Else
        vntValorLancCC = Split("0,0", ",", , vbBinaryCompare)
    End If
        
    If plngDebitoCredito = enumTipoDebitoCredito.Credito Then
        
        If UBound(vntValorLancCC) = 0 Then
            strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 12, True) & "00" & "-"
        Else
            strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 12, True) & fgCompletaString(vntValorLancCC(1), "0", 2, False) & "-"
        End If
    Else
        If UBound(vntValorLancCC) = 0 Then
            strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 13, True) & "00"
        Else
            strValorLancCC = fgCompletaString(vntValorLancCC(0), "0", 13, True) & fgCompletaString(vntValorLancCC(1), "0", 2, False)
        End If
    End If
    
    strCCC = fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_BANC").Text, "0", 4, True) & _
             fgCompletaString(xmlOperacao.documentElement.selectSingleNode("CO_AGEN").Text, "0", 4, True) & _
             fgCompletaString(xmlOperacao.documentElement.selectSingleNode("NU_CC").Text, "0", 12, True)

    Set xmlRemessaBG = CreateObject("MSXML2.DOMDocument.4.0")
    
    '-----------------------------------------------------------------
    'Pikachu - 22/06/2004
    'Obter parametros gerais
    
    strParametroGeral = fgSelectVarchar4000(0, False)

    Set xmlParametroGeral = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlParametroGeral.preserveWhiteSpace = True

    xmlParametroGeral.loadXML strParametroGeral

    xmlParametroGeral.loadXML xmlParametroGeral.documentElement.selectSingleNode("//PARM_CC_BG_ESTO").xml
               
    strOperacao = xmlParametroGeral.selectSingleNode("//dse_operationName").Text
    
    Call fgAppendNode(xmlRemessaBG, "", "requestMsg", vbNullString)
    Call fgAppendNode(xmlRemessaBG, "requestMsg", "dse_operationName", strOperacao)
    Call fgAppendNode(xmlRemessaBG, "requestMsg", "dse_formattedData", vbNullString)
    Call fgAppendNode(xmlRemessaBG, "dse_formattedData", "kColl", "")
        
    Call fgAppendAttribute(xmlRemessaBG, "kColl", "id", "entryData")
    
    '-----------------------------------------------------------------
    
    strUsuario = xmlParametroGeral.selectSingleNode("//Usuario").Text
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "Usuario")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strUsuario)
    
    '-----------------------------------------------------------------
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "CCC")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strCCC)
    
    '-----------------------------------------------------------------
    
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "IMPORTE")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strValorLancCC)
    
    '-----------------------------------------------------------------
    
    'Alfa(24)
    'NIO - Número de Identificação da Operação
    'Esse campo será preenchido da seguinte forma:
    'Codigo Aplicação                             (Alfa 2)
    'Usuário que executou a integração - Terminal (Alfa 4)
    'Data no formato AAAA-MM-DD                   (Alfa 10)
    'Hora no formato HH.MM.SS                     (Alfa 8)

'    strNIO = "A8"
'    strNIO = strNIO & Mid(fgUsuarioRede, 1, 4)
'    strNIO = strNIO & Format(fgDataHoraServidor(DataAux), "yyyymmdd")
'    strNIO = strNIO & Format(fgDataHoraServidor(HoraAux), "hhmmss")
'
'    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
'    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NIO")
'    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strNIO)

'    NIO não pode ser enviado no estorno
'    Comunicação Fabiana - 16/04/2008
'    Alteração Cas.
     '-----------------------------------------------------------------
    
    'Num(9)
    strNumeroMovimento = flObterNumeroMovimentoBG(xmlOperacao)
    strNumeroMovimento = fgCompletaString(strNumeroMovimento, "0", 9, True)
    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NUMMOV")
    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strNumeroMovimento)
        
    '-----------------------------------------------------------------
        
'    Node repetido
'    Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
'    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NUMMOV")
'    Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", strNumeroMovimento)
    
    '-----------------------------------------------------------------
            
    If xmlParametroGeral.selectSingleNode("//DIVISA/@OBRIG").Text = "S" Then
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "DIVISA")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", xmlParametroGeral.selectSingleNode("//DIVISA").Text)
    End If
    '-----------------------------------------------------------------
        
    If xmlParametroGeral.selectSingleNode("//NOM-TITULAR1/@OBRIG").Text = "S" Then
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NOM-TITULAR1")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", xmlParametroGeral.selectSingleNode("//NOM-TITULAR1").Text)
    End If
    
    '-----------------------------------------------------------------
        
    If xmlParametroGeral.selectSingleNode("//NOM-APE1-TITULAR1/@OBRIG").Text = "S" Then
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NOM-APE1-TITULAR1")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", xmlParametroGeral.selectSingleNode("//NOM-APE1-TITULAR1").Text)
    End If
    '-----------------------------------------------------------------
        
    If xmlParametroGeral.selectSingleNode("//NOM-APE2-TITULAR1/@OBRIG").Text = "S" Then
        Call fgAppendNode(xmlRemessaBG, "kColl", "field", "")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "id", "NOM-APE2-TITULAR1")
        Call fgAppendAttributeLastNode(xmlRemessaBG, "field", "value", xmlParametroGeral.selectSingleNode("//NOM-APE2-TITULAR1").Text)
    End If
    '-----------------------------------------------------------------
        
    Call fgAppendXML(xmlOperacao, "MESG", xmlRemessaBG.xml)
    
    Set xmlRemessaBG = Nothing
    Set objLancamentoCCContabil = Nothing
    Set xmlParametroGeral = Nothing
    
    Exit Function
    
ErrorHandler:

    Set xmlRemessaBG = Nothing
    Set objLancamentoCCContabil = Nothing
    Set xmlParametroHistoricoCC = Nothing
    Set xmlParametroGeral = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontarEnvioEstornoBG_Parm Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Processo de exclusão de remessas disponíveis em conta corrente
Public Function ExcluirDisponivel(ByRef xmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim objLancamentoCCContabil                 As A8LQS.clsLancamentoCCContabil
Dim xmlLancamentoCC                         As MSXML2.DOMDocument40
Dim strLancamentoCC                         As String

On Error GoTo ErrorHandler

    If CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrente And _
        CLng(xmlOperacao.documentElement.selectSingleNode("CO_FORM_LIQU").Text) <> enumFormaLiquidacao.ContaCorrenteComTributacao Then
        Exit Function
    End If

    Set objLancamentoCCContabil = CreateObject("A8LQS.clsLancamentoCCContabil")
    strLancamentoCC = objLancamentoCCContabil.Ler(xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                                  enumTipoLancamentoIntegracao.Normal)

    If strLancamentoCC = vbNullString Then
        Exit Function
    End If

    Set xmlLancamentoCC = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLancamentoCC.loadXML strLancamentoCC

    If Val(xmlLancamentoCC.documentElement.selectSingleNode("//CO_ULTI_SITU_PROC").Text) <> enumStatusIntegracao.Disponível Then
        ' 3054 - Operação com status de Conta Corrente inválido para cancelamento.
        lngCodigoErroNegocio = 3054
        GoTo ErrorHandler
    End If

    flExcluirCC xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                enumTipoLancamentoIntegracao.Normal

    objLancamentoCCContabil.Excluir xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text

    Set objLancamentoCCContabil = Nothing
    Set xmlLancamentoCC = Nothing

    Exit Function

ErrorHandler:

    Set objLancamentoCCContabil = Nothing
    Set xmlLancamentoCC = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExcluirDisponivel Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function BuscarGradeHorario(ByVal plngCodigoEmpresa As Long, _
                                    ByVal pstrSiglaSistema As String) As Date

Dim objHoraLimite                           As ADODB.Recordset
Dim strSQL                                  As String
Dim datHoraLimite                           As Date

On Error GoTo ErrorHandler

    strSQL = " SELECT   HO_LIMI_ENVI_ITGR         " & vbNewLine & _
             " FROM     A8.TB_HORA_LIMI_ITGR_SIST " & vbNewLine & _
             " WHERE    SG_SIST  = '" & pstrSiglaSistema & "'" & _
             " AND      CO_EMPR = " & plngCodigoEmpresa

    Set objHoraLimite = fgQuerySQL(strSQL)

    If objHoraLimite.RecordCount = 0 Then
        datHoraLimite = TimeSerial(0, 0, 0)
    Else
        datHoraLimite = TimeSerial(Hour(objHoraLimite!HO_LIMI_ENVI_ITGR), Minute(objHoraLimite!HO_LIMI_ENVI_ITGR), Second(objHoraLimite!HO_LIMI_ENVI_ITGR))
    End If
    
    BuscarGradeHorario = datHoraLimite

    Exit Function
ErrorHandler:

    BuscarGradeHorario = Null

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "BuscarGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recebe como parâmetro de entrada o registro de um lançamento de conta de corrente,
'e partir dos dados desse registro verifica o horario limite da teimosinha, e retorna True caso ele esteja excedido
Public Function ExcedeuHorarioLimiteTeimosinha(ByVal pvntSequenciaOperacao As Variant, _
                                      Optional ByVal plngTipoLancamentoIntegracao As enumTipoLancamentoIntegracao = enumTipoLancamentoIntegracao.Normal, _
                                      Optional ByVal pNumSeqLancamento As Long = 0) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim datHoraServidor                         As Date
Dim datHoraLimiteTeimosinha                 As Date

On Error GoTo ErrorHandler

    strSQL = " SELECT  CO_EMPR           " & _
             "        ,TP_BKOF           " & _
             "        ,TP_OPER           " & _
             "        ,IN_LANC_DEBT_CRED " & _
             "        ,TP_CNAL_VEND      " & _
             " FROM   A8.TB_LANC_CC_CNTB " & _
             " WHERE  NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao & _
             "   AND  TP_LANC_ITGR      = " & plngTipoLancamentoIntegracao & _
             "   AND  NR_SEQU_LANC      = " & pNumSeqLancamento

    Set objRS = fgQuerySQL(strSQL)

    If objRS.RecordCount > 0 Then
    
        If Not (IsNull(objRS!CO_EMPR) Or IsNull(objRS!TP_BKOF) Or IsNull(objRS!TP_OPER) Or IsNull(objRS!IN_LANC_DEBT_CRED) Or IsNull(objRS!TP_CNAL_VEND)) Then
    
            strSQL = " SELECT HR_LIMI_ITGR " & _
                     " FROM   A8.TB_PARM_REPR_CNTA_CRRT " & _
                     " WHERE  CD_EMPR           = " & objRS!CO_EMPR & _
                     "   AND  TP_BKOF           = " & objRS!TP_BKOF & _
                     "   AND  TP_OPER           = " & objRS!TP_OPER & _
                     "   AND  IN_LANC_DEBT_CRED = " & objRS!IN_LANC_DEBT_CRED & _
                     "   AND  TP_CNAL_VEND      = " & objRS!TP_CNAL_VEND
    
            Set objRS = fgQuerySQL(strSQL)
    
            If objRS.RecordCount > 0 Then
    
                If Not IsNull(objRS!HR_LIMI_ITGR) Then
                    
                    datHoraLimiteTeimosinha = TimeSerial(Hour(objRS!HR_LIMI_ITGR), Minute(objRS!HR_LIMI_ITGR), Second(objRS!HR_LIMI_ITGR))
                    datHoraServidor = Format(fgDataHoraServidor(DataHoraAux), "hh:mm")
    
                    If datHoraServidor > objRS!HR_LIMI_ITGR Then
                        ExcedeuHorarioLimiteTeimosinha = True
                    End If
                End If
            
            End If
    
        End If
        
    End If
    
    objRS.Close
    Set objRS = Nothing

Exit Function
ErrorHandler:
    
    objRS.Close
    Set objRS = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExcedeuHorarioLimiteTeimosinha Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
