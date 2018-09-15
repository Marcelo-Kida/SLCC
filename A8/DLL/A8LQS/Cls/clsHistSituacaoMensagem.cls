VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistSituacaoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável por fornecer informações históricas das mensagens cadastradas

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

'Recuperar registros da tabela A8.TB_HIST_SITU_MESG_SPB

Public Function LerTodos(Optional ByVal pstrNumeroControleIF As String, _
                         Optional ByVal pstrDataRegistroMensagem As String, _
                         Optional ByVal pstrNumeroSequenciaControleRepeticao As String = "1") As String

Dim strSQL                                  As String
Dim datDataCorte                            As Date

    On Error GoTo ErrorHandler
    
    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo

    strSQL = "Select    A.NU_CTRL_IF,                          " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                    " & vbNewLine & _
             "          A.NU_SEQU_CNTR_REPE,                   " & vbNewLine & _
             "          A.DH_SITU_ACAO_MESG_SPB,               " & vbNewLine & _
             "          A.CO_SITU_PROC,                        " & vbNewLine & _
             "          B.DE_SITU_PROC,                        " & vbNewLine & _
             "          A.TP_ACAO_MESG_SPB,                    " & vbNewLine
    
    strSQL = strSQL & _
             "          DECODE(A.TP_ACAO_MESG_SPB,             " & vbNewLine & _
             "          " & enumTipoAcao.AlteracaoHorarioAgendamento & ",     'Alteração Horário Agendamento',       " & vbNewLine & _
             "          " & enumTipoAcao.AlteracaoTipoCompromisso & ",        'Alteração Tipo Compromisso',          " & vbNewLine & _
             "          " & enumTipoAcao.CancelamentoSolicitado & ",          'Cancelamento Solicitado',             " & vbNewLine & _
             "          " & enumTipoAcao.CancelamentoEnviado & ",             'Cancelamento Enviado',                " & vbNewLine & _
             "          " & enumTipoAcao.EstornoSolicitado & ",               'Estorno Solicitado',                  " & vbNewLine & _
             "          " & enumTipoAcao.EstornoEnviado & ",                  'Estorno Enviado',                     " & vbNewLine & _
             "          " & enumTipoAcao.RejeicaoConcordancia & ",            'Rejeição Concordância',               " & vbNewLine & _
             "          " & enumTipoAcao.RejeicaoDiscordancia & ",            'Rejeição Discordância',               " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLDL1002 & ",                  'Enviada LDL1002',                     " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaSEL1023 & ",                  'Enviada SEL1023',                     " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLDL1016 & ",                  'Enviada LDL1016',                     " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLDL1006 & ",                  'Enviada LDL1006',                     " & vbNewLine
             
    strSQL = strSQL & _
             "          " & enumTipoAcao.RejeicaoConcordanciaAdmArea & ",     'Rejeição Concordância Adm. Área',     " & vbNewLine & _
             "          " & enumTipoAcao.RejeicaoDiscordanciaAdmArea & ",     'Rejeição Discordância Adm. Área',     " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLDL0003Concordancia & ",      'Enviada LDL0003 Concordância',        " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLDL0003Discordancia & ",      'Enviada LDL0003 Discordância',        " & vbNewLine & _
             "          " & enumTipoAcao.EnviadoPagamento & ",                'Enviado Pagamento',                   " & vbNewLine & _
             "          " & enumTipoAcao.EnviadoPagamentoContingencia & ",    'Enviado Pagamento Contingência',      " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLTR0002Concordancia & ",      'Enviada LTR0002 Concordância',        " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLTR0002Discordancia & ",      'Enviada LTR0002 Discordância',        " & vbNewLine & _
             "          " & enumTipoAcao.EnviadoPagamentoBACEN & ",           'Enviado Pagamento BACEN',             " & vbNewLine & _
             "          " & enumTipoAcao.EnviadoPagamentoSTR & ",             'Enviado Pagamento STR',               " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLTR0008Concordancia & ",      'Enviada LTR0008 Concordância',        " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaLTR0008Discordancia & ",      'Enviada LTR0008 Discordância',        " & vbNewLine & _
             "          " & enumTipoAcao.ConcordanciaEmSer & ",               'Concordância Em Ser',                 " & vbNewLine & _
             "          " & enumTipoAcao.ConcordanciaManualEmSer & ",         'Concordância Manual Em Ser',          " & vbNewLine & _
             "          " & enumTipoAcao.RejeicaoConcordanciaEmSer & ",       'Rejeição Concordância Em Ser',        " & vbNewLine & _
             "          " & enumTipoAcao.RejeicaoConcordanciaManualEmSer & ", 'Rejeição Concordância Manual Em Ser', " & vbNewLine & _
             "          " & enumTipoAcao.Liberacao & ",                       'Liberação',                           " & vbNewLine & _
             "          " & enumTipoAcao.LiberacaoAntecipada & ",             'Liberação Antecipada',                " & vbNewLine & _
             "          " & enumTipoAcao.Liquidacao & ",                      'Liquidação',                          " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaSTR0004Pagamento & ",         'Enviada STR0004',                     " & vbNewLine & _
             "          " & enumTipoAcao.EnviadaSTR0007 & ",                  'Enviada STR0007',                     " & vbNewLine & _
             "          'AÇÃO NÃO CONFIGURADA ('||A.TP_ACAO_MESG_SPB||')' ) AS DE_ACAO_MESG_SPB,                     " & vbNewLine
            
    strSQL = strSQL & _
             "          A.TX_CNTD_ANTE_ACAO,                   " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ,                   " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ               " & vbNewLine & _
             " From     A8.TB_HIST_SITU_ACAO_MESG_SPB        A," & vbNewLine & _
             "          A8.TB_SITU_PROC                      B " & vbNewLine & _
             " Where    A.CO_SITU_PROC = B.CO_SITU_PROC  (+) "

    If pstrNumeroControleIF <> vbNullString Then
        strSQL = strSQL & vbNewLine & _
            " And       A.NU_CTRL_IF = '" & pstrNumeroControleIF & "'    "
    End If

    If pstrDataRegistroMensagem <> vbNullString Then
    
        If fgDtHrStr_To_DateTime(pstrDataRegistroMensagem) < datDataCorte Then
            strSQL = Replace$(strSQL, "A8.TB_HIST_SITU_ACAO_MESG_SPB", "A8HIST.TB_HIST_SITU_ACAO_MESG_SPB")
        End If
    
        strSQL = strSQL & vbNewLine & _
            " And       A.DH_REGT_MESG_SPB      = " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem)
    End If

    strSQL = strSQL & vbNewLine & _
        " And       A.NU_SEQU_CNTR_REPE      = " & pstrNumeroSequenciaControleRepeticao

    strSQL = strSQL & _
            " ORDER BY A.DH_REGT_MESG_SPB DESC "


    LerTodos = fgQueryXMLLerTodos("SituacaoMensagem", strSQL, App.EXEName & "." & TypeName(Me))

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

'Incluir um registro na tabela A8.TB_HIST_SITU_MESG_SPB

Public Function Incluir(ByVal pstrNumeroControleIF As String, _
                        ByVal pstrDTRegistroMensagemSPB As String, _
               Optional ByVal pintStatusMensagem As Integer = 0, _
               Optional ByVal pintTipoAcao As Integer = 0, _
               Optional ByVal pstrTextoConteudoAnteriorAcao As String = vbNullString, _
               Optional ByVal plngSeqControleRepeticao As Long = 1) As Boolean

Dim strSQL                                  As String
Dim strDataGravacao                         As String

On Error GoTo ErrorHandler

    If plngSeqControleRepeticao = 0 Then plngSeqControleRepeticao = 1
    
    strDataGravacao = flObterDataGravacao(pstrNumeroControleIF, _
                                          pstrDTRegistroMensagemSPB, _
                                          fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora)))

    strSQL = " Insert Into A8.TB_HIST_SITU_ACAO_MESG_SPB (    " & vbNewLine & _
             "             NU_CTRL_IF,                        " & vbNewLine & _
             "             DH_REGT_MESG_SPB,                  " & vbNewLine & _
             "             NU_SEQU_CNTR_REPE,                 " & vbNewLine & _
             "             DH_SITU_ACAO_MESG_SPB,             " & vbNewLine & _
             "             CO_SITU_PROC,                      " & vbNewLine & _
             "             TP_ACAO_MESG_SPB,                  " & vbNewLine & _
             "             TX_CNTD_ANTE_ACAO,                 " & vbNewLine & _
             "             CO_USUA_ULTI_ATLZ,                 " & vbNewLine & _
             "             CO_ETCA_TRAB_ULTI_ATLZ)            " & vbNewLine & _
             " Values     ('" & pstrNumeroControleIF & "',    " & vbNewLine & _
             "             " & fgDtHrXML_To_Oracle(pstrDTRegistroMensagemSPB) & ", " & vbNewLine & _
             "             " & plngSeqControleRepeticao & ", " & vbNewLine & _
             "             " & fgDtHrXML_To_Oracle(strDataGravacao) & ", " & vbNewLine & _
             "             " & IIf(pintStatusMensagem = 0, "NULL", pintStatusMensagem) & ", " & vbNewLine & _
             "             " & IIf(pintTipoAcao = 0, "NULL", pintTipoAcao) & ", " & vbNewLine & _
             "             " & IIf(pstrTextoConteudoAnteriorAcao = vbNullString, "NULL", _
             "'" & pstrTextoConteudoAnteriorAcao & "'") & ", " & vbNewLine & _
             "             '" & fgUsuarioRede & "',  " & vbNewLine & _
             "             '" & fgEstacaoTrabalhoUsuario & "')"

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

'Obter data Hora do ultimo registro

Private Function flObterDataGravacao(ByVal pstrNumeroControleIF As String, _
                                     ByVal pstrDataGravacaoSPB As String, _
                                     ByVal pstrDataAcaoMensagem As String) As String

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim datDataGravacao                         As Date

    On Error GoTo ErrorHandler

    strSQL = " SELECT   MAX(DH_SITU_ACAO_MESG_SPB) AS DH_SITU_ACAO_MESG_SPB " & vbNewLine & _
             " FROM     A8.TB_HIST_SITU_ACAO_MESG_SPB    " & vbNewLine & _
             " WHERE    NU_CTRL_IF              =   '" & pstrNumeroControleIF & "'"

    Set objRS = fgQuerySQL(strSQL)

    If IsNull(objRS!DH_SITU_ACAO_MESG_SPB) Then
        datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
    Else
        If DateDiff("s", fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux), objRS!DH_SITU_ACAO_MESG_SPB) = 0 Then
            datDataGravacao = DateAdd("s", 1, fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux))
        Else
            datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
        End If
    End If

    flObterDataGravacao = fgDtHr_To_Xml(datDataGravacao)

    objRS.Close
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDataGravacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
