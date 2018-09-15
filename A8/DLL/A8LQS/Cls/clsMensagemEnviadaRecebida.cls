VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemEnviadaRecebida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável por controlar o acesso à tabela A8.TB_MESG_RECB_ENVI_SPB,
'' onde é armazenado todas as mensagens enviadas e recebidas da mensageria.

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

'--------------------------------------------------------------------------------------------------
'Ler as informações da primeira mensagem SPB da operação
'--------------------------------------------------------------------------------------------------

Public Function LerPrimeiraMensagem(ByVal pvntNumeroSequencia As Variant) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   NU_CTRL_IF, " & vbCrLf & _
             "          DH_REGT_MESG_SPB, " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV, " & vbCrLf & _
             "          TP_BKOF, " & vbCrLf & _
             "          CO_EMPR, " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB, " & vbCrLf & _
             "          CO_MESG_SPB, " & vbCrLf & _
             "          NU_COMD_OPER, " & vbCrLf & _
             "          CO_SITU_MESG_SPB, " & vbCrLf & _
             "          CO_TEXT_XML, " & vbCrLf & _
             "          HO_ENVI_MESG_SPB, " & vbCrLf & _
             "          CO_ULTI_SITU_PROC, " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ, " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ " & vbCrLf & _
             "          DH_ULTI_ATLZ, " & vbCrLf & _
             "          IN_ENTR_MANU, " & vbCrLf & _
             "          CO_LOCA_LIQU, " & vbCrLf & _
             "          CO_VEIC_LEGA, " & vbCrLf & _
             "          SG_SIST " & vbCrLf & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB " & vbCrLf & _
             " WHERE    NU_SEQU_OPER_ATIV  =   " & pvntNumeroSequencia & vbCrLf & _
             " AND      DH_REGT_MESG_SPB   = ( SELECT MIN(DH_REGT_MESG_SPB)  " & vbCrLf & _
             "                                 FROM A8.TB_MESG_RECB_ENVI_SPB " & vbCrLf & _
             "                                 WHERE NU_SEQU_OPER_ATIV = " & pvntNumeroSequencia & " ) " & vbCrLf

    LerPrimeiraMensagem = fgQueryXMLLer("MensagemRecebidaEnviada", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerPrimeiraMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Ler as informações da mensagem SPB enviada ou recebida pelo SLCC
'--------------------------------------------------------------------------------------------------

Public Function Ler(ByVal pstrNumeroControleIF As String, _
           Optional ByVal plngNumeroSequenciaRepeticao As Long = 0, _
           Optional ByVal pstrDHRegistroMensagem As String = "") As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   NU_CTRL_IF,                  " & vbCrLf & _
             "          DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          TP_BKOF,                     " & vbCrLf & _
             "          CO_EMPR,                     " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,       " & vbCrLf & _
             "          CO_MESG_SPB,                 " & vbCrLf & _
             "          NU_COMD_OPER,                " & vbCrLf & _
             "          CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          CO_TEXT_XML,                 " & vbCrLf & _
             "          HO_ENVI_MESG_SPB,            " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
             "          DH_ULTI_ATLZ,                " & vbCrLf & _
             "          IN_ENTR_MANU                 " & vbCrLf & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB     " & vbCrLf & _
             " WHERE    NU_CTRL_IF  =   '" & pstrNumeroControleIF & "'"

    If plngNumeroSequenciaRepeticao <> 0 Then
        strSQL = strSQL & " AND NU_SEQU_CNTR_REPE = " & plngNumeroSequenciaRepeticao
    End If
    
    If pstrDHRegistroMensagem <> "" Then
        strSQL = strSQL & " AND DH_REGT_MESG_SPB = " & fgDtHrXML_To_Oracle(pstrDHRegistroMensagem)
    End If
    
    Ler = fgQueryXMLLer("MensagemRecebidaEnviada", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Ler o xml da mensagem SPB enviada ou recebida pelo SLCC
'--------------------------------------------------------------------------------------------------

Public Function LerXMLMensagem(Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                               Optional ByVal pstrNumeroSequenciaRepeticao As String = "1", _
                               Optional ByVal plngDebitoCredito As Long = 0, _
                               Optional ByVal pstrCodigoMensagemSPB As String = vbNullString, _
                               Optional ByVal pstrNumeroControleCamara As String = vbNullString, _
                               Optional ByVal pstrDataRegistroMensagemSPB As String = vbNullString, _
                               Optional ByVal pstrNumeroSequenciaOperacao As String = vbNullString, _
                               Optional ByVal pstrDataMensagem As String = vbNullString, _
                               Optional ByVal pstrTipoInformacao As String = vbNullString, _
                               Optional ByVal plngLocalLiquidacao As Long = 0, _
                               Optional ByVal pstrNumeroControleIFDiferente As String = vbNullString, _
                               Optional ByVal pstrChACAM As String = vbNullString, _
                               Optional ByVal pstrContrSISBACEN As String = vbNullString, _
                               Optional ByVal pblnForUpdate As Boolean = True, _
                               Optional ByVal pstrNumeroComandoOperacao As String = vbNullString, _
                               Optional ByVal pstrNumeroOperacaoCambial2 As String = vbNullString) As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim strXML                                  As String
Dim rsMensagem                              As ADODB.Recordset
Dim strXMLBaseName                          As String
Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloNZ                          As udtProtocoloNZ

    On Error GoTo ErrorHandler

    Set rsMensagem = CreateObject("ADODB.Recordset")
    Set LerXMLMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    strSQL = " SELECT   A.NU_CTRL_IF,                  " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          A.CO_VEIC_LEGA,                " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          A.TP_BKOF,                     " & vbCrLf & _
             "          A.CO_EMPR,                     " & vbCrLf & _
             "          A.NU_COMD_OPER,                " & vbCrLf & _
             "          A.CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          A.CO_TEXT_XML,                 " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          A.CO_MESG_SPB,                 " & vbCrLf & _
             "          A.SG_SIST,                     " & vbCrLf & _
             "          A.CO_LOCA_LIQU,                " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          A.IN_ENTR_MANU,                " & vbCrLf & _
             "          A.NU_CTRL_CAMR,                " & vbCrLf & _
             "          A.IN_CONF_MESG_LTR,            " & vbCrLf & _
             "          A.CO_PARP_CAMR,                " & vbCrLf & _
             "          A.TP_ACAO_MESG_SPB_EXEC,       " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                " & vbCrLf
             
    strSQL = strSQL & _
             "          B.CO_ISPB_BANC_LIQU_CNPT,      " & vbCrLf & _
             "          B.TP_INFO_LDL,                 " & vbCrLf & _
             "          B.CO_PRAC,                     " & vbCrLf & _
             "          B.CO_MOED_ESTR,                " & vbCrLf & _
             "          B.PE_TAXA_NEGO,                " & vbCrLf & _
             "          B.VA_FINC,                     " & vbCrLf & _
             "          B.VA_MOED_ESTR,                " & vbCrLf & _
             "          B.DT_LIQU,                     " & vbCrLf & _
             "          B.CO_SISB_COTR,                " & vbCrLf & _
             "          B.IN_OPER_DEBT_CRED,           " & vbCrLf & _
             "          B.TP_LIQU_OPER_ATIV,           " & vbCrLf & _
             "          B.CD_ASSO_CAMB,                " & vbCrLf & _
             "          B.CO_CNTR_SISB                 " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB A,    " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL B     " & vbCrLf

    strSQL = strSQL & _
             " WHERE    A.NU_CTRL_IF         = B.NU_CTRL_IF        (+) " & vbCrLf & _
             "   AND    A.NU_SEQU_CNTR_REPE  = B.NU_SEQU_CNTR_REPE (+) " & vbCrLf & _
             "   AND    A.DH_REGT_MESG_SPB   = B.DH_REGT_MESG_SPB  (+) " & vbCrLf & _
             "   AND    A.NU_SEQU_CNTR_REPE  =  " & pstrNumeroSequenciaRepeticao & vbCrLf

    If plngDebitoCredito > 0 Then
        strSQL = strSQL & _
             "   AND    B.IN_OPER_DEBT_CRED  =  " & plngDebitoCredito & vbCrLf
    End If

    If pstrNumeroControleIF <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NU_CTRL_IF         = '" & pstrNumeroControleIF & "'" & vbCrLf
    End If

    If pstrNumeroControleIFDiferente <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NU_CTRL_IF         <> '" & pstrNumeroControleIFDiferente & "'" & vbCrLf
    End If

    If pstrNumeroSequenciaOperacao <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NU_SEQU_OPER_ATIV  = '" & pstrNumeroSequenciaOperacao & "'" & vbCrLf
    End If

    If pstrCodigoMensagemSPB <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.CO_MESG_SPB        = '" & pstrCodigoMensagemSPB & "'" & vbCrLf
    End If

    If pstrNumeroControleCamara <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NU_CTRL_CAMR        = '" & pstrNumeroControleCamara & "'" & vbCrLf
    End If

    If pstrDataRegistroMensagemSPB <> vbNullString Then
        If pstrCodigoMensagemSPB = "CAM0055" _
        Or pstrCodigoMensagemSPB = "CAM0005R2" Then
            strSQL = strSQL & _
                 "   AND    A.DH_REGT_MESG_SPB        < " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB) & vbCrLf
        Else
            strSQL = strSQL & _
                 "   AND    A.DH_REGT_MESG_SPB        = " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagemSPB) & vbCrLf
        End If
    End If

    If pstrDataMensagem <> vbNullString Then
        strSQL = strSQL & _
             "   AND    TRUNC(A.DH_REGT_MESG_SPB)       = " & fgDtXML_To_Oracle(pstrDataMensagem) & vbCrLf
    End If

    If pstrTipoInformacao <> vbNullString Then
        strSQL = strSQL & _
             "   AND    B.TP_INFO_LDL         = '" & pstrTipoInformacao & "'" & vbCrLf
    End If

    If plngLocalLiquidacao > 0 Then
        strSQL = strSQL & _
             "   AND    A.CO_LOCA_LIQU  =  " & plngLocalLiquidacao & vbCrLf
    End If

    If pstrChACAM <> vbNullString Then
        strSQL = strSQL & _
             "   AND    B.CD_ASSO_CAMB  =  '" & pstrChACAM & "'" & vbCrLf
    End If

    If pstrContrSISBACEN <> vbNullString Then
        strSQL = strSQL & _
             "   AND    B.CO_CNTR_SISB  =  " & pstrContrSISBACEN & vbCrLf
    End If
    
    If pstrNumeroComandoOperacao <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NU_COMD_OPER =  '" & pstrNumeroComandoOperacao & "'" & vbCrLf
    End If
    
    If pstrNumeroOperacaoCambial2 <> vbNullString Then
        strSQL = strSQL & _
             "   AND    A.NR_OPER_CAMB_2  =  " & CDbl(pstrNumeroOperacaoCambial2) & vbCrLf
    End If

    If pblnForUpdate And Not fgEstacaoDesenvolvimento Then
        strSQL = strSQL & " FOR UPDATE " & vbCrLf
    End If

    strSQL = strSQL & " ORDER BY A.DH_REGT_MESG_SPB DESC "
    
    Set rsMensagem = fgQuerySQL(strSQL)

    If rsMensagem.EOF Then
        Exit Function
    End If

    strXML = fgSelectVarchar4000(rsMensagem!CO_TEXT_XML)

    If Not LerXMLMensagem.loadXML(strXML) Then
        'Se não conseguir ler a STRING retornada pela função << fgSelectVarchar4000 >>,
        'tenta ler removendo os HEADERs (SLCC e NZ)
        LerXMLMensagem.loadXML (Mid(strXML, Len(udtProtocolo) + Len(udtProtocoloNZ) + 1))
    End If

    strXMLBaseName = "SISMSG"
    If Not LerXMLMensagem.childNodes(0) Is Nothing Then
        strXMLBaseName = LerXMLMensagem.childNodes(0).baseName
    End If
    
    If LerXMLMensagem.documentElement.selectSingleNode("NU_CTRL_IF") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_CTRL_IF", IIf(IsNull(rsMensagem!NU_CTRL_IF), "", rsMensagem!NU_CTRL_IF)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_CTRL_IF").Text = rsMensagem!NU_CTRL_IF
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("DH_REGT_MESG_SPB") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "DH_REGT_MESG_SPB", IIf(IsNull(rsMensagem!DH_REGT_MESG_SPB), "", fgDtHr_To_Xml(rsMensagem!DH_REGT_MESG_SPB))
    Else
        LerXMLMensagem.documentElement.selectSingleNode("DH_REGT_MESG_SPB").Text = fgDtHr_To_Xml(rsMensagem!DH_REGT_MESG_SPB)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("DH_ULTI_ATLZ") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "DH_ULTI_ATLZ", IIf(IsNull(rsMensagem!dh_ulti_atlz), "", fgDtHr_To_Xml(rsMensagem!dh_ulti_atlz))
    Else
        LerXMLMensagem.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text = fgDtHr_To_Xml(rsMensagem!dh_ulti_atlz)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_SEQU_OPER_ATIV", IIf(IsNull(rsMensagem!NU_SEQU_OPER_ATIV), "", rsMensagem!NU_SEQU_OPER_ATIV)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text = IIf(IsNull(rsMensagem!NU_SEQU_OPER_ATIV), "", rsMensagem!NU_SEQU_OPER_ATIV)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_VEIC_LEGA") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_VEIC_LEGA", IIf(IsNull(rsMensagem!CO_VEIC_LEGA), "", rsMensagem!CO_VEIC_LEGA)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_VEIC_LEGA").Text = IIf(IsNull(rsMensagem!CO_VEIC_LEGA), "", rsMensagem!CO_VEIC_LEGA)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_SEQU_CNTR_REPE", IIf(IsNull(rsMensagem!NU_SEQU_CNTR_REPE), "", rsMensagem!NU_SEQU_CNTR_REPE)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_CNTR_REPE").Text = rsMensagem!NU_SEQU_CNTR_REPE
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "TP_LIQU_OPER_ATIV", IIf(IsNull(rsMensagem!TP_LIQU_OPER_ATIV), "", rsMensagem!TP_LIQU_OPER_ATIV)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("TP_LIQU_OPER_ATIV").Text = IIf(IsNull(rsMensagem!TP_LIQU_OPER_ATIV), "", rsMensagem!TP_LIQU_OPER_ATIV)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("TP_BKOF") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "TP_BKOF", IIf(IsNull(rsMensagem!TP_BKOF), "", rsMensagem!TP_BKOF)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("TP_BKOF").Text = IIf(IsNull(rsMensagem!TP_BKOF), "", rsMensagem!TP_BKOF)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_EMPR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_EMPR", IIf(IsNull(rsMensagem!CO_EMPR), "", rsMensagem!CO_EMPR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_EMPR").Text = rsMensagem!CO_EMPR
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("NU_COMD_OPER") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_COMD_OPER", IIf(IsNull(rsMensagem!NU_COMD_OPER), "", rsMensagem!NU_COMD_OPER)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_COMD_OPER").Text = IIf(IsNull(rsMensagem!NU_COMD_OPER), "", rsMensagem!NU_COMD_OPER)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_SITU_MESG_SPB") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_SITU_MESG_SPB", IIf(IsNull(rsMensagem!CO_SITU_MESG_SPB), "", rsMensagem!CO_SITU_MESG_SPB)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_SITU_MESG_SPB").Text = IIf(IsNull(rsMensagem!CO_SITU_MESG_SPB), "", rsMensagem!CO_SITU_MESG_SPB)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_TEXT_XML") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_TEXT_XML", IIf(IsNull(rsMensagem!CO_TEXT_XML), "", rsMensagem!CO_TEXT_XML)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_TEXT_XML").Text = rsMensagem!CO_TEXT_XML
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_ULTI_SITU_PROC") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_ULTI_SITU_PROC", IIf(IsNull(rsMensagem!CO_ULTI_SITU_PROC), "", rsMensagem!CO_ULTI_SITU_PROC)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text = rsMensagem!CO_ULTI_SITU_PROC
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_MESG", IIf(IsNull(rsMensagem!CO_MESG_SPB), "", rsMensagem!CO_MESG_SPB)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_MESG").Text = Trim$(rsMensagem!CO_MESG_SPB)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("IN_ENTR_MANU") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "IN_ENTR_MANU", IIf(IsNull(rsMensagem!IN_ENTR_MANU), "", rsMensagem!IN_ENTR_MANU)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("IN_ENTR_MANU").Text = rsMensagem!IN_ENTR_MANU
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("SG_SIST") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "SG_SIST", IIf(IsNull(rsMensagem!SG_SIST), "", rsMensagem!SG_SIST)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("SG_SIST").Text = IIf(IsNull(rsMensagem!SG_SIST), "", rsMensagem!SG_SIST)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_LOCA_LIQU") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_LOCA_LIQU", IIf(IsNull(rsMensagem!CO_LOCA_LIQU), "", rsMensagem!CO_LOCA_LIQU)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_LOCA_LIQU").Text = IIf(IsNull(rsMensagem!CO_LOCA_LIQU), "", rsMensagem!CO_LOCA_LIQU)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_SEQU_CNCL_OPER_ATIV_MESG", IIf(IsNull(rsMensagem!NU_SEQU_CNCL_OPER_ATIV_MESG), "0", rsMensagem!NU_SEQU_CNCL_OPER_ATIV_MESG)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_SEQU_CNCL_OPER_ATIV_MESG").Text = IIf(IsNull(rsMensagem!NU_SEQU_CNCL_OPER_ATIV_MESG), "0", rsMensagem!NU_SEQU_CNCL_OPER_ATIV_MESG)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("NU_CTRL_CAMR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "NU_CTRL_CAMR", IIf(IsNull(rsMensagem!NU_CTRL_CAMR), "", rsMensagem!NU_CTRL_CAMR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("NU_CTRL_CAMR").Text = IIf(IsNull(rsMensagem!NU_CTRL_CAMR), "", rsMensagem!NU_CTRL_CAMR)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_ISPB_BANC_LIQU_CNPT", IIf(IsNull(rsMensagem!CO_ISPB_BANC_LIQU_CNPT), "", rsMensagem!CO_ISPB_BANC_LIQU_CNPT)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text = IIf(IsNull(rsMensagem!CO_ISPB_BANC_LIQU_CNPT), "", rsMensagem!CO_ISPB_BANC_LIQU_CNPT)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("IN_CONF_MESG_LTR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "IN_CONF_MESG_LTR", IIf(IsNull(rsMensagem!IN_CONF_MESG_LTR), "", rsMensagem!IN_CONF_MESG_LTR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("IN_CONF_MESG_LTR").Text = IIf(IsNull(rsMensagem!IN_CONF_MESG_LTR), "", rsMensagem!IN_CONF_MESG_LTR)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_PARP_CAMR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_PARP_CAMR", IIf(IsNull(rsMensagem!CO_PARP_CAMR), "", rsMensagem!CO_PARP_CAMR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_PARP_CAMR").Text = IIf(IsNull(rsMensagem!CO_PARP_CAMR), "", rsMensagem!CO_PARP_CAMR)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("TP_INFO_LDL") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "TP_INFO_LDL", IIf(IsNull(rsMensagem!TP_INFO_LDL), "", rsMensagem!TP_INFO_LDL)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("TP_INFO_LDL").Text = IIf(IsNull(rsMensagem!TP_INFO_LDL), "", rsMensagem!TP_INFO_LDL)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_PRAC") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_PRAC", IIf(IsNull(rsMensagem!CO_PRAC), "", rsMensagem!CO_PRAC)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_PRAC").Text = IIf(IsNull(rsMensagem!CO_PRAC), "", rsMensagem!CO_PRAC)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_MOED_ESTR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_MOED_ESTR", IIf(IsNull(rsMensagem!CO_MOED_ESTR), "", rsMensagem!CO_MOED_ESTR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_MOED_ESTR").Text = IIf(IsNull(rsMensagem!CO_MOED_ESTR), "", rsMensagem!CO_MOED_ESTR)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("PE_TAXA_NEGO") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "PE_TAXA_NEGO", IIf(IsNull(rsMensagem!PE_TAXA_NEGO), "", rsMensagem!PE_TAXA_NEGO)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("PE_TAXA_NEGO").Text = IIf(IsNull(rsMensagem!PE_TAXA_NEGO), "", rsMensagem!PE_TAXA_NEGO)
    End If

    'RATS 394
    LerXMLMensagem.documentElement.selectSingleNode("PE_TAXA_NEGO").Text = Replace(LerXMLMensagem.documentElement.selectSingleNode("PE_TAXA_NEGO").Text, ".", ",")

    If LerXMLMensagem.documentElement.selectSingleNode("VA_FINC") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "VA_FINC", IIf(IsNull(rsMensagem!VA_FINC), "", rsMensagem!VA_FINC)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("VA_FINC").Text = IIf(IsNull(rsMensagem!VA_FINC), "", rsMensagem!VA_FINC)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("VA_MOED_ESTR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "VA_MOED_ESTR", IIf(IsNull(rsMensagem!VA_MOED_ESTR), "", rsMensagem!VA_MOED_ESTR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("VA_MOED_ESTR").Text = IIf(IsNull(rsMensagem!VA_MOED_ESTR), "", rsMensagem!VA_MOED_ESTR)
    End If
    
    LerXMLMensagem.documentElement.selectSingleNode("VA_MOED_ESTR").Text = Replace(LerXMLMensagem.documentElement.selectSingleNode("VA_MOED_ESTR").Text, ".", ",")
    
    If LerXMLMensagem.documentElement.selectSingleNode("DT_LIQU") Is Nothing Then
        If IsNull(rsMensagem!DT_LIQU) Then
            fgAppendNode LerXMLMensagem, strXMLBaseName, "DT_LIQU", vbNullString
        Else
            fgAppendNode LerXMLMensagem, strXMLBaseName, "DT_LIQU", fgDt_To_Xml(rsMensagem!DT_LIQU)
        End If
    Else
        If IsNull(rsMensagem!DT_LIQU) Then
            LerXMLMensagem.documentElement.selectSingleNode("DT_LIQU").Text = vbNullString
        Else
            LerXMLMensagem.documentElement.selectSingleNode("DT_LIQU").Text = fgDt_To_Xml(rsMensagem!DT_LIQU)
        End If
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_SISB_COTR") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_SISB_COTR", IIf(IsNull(rsMensagem!CO_SISB_COTR), "", rsMensagem!CO_SISB_COTR)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_SISB_COTR").Text = IIf(IsNull(rsMensagem!CO_SISB_COTR), "", rsMensagem!CO_SISB_COTR)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("IN_OPER_DEBT_CRED") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "IN_OPER_DEBT_CRED", IIf(IsNull(rsMensagem!IN_OPER_DEBT_CRED), "", rsMensagem!IN_OPER_DEBT_CRED)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("IN_OPER_DEBT_CRED").Text = IIf(IsNull(rsMensagem!IN_OPER_DEBT_CRED), "", rsMensagem!IN_OPER_DEBT_CRED)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("TP_ACAO_MESG_SPB_EXEC") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "TP_ACAO_MESG_SPB_EXEC", IIf(IsNull(rsMensagem!TP_ACAO_MESG_SPB_EXEC), "", rsMensagem!TP_ACAO_MESG_SPB_EXEC)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("TP_ACAO_MESG_SPB_EXEC").Text = IIf(IsNull(rsMensagem!TP_ACAO_MESG_SPB_EXEC), "", rsMensagem!TP_ACAO_MESG_SPB_EXEC)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CD_ASSO_CAMB") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CD_ASSO_CAMB", IIf(IsNull(rsMensagem!CD_ASSO_CAMB), "", rsMensagem!CD_ASSO_CAMB)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CD_ASSO_CAMB").Text = IIf(IsNull(rsMensagem!CD_ASSO_CAMB), "", rsMensagem!CD_ASSO_CAMB)
    End If

    If LerXMLMensagem.documentElement.selectSingleNode("CO_CNTR_SISB") Is Nothing Then
        fgAppendNode LerXMLMensagem, strXMLBaseName, "CO_CNTR_SISB", IIf(IsNull(rsMensagem!CO_CNTR_SISB), "", rsMensagem!CO_CNTR_SISB)
    Else
        LerXMLMensagem.documentElement.selectSingleNode("CO_CNTR_SISB").Text = IIf(IsNull(rsMensagem!CO_CNTR_SISB), "", rsMensagem!CO_CNTR_SISB)
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerXMLMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Ler as informações da mensagem SPB enviada pelo SLCC
'--------------------------------------------------------------------------------------------------

Public Function LerMensagemIda(ByVal pstrNumeroControleIF As String, _
                      Optional ByVal pstrCodigoMensagem As String = vbNullString) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandLerMensagemIda

    strSQL = " SELECT   NU_CTRL_IF,                  " & vbCrLf & _
             "          DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          TP_BKOF,                     " & vbCrLf & _
             "          CO_EMPR,                     " & vbCrLf & _
             "          CO_VEIC_LEGA,                " & vbCrLf & _
             "          CO_LOCA_LIQU,                " & vbCrLf & _
             "          SG_SIST,                     " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,       " & vbCrLf & _
             "          CO_MESG_SPB,                 " & vbCrLf & _
             "          NU_COMD_OPER,                " & vbCrLf & _
             "          CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          CO_TEXT_XML,                 " & vbCrLf & _
             "          HO_ENVI_MESG_SPB,            " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
             "          DH_ULTI_ATLZ,                " & vbCrLf & _
             "          NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          IN_ENTR_MANU                 " & vbCrLf
             
    strSQL = strSQL & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB     " & vbCrLf & _
             " WHERE    NU_CTRL_IF          =   '" & pstrNumeroControleIF & "'" & vbCrLf & _
             " AND      DH_REGT_MESG_SPB    =   (Select Min(DH_REGT_MESG_SPB) " & vbCrLf & _
             "                                     From A8.TB_MESG_RECB_ENVI_SPB " & vbCrLf & _
             "                                    Where NU_CTRL_IF          =   '" & pstrNumeroControleIF & "')"
    
    If pstrCodigoMensagem <> vbNullString Then
        strSQL = strSQL & " AND CO_MESG_SPB = '" & pstrCodigoMensagem & "'"
    End If
    
    LerMensagemIda = fgQueryXMLLer("MensagemRecebidaEnviada", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandLerMensagemIda:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagemIda Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Salvar as informações da mensagem SPB enviada pelo SLCC
'--------------------------------------------------------------------------------------------------

Public Function Salvar(ByRef xmlMensagem As MSXML2.DOMDocument40, _
                       ByVal plngCodigoEmpresa As Long, _
                       ByVal pintTipoBackOffice As enumTipoBackOffice, _
                       ByVal plngCodigoLocalLiquidacao As Long, _
                       ByVal pstrCodigoVeiculoLegal As String, _
                       ByVal pstrSiglaSistema As String, _
              Optional ByVal pstrComandoOperacao As String = vbNullString, _
              Optional ByVal penumStatusMensagem As enumStatusMensagem = enumStatusMensagem.ManualEmSer, _
              Optional ByVal plngIndicadorEntradaManual As enumIndicadorSimNao = enumIndicadorSimNao.Sim, _
              Optional ByVal pvntSequenciaConciliacaoOperacao As Variant = 0, _
              Optional ByVal pstrAgendamento As String) As Boolean

Dim strNumCtrlIF                            As String
Dim strMensagem                             As String
Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim objProcessoMensagemBMA                  As A8LQS.clsProcessoMensagemBMA

On Error GoTo ErrorHandler

    'Complementar as informações necessárias para enviar a operação para o PJ
    udtProtocolo.CodigoEmpresa = Format(plngCodigoEmpresa, "00000")
    udtProtocolo.SiglaSistemaDestino = "NZ "
    udtProtocolo.SiglaSistemaOrigem = "A8 "
    
    If xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1069" Or _
        xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1066" Then
        udtProtocolo.TipoMensagem = fgCompletaString(xmlMensagem.selectSingleNode("//CodMsg").Text, " ", 9, False)
    Else
        udtProtocolo.TipoMensagem = Format(enumTipoMensagemBUS.MensagemSPBNaoTratada, "000000000") 'Verificar qual será o evento do BUS para esta função.
    End If
    LSet udtProtocoloAux = udtProtocolo

    'verificar se no xml tem a tag NumCtrlIF, caso não tenha chamar a função abaixo
    If Not xmlMensagem.selectSingleNode("//NumCtrlIF") Is Nothing Then
        strNumCtrlIF = xmlMensagem.selectSingleNode("//NumCtrlIF").Text
    End If
    
    strMensagem = udtProtocoloAux.String

    If xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1069" Or _
        xmlMensagem.selectSingleNode("//CodMsg").Text = "SEL1066" Then
        
        fgAppendNode xmlMensagem, "SISMSG", "TX_HEAD_NZ", fgMontaHeaderMensageNZ(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                                                 plngCodigoEmpresa, _
                                                                                 strNumCtrlIF)
    End If

    If Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "SEL" Or _
       Mid(xmlMensagem.selectSingleNode("//CodMsg").Text, 1, 3) = "RDC" Then
        If xmlMensagem.selectSingleNode("//NumCtrlIF") Is Nothing Then
            fgAppendNode xmlMensagem, "SISMSG", "NumCtrlIF", strNumCtrlIF
        End If
    End If

    strMensagem = strMensagem & fgMontaHeaderMensageNZ(xmlMensagem.selectSingleNode("//CodMsg").Text, _
                                                       plngCodigoEmpresa, _
                                                       strNumCtrlIF)

    strMensagem = strMensagem & xmlMensagem.xml
    
    If xmlMensagem.selectSingleNode("//CodMsg").Text = "STR0011" Then
        
        Set objProcessoMensagemBMA = CreateObject("A8LQS.clsProcessoMensagemBMA")
        Call objProcessoMensagemBMA.GerenciarMensagemAvisoInformacao(xmlMensagem, vbNullString)
        Set objProcessoMensagemBMA = Nothing
    
        pvntSequenciaConciliacaoOperacao = xmlMensagem.selectSingleNode("//NU_SEQU_CNCL_OPER_ATIV_MESG").Text
        pintTipoBackOffice = xmlMensagem.selectSingleNode("//TP_BKOF").Text
        plngCodigoEmpresa = xmlMensagem.selectSingleNode("//CO_EMPR").Text
        plngCodigoLocalLiquidacao = xmlMensagem.selectSingleNode("//CO_LOCA_LIQU").Text
        pstrCodigoVeiculoLegal = xmlMensagem.selectSingleNode("//CO_VEIC_LEGA").Text
        pstrSiglaSistema = xmlMensagem.selectSingleNode("//SG_SIST").Text
    
    End If
    
    Salvar = Incluir(strMensagem, _
                     strNumCtrlIF, _
                     0, _
                     pvntSequenciaConciliacaoOperacao, _
                     xmlMensagem.selectSingleNode("//CodMsg").Text, _
                     pintTipoBackOffice, _
                     plngCodigoEmpresa, _
                     pstrComandoOperacao, _
                     vbNullString, _
                     penumStatusMensagem, _
                     plngIndicadorEntradaManual, _
                     pstrAgendamento, _
                     plngCodigoLocalLiquidacao, _
                     pstrCodigoVeiculoLegal, _
                     pstrSiglaSistema)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Cancelar a mensagem SPB gerada via entrada manual do SLCC
'--------------------------------------------------------------------------------------------------

Public Function CancelarEntradaManual(ByVal pstrNumeroControleIF As String, _
                                      ByVal pstrDTRRegistrMensagemSPB As String, _
                                      ByVal pintStatus As Integer, _
                                      ByVal pstrDataUltimaAtualizacao As String)
    

On Error GoTo ErrorHandler
    
    Call AlterarStatus(pstrNumeroControleIF, _
                       pstrDTRRegistrMensagemSPB, _
                       enumStatusMensagem.CanceladaOrigem, _
                       pstrDataUltimaAtualizacao)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarEntradaManual Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Incluir as informações da mensagem SPB enviada pelo SLCC
'--------------------------------------------------------------------------------------------------

Public Function Incluir(ByRef pstrMensagem As String, _
                        ByVal pstrNumeroControleIF As String, _
                        ByVal pvntSequenciaOperacao As Variant, _
                        ByVal pvntSequenciaConciliacaoOperacao As Variant, _
                        ByVal pstrCodigoMensagemSPB As String, _
                        ByVal penumTipoBackOffice As enumTipoBackOffice, _
                        ByVal plngCodigoEmpresa As Long, _
                        ByVal pstrComandoOperacao As String, _
                        ByVal pstrSituacaoMensagemSPB As String, _
                        ByVal penumStatusMensagem As enumStatusMensagem, _
                        ByVal plngIndicadorEntradaManual As Long, _
                        ByVal pstrAgendamento As String, _
                        ByVal plngCodigoLocalLiquidacao As Long, _
                        ByVal pstrCodigoVeiculoLegal As String, _
                        ByVal pstrSiglaSistema As String, _
               Optional ByVal plngSeqControleRepeticao As Long = 1, _
               Optional ByRef pstrDTRegistroMensagemSPB As String = vbNullString, _
               Optional ByVal pvntProtocoloAlcada As Variant = 0, _
               Optional ByVal pstrCodigoPartCamara As String = vbNullString, _
               Optional ByVal pstrNumeroControleCamara As String = vbNullString, _
               Optional ByVal pintCanalVenda As Integer = 0, _
               Optional ByVal pstrDataOperacaoCambioSisbacen As String = vbNullString)

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long
Dim strDataRegistroMensagem                 As String
Dim objHistSituacaoMensagem                 As A8LQS.clsHistSituacaoMensagem
Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strDtOpRet                              As String
Dim lngNumOpSelOr                           As Long
Dim strIdentdTitSEL                         As String
Dim strPU                                   As String
Dim strQtdTit                               As String
Dim strDtVenc                               As String

Dim intDebitoCredito                        As Integer
Dim lngContaCustodiaSelic                   As Long

    On Error GoTo ErrorHandler

    'Obter a data para gravação da mensagem e do historico
    If pstrDTRegistroMensagemSPB = "" Then
        strDataRegistroMensagem = flObterDataGravacao(pstrNumeroControleIF)
        pstrDTRegistroMensagemSPB = strDataRegistroMensagem
    Else
        strDataRegistroMensagem = pstrDTRegistroMensagemSPB
    End If

    lngCodigoTxtXML = fgInsertVarchar4000(pstrMensagem)

    strSQL = "INSERT INTO " & vbCrLf & _
             "  A8.TB_MESG_RECB_ENVI_SPB (" & vbCrLf & _
             "          NU_CTRL_IF,                    " & vbCrLf & _
             "          DH_REGT_MESG_SPB,              " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,             " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,             " & vbCrLf & _
             "          TP_BKOF,                       " & vbCrLf & _
             "          CO_EMPR,                       " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,         " & vbCrLf & _
             "          CO_MESG_SPB,                   " & vbCrLf & _
             "          NU_COMD_OPER,                  " & vbCrLf & _
             "          CO_SITU_MESG_SPB,              " & vbCrLf & _
             "          CO_TEXT_XML,                   " & vbCrLf

    strSQL = strSQL & _
             "          HO_ENVI_MESG_SPB,              " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,             " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,             " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,        " & vbCrLf & _
             "          DH_ULTI_ATLZ,                  " & vbCrLf & _
             "          IN_ENTR_MANU,                  " & vbCrLf & _
             "          CO_LOCA_LIQU,                  " & vbCrLf & _
             "          CO_VEIC_LEGA,                  " & vbCrLf & _
             "          SG_SIST,                       " & vbCrLf & _
             "          NU_SEQU_CNCL_OPER_ATIV_MESG,   " & vbCrLf & _
             "          NU_CTRL_CAMR,                  " & vbCrLf & _
             "          NU_PRTC_MESG_LG,               " & vbCrLf & _
             "          CO_PARP_CAMR ,                 " & vbCrLf & _
             "          TP_CNAL_VEND,                  " & vbCrLf & _
             "          DT_OPER_CAMB_SISBACEN )        " & vbCrLf
    
    
    strSQL = strSQL & _
             "  VALUES (" & _
             "'" & pstrNumeroControleIF & "', " & vbCrLf & _
             fgDtHrXML_To_Oracle(strDataRegistroMensagem) & ", " & vbCrLf & _
             plngSeqControleRepeticao & ", " & vbCrLf & _
             IIf(Val(pvntSequenciaOperacao) = 0, "NULL", pvntSequenciaOperacao) & ", " & vbCrLf & _
             IIf(penumTipoBackOffice = enumTipoBackOffice.Todos Or penumTipoBackOffice = 0, "Null", penumTipoBackOffice) & ", " & vbCrLf & _
             plngCodigoEmpresa & ", " & vbCrLf & _
             fgDataHoraServidor_To_Oracle & ", " & vbCrLf & _
             "'" & pstrCodigoMensagemSPB & "', " & vbCrLf & _
             "'" & pstrComandoOperacao & "', " & vbCrLf & _
             IIf(pstrSituacaoMensagemSPB = vbNullString, "Null", " '" & pstrSituacaoMensagemSPB & "'") & ", " & vbCrLf & _
             lngCodigoTxtXML & ", " & vbCrLf

    If Val(pstrAgendamento) <> 0 Then
       strSQL = strSQL & _
                fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(enumFormatoDataHora.DataHora), "YYYYMMDD") & _
                fgCompletaString(pstrAgendamento, "0", 4, True) & _
                "00") & ", " & vbCrLf
    Else
       strSQL = strSQL & _
             " NULL " & ", " & vbCrLf
    End If

    strSQL = strSQL & _
             penumStatusMensagem & ", " & vbCrLf & _
             " '" & fgUsuarioRede & "', " & vbCrLf & _
             " '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "Sysdate , " & vbCrLf & _
             plngIndicadorEntradaManual & " , " & vbCrLf & _
             IIf(plngCodigoLocalLiquidacao = 0, " NULL ", plngCodigoLocalLiquidacao) & " , " & vbCrLf & _
             IIf(pstrCodigoVeiculoLegal = vbNullString, " NULL ", " '" & pstrCodigoVeiculoLegal & "' ") & " , " & vbCrLf & _
             IIf(pstrSiglaSistema = vbNullString, " NULL ", " '" & pstrSiglaSistema & "' ") & " , " & vbCrLf & _
             IIf(Val(pvntSequenciaConciliacaoOperacao) = 0, "NULL", pvntSequenciaConciliacaoOperacao) & ", " & vbCrLf & _
             " '" & pstrNumeroControleCamara & "', " & vbCrLf & _
             pvntProtocoloAlcada & ", " & vbCrLf & _
             " '" & pstrCodigoPartCamara & "'," & vbCrLf & _
            pintCanalVenda & "," & vbCrLf & _
            IIf(pstrDataOperacaoCambioSisbacen = vbNullString, " NULL ", pstrDataOperacaoCambioSisbacen) & ")"

    fgExecuteSQL strSQL

    Set objHistSituacaoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    objHistSituacaoMensagem.Incluir pstrNumeroControleIF, _
                                    strDataRegistroMensagem, _
                                    penumStatusMensagem
    Set objHistSituacaoMensagem = Nothing

    'Gravação na tabela que será utilizada para conciliação
    If pstrCodigoMensagemSPB = "SEL1611" Then
        Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        xmlMensagem.loadXML pstrMensagem

        If xmlMensagem.documentElement.selectSingleNode("//DtMovto") Is Nothing Then
            strDtOpRet = vbNullString
        Else
            strDtOpRet = xmlMensagem.documentElement.selectSingleNode("//DtMovto").Text
        End If
        If xmlMensagem.documentElement.selectSingleNode("//NumOpSELOr") Is Nothing Then
            lngNumOpSelOr = 0
        Else
            lngNumOpSelOr = CLng("0" & xmlMensagem.documentElement.selectSingleNode("//NumOpSELOr").Text)
        End If
        If xmlMensagem.documentElement.selectSingleNode("//IdentdTitSEL") Is Nothing Then
            strIdentdTitSEL = vbNullString
        Else
            strIdentdTitSEL = xmlMensagem.documentElement.selectSingleNode("//IdentdTitSEL").Text
        End If
        If xmlMensagem.documentElement.selectSingleNode("//PU") Is Nothing Then
            strPU = vbNullString
        Else
            strPU = xmlMensagem.documentElement.selectSingleNode("//PU").Text
        End If
        If xmlMensagem.documentElement.selectSingleNode("//QtdTit") Is Nothing Then
            strQtdTit = vbNullString
        Else
            strQtdTit = xmlMensagem.documentElement.selectSingleNode("//QtdTit").Text
        End If
        If xmlMensagem.documentElement.selectSingleNode("//DtVenc") Is Nothing Then
            strDtVenc = vbNullString
        Else
            strDtVenc = xmlMensagem.documentElement.selectSingleNode("//DtVenc").Text
        End If
        
        intDebitoCredito = 0
        If Not xmlMensagem.documentElement.selectSingleNode("//TpDeb_Cred") Is Nothing Then
            intDebitoCredito = IIf(xmlMensagem.documentElement.selectSingleNode("//TpDeb_Cred").Text = "D", _
                                   enumTipoDebitoCredito.Debito, _
                                   enumTipoDebitoCredito.Credito)
        End If

        If Not xmlMensagem.documentElement.selectSingleNode("//CtCed") Is Nothing Then
            lngContaCustodiaSelic = CLng("0" & xmlMensagem.documentElement.selectSingleNode("//CtCed").Text)
        End If
        
        If intDebitoCredito = enumTipoDebitoCredito.Credito Then
            If Not xmlMensagem.documentElement.selectSingleNode("//CtCes") Is Nothing Then
                lngContaCustodiaSelic = CLng("0" & xmlMensagem.documentElement.selectSingleNode("//CtCes").Text)
            End If
        End If
        
        objConciliacao.Incluir pstrNumeroControleIF, _
                               strDataRegistroMensagem, _
                               strDtOpRet, _
                               xmlMensagem.documentElement.selectSingleNode("//CodOpSEL").Text, _
                               lngNumOpSelOr, _
                               strIdentdTitSEL, _
                               strPU, _
                               strQtdTit, _
                               xmlMensagem.documentElement.selectSingleNode("//VlrFinanc").Text, _
                               , , _
                               intDebitoCredito, _
                               strDtVenc, _
                               , , _
                               lngContaCustodiaSelic

        Set objConciliacao = Nothing
        Set xmlMensagem = Nothing
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objHistSituacaoMensagem = Nothing
    Set objConciliacao = Nothing
    Set xmlMensagem = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Alterar o horário de agendamento de envio da mensagem SPB
'--------------------------------------------------------------------------------------------------

Public Function AlterarAgendamento(ByVal pstrNumeroControleIF As String, _
                                   ByVal pstrDTRegistroMensagemSPB As String, _
                                   ByVal pstrHorarioAgendamento As String, _
                                   ByVal pstrDHUltimaAtualizacao As String, _
                                   ByVal plngCodigoMensagemXML As Long, _
                                   ByVal pstrCodigoMensagem As String, _
                                   ByVal pintStatusMensagem As enumStatusMensagem) As Boolean

Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim xmlGradeHorario                         As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim lngSequencialISPBDestino                As Long
Dim strHorarioAbertura                      As String
Dim strHorarioEncerramento                  As String

    On Error GoTo ErrorHandler
    
    If pstrHorarioAgendamento = vbNullString Then pstrHorarioAgendamento = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data)) & "000000"

    Set xmlGradeHorario = CreateObject("MSXML2.DOMDocument.4.0")
    
    '-------------------------------------------------------------------------------------------------------
    '>>> Verificar Grade de Horário
    '-------------------------------------------------------------------------------------------------------
    '    Captura ISPB Destino
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    lngSequencialISPBDestino = objMensagem.ObterSequencialISPBDestino(pstrCodigoMensagem, _
                                                                      plngCodigoMensagemXML)
    Set objMensagem = Nothing
    '-------------------------------------------------------------------------------------------------------
    '    Captura a Grade de Horário
    Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
    Call xmlGradeHorario.loadXML(objGradeHorario.ObterGradeHorario(pstrCodigoMensagem, lngSequencialISPBDestino))
    Set objGradeHorario = Nothing
    
    strHorarioAbertura = Mid(xmlGradeHorario.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioAbertura").Text, 9, 6)
    strHorarioEncerramento = Mid(xmlGradeHorario.selectSingleNode("Repeat_GradeHorario/Grupo_GradeHorario/HorarioEncerramento").Text, 9, 6)
    '-------------------------------------------------------------------------------------------------------
    'Verifica se o horário está fora da Grade
    If TimeSerial(Mid(pstrHorarioAgendamento, 9, 2), Mid(pstrHorarioAgendamento, 11, 2), Mid(pstrHorarioAgendamento, 13, 2)) < _
       TimeSerial(Mid(strHorarioAbertura, 1, 2), Mid(strHorarioAbertura, 3, 2), Mid(strHorarioAbertura, 5, 2)) Or _
       TimeSerial(Mid(pstrHorarioAgendamento, 9, 2), Mid(pstrHorarioAgendamento, 11, 2), Mid(pstrHorarioAgendamento, 13, 2)) > _
       TimeSerial(Mid(strHorarioEncerramento, 1, 2), Mid(strHorarioEncerramento, 3, 2), Mid(strHorarioEncerramento, 5, 2)) Then
       
        'Agendamento fora da Grade de Horário estipulada.
        lngCodigoErroNegocio = 3020
        GoTo ErrorHandler
    End If
    
    Set xmlGradeHorario = Nothing
    
    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB                                                         " & vbNewLine & _
             " Set      HO_ENVI_MESG_SPB        = " & fgDtHrXML_To_Oracle(pstrHorarioAgendamento) & ",   " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "',                               " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "',                    " & vbNewLine & _
             "          DH_ULTI_ATLZ            = Sysdate                                                " & vbNewLine & _
             " Where    NU_CTRL_IF              = '" & pstrNumeroControleIF & "'                         " & vbNewLine & _
             " And      DH_REGT_MESG_SPB        = " & fgDtHrXML_To_Oracle(pstrDTRegistroMensagemSPB) & " " & vbNewLine & _
             " And      DH_ULTI_ATLZ            = " & fgDtHrXML_To_Oracle(pstrDHUltimaAtualizacao)

    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
    End If

    'Armazenar o Histórico do Agendamento da Mensagem
    Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    objHistoricoMensagem.Incluir pstrNumeroControleIF, _
                                 pstrDTRegistroMensagemSPB, _
                                 pintStatusMensagem, _
                                 enumTipoAcao.AlteracaoHorarioAgendamento, _
                                 pstrHorarioAgendamento
    Set objHistoricoMensagem = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objHistoricoMensagem = Nothing
    Set objGradeHorario = Nothing
    Set objMensagem = Nothing
    Set xmlGradeHorario = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Confirmar a mensagem SPB deixando a mesma com o status de concordância
'--------------------------------------------------------------------------------------------------

Public Function Confirmar(ByVal pstrNumeroControleIF As String, _
                          ByVal pstrDTRegistroMensagemSPB As String, _
                          ByVal pintStatusMensagem As enumStatusMensagem, _
                          ByVal pstrDHUltimaAtualizacao As String, _
                          ByVal pvntProtocolo As Variant) As Boolean
    
On Error GoTo ErrorHandler
   
    Call AlterarStatus(pstrNumeroControleIF, _
                       pstrDTRegistroMensagemSPB, _
                       enumStatusMensagem.Concordancia, _
                       pstrDHUltimaAtualizacao)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function
    
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Confirmar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Ler o xml da mensagem SPB
'--------------------------------------------------------------------------------------------------

Public Function ObterMensagemXML(ByVal pstrNumeroControleIF As String, _
                        Optional ByRef pstrDataHoraAgendamento As String) As String

Dim vntCodTextXml                           As Variant
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String

    strMensagem = Ler(pstrNumeroControleIF)

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMensagem.loadXML(strMensagem) Then
        Call fgErroLoadXML(xmlMensagem, App.EXEName, TypeName(Me), "ObterMensagemXML")
    End If

    vntCodTextXml = xmlMensagem.selectSingleNode("//CO_TEXT_XML").Text
    ObterMensagemXML = fgSelectVarchar4000(vntCodTextXml)

    pstrDataHoraAgendamento = xmlMensagem.selectSingleNode("//HO_ENVI_MESG_SPB").Text

End Function

'--------------------------------------------------------------------------------------------------
'Alterar o tipo de Backoffice da mensagem SPB
'--------------------------------------------------------------------------------------------------

Public Function AlterarTipoBackOffice(ByVal pstrNumeroControleIF As String, _
                                      ByVal plngTipoBackOffice As Long, _
                             Optional ByVal pvntNumeroSequenciaRepeticao As Variant = 0, _
                             Optional ByVal pstrDataRegistroMensagem As String = "") As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB " & vbNewLine & _
             " Set      TP_BKOF                 =   " & plngTipoBackOffice & ", " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate " & vbNewLine & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'" & vbNewLine

    If pvntNumeroSequenciaRepeticao <> 0 Then
        strSQL = strSQL & " AND   NU_SEQU_CNTR_REPE   = " & pvntNumeroSequenciaRepeticao & vbNewLine
    End If
    
    If pstrDataRegistroMensagem <> "" Then
        strSQL = strSQL & " AND      DH_REGT_MESG_SPB            =  " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbNewLine
    End If
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarTipoBackOffice Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Alterar o número da sequencia de conciliação de uma mensagem SPB após a sua conciliação com a operação
'--------------------------------------------------------------------------------------------------

Public Function AlterarSequenciaConciliacao(ByVal pstrNumeroControleIF As String, _
                                            ByVal pdatDTRegistroMensagemSPB As Date, _
                                            ByVal pvntSequenciaConciliacaoOperacao As Variant, _
                                            ByVal plngNumeroSequenciaRepeticao As Long, _
                                   Optional ByVal pintTipoAcao As Integer = 0) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB                                                  " & vbNewLine & _
             " Set      NU_SEQU_CNCL_OPER_ATIV_MESG =  " & pvntSequenciaConciliacaoOperacao & ",  " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ           = '" & fgUsuarioRede & "',                    " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ      = '" & fgEstacaoTrabalhoUsuario & "',         " & vbNewLine & _
             "          DH_ULTI_ATLZ                =      Sysdate                                " & vbNewLine

    If pintTipoAcao <> 0 Then
        strSQL = strSQL & _
             ",         TP_ACAO_MESG_SPB_EXEC       =  " & pintTipoAcao & "                       " & vbNewLine
    Else
        strSQL = strSQL & _
             ",         TP_ACAO_MESG_SPB_EXEC       =  NULL                                       " & vbNewLine
    End If
             
    strSQL = strSQL & _
             " Where    NU_CTRL_IF                  = '" & pstrNumeroControleIF & "'              " & vbNewLine & _
             " And      DH_REGT_MESG_SPB            =  " & fgDtHrXML_To_Oracle(fgDtHr_To_Xml(pdatDTRegistroMensagemSPB)) & vbNewLine
    
    If plngNumeroSequenciaRepeticao <> 0 Then
        strSQL = strSQL & _
             " And      NU_SEQU_CNTR_REPE       =   " & plngNumeroSequenciaRepeticao
    End If
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarSequenciaConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Alterar status de uma mensagem SPB
'--------------------------------------------------------------------------------------------------

Public Function AlterarStatus(ByVal pstrNumeroControleIF As String, _
                              ByVal pstrDTRegistroMensagemSPB As String, _
                              ByVal pintStatusMensagem As enumStatusMensagem, _
                              ByVal pstrDHUltimaAtualizacao As String, _
                     Optional ByVal pstrDataRecebEnvioMesgSPB As String = vbNullString, _
                     Optional ByVal pblnComparaDHUltimaAtualizacao As Boolean = True, _
                     Optional ByVal plngNumeroSequenciaContadorRepeticao As Long = 1, _
                     Optional ByVal pstrCodigoMensagem As String = vbNullString, _
                     Optional ByVal plngTipoAcao As Long = 0, _
                     Optional ByVal pvntNumeroSequenciaConci As Variant = 0, _
                     Optional ByVal pblnAlterarAcao As Boolean = True, _
                     Optional ByVal plngNumeroSequencialOperacao As Long = 0, _
                     Optional ByVal pstrCodigoVeiculoLegal As String = vbNullString, _
                     Optional ByVal pstrSiglaSistema As String = vbNullString, _
                     Optional ByVal pblnIncluirHistorico As Boolean = True) As Boolean

Dim strSQL                                  As String
Dim objHistoricoMensagem                    As A8LQS.clsHistSituacaoMensagem

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB                                                           " & vbNewLine & _
             " Set      CO_ULTI_SITU_PROC       =   " & pintStatusMensagem & ",                            " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "',                               " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "',                    " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate                                                " & vbNewLine
             
    If pstrCodigoMensagem <> vbNullString Then
        strSQL = strSQL & _
            " ,         CO_MESG_SPB   = '" & pstrCodigoMensagem & "'  " & vbNewLine
    End If
             
    If pvntNumeroSequenciaConci <> 0 Then
        strSQL = strSQL & _
            " ,         NU_SEQU_CNCL_OPER_ATIV_MESG   = " & pvntNumeroSequenciaConci & "  " & vbNewLine
    End If
             
    If plngNumeroSequencialOperacao <> 0 Then
        strSQL = strSQL & _
            " ,         NU_SEQU_OPER_ATIV   = " & plngNumeroSequencialOperacao & "  " & vbNewLine
    End If
             
    If pstrDataRecebEnvioMesgSPB <> vbNullString Then
        strSQL = strSQL & _
            " ,         DH_RECB_ENVI_MESG_SPB   = " & fgDtHrXML_To_Oracle(pstrDataRecebEnvioMesgSPB) & "    " & vbNewLine
    End If
    
    If pstrCodigoVeiculoLegal <> vbNullString Then
        strSQL = strSQL & _
            " ,         CO_VEIC_LEGA   = '" & pstrCodigoVeiculoLegal & "'  " & vbNewLine
    End If
             
    If pstrSiglaSistema <> vbNullString Then
        strSQL = strSQL & _
            " ,         SG_SIST   = '" & pstrSiglaSistema & "'  " & vbNewLine
    End If
             
    If pblnAlterarAcao Then
        If plngTipoAcao <> 0 Then
            strSQL = strSQL & _
                " ,         TP_ACAO_MESG_SPB_EXEC  =  " & plngTipoAcao & "                            " & vbNewLine
        Else
            strSQL = strSQL & _
                " ,         TP_ACAO_MESG_SPB_EXEC  =  NULL                                            " & vbNewLine
        End If
    End If
             
    strSQL = strSQL & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'                           " & vbNewLine & _
             " And      DH_REGT_MESG_SPB        =   " & fgDtHrXML_To_Oracle(pstrDTRegistroMensagemSPB) & "  "
    
    If plngNumeroSequenciaContadorRepeticao <> 0 Then
        strSQL = strSQL & _
             " And      NU_SEQU_CNTR_REPE       =   " & plngNumeroSequenciaContadorRepeticao & "            " & vbNewLine
    End If

    If pblnComparaDHUltimaAtualizacao Then
        strSQL = strSQL & vbNewLine & _
             " And      DH_ULTI_ATLZ            =   " & fgDtHrXML_To_Oracle(pstrDHUltimaAtualizacao)
    End If
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
    End If

    If pblnIncluirHistorico Then
        'Armazenar o Histórico da Situacao da Operação
        Set objHistoricoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
        objHistoricoMensagem.Incluir pstrNumeroControleIF, _
                                     pstrDTRegistroMensagemSPB, _
                                     pintStatusMensagem, _
                                     plngTipoAcao, _
                                     , _
                                     plngNumeroSequenciaContadorRepeticao
                                     
        Set objHistoricoMensagem = Nothing
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    
    Set objHistoricoMensagem = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Date:" & pstrDTRegistroMensagemSPB)

End Function

'--------------------------------------------------------------------------------------------------
'Alterar dados da tabela de mensagem
'--------------------------------------------------------------------------------------------------

Public Function Alterar(ByVal pstrNumeroControleIF As String, _
                        ByVal pstrDTRegistroMensagemSPB As String, _
                        ByVal plngNumeroSequenciaContadorRepeticao As Long, _
               Optional ByVal pintStatusMensagem As Long = 0, _
               Optional ByVal pstrDataRecebEnvioMesgSPB As String = vbNullString, _
               Optional ByVal pstrCodigoMensagem As String = vbNullString, _
               Optional ByVal plngTipoAcao As Long = 0, _
               Optional ByVal pvntNumeroSequenciaConci As Variant = 0, _
               Optional ByVal pblnAlterarAcao As Boolean = True, _
               Optional ByVal plngNumeroSequencialOperacao As Long = 0, _
               Optional ByVal pstrCodigoVeiculoLegal As String = vbNullString, _
               Optional ByVal pstrSiglaSistema As String = vbNullString, _
               Optional ByVal plngTipoBackOffice As Long = 0) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB                                                           " & vbNewLine & _
             " Set      CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "',                               " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "',                    " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate                                                " & vbNewLine
             
    If pintStatusMensagem <> 0 Then
        strSQL = strSQL & _
            " ,         CO_ULTI_SITU_PROC   = " & pintStatusMensagem & "  " & vbNewLine
    End If
             
    If pstrCodigoMensagem <> vbNullString Then
        strSQL = strSQL & _
            " ,         CO_MESG_SPB   = '" & pstrCodigoMensagem & "'  " & vbNewLine
    End If
             
    If pvntNumeroSequenciaConci <> 0 Then
        strSQL = strSQL & _
            " ,         NU_SEQU_CNCL_OPER_ATIV_MESG   = " & pvntNumeroSequenciaConci & "  " & vbNewLine
    End If
             
    If plngNumeroSequencialOperacao <> 0 Then
        strSQL = strSQL & _
            " ,         NU_SEQU_OPER_ATIV   = " & plngNumeroSequencialOperacao & "  " & vbNewLine
    End If
             
    If pstrDataRecebEnvioMesgSPB <> vbNullString Then
        strSQL = strSQL & _
            " ,         DH_RECB_ENVI_MESG_SPB   = " & fgDtHrXML_To_Oracle(pstrDataRecebEnvioMesgSPB) & "    " & vbNewLine
    End If
    
    If pstrCodigoVeiculoLegal <> vbNullString Then
        strSQL = strSQL & _
            " ,         CO_VEIC_LEGA   = '" & pstrCodigoVeiculoLegal & "'  " & vbNewLine
    End If
             
    If pstrSiglaSistema <> vbNullString Then
        strSQL = strSQL & _
            " ,         SG_SIST   = '" & pstrSiglaSistema & "'  " & vbNewLine
    End If
             
    If pblnAlterarAcao Then
        If plngTipoAcao <> 0 Then
            strSQL = strSQL & _
                " ,         TP_ACAO_MESG_SPB_EXEC  =  " & plngTipoAcao & "                            " & vbNewLine
        Else
            strSQL = strSQL & _
                " ,         TP_ACAO_MESG_SPB_EXEC  =  NULL                                            " & vbNewLine
        End If
    End If
             
    If plngTipoBackOffice <> 0 Then
        strSQL = strSQL & _
            " ,         TP_BKOF   = " & plngTipoBackOffice & "  " & vbNewLine
    End If
             
    strSQL = strSQL & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'                           " & vbNewLine & _
             " And      DH_REGT_MESG_SPB        =   " & fgDtHrXML_To_Oracle(pstrDTRegistroMensagemSPB) & "  "
    
    If plngNumeroSequenciaContadorRepeticao <> 0 Then
        strSQL = strSQL & _
             " And      NU_SEQU_CNTR_REPE       =   " & plngNumeroSequenciaContadorRepeticao & "            " & vbNewLine
    End If

    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Alterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Alterar dados da tabela de Mensagem / Conciliação
'--------------------------------------------------------------------------------------------------

Public Function AlterarTabelaConciliacao(ByVal pstrNumeroControleIF As String, _
                                         ByVal pstrDTRegistroMensagemSPB As String, _
                                         ByVal pstrNumeroSequenciaContadorRepeticao As String, _
                                Optional ByVal pstrCanalOperacaoInterna As String = vbNullString) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " UPDATE   A8.TB_MESG_RECB_SPB_CNCL                                                           " & vbNewLine & _
             " SET"
             
    If pstrCanalOperacaoInterna <> vbNullString Then
        strSQL = strSQL & _
            ",          CO_CNAL_OPER_INTE   = '" & pstrCanalOperacaoInterna & "'  " & vbNewLine
    End If
             
    strSQL = Replace$(strSQL, "SET,", "SET")
    
    strSQL = strSQL & _
             " WHERE    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'                           " & vbNewLine & _
             " AND      DH_REGT_MESG_SPB        =   " & fgDtHrXML_To_Oracle(pstrDTRegistroMensagemSPB) & "  " & vbNewLine & _
             " AND      NU_SEQU_CNTR_REPE       =   " & pstrNumeroSequenciaContadorRepeticao & "            " & vbNewLine

    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarTabelaConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar a aprovação de alçada

Private Sub flAprovarAlcada(ByRef pxmlremessa As MSXML2.IXMLDOMNode)

Dim objPerfil                               As A6A7A8.clsPerfil

    On Error GoTo ErrorHandler

    Set objPerfil = CreateObject("A6A7A8.clsPerfil")
    Call objPerfil.AprovarAlcada(fgUsuarioRede, _
                                 pxmlremessa.selectSingleNode("ValorMensagem").Text, _
                                 pxmlremessa.selectSingleNode("ProtocoloAlcada").Text)
    Set objPerfil = Nothing

    Exit Sub

ErrorHandler:
    Set objPerfil = Nothing
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAprovarAlcada", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Controla a liberação mensagem STR para o SPB.

Public Function LiberarMensagemControleAlcada(ByRef xmlMensagem As MSXML2.IXMLDOMNode, _
                          Optional ByVal pblnIgnoraGradeHorario As Boolean = False, _
                          Optional ByVal pxmlOperacao As MSXML2.DOMDocument40) As Boolean

Dim xmlMensagemEnvio                        As MSXML2.DOMDocument40
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim objOperacao                             As A8LQS.clsOperacao
Dim objWorkflow                             As A8LQS.clsWorkflow

Dim strMensagemEnvio                        As String

    On Error GoTo ErrorHandler

    If pblnIgnoraGradeHorario = False Then
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlMensagem.selectSingleNode("CodigoMensagem").Text, 0, enumLocalLiquidacao.BMA) Then
             '3023 - Horário limite para envio da Operação excedido.
            lngCodigoErroNegocio = 3023
            GoTo ErrorHandler
        End If
    End If

    Call flAprovarAlcada(xmlMensagem)
    
    Set xmlMensagemEnvio = LerXMLMensagem(xmlMensagem.selectSingleNode("NumeroControleIF").Text)
    If Not xmlMensagemEnvio.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val("0" & xmlMensagemEnvio.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemLQS.EnvioPagDespesas Then
            fgAppendNode xmlMensagemEnvio, "SISMSG", "TP_CONT_DEBT", pxmlOperacao.documentElement.selectSingleNode("TP_CONT_DEBT").Text
            fgAppendNode xmlMensagemEnvio, "SISMSG", "NU_CONT_DEBT", pxmlOperacao.documentElement.selectSingleNode("NU_CONT_DEBT").Text
            fgAppendNode xmlMensagemEnvio, "SISMSG", "TP_PESS_DEBT", pxmlOperacao.documentElement.selectSingleNode("TP_PESS_DEBT").Text
            fgAppendNode xmlMensagemEnvio, "SISMSG", "CO_CNPJ_CPF_DEBT_1", pxmlOperacao.documentElement.selectSingleNode("CO_CNPJ_CPF_DEBT_1").Text
            fgAppendNode xmlMensagemEnvio, "SISMSG", "NO_TITU_DEBT_1", pxmlOperacao.documentElement.selectSingleNode("NO_TITU_DEBT_1").Text
            fgAppendNode xmlMensagemEnvio, "SISMSG", "CO_AGEN_DEBT", pxmlOperacao.documentElement.selectSingleNode("CO_AGEN_DEBT").Text
        
        End If
    End If
    
    strMensagemEnvio = xmlMensagemEnvio.xml
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    
    If Mid(xmlMensagemEnvio.selectSingleNode("//CO_MESG").Text, 1, 3) = "STR" Or _
       Mid(xmlMensagemEnvio.selectSingleNode("//CO_MESG").Text, 1, 3) = "PAG" Then
        strMensagemEnvio = objMensagem.MontarMensagemPZ(xmlMensagemEnvio, _
                                              IIf(Trim$(xmlMensagemEnvio.selectSingleNode("//TP_BKOF").Text) = "", 0, xmlMensagemEnvio.selectSingleNode("//TP_BKOF").Text))
    Else
        strMensagemEnvio = objMensagem.MontarMensagemPorMensagemSPB(xmlMensagemEnvio, _
                                                                    True, _
                                                                    xmlMensagem.selectSingleNode("NumeroControleIF").Text)
    End If
            
    Set objMensagem = Nothing
    
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    objProcessoOperacao.EnviarMensagemMQ strMensagemEnvio, enumIdentificadorFila.BUS, False, True
    Set objProcessoOperacao = Nothing

    Call AlterarStatus(xmlMensagem.selectSingleNode("NumeroControleIF").Text, _
                       xmlMensagem.selectSingleNode("DTRegistroMensagemSPB").Text, _
                       enumStatusMensagem.EnviadaBUS, _
                       xmlMensagem.selectSingleNode("DHUltimaAtualizacao").Text, _
                       , _
                       True, _
                       , _
                       , _
                       enumTipoAcao.Liberacao)
    
    If Not xmlMensagemEnvio.selectSingleNode("//TP_MESG") Is Nothing Then
        If Val(xmlMensagemEnvio.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioTEDClientes Then
        
            Set objOperacao = CreateObject("A8LQS.clsOperacao")
            Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
            
            Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMensagemEnvio.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)
    
            Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.PagamentoLib, _
                                               vbNullString, _
                                               , , _
                                               False)
                                               
            xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.PagamentoLib
            
            Call objWorkflow.GerenciarChamadas(xmlOperacao, "TEDLiberacaoAlcada")
            
            Set objOperacao = Nothing
            Set objWorkflow = Nothing
            Set xmlOperacao = Nothing
        
        End If
        
'Pagamento de Despesas
        If Not xmlMensagemEnvio.selectSingleNode("//TP_MESG") Is Nothing Then
            If Val(xmlMensagemEnvio.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
                If (Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0006PagDespesas Or _
                Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0008PagDespesas Or _
                Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasIsenta Or _
                Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0006PagDespesasTrib Or _
                Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasIsenta Or _
                Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0008PagDespesasTrib) Then

                    Set objOperacao = CreateObject("A8LQS.clsOperacao")
                    Set objWorkflow = CreateObject("A8LQS.clsWorkflow")

                    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMensagemEnvio.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

                    Call objOperacao.AlterarStatusAcao(xmlOperacao.selectSingleNode("//NU_SEQU_OPER_ATIV").Text, _
                                               enumStatusOperacao.Liberada, _
                                               vbNullString, _
                                               , , _
                                               False)

                    xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada

                    Call objWorkflow.GerenciarChamadas(xmlOperacao, "TEDLiberacaoAlcada")
                
                
                'ElseIf Val(xmlMensagemEnvio.selectSingleNode("//TP_MESG").Text) = enumTipoMensagemBUS.EnvioPagDespesas Then
                
                 '   If Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0007PagDespesas Or _
                 '       Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasIsenta Or _
                 '       Val(xmlMensagemEnvio.selectSingleNode("//TP_OPER").Text) = enumTipoOperacaoLQS.EnvioSTR0007PagDespesasTrib Then
                 '       Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")
                 '       xmlOperacao.loadXML (xmlMensagemEnvio.xml)
                 '       xmlOperacao.selectSingleNode("//CO_ULTI_SITU_PROC").Text = enumStatusOperacao.Liberada
                 '       fgAppendNode xmlOperacao, "SISMSG", "MESG", ""
                 '
                 '       fgAppendNode xmlOperacao, "SISMSG", "IN_SIST_SITU_CNTG", enumIndicadorSimNao.Nao
                 '
                 '       Set objWorkflow = CreateObject("A8LQS.clsWorkflow")
                 '
                 '       Call objWorkflow.GerenciarChamadas(xmlOperacao, "TEDLiberacaoAlcada")
                 '   End If
                End If
            End If
            Set objOperacao = Nothing
            Set objWorkflow = Nothing
            Set xmlOperacao = Nothing
        End If
    
    End If
    
    Set xmlMensagemEnvio = Nothing
    Set objGradeHorario = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagemEnvio = Nothing
    Set objGradeHorario = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarMensagemControleAlcada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Controlar a liberação do envio de uma mensagem SPB verificando a grade de horário
'--------------------------------------------------------------------------------------------------

Public Function Liberar(ByVal pstrNumeroControleIF As String, _
                        ByVal pstrDTRegistroMensagemSPB As String, _
                        ByVal pvntNumeroSequencia As Variant, _
                        ByVal pintStatusMensagem As Integer, _
                        ByVal pstrDHUltimaAtualizacao As String, _
                        ByVal pblnIgnoraGradeHorario As Boolean)

Dim strHeaderXMLMensagem                    As String
Dim strXMLMensagem                          As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim strProtocolo                            As String
Dim intStatusMensagem                       As Integer
Dim strDataHoraAgendamento                  As String
Dim strDataHoraLocal                        As String
Dim blnEnviar                               As Boolean
Dim strCodMsg                               As String

On Error GoTo ErrorHandler

    strHeaderXMLMensagem = ObterMensagemXML(pstrNumeroControleIF, strDataHoraAgendamento)
    strXMLMensagem = Mid(strHeaderXMLMensagem, 221)
    strProtocolo = Mid(strHeaderXMLMensagem, 1, 20)

    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlDomLeitura.loadXML(strXMLMensagem)
    
    If Not xmlDomLeitura.selectSingleNode("//CodMsg") Is Nothing Then
        strCodMsg = xmlDomLeitura.selectSingleNode("//CodMsg").Text
    Else
        If Not xmlDomLeitura.selectSingleNode("//CO_MESG") Is Nothing Then
            strCodMsg = xmlDomLeitura.selectSingleNode("//CO_MESG").Text
        End If
    End If

    If strCodMsg = "STR0011" Then
        pblnIgnoraGradeHorario = True
    End If
    
    If Not pblnIgnoraGradeHorario Then
        'Verificar a Grade de Horário
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        Set objMensagem = CreateObject("A8LQS.clsMensagem")
        
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(strCodMsg, _
                                                                  objMensagem.ObterSequencialISPBDestino(strCodMsg, _
                                                                                                         , _
                                                                                                         , _
                                                                                                         strXMLMensagem)) Then
            ' 3023 - Horário limite para envio da Operação excedido.
            lngCodigoErroNegocio = 3023
            GoTo ErrorHandler
        End If

        Set objGradeHorario = Nothing
        Set objMensagem = Nothing
    End If

    strDataHoraLocal = fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora))

    If strDataHoraLocal > strDataHoraAgendamento Then
        intStatusMensagem = enumStatusMensagem.EnviadaBUS
        blnEnviar = True
    Else
        intStatusMensagem = enumStatusMensagem.MensagemAgendada
        blnEnviar = False
    End If

    'Pikachu - 17/01/2005
    'Mensageria CETIP - CTP0100
    If Not xmlDomLeitura.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If xmlDomLeitura.selectSingleNode("//CO_LOCA_LIQU").Text = enumLocalLiquidacao.CETIP Then
            Set objMensagem = CreateObject("A8LQS.clsMensagem")
            Call objMensagem.MontarMensagemCETIP(xmlDomLeitura)
            strXMLMensagem = xmlDomLeitura.xml
            Set objMensagem = Nothing
        End If
    End If

    If strCodMsg = "SEL1069" _
    Or strCodMsg = "SEL1066" Then
        strHeaderXMLMensagem = strProtocolo & xmlDomLeitura.xml
    End If

    If Not xmlDomLeitura.selectSingleNode("//COM_REGRA") Is Nothing Then
        If xmlDomLeitura.selectSingleNode("//COM_REGRA").Text = "S" Then
            strHeaderXMLMensagem = Mid(strHeaderXMLMensagem, 1, 20) & strXMLMensagem
        End If
    End If

    If blnEnviar Then
        Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
        objProcessoOperacao.EnviarMensagemMQ strHeaderXMLMensagem, enumIdentificadorFila.BUS, False, True
        Set objProcessoOperacao = Nothing

        'Alterar o status/ação da operação
        flVerificarOperacao xmlDomLeitura

    End If

    Set xmlDomLeitura = Nothing
    
    Call AlterarStatus(pstrNumeroControleIF, _
                       pstrDTRegistroMensagemSPB, _
                       intStatusMensagem, _
                       pstrDHUltimaAtualizacao, _
                       fgDtHr_To_Xml(fgDataHoraServidor(enumFormatoDataHora.DataHora)))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:
    
    Set objProcessoOperacao = Nothing
    Set objGradeHorario = Nothing
    Set objMensagem = Nothing
    Set xmlDomLeitura = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flLiberar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Controlar a liberacao da mensagem BMC0015 em Contingencia verificando a grade de horário
'--------------------------------------------------------------------------------------------------

Public Function LiberarContingencia(ByVal pstrNumeroControleIF As String, _
                                    ByVal pstrDTRegistroMensagemSPB As String, _
                                    ByVal pvntNumeroSequencia As Variant, _
                                    ByVal pintStatusMensagem As Integer, _
                                    ByVal pstrDHUltimaAtualizacao As String, _
                                    ByVal pblnIgnoraGradeHorario As Boolean)

Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim objMensagem                             As A8LQS.clsMensagem
Dim strHeaderXMLMensagem                    As String
Dim strXMLMensagem                          As String
Dim strProtocolo                            As String
Dim strDataHoraAgendamento                  As String
Dim lngStatusMensagem                       As Long
Dim blnAux                                  As Boolean

On Error GoTo ErrorHandler

    Set xmlDomLeitura = LerXMLMensagem(Trim$(pstrNumeroControleIF), , , "BMC0015")

    If Not pblnIgnoraGradeHorario Then
        'Verificar a Grade de Horário
        Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
        Set objMensagem = CreateObject("A8LQS.clsMensagem")
        
        If Not objGradeHorario.VerificarGradeHorarioEnvioMensagem(xmlDomLeitura.selectSingleNode("//CO_MESG").Text, _
                                                                  objMensagem.ObterSequencialISPBDestino(xmlDomLeitura.selectSingleNode("//CO_MESG").Text, , , _
                                                                                                         strXMLMensagem)) Then
            ' 3023 - Horário limite para envio da Operação excedido.
            lngCodigoErroNegocio = 3023
            GoTo ErrorHandler
        End If

        Set objGradeHorario = Nothing
        Set objMensagem = Nothing
    End If

    'Atualiza status da mensagem BMC0015
    lngStatusMensagem = enumStatusMensagem.ConciliadaContingencia
    
    Call AlterarStatus(xmlDomLeitura.selectSingleNode("//NU_CTRL_IF").Text, _
                       xmlDomLeitura.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                       lngStatusMensagem, _
                       xmlDomLeitura.selectSingleNode("//DH_ULTI_ATLZ").Text, _
                       , _
                       True, _
                       xmlDomLeitura.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)

    'Envia Mensagem CAM0002
    Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
    Call objProcessoOperacao.GerenciarChamadaViaTela(xmlDomLeitura.selectSingleNode("//NU_CTRL_IF").Text, _
                                                     enumFuncaoSistema.Liberar, , , _
                                                     True, , _
                                                     xmlDomLeitura.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)


    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Set objProcessoOperacao = Nothing
    Set xmlDomLeitura = Nothing

    Exit Function
ErrorHandler:
    
    Set objProcessoOperacao = Nothing
    Set objGradeHorario = Nothing
    Set objMensagem = Nothing
    Set xmlDomLeitura = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarlContingencia Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Obter os valores das mensagens SPB disponíveis para conciliação
'--------------------------------------------------------------------------------------------------

Public Function ObterValoresConciliacao(ByVal pstrXMLDataOperacao As String, _
                                        ByVal plngComandoOperacao As Long, _
                                        ByVal plngCodigoOperacaoSelic As Long, _
                                        ByVal pblnComandoOperacaoMsg As Boolean, _
                                        ByVal plngCodigoEmpresa As Long, _
                               Optional ByVal pstrIdentificadorTituloSelic As String, _
                               Optional ByVal pstrDataVencimento As String = vbNullString, _
                               Optional ByVal pstrContaCedente As String = vbNullString, _
                               Optional ByVal pvntValor As Variant = 0, _
                               Optional ByVal plngQuantidade As Long = 0, _
                               Optional ByVal pintIndicadorDebitoCredito As Integer = 0) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    If pblnComandoOperacaoMsg Then
        strSQL = "  SELECT  /*+ INDEX(b) */ B.NU_CTRL_IF, " & vbCrLf & _
                 "          B.DH_REGT_MESG_SPB, " & vbCrLf & _
                 "          B.DT_OPER, " & vbCrLf & _
                 "          B.NU_COMD_OPER_ORIG, " & vbCrLf & _
                 "          B.CO_OPER_SELIC, " & vbCrLf & _
                 "          B.VA_FINC, " & vbCrLf & _
                 "          B.PU_ATIV_MERC, " & vbCrLf & _
                 "          B.QT_ATIV_MERC, " & vbCrLf & _
                 "          B.NU_ATIV_MERC, " & vbCrLf & _
                 "          B.DT_VENC, " & vbCrLf & _
                 "          B.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbCrLf & _
                 "          A.NU_COMD_OPER, " & vbCrLf & _
                 "          A.DH_ULTI_ATLZ " & vbCrLf & _
                 "   FROM   A8.TB_MESG_RECB_ENVI_SPB    A, " & vbCrLf & _
                 "          A8.TB_MESG_RECB_SPB_CNCL    B  " & vbCrLf & _
                 "  WHERE   A.NU_COMD_OPER          =   '" & plngComandoOperacao & "'" & vbCrLf & _
                 "    AND   A.CO_EMPR               =    " & plngCodigoEmpresa & vbCrLf & _
                 "    AND   A.CO_ULTI_SITU_PROC     =    " & enumStatusMensagem.AConciliar & vbCrLf & _
                 "    AND   A.NU_CTRL_IF            =   B.NU_CTRL_IF " & vbCrLf & _
                 "    AND   A.DH_REGT_MESG_SPB      =   B.DH_REGT_MESG_SPB " & vbCrLf & _
                 "    AND   A.NU_SEQU_CNTR_REPE     =   B.NU_SEQU_CNTR_REPE " & vbCrLf & _
                 "    AND   B.DT_OPER               =    " & fgDtXML_To_Oracle(pstrXMLDataOperacao) & vbCrLf & _
                 "    AND   B.CO_OPER_SELIC     =    " & plngCodigoOperacaoSelic
    Else
        strSQL = "  SELECT  /*+ INDEX(a) */ A.NU_CTRL_IF,        " & vbCrLf & _
                 "          A.DH_REGT_MESG_SPB,  " & vbCrLf & _
                 "          A.DT_OPER,           " & vbCrLf & _
                 "          A.NU_COMD_OPER_ORIG, " & vbCrLf & _
                 "          A.CO_OPER_SELIC, " & vbCrLf & _
                 "          A.VA_FINC,           " & vbCrLf & _
                 "          A.PU_ATIV_MERC,      " & vbCrLf & _
                 "          A.QT_ATIV_MERC,      " & vbCrLf & _
                 "          A.NU_ATIV_MERC, " & vbCrLf & _
                 "          A.DT_VENC, " & vbCrLf & _
                 "          A.CO_CNTA_CUTD_SELIC_VEIC_LEGA, " & vbCrLf & _
                 "          B.NU_COMD_OPER, " & vbCrLf & _
                 "          B.DH_ULTI_ATLZ " & vbCrLf & _
                 "   FROM   A8.TB_MESG_RECB_SPB_CNCL    A, " & vbCrLf & _
                 "          A8.TB_MESG_RECB_ENVI_SPB    B  " & vbCrLf & _
                 "  WHERE   A.DT_OPER               =    " & fgDtXML_To_Oracle(pstrXMLDataOperacao) & vbCrLf
        
        If plngComandoOperacao <> 0 Then
            strSQL = strSQL & "    AND   A.NU_COMD_OPER_ORIG    =    " & plngComandoOperacao & vbCrLf
        End If
        
        If pstrIdentificadorTituloSelic <> vbNullString Then
            strSQL = strSQL & "    AND   A.NU_ATIV_MERC         =    '" & pstrIdentificadorTituloSelic & "'" & vbCrLf
        End If

        If pstrDataVencimento <> vbNullString Then
            strSQL = strSQL & "    AND   A.DT_VENC    =    " & fgDtXML_To_Oracle(pstrDataVencimento) & vbCrLf
        End If

        'Bruno Oliveira - 08/jul/10 - retirado devido reestruturação das contas selic
        'If pstrContaCedente <> vbNullString Then
            'strSQL = strSQL & "    AND   SUBSTR(LPAD(A.CO_CNTA_CUTD_SELIC_VEIC_LEGA,9,0),1,4)   =    " & pstrContaCedente & vbCrLf
        'End If

        If pvntValor > 0 Then
            strSQL = strSQL & "    AND   A.VA_FINC    =    " & fgVlrToDBServer(pvntValor) & vbCrLf
        End If

        If plngQuantidade > 0 Then
            strSQL = strSQL & "    AND   A.QT_ATIV_MERC    =    " & plngQuantidade & vbCrLf
        End If

        'Bruno Oliveira - 08/jul/10 - incluído devido reestruturação das contas selic
        If pintIndicadorDebitoCredito <> 0 Then
            strSQL = strSQL & "    AND   A.IN_OPER_DEBT_CRED    =    " & IIf(pintIndicadorDebitoCredito = enumTipoDebitoCredito.Credito, enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito) & vbCrLf
        End If

        strSQL = strSQL & _
                 "    AND   A.CO_OPER_SELIC     =   " & plngCodigoOperacaoSelic & vbCrLf & _
                 "    AND   A.NU_CTRL_IF            =   B.NU_CTRL_IF " & vbCrLf & _
                 "    AND   A.DH_REGT_MESG_SPB      =   B.DH_REGT_MESG_SPB " & vbCrLf & _
                 "    AND   A.NU_SEQU_CNTR_REPE     =   B.NU_SEQU_CNTR_REPE " & vbCrLf & _
                 "    AND   B.CO_ULTI_SITU_PROC     =    " & enumStatusMensagem.AConciliar & _
                 "    AND   B.CO_EMPR               =    " & plngCodigoEmpresa
    End If

    ObterValoresConciliacao = fgQueryXMLLerTodos("ValorConciliacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
       'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterValoresConciliacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Controlar a liberação automática de mensagens SPB agendadas
'--------------------------------------------------------------------------------------------------

Public Function LiberarAgendamento() As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim objOperacao                             As A8LQS.clsOperacao
Dim objProcessoOperacao                     As A8LQS.clsProcessoOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40
Dim datServidor                             As Date
Dim strXMLMensagem                          As String
Dim strCabecalho                            As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    '--------------------------------------------------------------------------------------------
    '>>> Captura a data do Servidor para comparar as mensagens agendadas e não enviadas
    '--------------------------------------------------------------------------------------------
    datServidor = fgDataHoraServidor(DataHoraAux)
    
    strSQL = " SELECT   NU_CTRL_IF,                                                     " & vbNewLine & _
             "          NU_SEQU_OPER_ATIV,                                              " & vbNewLine & _
             "          DH_REGT_MESG_SPB,                                               " & vbNewLine & _
             "          CO_TEXT_XML,                                                    " & vbNewLine & _
             "          DH_ULTI_ATLZ                                                    " & vbNewLine & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB                                        " & vbNewLine & _
             " WHERE    CO_ULTI_SITU_PROC         = " & _
                            enumStatusMensagem.MensagemAgendada & "                     " & vbNewLine & _
             " AND      HO_ENVI_MESG_SPB         <= " & _
                            fgDtHrXML_To_Oracle(fgDtHr_To_Xml(datServidor))

    Set objRS = fgQuerySQL(strSQL)
    Set xmlOperacao = CreateObject("MSXML2.DOMDocument.4.0")

    Do While Not objRS.EOF
        'Carlos 02/06/2004 - Liberação do agendamento de msg SPB
        If IsNull(objRS!NU_SEQU_OPER_ATIV) Then
            Liberar objRS!NU_CTRL_IF, fgDtHr_To_Xml(objRS!DH_REGT_MESG_SPB), 0, 0, fgDtHr_To_Xml(objRS!dh_ulti_atlz), True
        Else
            strXMLMensagem = fgSelectVarchar4000(objRS("CO_TEXT_XML"))
    
            xmlOperacao.loadXML strXMLMensagem
    
            strCabecalho = xmlOperacao.documentElement.selectSingleNode("CABEC").Text

            If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
                'Transformar o Tipo de Débito/Crédito para Mensageria
                'RATS 399 - Ocor 01
                For Each xmlNode In xmlOperacao.selectNodes("//IN_OPER_DEBT_CRED")
                    xmlNode.Text = IIf(CLng(xmlNode.Text) = enumTipoDebitoCredito.Debito, "D", "C")
                Next
            End If

            strXMLMensagem = xmlOperacao.xml

            If xmlOperacao.documentElement.selectSingleNode("CO_LOCA_LIQU").Text <> enumLocalLiquidacao.CETIP Then
                'Transformar o Tipo de Débito/Crédito para Mensageria
                'RATS 399 - Ocor 01
                For Each xmlNode In xmlOperacao.selectNodes("//IN_OPER_DEBT_CRED")
                    xmlNode.Text = IIf(xmlNode.Text = "D", enumTipoDebitoCredito.Debito, enumTipoDebitoCredito.Credito)
                Next
            End If

            '--------------------------------------------------------------------------------------------
            '>>> Envia a mensagem para a fila do BUS e altera seu STATUS para EnviadaBus
            '--------------------------------------------------------------------------------------------
            Set objProcessoOperacao = CreateObject("A8LQS.clsProcessoOperacao")
            objProcessoOperacao.EnviarMensagemMQ strCabecalho & strXMLMensagem, enumIdentificadorFila.BUS, True, True
            Set objProcessoOperacao = Nothing

            Call AlterarStatus(objRS("NU_CTRL_IF"), _
                               fgDtHr_To_Xml(objRS("DH_REGT_MESG_SPB")), _
                               enumStatusMensagem.EnviadaBUS, _
                               fgDtHr_To_Xml(objRS("DH_ULTI_ATLZ")), _
                               fgDtHr_To_Xml(fgDataHoraServidor(DataHoraAux)))

        End If

        objRS.MoveNext
    Loop

    objRS.Close
    Set objRS = Nothing

    LiberarAgendamento = True

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objOperacao = Nothing
    Set objProcessoOperacao = Nothing
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LiberarAgendamento Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Obter a data para a gravação de um novo registro na tabela de mensagens enviadas/recebidas
'--------------------------------------------------------------------------------------------------

Private Function flObterDataGravacao(ByVal strNumeroControleIF As String) As String

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim datDataGravacao                         As Date

    On Error GoTo ErrorHandler

    strSQL = " SELECT   MAX(DH_REGT_MESG_SPB) AS DH_REGT_MESG_SPB " & vbNewLine & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB     " & vbNewLine & _
             " WHERE    NU_CTRL_IF          =       '" & strNumeroControleIF & "'"

    Set objRS = fgQuerySQL(strSQL)

    If IsNull(objRS!DH_REGT_MESG_SPB) Then
        datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
    Else
        If DateDiff("s", fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux), objRS!DH_REGT_MESG_SPB) = 0 Then
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

'--------------------------------------------------------------------------------------------------
'Incluir as informações das mensagens SPB do grupo BMA
'--------------------------------------------------------------------------------------------------

Public Function IncluirBMA(ByRef pstrMensagem As String, _
                           ByVal pstrNumeroControleIF As String, _
                           ByVal pvntSequenciaOperacao As Variant, _
                           ByVal pvntSequenciaConciliacaoOperacao As Variant, _
                           ByVal pstrCodigoMensagemSPB As String, _
                           ByVal penumTipoBackOffice As enumTipoBackOffice, _
                           ByVal plngCodigoEmpresa As Long, _
                           ByVal pstrComandoOperacao As String, _
                           ByVal pstrSituacaoMensagemSPB As String, _
                           ByVal penumStatusMensagem As enumStatusMensagem, _
                           ByVal plngIndicadorEntradaManual As Long, _
                           ByVal pstrAgendamento As String, _
                           ByVal plngCodigoLocalLiquidacao As Long, _
                           ByVal pstrCodigoVeiculoLegal As String, _
                           ByVal pstrSiglaSistema As String, _
                  Optional ByVal plngSeqControleRepeticao As Long = 1, _
                  Optional ByVal pstrNumControleCamara As String, _
                  Optional ByVal pstrDataRegistroMensagem As String = vbNullString, _
                  Optional ByVal plngIndicadorConfirmacaoMesg As Long = 0, _
                  Optional ByVal pvntCodPartipanteCamara As Variant = vbNullString, _
                  Optional ByVal pstrDataHoraRecebMensagem As String = vbNullString, _
                  Optional ByVal plngTipoAcaoMensagem As Long = 0, _
                  Optional ByVal pstrISPBPartipanteCamara As String = vbNullString) As String

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long
Dim objHistSituacaoMensagem                 As A8LQS.clsHistSituacaoMensagem
Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    'Obter a data para gravação da mensagem e do historico
    If pstrDataRegistroMensagem = vbNullString Then
        pstrDataRegistroMensagem = flObterDataGravacao(pstrNumeroControleIF)
    End If
    
    lngCodigoTxtXML = fgInsertVarchar4000(pstrMensagem)

    strSQL = "INSERT INTO " & vbCrLf & _
             "  A8.TB_MESG_RECB_ENVI_SPB (" & vbCrLf & _
             "          NU_CTRL_IF,                    " & vbCrLf & _
             "          DH_REGT_MESG_SPB,              " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,             " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,             " & vbCrLf & _
             "          TP_BKOF,                       " & vbCrLf & _
             "          CO_EMPR,                       " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,         " & vbCrLf & _
             "          CO_MESG_SPB,                   " & vbCrLf & _
             "          NU_COMD_OPER,                  " & vbCrLf & _
             "          CO_SITU_MESG_SPB,              " & vbCrLf & _
             "          CO_TEXT_XML,                   " & vbCrLf & _
             "          HO_ENVI_MESG_SPB,              " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,             " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,             " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,        " & vbCrLf & _
             "          DH_ULTI_ATLZ,                  " & vbCrLf & _
             "          IN_ENTR_MANU,                  " & vbCrLf & _
             "          CO_LOCA_LIQU,                  " & vbCrLf & _
             "          CO_VEIC_LEGA,                  " & vbCrLf & _
             "          SG_SIST,                       " & vbCrLf & _
             "          NU_SEQU_CNCL_OPER_ATIV_MESG,   " & vbCrLf & _
             "          NU_CTRL_CAMR,                  " & vbCrLf & _
             "          IN_CONF_MESG_LTR,              " & vbCrLf

    strSQL = strSQL & _
             "          TP_ACAO_MESG_SPB_EXEC,         " & vbCrLf & _
             "          CO_PARP_CAMR,                  " & vbCrLf & _
             "          CO_ISPB_PART_CAMR             )" & vbCrLf & _
             "  VALUES (" & _
             "'" & pstrNumeroControleIF & "', " & vbCrLf & _
             fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & ", " & vbCrLf & _
             plngSeqControleRepeticao & ", " & vbCrLf & _
             IIf(Val(pvntSequenciaOperacao) = 0, "Null", pvntSequenciaOperacao) & ", " & vbCrLf & _
             IIf(penumTipoBackOffice = enumTipoBackOffice.Todos Or penumTipoBackOffice = 0, "Null", penumTipoBackOffice) & ", " & vbCrLf & _
             plngCodigoEmpresa & ", " & vbCrLf & _
             IIf(pstrDataHoraRecebMensagem = vbNullString, fgDataHoraServidor_To_Oracle, fgDtHrXML_To_Oracle(pstrDataHoraRecebMensagem)) & ", " & vbCrLf & _
             "'" & pstrCodigoMensagemSPB & "', " & vbCrLf & _
             "'" & pstrComandoOperacao & "', " & vbCrLf & _
             IIf(pstrSituacaoMensagemSPB = vbNullString, "Null", " '" & pstrSituacaoMensagemSPB & "'") & ", " & vbCrLf & _
             lngCodigoTxtXML & ", " & vbCrLf

    If Val(pstrAgendamento) <> 0 Then
       strSQL = strSQL & _
             fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(enumFormatoDataHora.DataHora), "YYYYMMDD") & pstrAgendamento & "00") & ", " & vbCrLf
    Else
       strSQL = strSQL & _
             " NULL " & ", " & vbCrLf
    End If

    strSQL = strSQL & _
             penumStatusMensagem & ", " & vbCrLf & _
             " '" & fgUsuarioRede & "', " & vbCrLf & _
             " '" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
             "Sysdate , " & vbCrLf & _
             plngIndicadorEntradaManual & " , " & vbCrLf & _
             IIf(plngCodigoLocalLiquidacao = 0, " NULL ", plngCodigoLocalLiquidacao) & " , " & vbCrLf & _
             IIf(pstrCodigoVeiculoLegal = vbNullString, " NULL ", " '" & pstrCodigoVeiculoLegal & "' ") & " , " & vbCrLf & _
             IIf(pstrSiglaSistema = vbNullString, " NULL ", " '" & pstrSiglaSistema & "' ") & " , " & vbCrLf & _
             IIf(pvntSequenciaConciliacaoOperacao = 0, " NULL ", pvntSequenciaConciliacaoOperacao) & ", " & vbCrLf & _
             IIf(pstrNumControleCamara = vbNullString, " NULL ", " '" & pstrNumControleCamara & "' ") & ", " & vbCrLf & _
             IIf(plngIndicadorConfirmacaoMesg = 0, " NULL ", " '" & plngIndicadorConfirmacaoMesg & "' ") & ", " & vbCrLf & _
             IIf(plngTipoAcaoMensagem = 0, " NULL ", plngTipoAcaoMensagem) & " , " & vbCrLf & _
             IIf(pvntCodPartipanteCamara = vbNullString, " NULL ", " '" & pvntCodPartipanteCamara & "' ") & ", " & vbCrLf & _
             " '" & pstrISPBPartipanteCamara & "')"
             
    fgExecuteSQL strSQL

    Set objHistSituacaoMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
    
    objHistSituacaoMensagem.Incluir pstrNumeroControleIF, _
                                    pstrDataRegistroMensagem, _
                                    penumStatusMensagem, _
                                    plngTipoAcaoMensagem, _
                                    , _
                                    plngSeqControleRepeticao
    
    Set objHistSituacaoMensagem = Nothing

    IncluirBMA = pstrDataRegistroMensagem

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objHistSituacaoMensagem = Nothing
    Set xmlMensagem = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Alterar o status das mensagens BMA0002 e/ou LDL0001 para o status de A Conciliar
'--------------------------------------------------------------------------------------------------

Public Function AlterarStatusBMA0002LDL0001(ByVal pvntSequenciaOperacao As Variant, _
                                   Optional ByVal plngStatusBMA0002 As Long = enumStatusMensagem.MensagemCancelada, _
                                   Optional ByVal plngStatusLDL0001 As Long = enumStatusMensagem.AConciliar) As Boolean

Dim objConciliacao                          As A8LQS.clsConciliacao
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    Set objConciliacao = CreateObject("A8LQS.clsConciliacao")
    strMensagem = vbNullString
    strMensagem = objConciliacao.LerTodos(, pvntSequenciaOperacao)

    If strMensagem = vbNullString Then
        Exit Function
    End If

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    xmlMensagem.loadXML strMensagem

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB " & vbNewLine & _
             " Set      CO_ULTI_SITU_PROC       =   " & plngStatusBMA0002 & ", " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate  " & vbNewLine & _
             " Where    NU_CTRL_IF              =   '" & xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text & "'" & vbNewLine & _
             " And      CO_MESG_SPB             =   'BMA0002' "

    fgExecuteSQL strSQL

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB " & vbNewLine & _
             " Set      CO_ULTI_SITU_PROC       =   " & plngStatusLDL0001 & ", " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate  " & vbNewLine & _
             " Where    NU_CTRL_IF              =   '" & xmlMensagem.documentElement.selectSingleNode("//NU_CTRL_IF").Text & "'" & vbNewLine & _
             " And      CO_MESG_SPB             =   'LDL0001' "

    fgExecuteSQL strSQL
    
    
    Set objConciliacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objConciliacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarStatusBMA0002LDL0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'--------------------------------------------------------------------------------------------------
'Ler mensagens LDL0001 com mesmo status e mesmo tipo de operação
'--------------------------------------------------------------------------------------------------

Public Function ObterMensagemLDL0001Conflitante(ByVal plngEmpresa As Long, _
                                                ByVal plngLocalLiquidacao As Long) As String

    On Error GoTo ErrorHandler

    'Não pode haver ao mesmo tempo:
    '       LDL0001
    '       Definitiva
    '       Mesma data
    '       Status: a conciliar
    
    Dim xmlFiltro                           As MSXML2.DOMDocument40
    Dim xmlMensagem                         As MSXML2.DOMDocument40
    Dim xmlHistorico                        As MSXML2.DOMDocument40
    Dim objMensagemSLCC                     As A8LQS.clsMensagemSLCC
    Dim strRet                              As String
    
    Dim blnAcaoJaPagou                      As Boolean
    Dim objHistMensagem                     As A8LQS.clsHistSituacaoMensagem
    
    
    ObterMensagemLDL0001Conflitante = ""

    Set xmlFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlHistorico = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagemSLCC = CreateObject("A8LQS.clsMensagemSLCC")
    
    Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")

    xmlFiltro.Load ""
    Call fgAppendNode(xmlFiltro, "", "Repeat_Filtros", "", "")

    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Status", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Status", "Status", enumStatusMensagem.AConciliar)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
    Call fgAppendNode(xmlFiltro, "Grupo_BancoLiquidante", "Empresa", plngEmpresa)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
    Call fgAppendNode(xmlFiltro, "Grupo_LocalLiquidacao", "LocalLiquidacao", plngLocalLiquidacao)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Mensagem", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Mensagem", "CodMensagem", "LDL0001")
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_TipoInformacao", "")
    Call fgAppendNode(xmlFiltro, "Grupo_TipoInformacao", "TipoInf", "D")
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_SequenciaControleRepeticao", "")
    Call fgAppendNode(xmlFiltro, "Grupo_SequenciaControleRepeticao", "Igual", 1)
    
    Call fgAppendNode(xmlFiltro, "Repeat_Filtros", "Grupo_Data", "")
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataIni", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux), "YYYYMMDD") & "000000"))
    Call fgAppendNode(xmlFiltro, "Grupo_Data", "DataFim", fgDtHrXML_To_Oracle(Format(fgDataHoraServidor(DataAux) + 1, "YYYYMMDD") & "000000"))
    
    strRet = objMensagemSLCC.ObterDetalheMensagemCamara(xmlFiltro, True)

    If strRet = "" Then
        Exit Function
    End If

    If xmlMensagem.loadXML(strRet) Then
        'Achou mensagem, verifica se já pagou
         
        Set objHistMensagem = CreateObject("A8LQS.clsHistSituacaoMensagem")
         
        'parei aqui
        strRet = objHistMensagem.LerTodos(xmlMensagem.selectSingleNode("//NU_CTRL_IF").Text, _
                                          xmlMensagem.selectSingleNode("//DH_REGT_MESG_SPB").Text, _
                                          xmlMensagem.selectSingleNode("//NU_SEQU_CNTR_REPE").Text)
         
        xmlHistorico.loadXML strRet
                        
        blnAcaoJaPagou = xmlHistorico.selectNodes("//TP_ACAO_MESG_SPB[../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamentoContingencia & "' or ../TP_ACAO_MESG_SPB='" & enumTipoAcao.EnviadoPagamento & "']").length > 0

        If blnAcaoJaPagou Then
            fgAppendNode xmlMensagem, "Grupo_DetalheMensagemCamara", "JaEnviouPagamento", "sim"
        End If
    
        ObterMensagemLDL0001Conflitante = xmlMensagem.xml
    End If

    Set xmlFiltro = Nothing
    Set xmlMensagem = Nothing
    Set objHistMensagem = Nothing
    Set objMensagemSLCC = Nothing

    If Not objContext Is Nothing Then
       'ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set xmlFiltro = Nothing
    Set xmlMensagem = Nothing
    Set objHistMensagem = Nothing
    Set objMensagemSLCC = Nothing

    If Not objContext Is Nothing Then
        'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagemLDL0001Conflitante Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Alterar a sequencia da operação na mensagem SPB
'--------------------------------------------------------------------------------------------------

Public Function AlterarSequenciaOperacao(ByVal pstrNumeroControleIF As String, _
                                         ByVal pdatDTRegistroMensagemSPB As Date, _
                                         ByVal pvntSequenciaOperacao As Variant, _
                                         ByVal plngNumeroSequenciaRepeticao As Long) As Boolean
 
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB                                                              " & vbNewLine & _
             " Set      NU_SEQU_OPER_ATIV           =  " & pvntSequenciaOperacao & ",              " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ           = '" & fgUsuarioRede & "',                                " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ      = '" & fgEstacaoTrabalhoUsuario & "',                     " & vbNewLine & _
             "          DH_ULTI_ATLZ                =      Sysdate                                            " & vbNewLine & _
             " Where    NU_CTRL_IF                  = '" & pstrNumeroControleIF & "'                          " & vbNewLine & _
             " And      DH_REGT_MESG_SPB            =  " & fgDtHrXML_To_Oracle(fgDtHr_To_Xml(pdatDTRegistroMensagemSPB)) & vbNewLine & _
             " And      NU_SEQU_CNTR_REPE           =  " & plngNumeroSequenciaRepeticao
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarSequenciaOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Ler as informações da tabela mensagem recebida/enviada SPB conforme o filtro definido
'--------------------------------------------------------------------------------------------------

Public Function LerTodos(Optional ByVal pstrNumeroControleIF As String = vbNullString, _
                         Optional ByVal pvntSequenciaOperacao As Variant = 0, _
                         Optional ByVal pstrNumeroControleCamara As String = vbNullString, _
                         Optional ByVal pstrCodigoMensagem As String = vbNullString, _
                         Optional ByVal pintLocalLiquidacao As Integer = 0, _
                         Optional ByVal pstrDataOperacao As String) As String
                         
Dim strSQL                                  As String
Dim blnTemWhere                             As Boolean

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   NU_CTRL_IF,                  " & vbCrLf & _
             "          NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          CO_TEXT_XML,                 " & vbCrLf & _
             "          CO_VEIC_LEGA,                " & vbCrLf & _
             "          CO_USUA_ULTI_ATLZ,           " & vbCrLf & _
             "          SG_SIST,                     " & vbCrLf & _
             "          CO_ETCA_TRAB_ULTI_ATLZ,      " & vbCrLf & _
             "          DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          CO_MESG_SPB,                 " & vbCrLf & _
             "          TP_BKOF,                     " & vbCrLf & _
             "          CO_EMPR,                     " & vbCrLf & _
             "          NU_COMD_OPER,                " & vbCrLf & _
             "          CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          DH_RECB_ENVI_MESG_SPB,       " & vbCrLf & _
             "          HO_ENVI_MESG_SPB,            " & vbCrLf & _
             "          DH_ULTI_ATLZ,                " & vbCrLf & _
             "          IN_ENTR_MANU,                " & vbCrLf & _
             "          NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          CO_LOCA_LIQU,                " & vbCrLf & _
             "          NU_CTRL_CAMR,                " & vbCrLf & _
             "          IN_CONF_MESG_LTR             " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB "


    blnTemWhere = False

    If Trim(pstrNumeroControleIF) <> vbNullString Then
        strSQL = strSQL & " WHERE    NU_CTRL_IF = '" & pstrNumeroControleIF & "'"
        blnTemWhere = True
    End If

    If pvntSequenciaOperacao > 0 Then
        If blnTemWhere Then
             strSQL = strSQL & " AND      NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao
        Else
             strSQL = strSQL & " WHERE    NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao
        End If
        blnTemWhere = True
    End If

    If Trim(pstrNumeroControleCamara) <> vbNullString Then
        If blnTemWhere Then
             strSQL = strSQL & " AND      NU_CTRL_CAMR = '" & Trim(pstrNumeroControleCamara) & "'"
        Else
             strSQL = strSQL & " WHERE    NU_CTRL_CAMR = '" & Trim(pstrNumeroControleCamara) & "'"
        End If
        blnTemWhere = True
    End If

    If Trim(pstrCodigoMensagem) <> vbNullString Then
        If blnTemWhere Then
             strSQL = strSQL & " AND      CO_MESG_SPB = '" & Trim(pstrCodigoMensagem) & "'"
        Else
             strSQL = strSQL & " WHERE    CO_MESG_SPB = '" & Trim(pstrCodigoMensagem) & "'"
        End If
        blnTemWhere = True
    End If

    If pintLocalLiquidacao <> 0 Then
        If blnTemWhere Then
             strSQL = strSQL & " AND      CO_LOCA_LIQU = " & pintLocalLiquidacao
        Else
             strSQL = strSQL & " WHERE    CO_LOCA_LIQU = " & pintLocalLiquidacao
        End If
        blnTemWhere = True
    End If
    
    'KIDA 11/12/2009
    'RATS 696
    If Trim(pstrDataOperacao) <> vbNullString Then
        If blnTemWhere Then
             strSQL = strSQL & " AND      TRUNC(DH_REGT_MESG_SPB)   = " & fgDtXML_To_Oracle(pstrDataOperacao)
        Else
             strSQL = strSQL & " WHERE    TRUNC(DH_REGT_MESG_SPB)   = " & fgDtXML_To_Oracle(pstrDataOperacao)
        End If
        blnTemWhere = True
    End If
    
    
    
    LerTodos = fgQueryXMLLerTodos("MensagemEnviadaRecebida", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Ler as informações da tabela mensagem recebida/enviada SPB  e da tabela de conciliação conforme o filtro definido
'--------------------------------------------------------------------------------------------------

Public Function LerMensagemConciliada(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal pstrCodigoMensagem As String, _
                             Optional ByVal pstrChACAM As String = vbNullString, _
                             Optional ByVal pstrContrSISBACEN As String = vbNullString) As String

Dim strSQL                                  As String
Dim blnTemWhere                             As Boolean

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   A.NU_CTRL_IF,                   " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,            " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,            " & vbCrLf & _
             "          A.CO_TEXT_XML,                  " & vbCrLf & _
             "          A.CO_VEIC_LEGA,                 " & vbCrLf & _
             "          A.CO_USUA_ULTI_ATLZ,            " & vbCrLf & _
             "          A.SG_SIST,                      " & vbCrLf & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,       " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,             " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,            " & vbCrLf & _
             "          A.CO_MESG_SPB,                  " & vbCrLf & _
             "          A.TP_BKOF,                      " & vbCrLf & _
             "          A.CO_EMPR,                      " & vbCrLf & _
             "          A.NU_COMD_OPER,                 " & vbCrLf & _
             "          A.CO_SITU_MESG_SPB,             " & vbCrLf & _
             "          A.DH_RECB_ENVI_MESG_SPB,        " & vbCrLf & _
             "          A.HO_ENVI_MESG_SPB,             " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                 " & vbCrLf & _
             "          A.IN_ENTR_MANU,                 " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,  " & vbCrLf & _
             "          A.CO_LOCA_LIQU,                 " & vbCrLf & _
             "          A.NU_CTRL_CAMR,                 " & vbCrLf & _
             "          A.IN_CONF_MESG_LTR,             " & vbCrLf
    strSQL = strSQL & _
             "          C.CD_ASSO_CAMB,                 " & vbCrLf & _
             "          C.CO_CNTR_SISB                  " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB    A,  " & vbCrLf & _
             "          A8.TB_CNCL_OPER_ATIV        B,  " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL    C   " & vbCrLf & _
             " WHERE    B.NU_SEQU_OPER_ATIV     =       " & pvntSequenciaOperacao & vbCrLf & _
             "   AND    A.NU_CTRL_IF            =       B.NU_CTRL_IF   " & vbCrLf & _
             "   AND    A.NU_CTRL_IF            =       C.NU_CTRL_IF   " & vbCrLf & _
             "   AND    A.CO_MESG_SPB = '" & Trim(pstrCodigoMensagem) & "'"

    If Trim(pstrChACAM) <> vbNullString Then
        strSQL = strSQL & "   AND    C.CD_ASSO_CAMB = '" & Trim(pstrChACAM) & "'"
    End If

    If Trim(pstrContrSISBACEN) <> vbNullString Then
        strSQL = strSQL & "   AND    C.CO_CNTR_SISB = " & Trim(pstrContrSISBACEN)
    End If

    LerMensagemConciliada = fgQueryXMLLerTodos("MensagemEnviadaRecebida", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagemConciliada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Altera o código de veículo legal da mensagem SPB recebida
'--------------------------------------------------------------------------------------------------

Public Function AlterarVeiculoLegal(ByVal pstrNumeroControleIF As String, _
                                    ByVal pstrDataRegistroMensagem As String, _
                                    ByVal pvntNumeroSequenciaRepeticao As Variant, _
                                    ByVal pstrCodigoVeiculoLegal As String, _
                                    ByVal pstrSiglaSistema As String) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB " & vbNewLine & _
             " Set      CO_VEIC_LEGA            =   '" & pstrCodigoVeiculoLegal & "', " & vbNewLine & _
             "          SG_SIST                 =   '" & pstrSiglaSistema & "', " & vbNewLine & _
             "          CO_USUA_ULTI_ATLZ       =   '" & fgUsuarioRede & "', " & vbNewLine & _
             "          CO_ETCA_TRAB_ULTI_ATLZ  =   '" & fgEstacaoTrabalhoUsuario & "', " & vbNewLine & _
             "          DH_ULTI_ATLZ            =   Sysdate " & vbNewLine & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'" & vbNewLine & _
             " AND      DH_REGT_MESG_SPB        =   " & fgDtHrXML_To_Oracle(pstrDataRegistroMensagem) & vbNewLine & _
             " AND      NU_SEQU_CNTR_REPE       =   " & pvntNumeroSequenciaRepeticao
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarVeiculoLegal Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'--------------------------------------------------------------------------------------------------
'Altera o número de controle da câmara da mensagem SPB recebida
'--------------------------------------------------------------------------------------------------

Public Function AlterarNumeroControleCamara(ByVal pstrNumeroControleIF As String, _
                                            ByVal pstrNumeroControleCamara As String) As Boolean

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " Update   A8.TB_MESG_RECB_ENVI_SPB " & vbNewLine & _
             " Set      NU_CTRL_CAMR            =  '" & pstrNumeroControleCamara & "' " & vbNewLine & _
             " Where    NU_CTRL_IF              =  '" & pstrNumeroControleIF & "'"
    
    If fgExecuteSQL(strSQL) = 0 Then
        'Mensagem já processada por outro usuário.
        lngCodigoErroNegocio = 3014
        GoTo ErrorHandler
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AlterarNumeroControleCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Alterar a ação da operação que na liberação da mensagem SPB seja de cancelamento
Private Function flVerificarOperacao(ByRef xmlMensagem As MSXML2.DOMDocument40) As String

Dim objOperacao                             As A8LQS.clsOperacao
Dim xmlOperacao                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    If xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV") Is Nothing Then
        Exit Function
    End If

    If Val("0" & xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text) = 0 Then
        Exit Function
    End If

    Set objOperacao = CreateObject("A8LQS.clsOperacao")
    Set xmlOperacao = objOperacao.LerXMLOperacao(xmlMensagem.selectSingleNode("//NU_SEQU_OPER_ATIV").Text)

    If xmlOperacao.xml = vbNullString Then
        Exit Function
    End If

    If xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = enumTipoAcao.CancelamentoSolicitado Or _
        xmlOperacao.documentElement.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text = enumTipoAcao.CancelamentoSolicitadoComMensagem Then
        objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                      xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                      xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                       , _
                                      enumTipoAcao.CancelamentoEnviado
    End If
    
    If xmlOperacao.documentElement.selectSingleNode("TP_MESG").Text = enumTipoMensagemLQS.RegistroOperacoesBMC Then
        If Not xmlMensagem.documentElement.selectSingleNode("CO_MESG") Is Nothing Then
            If xmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = "BMC0002" Then
                objOperacao.AlterarStatusAcao xmlOperacao.documentElement.selectSingleNode("NU_SEQU_OPER_ATIV").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("CO_ULTI_SITU_PROC").Text, _
                                              xmlOperacao.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text, _
                                              , _
                                              enumTipoAcao.EnviadaBMC0002
            End If
        End If
    End If
    
    Set objOperacao = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    Set objOperacao = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarOperacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


'--------------------------------------------------------------------------------------------------
'Ler o xml da mensagem SPB enviada ou recebida pelo SLCC
'--------------------------------------------------------------------------------------------------
Public Function LerXMLMensagensPagamentoCorretoras() As ADODB.Recordset

Dim strSQL                                  As String
Dim rsMensagem                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsMensagem = CreateObject("ADODB.Recordset")

    strSQL = " SELECT   A.NU_CTRL_IF,                  " & vbCrLf & _
             "          A.DH_REGT_MESG_SPB,            " & vbCrLf & _
             "          A.NU_SEQU_OPER_ATIV,           " & vbCrLf & _
             "          A.CO_VEIC_LEGA,                " & vbCrLf & _
             "          A.NU_SEQU_CNTR_REPE,           " & vbCrLf & _
             "          A.TP_BKOF,                     " & vbCrLf & _
             "          A.CO_EMPR,                     " & vbCrLf & _
             "          A.NU_COMD_OPER,                " & vbCrLf & _
             "          A.CO_SITU_MESG_SPB,            " & vbCrLf & _
             "          A.CO_TEXT_XML,                 " & vbCrLf & _
             "          A.CO_ULTI_SITU_PROC,           " & vbCrLf & _
             "          A.CO_MESG_SPB,                 " & vbCrLf & _
             "          A.SG_SIST,                     " & vbCrLf & _
             "          A.CO_LOCA_LIQU,                " & vbCrLf & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG, " & vbCrLf & _
             "          A.IN_ENTR_MANU,                " & vbCrLf & _
             "          A.NU_CTRL_CAMR,                " & vbCrLf & _
             "          A.IN_CONF_MESG_LTR,            " & vbCrLf & _
             "          A.CO_PARP_CAMR,                " & vbCrLf & _
             "          A.TP_ACAO_MESG_SPB_EXEC,       " & vbCrLf & _
             "          A.DH_ULTI_ATLZ,                " & vbCrLf
             
    strSQL = strSQL & _
             "          B.CO_ISPB_BANC_LIQU_CNPT,      " & vbCrLf & _
             "          B.TP_INFO_LDL,                 " & vbCrLf & _
             "          B.CO_PRAC,                     " & vbCrLf & _
             "          B.CO_MOED_ESTR,                " & vbCrLf & _
             "          B.PE_TAXA_NEGO,                " & vbCrLf & _
             "          B.VA_FINC,                     " & vbCrLf & _
             "          B.VA_MOED_ESTR,                " & vbCrLf & _
             "          B.DT_LIQU,                     " & vbCrLf & _
             "          B.CO_SISB_COTR,                " & vbCrLf & _
             "          B.IN_OPER_DEBT_CRED,           " & vbCrLf & _
             "          B.TP_LIQU_OPER_ATIV,           " & vbCrLf & _
             "          B.CD_ASSO_CAMB,                " & vbCrLf & _
             "          B.CO_CNTR_SISB, B.CO_CNPJ_CNPT , B.NO_CNPT , B.CO_AGEN_COTR , B.NU_CC_COTR " & vbCrLf & _
             "  FROM    A8.TB_MESG_RECB_ENVI_SPB A,    " & vbCrLf & _
             "          A8.TB_MESG_RECB_SPB_CNCL B     " & vbCrLf

    'relacionamentos
    strSQL = strSQL & _
             " WHERE    A.NU_CTRL_IF         = B.NU_CTRL_IF        (+) " & vbCrLf & _
             "   AND    A.NU_SEQU_CNTR_REPE  = B.NU_SEQU_CNTR_REPE (+) " & vbCrLf & _
             "   AND    A.DH_REGT_MESG_SPB   = B.DH_REGT_MESG_SPB  (+) " & vbCrLf
    'filtros
    strSQL = strSQL & _
             "   AND    A.CO_MESG_SPB             IN ('STR0006R2','STR0007R2','STR0008R2','PAG0108R2')" & vbCrLf & _
             "   AND    TRUNC(A.DH_REGT_MESG_SPB) = TRUNC(SYSDATE) " & vbCrLf & _
             "   AND    A.CO_ULTI_SITU_PROC       = " & enumStatusMensagem.AConciliar & vbCrLf

    strSQL = strSQL & " ORDER BY A.DH_REGT_MESG_SPB "
    
    Set rsMensagem = fgQuerySQL(strSQL)

    If rsMensagem.EOF Then
        Exit Function
    Else
        Set LerXMLMensagensPagamentoCorretoras = rsMensagem
    End If

Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerXMLMensagensPagamentoCorretoras Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function



