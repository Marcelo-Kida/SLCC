VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentoCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Obter informações necessárias à movimentação do caixa, a partir de mensagens do sistema A7

Option Explicit

'Obter Item caixa

Public Sub ObterItemCaixa(ByRef pstrRemessa As String, ByVal pintTipoCaixa As enumTipoCaixa)

Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim strSQL                                  As String

Dim strCodProduto                           As String
Dim strItemCaixa                            As String
Dim strTipoOper                             As String
Dim strTipoContraparte                      As String
Dim strIndEntrSaida                         As String
Dim strCodVeicLegal                         As String
Dim strSiglaSistema                         As String

Dim strFilter                               As String
Dim rsItemCaixa                             As ADODB.Recordset

On Error GoTo ErrorHandler

    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlRemessa.loadXML(pstrRemessa)
    
    With xmlRemessa.documentElement
        If rsItemCaixa Is Nothing Then
            
            strSQL = " SELECT   A.CO_PROD,                                                  " & vbNewLine & _
                     "          A.TP_OPER,                                                  " & vbNewLine & _
                     "          B.TP_CNPT,                                                  " & vbNewLine & _
                     "          B.IN_ENTR_SAID,                                             " & vbNewLine & _
                     "          B.CO_ITEM_CAIX,                                             " & vbNewLine & _
                     "          C.TP_BKOF,                                                  " & vbNewLine & _
                     "          C.TP_CAIX,                                                  " & vbNewLine & _
                     "          D.CO_GRUP_VEIC_LEGA                                         " & vbNewLine & _
                     " FROM     A6.TB_TIPO_OPER_PROD_PJ        A,                           " & vbNewLine & _
                     "          A6.TB_ITEM_CAIX_TIPO_OPER      B,                           " & vbNewLine & _
                     "          A6.TB_ITEM_CAIX_OPER_ATIV      C,                           " & vbNewLine & _
                     "          A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA D                            " & vbNewLine & _
                     " WHERE    A.TP_OPER                    = B.TP_OPER                    " & vbNewLine & _
                     " AND      B.CO_ITEM_CAIX               = C.CO_ITEM_CAIX               " & vbNewLine & _
                     " AND      SUBSTR(B.CO_ITEM_CAIX, 1, 4) = SUBSTR(D.CO_ITEM_CAIX, 1, 4) " & vbNewLine & _
                     " AND      A.DT_INIC_VIGE     <= " & fgDataHoraServidor_To_Oracle & "  " & vbNewLine & _
                     " AND     (A.DT_FIM_VIGE      >= " & fgDataHoraServidor_To_Oracle & "  " & vbNewLine & _
                     " OR       A.DT_FIM_VIGE       IS NULL)                                "
    
            Set rsItemCaixa = QuerySQL(strSQL)
            
        End If
    
        strCodProduto = .selectSingleNode("CO_PROD").Text
        strCodVeicLegal = .selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = .selectSingleNode("SG_SIST_ORIG").Text
        
        If Not .selectSingleNode("IN_MOVI_ENTR_SAID") Is Nothing Then
            strIndEntrSaida = .selectSingleNode("IN_MOVI_ENTR_SAID").Text
        Else
            strIndEntrSaida = .selectSingleNode("IN_OPER_DEBT_CRED").Text
        End If
        
        If .selectSingleNode("TP_CNPT") Is Nothing Then
            strTipoContraparte = "NULL"
        Else
            strTipoContraparte = IIf(.selectSingleNode("TP_CNPT").Text = vbNullString, "NULL", .selectSingleNode("TP_CNPT").Text)
        End If
        
        With rsItemCaixa
            strFilter = "(CO_PROD           = " & strCodProduto & "             AND " & vbNewLine & _
                        " IN_ENTR_SAID      = " & strIndEntrSaida & "           AND " & vbNewLine & _
                        " TP_BKOF           = " & glngTipoBackOffice & "        AND " & vbNewLine & _
                        " TP_CAIX           = " & pintTipoCaixa & "             AND " & vbNewLine & _
                        " CO_GRUP_VEIC_LEGA = " & gstrGrupoVeiculoLegal
            
            If strTipoContraparte <> "NULL" Then
                strFilter = strFilter & " AND TP_CNPT = " & strTipoContraparte
            End If

            strFilter = strFilter & ")"
            
            .Filter = strFilter
            
            If .RecordCount = 0 Then
                strItemCaixa = fgObterItemCaixaGenerico(pintTipoCaixa, glngTipoBackOffice)
                strTipoOper = vbNullString
            Else
                .MoveFirst
                strItemCaixa = !CO_ITEM_CAIX
                strTipoOper = !TP_OPER
            End If
        End With
        
        ' Estes Nodes estão sendo acrescentados ao DOMXML original para que seja aproveitada a mesma rotina de
        ' inclusão de Movimento Caixa
        
        If pintTipoCaixa = CaixaFuturo Then
            Call fgAppendNode(xmlRemessa, "MESG", "SG_SIST", strSiglaSistema)
            Call fgAppendNode(xmlRemessa, "MESG", "CO_ITEM_CAIX", strItemCaixa)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_GERA_MOVI", enumTipoRemessa.Remessa)
            Call fgAppendNode(xmlRemessa, "MESG", "TX_MOTI_GERA_MOVI", vbNullString)
            
            If Not .selectSingleNode("CO_TIPO_LIQU") Is Nothing Then
                Call fgAppendNode(xmlRemessa, "MESG", "DE_TIPO_LIQU", .selectSingleNode("CO_TIPO_LIQU").Text)
            Else
                Call fgAppendNode(xmlRemessa, "MESG", "DE_TIPO_LIQU", .selectSingleNode("DE_TIPO_LIQU").Text)
            End If
        Else
            Call fgAppendNode(xmlRemessa, "MESG", "CO_ITEM_CAIX_SUB_RESE", strItemCaixa)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_OPER", strTipoOper)
            Call fgAppendNode(xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.Previsto)
            Call fgAppendNode(xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", vbNullString)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_GERA_MOVI", enumTipoRemessa.Remessa)
            Call fgAppendNode(xmlRemessa, "MESG", "TX_MOTI_GERA_MOVI", vbNullString)
        End If
    End With
    
    pstrRemessa = xmlRemessa.xml
    
    Set rsItemCaixa = Nothing
    Set xmlRemessa = Nothing
    
    Exit Sub

ErrorHandler:
    Set rsItemCaixa = Nothing
    Set xmlRemessa = Nothing

    
    fgRaiseError App.EXEName, TypeName(Me), "ObterItemCaixa", 0

End Sub

'Obter Item caixa OnLine

Public Sub ObterItemCaixaOnLine(ByRef pstrRemessa As String, _
                                ByVal pintTipoCaixa As enumTipoCaixa, _
                                ByVal pstrGrupoVeiculoLegal As String, _
                                ByVal plngTipoBackOffice As Long)

Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim strSQL                                  As String

Dim strItemCaixa                            As String
Dim strTipoOper                             As String
Dim strTipoContraparte                      As String
Dim strIndEntrSaida                         As String
Dim strCodVeicLegal                         As String
Dim strSiglaSistema                         As String

Dim strFilter                               As String
Dim rsItemCaixa                             As ADODB.Recordset
    
On Error GoTo ErrorHandler

    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlRemessa.loadXML(pstrRemessa)
    
'    fgGravaArquivo "LOGA6_ObterItemCaixaOnLine_", vbNullString
    
    With xmlRemessa.documentElement
        If rsItemCaixa Is Nothing Then
            
            strSQL = " SELECT   A.TP_OPER,                                                  " & vbNewLine & _
                     "          A.TP_CNPT,                                                  " & vbNewLine & _
                     "          A.IN_ENTR_SAID,                                             " & vbNewLine & _
                     "          B.TP_BKOF,                                                  " & vbNewLine & _
                     "          B.TP_CAIX,                                                  " & vbNewLine & _
                     "          C.CO_GRUP_VEIC_LEGA,                                        " & vbNewLine & _
                     "          A.CO_ITEM_CAIX                                              " & vbNewLine & _
                     " FROM     A6.TB_ITEM_CAIX_TIPO_OPER      A,                           " & vbNewLine & _
                     "          A6.TB_ITEM_CAIX_OPER_ATIV      B,                           " & vbNewLine & _
                     "          A6.TB_ITEM_CAIX_GRUP_VEIC_LEGA C                            " & vbNewLine & _
                     " WHERE    A.CO_ITEM_CAIX               = B.CO_ITEM_CAIX               " & vbNewLine & _
                     " AND      SUBSTR(A.CO_ITEM_CAIX, 1, 4) = SUBSTR(C.CO_ITEM_CAIX, 1, 4) "
    
            Set rsItemCaixa = QuerySQL(strSQL)
            
        End If
    
        strTipoOper = .selectSingleNode("TP_OPER").Text
        strCodVeicLegal = .selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = .selectSingleNode("SG_SIST_ORIG").Text
        
        If Not .selectSingleNode("IN_MOVI_ENTR_SAID") Is Nothing Then
            strIndEntrSaida = .selectSingleNode("IN_MOVI_ENTR_SAID").Text
        Else
            strIndEntrSaida = .selectSingleNode("IN_OPER_DEBT_CRED").Text
        End If
        
        If .selectSingleNode("TP_CNPT") Is Nothing Then
            strTipoContraparte = "NULL"
        Else
            strTipoContraparte = IIf(.selectSingleNode("TP_CNPT").Text = vbNullString, "NULL", .selectSingleNode("TP_CNPT").Text)
        End If
        
        If Trim$(pstrGrupoVeiculoLegal) = vbNullString Or plngTipoBackOffice = 0 Then
            fgGravaArquivo "LOGA6_PerdeuOVeiculoLegal", vbNullString
    
'            Call fgExisteVeiculoLegal(.selectSingleNode("SG_SIST_ORIG").Text, _
'                                      .selectSingleNode("//CO_VEIC_LEGA").Text, _
'                                      fgDataHoraServidor(enumFormatoDataHora.Data), _
'                                      .selectSingleNode("CO_EMPR").Text)
        End If
        
        With rsItemCaixa
            strFilter = "(TP_OPER           = " & strTipoOper & "               AND " & vbNewLine & _
                        " IN_ENTR_SAID      = " & strIndEntrSaida & "           AND " & vbNewLine & _
                        " TP_BKOF           = " & plngTipoBackOffice & "        AND " & vbNewLine & _
                        " TP_CAIX           = " & pintTipoCaixa & "             AND " & vbNewLine & _
                        " CO_GRUP_VEIC_LEGA = " & pstrGrupoVeiculoLegal
            
            If strTipoContraparte <> "NULL" Then
                strFilter = strFilter & " AND TP_CNPT = " & strTipoContraparte
            End If

            strFilter = strFilter & ")"
            
            .Filter = strFilter
            
            If .EOF Then
'                fgGravaArquivo "LOGA6_EOF", vbNullString
                strItemCaixa = fgObterItemCaixaGenerico(pintTipoCaixa, plngTipoBackOffice)
            Else
                If .RecordCount = 0 Then
                    strItemCaixa = fgObterItemCaixaGenerico(pintTipoCaixa, plngTipoBackOffice)
                Else
                    .MoveFirst
                    strItemCaixa = !CO_ITEM_CAIX
                End If
            End If
        
        End With
        
        ' Estes Nodes estão sendo acrescentados ao DOMXML original para que seja aproveitada a mesma rotina de
        ' inclusão de Movimento Caixa
        
        If pintTipoCaixa = CaixaFuturo Then
            Call fgAppendNode(xmlRemessa, "MESG", "SG_SIST", strSiglaSistema)
            Call fgAppendNode(xmlRemessa, "MESG", "CO_ITEM_CAIX", strItemCaixa)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_GERA_MOVI", enumTipoRemessa.Remessa)
            Call fgAppendNode(xmlRemessa, "MESG", "TX_MOTI_GERA_MOVI", vbNullString)
            
            If Not .selectSingleNode("CO_TIPO_LIQU") Is Nothing Then
                Call fgAppendNode(xmlRemessa, "MESG", "DE_TIPO_LIQU", .selectSingleNode("CO_TIPO_LIQU").Text)
            Else
                Call fgAppendNode(xmlRemessa, "MESG", "DE_TIPO_LIQU", .selectSingleNode("DE_TIPO_LIQU").Text)
            End If
        Else
            Call fgAppendNode(xmlRemessa, "MESG", "CO_ITEM_CAIX_SUB_RESE", strItemCaixa)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_OPER", strTipoOper)
            Call fgAppendNode(xmlRemessa, "MESG", "CO_SITU_MOVI_CAIX_SUB_RESE", enumTipoMovimento.Previsto)
            Call fgAppendNode(xmlRemessa, "MESG", "NU_SEQU_OPER_ATIV", vbNullString)
            Call fgAppendNode(xmlRemessa, "MESG", "TP_GERA_MOVI", enumTipoRemessa.Remessa)
            Call fgAppendNode(xmlRemessa, "MESG", "TX_MOTI_GERA_MOVI", vbNullString)
        End If
    End With
    
    pstrRemessa = xmlRemessa.xml
        
    Set xmlRemessa = Nothing
    Set rsItemCaixa = Nothing
    Exit Sub

ErrorHandler:
    Set xmlRemessa = Nothing
    Set rsItemCaixa = Nothing

    
    fgRaiseError App.EXEName, TypeName(Me), "ObterItemCaixaOnLine", 0, 0, strSQL & " " & strFilter

End Sub

Public Function VerificarMovimentoSubReserva(ByVal pstrRemessa As String) As Boolean

Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim strSQL                                  As String

Dim strCodVeicLegal                         As String
Dim strSiglaSistema                         As String
Dim strDataMovimento                        As String
Dim datDataD0                               As Date

On Error GoTo ErrorHandler

    VerificarMovimentoSubReserva = False
    
    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlRemessa.loadXML(pstrRemessa)
    
    With xmlRemessa.documentElement
        strCodVeicLegal = .selectSingleNode("//CO_VEIC_LEGA").Text
        strSiglaSistema = .selectSingleNode("SG_SIST_ORIG").Text
        strDataMovimento = .selectSingleNode("DT_LIQU_OPER").Text
        
        datDataD0 = fgAdicionarDiasUteis(fgDtXML_To_Date(.selectSingleNode("DT_FECH_PROC").Text), 1, Proximo)
        
        If strDataMovimento = fgDt_To_Xml(datDataD0) Then
            VerificarMovimentoSubReserva = True
        End If
    End With

    Set xmlRemessa = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlRemessa = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "VerificarMovimentoSubReserva", 0

End Function
