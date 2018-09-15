VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela apresentação dos dados da tela de monitoração do sistema A7.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Private ObjectContext                       As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

'Obter a quantidade de mensagens por status de acordo com os parâmetros de filtro informados.
Public Function ObterQtd(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String, _
                         ByVal pintEmpresa As Integer, _
                         ByVal pstrSistema As String, _
                         ByVal pstrTipoMensagem As String, _
                         ByVal pstrIDMensagem As String) As String
    
Dim strSQL                                  As String
Dim strOwner                                As String
    
On Error GoTo ErrorHandler

    If fgDtXML_To_Date(Left$(pstrDataHoraDe, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        strOwner = "A7HIST"
    Else
        strOwner = "A7"
    End If

    'KIDA - 23/04/2009
    'RATS - 905
'    strSQL = " SELECT  C.CO_OCOR_MESG,          " & vbCrLf & _
                      "COUNT(*) as QTD          " & vbCrLf & _
             " FROM     " & strOwner & ".TB_MESG A,            " & vbCrLf & _
                       "A7.TB_TIPO_MESG B,       " & vbCrLf & _
                       "A7.TB_OCOR_MESG C,       " & vbCrLf & _
                       strOwner & ".TB_SITU_MESG D,       " & vbCrLf & _
                       "A7.TB_SIST      E        " & vbCrLf & _
             " WHERE    A.TP_MESG      = B.TP_MESG      " & vbCrLf & _
             "   AND    A.CO_MESG      = D.CO_MESG      " & vbCrLf & _
             "   AND    C.CO_OCOR_MESG = D.CO_OCOR_MESG " & vbCrLf & _
             "   AND    A.SG_SIST_ORIG = E.SG_SIST " & vbCrLf & _
             "   AND    A.CO_EMPR_ORIG = E.CO_EMPR " & vbCrLf

    strSQL = " SELECT  C.CO_OCOR_MESG,          " & vbCrLf & _
                      "COUNT(*) as QTD          " & vbCrLf & _
             " FROM     " & strOwner & ".TB_MESG A,            " & vbCrLf & _
                       "A7.TB_OCOR_MESG C,       " & vbCrLf & _
                       strOwner & ".TB_SITU_MESG D,       " & vbCrLf & _
                       "A7.TB_SIST      E        " & vbCrLf & _
             " WHERE    A.CO_MESG      = D.CO_MESG      " & vbCrLf & _
             "   AND    C.CO_OCOR_MESG = D.CO_OCOR_MESG " & vbCrLf & _
             "   AND    A.SG_SIST_ORIG = E.SG_SIST " & vbCrLf & _
             "   AND    A.CO_EMPR_ORIG = E.CO_EMPR " & vbCrLf & _
             "   AND    C.CO_OCOR_MESG In (1, 2, 3, 4, 6) " & vbCrLf
             
    strSQL = strSQL & "  AND     D.CO_OCOR_MESG = ( SELECT MAX(CO_OCOR_MESG) " & _
                                                    " FROM " & strOwner & ".TB_SITU_MESG " & _
                                                    " WHERE CO_MESG = A.CO_MESG)"
    
    strSQL = strSQL & "  AND     A.DH_MESG BETWEEN " & _
                                                    fgDtHrXML_To_Oracle(pstrDataHoraDe) & " AND " & _
                                                    fgDtHrXML_To_Oracle(pstrDataHoraAte) & vbCrLf
    
    If pintEmpresa <> 0 Then
        strSQL = strSQL & " AND A.CO_EMPR_ORIG = " & pintEmpresa & vbCrLf
    End If
    
    If pstrSistema <> "" Then
        strSQL = strSQL & " AND A.SG_SIST_ORIG = '" & pstrSistema & "'" & vbCrLf
    End If
    
    If Len(pstrTipoMensagem) <> 0 Then
        strSQL = strSQL & " AND A.TP_MESG = '" & pstrTipoMensagem & "'" & vbCrLf
    End If
    
    If pstrIDMensagem <> "" Then
        strSQL = strSQL & " AND CO_CMPO_ATRB_IDEF_MESG like '%" & Trim(pstrIDMensagem) & "%'" & vbCrLf
    End If
    
    strSQL = strSQL & " GROUP BY  C.CO_OCOR_MESG "

    ObterQtd = fgQueryXMLLerTodos("Monitoracao", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterQtd Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todas as mensagens transitadas no sistema A7 de acordo com os parâmetos de filtro informados.
Public Function LerTodos(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String, _
                         ByVal pintEmpresa As Integer, _
                         ByVal pstrSistema As String, _
                         ByVal pstrTipoMensagem As String, _
                         ByVal pstrCodOcorrencia As String, _
                         ByVal pstrIDMensagem As String) As String
    
Dim strSQL                                  As String
Dim strOwner                                As String
    
On Error GoTo ErrorHandler
    
    If fgDtXML_To_Date(Left$(pstrDataHoraDe, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        strOwner = "A7HIST"
    Else
        strOwner = "A7"
    End If

    strSQL = " SELECT  '" & strOwner & "' AS OWNER, " & vbCrLf & _
                       "A.CO_MESG,               " & vbCrLf & _
                       "A.CO_CMPO_ATRB_IDEF_MESG," & vbCrLf & _
                       "A.DH_MESG,               " & vbCrLf & _
                       "A.SG_SIST_ORIG,          " & vbCrLf & _
                       "A.TP_MESG,               " & vbCrLf & _
                       "A.TP_FORM_MESG_SAID,     " & vbCrLf & _
                       "D.CO_OCOR_MESG           " & vbCrLf & _
             " FROM     " & strOwner & ".TB_MESG A, " & vbCrLf & _
                       strOwner & ".TB_SITU_MESG D  " & vbCrLf & _
             " WHERE    A.CO_MESG      = D.CO_MESG  " & vbCrLf

    strSQL = strSQL & "  AND     D.CO_OCOR_MESG = ( SELECT MAX(CO_OCOR_MESG) " & _
                                                    " FROM " & strOwner & ".TB_SITU_MESG " & _
                                                    " WHERE CO_MESG = A.CO_MESG)"
    
    strSQL = strSQL & "  AND     A.DH_MESG BETWEEN " & _
                                                    fgDtHrXML_To_Oracle(pstrDataHoraDe) & " AND " & _
                                                    fgDtHrXML_To_Oracle(pstrDataHoraAte) & vbCrLf
    
    If pintEmpresa <> 0 Then
        strSQL = strSQL & " AND A.CO_EMPR_ORIG = " & pintEmpresa & vbCrLf
    End If
    
    If pstrSistema <> "" Then
        strSQL = strSQL & " AND A.SG_SIST_ORIG = '" & pstrSistema & "'" & vbCrLf
    End If
    
    If Len(pstrTipoMensagem) <> 0 Then
        strSQL = strSQL & " AND A.TP_MESG = '" & pstrTipoMensagem & "'" & vbCrLf
    End If
    
    If pstrIDMensagem <> "" Then
        strSQL = strSQL & " AND CO_CMPO_ATRB_IDEF_MESG like '%" & Trim(pstrIDMensagem) & "%'" & vbCrLf
    End If
    
    strSQL = strSQL & " AND D.CO_OCOR_MESG In (" & pstrCodOcorrencia & ")" & _
                      " ORDER BY  A.DH_MESG DESC"

    LerTodos = fgQueryXMLLerTodos("Monitoracao", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obtém o detalhe da mensagem informada.
Public Function ObterDetalheMensagem(ByVal plngCodigoMensagem As Long) As String
    
Dim strSQL                                  As String
Dim objDomDetalhe                           As MSXML2.DOMDocument40
Dim strDetalhe                              As String
Dim strMensagem                             As String
    
Dim strOwner                                As String
Dim lngCoTextXML                            As Long
    
On Error GoTo ErrorHandler

    If plngCodigoMensagem < 0 Then
        strOwner = "A7HIST"
    Else
        strOwner = "A7"
    End If

    strSQL = " Select                             " & vbCrLf & _
             "           A.TP_MESG,               " & vbCrLf & _
             "           A.CO_MESG,               " & vbCrLf & _
             "           A.CO_CMPO_ATRB_IDEF_MESG," & vbCrLf & _
             "           A.CO_EMPR_ORIG,          " & vbCrLf & _
             "           A.SG_SIST_ORIG,          " & vbCrLf & _
             "           '' as TX_CNTD_ENTR,      " & vbCrLf & _
             "           '' as TX_CNTD_SAID,    " & vbCrLf & _
             "           A.CO_TEXT_XML_ENTR, " & vbCrLf & _
             "           A.CO_TEXT_XML_SAID " & vbCrLf & _
             " From      " & strOwner & ".TB_MESG       A       " & vbCrLf & _
             " Where     A.CO_MESG           =           " & IIf(strOwner = "A7HIST", plngCodigoMensagem * -1, plngCodigoMensagem)

    strDetalhe = fgQueryXMLLer("DetalheMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Set objDomDetalhe = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(objDomDetalhe, "", "DetalheMonitoracao", "")
    
    Call fgAppendXML(objDomDetalhe, "DetalheMonitoracao", strDetalhe)
        
    If Trim(objDomDetalhe.selectSingleNode("//DetalheMensagem/CO_TEXT_XML_ENTR").Text) <> vbNullString Then
        lngCoTextXML = CLng(objDomDetalhe.selectSingleNode("//DetalheMensagem/CO_TEXT_XML_ENTR").Text)
        lngCoTextXML = IIf(strOwner = "A7HIST", lngCoTextXML * -1, lngCoTextXML)
        strMensagem = fgSelectVarchar4000(lngCoTextXML)
        objDomDetalhe.selectSingleNode("//DetalheMensagem/TX_CNTD_ENTR").appendChild fgCreateCDATASection(strMensagem)
    End If
        
    If Trim(objDomDetalhe.selectSingleNode("//DetalheMensagem/CO_TEXT_XML_SAID").Text) <> vbNullString And _
       Trim(objDomDetalhe.selectSingleNode("//DetalheMensagem/CO_TEXT_XML_SAID").Text) <> "0" Then
        lngCoTextXML = CLng(objDomDetalhe.selectSingleNode("//DetalheMensagem/CO_TEXT_XML_SAID").Text)
        lngCoTextXML = IIf(strOwner = "A7HIST", lngCoTextXML * -1, lngCoTextXML)
        strMensagem = fgSelectVarchar4000(lngCoTextXML)
        objDomDetalhe.selectSingleNode("//DetalheMensagem/TX_CNTD_SAID").appendChild fgCreateCDATASection(strMensagem)
    End If
        
    Call fgAppendXML(objDomDetalhe, "DetalheMonitoracao", ObterHistoricoMernsagem(plngCodigoMensagem))

    ObterDetalheMensagem = objDomDetalhe.xml
    
    Set objDomDetalhe = Nothing
    
    Exit Function

ErrorHandler:
    
    Set objDomDetalhe = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obtém o histórico de situações da mensagem informada.
Private Function ObterHistoricoMernsagem(ByVal plngCodigoMensagem As Long) As String
    
Dim strSQL                                  As String
Dim strOwner                                As String
    
On Error GoTo ErrorHandler

    If plngCodigoMensagem < 0 Then
        strOwner = "A7HIST"
    Else
        strOwner = "A7"
    End If

    strSQL = " Select    A.DH_OCOR_MESG,     " & vbCrLf & _
             "           B.DE_OCOR_MESG,     " & vbCrLf & _
             "           B.DE_ABRV_OCOR_MESG, " & vbCrLf & _
             "           A.TX_DTLH_OCOR_ERRO " & vbCrLf & _
             " From      " & strOwner & ".TB_SITU_MESG A,  " & vbCrLf & _
             "           A7.TB_OCOR_MESG B   " & vbCrLf & _
             " Where     A.CO_OCOR_MESG  = B.CO_OCOR_MESG " & vbCrLf & _
             " And       A.CO_MESG       = " & IIf(strOwner = "A7HIST", plngCodigoMensagem * -1, plngCodigoMensagem) & vbCrLf & _
             " Order by  A.NU_SEQU_SITU_MESG DESC"

    ObterHistoricoMernsagem = fgQueryXMLLerTodos("SituacaoMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoMernsagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

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



' Obtém o histórico de situações das mensagens de ECO GEN0001
Public Function ObterHistoricoGEN0001() As String
    
Dim strSQL                                  As String
Dim objRSGEN0001                            As ADODB.Recordset
Dim objRSGEN0001R1                          As ADODB.Recordset
Dim strDataDe                               As String
Dim strDataAte                              As String
Dim xmlStatusGEN0001                        As MSXML2.DOMDocument40
Dim xmlGrupoGEN0001                         As MSXML2.DOMDocument40
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim strNumCtrlifReq                         As String
Dim strNumCtrlifR1                          As String
Dim blnTemR1                                As Boolean

Dim udtProtocoloNZ                          As udtProtocoloNZ
Dim udtProtocoloNZAux                       As udtProtocoloNZAux

On Error GoTo ErrorHandler

    strDataDe = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDataAte = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")

    strSQL = " SELECT   CO_TEXT_XML_SAID, " & _
             "          CO_TEXT_XML_ENTR, " & _
             "          DH_MESG, " & _
             "          CO_CMPO_ATRB_IDEF_MESG " & _
             "   FROM   A7.TB_MESG " & _
             "  WHERE   TP_MESG = 'GEN0001'" & _
             "    AND   DH_MESG BETWEEN " & strDataDe & " AND " & strDataAte
    
    Set objRSGEN0001 = fgQuerySQL(strSQL)
    
    If objRSGEN0001.EOF Then
        objRSGEN0001.Close
        Exit Function
    End If
    
    Set xmlStatusGEN0001 = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlStatusGEN0001, "", "REPET_GEN0001", "")
    
    While Not objRSGEN0001.EOF
                
        blnTemR1 = False
        
        Set xmlGrupoGEN0001 = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlGrupoGEN0001, "", "GRUP_GEN0001", "")
        Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "DH_MESG", fgDtHr_To_Xml(objRSGEN0001!DH_MESG))
        
        'Obter ISPB Destino -------------------------------------------------------------------------
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
        strMensagem = fgSelectVarchar4000(objRSGEN0001!CO_TEXT_XML_ENTR)
        xmlMensagem.loadXML Mid(strMensagem, 21)
        
        Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "CO_ISPB_DEST", xmlMensagem.selectSingleNode("//ISPBDestinatario").Text)
        Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "CO_ISPB_EMIS", xmlMensagem.selectSingleNode("//ISPBEmissor").Text)
        
        
        'Obter Numero Controle IF da mensagem GEN0001
        strNumCtrlifReq = Trim(xmlMensagem.selectSingleNode("//NumCtrlIF").Text)
        
        Set xmlMensagem = Nothing
        '---------------------------------------------------------------------------------------------
        
        'Verifica se existe GEN0001R1 ----------------------------------------------------------------
        strSQL = " SELECT   CO_TEXT_XML_SAID," & _
                 "          DH_MESG   " & _
                 "   FROM   A7.TB_MESG " & _
                 "  WHERE   CO_CMPO_ATRB_IDEF_MESG = 'GEN0001R1'" & _
                 "    AND   DH_MESG BETWEEN " & strDataDe & " AND " & strDataAte

        Set objRSGEN0001R1 = fgQuerySQL(strSQL)
                
        If objRSGEN0001R1.EOF Then
            'Se a mensagem R1 não chegar em ate 3 minutos Pintar em vermelho
            If DateDiff("n", objRSGEN0001!DH_MESG, Now) > 3 Then
                Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "COR", "RED")
            Else
                Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "COR", "BLACK")
            End If
            
            Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "DE_STAT_MESG", "Mensagem enviada")
        Else
                        
            'Procurar a R1 correspondente
                        
            While Not objRSGEN0001R1.EOF
                                            
                ''Obter Numero Controle IF da mensagem GEN0001R1
                Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
                strMensagem = fgSelectVarchar4000(objRSGEN0001R1!CO_TEXT_XML_SAID)
                xmlMensagem.loadXML strMensagem
                            
                udtProtocoloNZAux.String = xmlMensagem.selectSingleNode("//TX_MESG").Text
                LSet udtProtocoloNZ = udtProtocoloNZAux
                                                                                
                strNumCtrlifR1 = Trim(udtProtocoloNZ.ControleRemessaNZ)
                                                                                
                Set xmlMensagem = Nothing
                '------------------------------------------------------------------------------
                
                If strNumCtrlifReq = strNumCtrlifR1 Then
                    'Se a mensagem R1 existir Pintar em azul
                    Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "COR", "BLUE")
                    Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "DE_STAT_MESG", "Mensagem resposta recebida")
                    Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "DH_MESG_R1", fgDtHr_To_Xml(objRSGEN0001R1!DH_MESG))
                    blnTemR1 = True
                End If
                
                objRSGEN0001R1.MoveNext
            Wend
                        
            If Not blnTemR1 Then
                If DateDiff("n", objRSGEN0001!DH_MESG, Now) > 3 Then
                    Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "COR", "RED")
                Else
                    Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "COR", "BLACK")
                End If
                
                Call fgAppendNode(xmlGrupoGEN0001, "GRUP_GEN0001", "DE_STAT_MESG", "Mensagem enviada")
            End If
            
            
            
        End If
        
        Call fgAppendXML(xmlStatusGEN0001, "REPET_GEN0001", xmlGrupoGEN0001.xml)
        
        objRSGEN0001R1.Close
        Set xmlGrupoGEN0001 = Nothing
                        
        '---------------------------------------------------------------------------------------------
                        
        objRSGEN0001.MoveNext
    Wend
                    
    objRSGEN0001.Close
    
    ObterHistoricoGEN0001 = xmlStatusGEN0001.xml

    Set xmlStatusGEN0001 = Nothing

    Exit Function

ErrorHandler:
    
    Set xmlMensagem = Nothing
    Set xmlStatusGEN0001 = Nothing
    Set xmlGrupoGEN0001 = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterHistoricoGEN0001 Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

