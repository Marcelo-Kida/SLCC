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
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF75470030F"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Objeto responsável pela persistência de mensagem transitadas no sistema A7 bem como controle dos dados pertinentes a ela.
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

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

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_MESG.
Public Function ObterPropriedades()

Dim strSQL                                   As String

On Error GoTo ErrorHandler
  
    strSQL = " SELECT   CO_MESG,                " & _
            "           TP_FORM_MESG_SAID,           " & _
            "           DH_MESG,                " & _
            "           TP_MESG,                " & _
            "           CO_EMPR_ORIG,           " & _
            "           SG_SIST_ORIG,           " & _
            "           '' AS SG_SIST_DEST,     " & _
            "           DH_INIC_VIGE_REGR_TRAP, " & _
            "           CO_CMPO_ATRB_IDEF_MESG, " & _
            "           CO_MESG_MQSE,           " & _
            "           CO_TEXT_XML_ENTR,       " & _
            "           CO_TEXT_XML_SAID,       " & _
            "           '' as TX_DTLH_OCOR_ERRO," & _
            "           '' as TX_CNTD_ENTR,     " & _
            "           '' as TX_CNTD_SAID      " & _
            " FROM A7.TB_MESG                   "
    
    ObterPropriedades = fgPropriedades("Grupo_Mensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar a mensagem de saída após o transito pelo sistema A7.
Public Function SalvarMensagemDadosSaida(ByRef pxmlMensagem As MSXML2.DOMDocument40) As Long

Dim lngCodigoMensagem                       As Long
Dim lngCodCntdAtrbXml                       As Long

Dim strSQL                                  As String

On Error GoTo ErrorHandler
            
    With pxmlMensagem.documentElement
            
        lngCodigoMensagem = CLng(.selectSingleNode("CO_MESG").Text)
                    
        lngCodCntdAtrbXml = fgInsertVarchar4000(.selectSingleNode("TX_CNTD_SAID").Text)
        
        strSQL = " UPDATE A7.TB_MESG " & _
                 "    SET CO_TEXT_XML_SAID   =  " & lngCodCntdAtrbXml & "," & _
                 "        CO_CMPO_ATRB_IDEF_MESG  = '" & Mid(.selectSingleNode("CO_CMPO_ATRB_IDEF_MESG").Text, 1, 50) & "'" & _
                 "  WHERE CO_MESG                 =  " & lngCodigoMensagem
                 
        fgExecuteSQL strSQL
    
    End With
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarMensagemDadosSaida Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar a mensagem de entrada do sistema A7.
Public Function SalvarMensagem(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                               ByVal plngCodigoOcorrencia As Long) As Long


Dim objSituacaoMensagem                     As A7Server.clsSituacaoMensagem

Dim strSQL                                  As String

Dim lngCodigoMensagem                       As Long

Dim lngCodCntdAtrbXmlEntrada                As Long
Dim lngCodCntdAtrbXmlSaida                  As Long

Dim strCoTextXMLSaida                       As String
Dim strMesgID                               As String
Dim strCoOperAtiv                           As String

On Error GoTo ErrorHandler
                
    lngCodCntdAtrbXmlEntrada = fgInsertVarchar4000(pxmlMensagem.selectSingleNode("//Grupo_Mensagem/TX_CNTD_ENTR").Text)
    
    strCoTextXMLSaida = "NULL"
    strMesgID = "NULL"
    strCoOperAtiv = "NULL"
    
    If Not pxmlMensagem.selectSingleNode("//Grupo_Mensagem/TX_CNTD_ENTR") Is Nothing Then
        If Trim(pxmlMensagem.selectSingleNode("//Grupo_Mensagem/TX_CNTD_ENTR").Text) <> vbNullString Then
            lngCodCntdAtrbXmlSaida = fgInsertVarchar4000(pxmlMensagem.selectSingleNode("//Grupo_Mensagem/TX_CNTD_SAID").Text)
            strCoTextXMLSaida = lngCodCntdAtrbXmlSaida
        End If
    End If
    
    If Not pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_MESG_MQSE") Is Nothing Then
        If Trim(pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_MESG_MQSE").Text) <> vbNullString Then
            strMesgID = "'" & pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_MESG_MQSE").Text & "'"
        End If
    End If
    
    If Not pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_CMPO_ATRB_IDEF_MESG") Is Nothing Then
        If Trim(pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_CMPO_ATRB_IDEF_MESG").Text) <> vbNullString Then
            strCoOperAtiv = "'" & pxmlMensagem.selectSingleNode("//Grupo_Mensagem/CO_CMPO_ATRB_IDEF_MESG").Text & "'"
        End If
    End If
                
    
                
    lngCodigoMensagem = fgExecuteSequence("A7.SQ_A7_CO_MESG")
    
    With pxmlMensagem.documentElement
    
        strSQL = " INSERT INTO A7.TB_MESG ( " & _
                "        CO_MESG , " & _
                "        SG_SIST_ORIG , " & _
                "        CO_MESG_MQSE , " & _
                "        TP_MESG      , " & _
                "        CO_EMPR_ORIG , " & _
                "        DH_INIC_VIGE_REGR_TRAP , " & _
                "        DH_MESG ,       " & _
                "        CO_CMPO_ATRB_IDEF_MESG,   " & _
                "        CO_TEXT_XML_ENTR, " & _
                "        CO_TEXT_XML_SAID, " & _
                "        TP_FORM_MESG_SAID      " & _
                " ) VALUES ( " & _
                lngCodigoMensagem & "," & _
                "'" & .selectSingleNode("SG_SIST_ORIG").Text & "', " & _
                strMesgID & ", " & _
                "'" & IIf(IsNumeric(.selectSingleNode("TP_MESG").Text), Val(.selectSingleNode("TP_MESG").Text), Trim(.selectSingleNode("TP_MESG").Text)) & "', " & _
                CLng(.selectSingleNode("CO_EMPR_ORIG").Text) & ", " & _
                fgDtHrXML_To_Oracle(.selectSingleNode("DH_INIC_VIGE_REGR_TRAP").Text) & ", " & _
                "SYSDATE , " & _
                strCoOperAtiv & " , " & _
                lngCodCntdAtrbXmlEntrada & ", " & _
                strCoTextXMLSaida & ", " & _
                .selectSingleNode("TP_FORM_MESG_SAID").Text & ")"
    End With
    
    Call fgExecuteSQL(strSQL)
               
    pxmlMensagem.documentElement.selectSingleNode("CO_MESG").Text = lngCodigoMensagem
    
    Set objSituacaoMensagem = CreateObject("A7Server.clsSituacaoMensagem")
    Call objSituacaoMensagem.SalvarSituacaoMensagem(pxmlMensagem, enumOcorrencia.RecebimentoBemSucedido)
    Set objSituacaoMensagem = Nothing
               
    Set objSituacaoMensagem = CreateObject("A7Server.clsSituacaoMensagem")
    Call objSituacaoMensagem.SalvarSituacaoMensagem(pxmlMensagem, plngCodigoOcorrencia)
    Set objSituacaoMensagem = Nothing
        
    SalvarMensagem = lngCodigoMensagem

    Exit Function
ErrorHandler:
    
    Set objSituacaoMensagem = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar o MessageID da mensagem enviada.
Public Function AtualizaMessageID(ByVal plngCodigoMensagem As Long, _
                                  ByVal pstrMessageIdHex As String) As Long

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = " UPDATE A7.TB_MESG  " & _
            "     SET CO_MESG_MQSE  ='" & pstrMessageIdHex & "'" & _
            "   WHERE CO_MESG       = " & plngCodigoMensagem
                
    fgExecuteSQL strSQL

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AtualizaMessageID Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ExisteMensagem(ByVal pstrTipoMensagem As String, _
                               ByVal pstrSistemaOrigem As String, _
                               ByVal plngCodigoEmpresa As Long, _
                               ByVal pstrDataHoraInicioVigencia As String) As Boolean

Dim strSQL                     As String
Dim Rs                         As Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT ( " & _
             " SELECT COUNT (*) " & _
             " FROM   A7.TB_MESG " & _
             " WHERE  TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
             " AND    SG_SIST_ORIG = '" & pstrSistemaOrigem & "'" & _
             " AND    CO_EMPR_ORIG = " & plngCodigoEmpresa & _
             " AND    DH_INIC_VIGE_REGR_TRAP = " & fgDtHrXML_To_Oracle(pstrDataHoraInicioVigencia)  '& ") "
    
    strSQL = strSQL & " ) + ( " & _
             " SELECT COUNT (*) " & _
             " FROM   A7HIST.TB_MESG " & _
             " WHERE  TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
             " AND    SG_SIST_ORIG = '" & pstrSistemaOrigem & "'" & _
             " AND    CO_EMPR_ORIG = " & plngCodigoEmpresa & _
             " AND    DH_INIC_VIGE_REGR_TRAP = " & fgDtHrXML_To_Oracle(pstrDataHoraInicioVigencia) '& ") "

    strSQL = strSQL & " ) AS QTD FROM DUAL "
    
    Set Rs = fgQuerySQL(strSQL)
    
    ExisteMensagem = (Rs!QTD > 0)
    
    Rs.Close
    
    Set Rs = Nothing

    Exit Function
ErrorHandler:
    Set Rs = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExisteMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter o layout da mensagem SPB.
Public Function ObterEstruturaMensagemSPB(ByVal pstrTipoMensagem As String) As String

Dim Rs                                      As ADODB.Recordset
Dim strSQL                                  As String
Dim strMensagemSPB                          As String
Dim strNomeTagPrincipal                     As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   A.NO_ATRB_MESG, " & _
             "          B.NO_TITU_MESG  " & _
             "   FROM   A7.TB_TIPO_MESG_ATRB A, " & _
             "          A7.TB_TIPO_MESG B       " & _
             "  WHERE   A.TP_MESG = B.TP_MESG   " & _
             "    AND   A.TP_MESG = '" & Trim(pstrTipoMensagem) & "'" & _
             "    AND   A.TP_FORM_MESG = " & enumTipoParteSaida.ParteXML & _
             " ORDER BY NU_ORDE_AGRU_ATRB "
    
    Set Rs = fgQuerySQL(strSQL)
    
    strMensagemSPB = "<SISMSG>"
    
    While Not Rs.EOF
            
        If strNomeTagPrincipal = vbNullString Then
            strNomeTagPrincipal = Trim(Rs!NO_TITU_MESG)
            strMensagemSPB = strMensagemSPB & "<" & strNomeTagPrincipal & ">"
        End If
        
        If Mid(Trim(Rs!NO_ATRB_MESG), 1, 6) <> "Grupo_" And Mid(Trim(Rs!NO_ATRB_MESG), 1, 7) <> "/Grupo_" Then
            strMensagemSPB = strMensagemSPB & "<" & Trim(Rs!NO_ATRB_MESG) & "/>"
        Else
            strMensagemSPB = strMensagemSPB & "<" & Trim(Rs!NO_ATRB_MESG) & ">"
        End If
    
        Rs.MoveNext
    Wend
    
    strMensagemSPB = strMensagemSPB & "</" & strNomeTagPrincipal & ">"
    strMensagemSPB = strMensagemSPB & "</SISMSG>"
    
    ObterEstruturaMensagemSPB = strMensagemSPB
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterEstruturaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Enviar teste de conexão a câmaras participantes do SPB através da mensagem GEN0001.
Public Sub EnviarTesteConexao(ByVal pstrXMLTesteConectividade As String)


Dim udtProtocolo                            As udtProtocolo
Dim udtProtocoloAux                         As udtProtocoloAux
Dim strHeaderNZ                             As String
Dim strNumeroControleIF                     As String
Dim xmlTesteConectividade                   As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim objGerenciadorEnvio                     As A7Server.clsGerenciadorEnvio

On Error GoTo ErrorHandler
    
    Set xmlTesteConectividade = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlTesteConectividade.loadXML pstrXMLTesteConectividade
    
    For Each xmlNode In xmlTesteConectividade.selectNodes("//TESTE_CONECTIVIDADE/*")
        
        strNumeroControleIF = ""
        udtProtocolo.CodigoEmpresa = fgCompletaString(xmlNode.selectSingleNode("CO_EMPR").Text, "0", 5, True)
        udtProtocolo.SiglaSistemaDestino = "NZ "
        udtProtocolo.SiglaSistemaOrigem = "A8 "
        udtProtocolo.TipoMensagem = "GEN0001"
        LSet udtProtocoloAux = udtProtocolo
    
        strHeaderNZ = fgMontaHeaderMensageNZ("GEN0001", _
                                             xmlNode.selectSingleNode("CO_EMPR").Text, _
                                             strNumeroControleIF)
         
        Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
       
        fgAppendNode xmlMensagem, "", "MESG", ""
        fgAppendNode xmlMensagem, "MESG", "CodMsg", "GEN0001"
        fgAppendNode xmlMensagem, "MESG", "ISPBEmissor", xmlNode.selectSingleNode("CO_ISPB_EMIS").Text
        fgAppendNode xmlMensagem, "MESG", "ISPBDestinatario", xmlNode.selectSingleNode("CO_ISPB_DEST").Text
        fgAppendNode xmlMensagem, "MESG", "MsgECO", "Teste de Conectividade SLCC"
        fgAppendNode xmlMensagem, "MESG", "TX_HEAD_NZ", strHeaderNZ
        fgAppendNode xmlMensagem, "MESG", "NumCtrlIF", strNumeroControleIF
        
        strMensagem = udtProtocoloAux.String & xmlMensagem.xml
            
        Set xmlMensagem = Nothing
        
        Set objGerenciadorEnvio = CreateObject("A7Server.clsGerenciadorEnvio")
        Call objGerenciadorEnvio.EnviarTesteConexao(strMensagem)
        Set objGerenciadorEnvio = Nothing
    
    Next
        
    Set xmlTesteConectividade = Nothing
     
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If
    
    Exit Sub
ErrorHandler:
    
    Set xmlTesteConectividade = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.DisableCommit
    End If
    
    Set objGerenciadorEnvio = Nothing
    Set xmlMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "EnviarTesteConexao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub



