VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemSLCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas às
'' operações/Mensagens processadas pelo A8.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Enum enumTipoControleDominio
    DominioInterno = 1
    DominioSPB = 2
    DominioTabela = 3
End Enum

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

'Seleciona todas as mensagens de uma empresa e de um determinado tipo
Public Function LerMensagem(ByVal plngTipoMensagem As Long, _
                   Optional ByVal plngTipoOperacao As Long = 0) As String

Dim strSQL                                  As String
Dim strMensagem                             As String
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
   
    strSQL = " SELECT   A.TP_MESG,                                    " & vbNewLine & _
             "          A.NO_TIPO_MESG,                               " & vbNewLine & _
             "          B.NU_ORDE_AGRU_ATRB,                          " & vbNewLine & _
             "          B.IN_OBRI_ATRB,                               " & vbNewLine & _
             "          B.TP_FORM_MESG,                               " & vbNewLine & _
             "          B.NU_NIVE_MESG_ATRB,                          " & vbNewLine & _
             "          B.QT_REPE,                                    " & vbNewLine & _
             "          C.NO_ATRB_MESG,                               " & vbNewLine & _
             "          C.NO_TRAP_ATRB,                               " & vbNewLine & _
             "          C.TP_DADO_ATRB_MESG,                          " & vbNewLine & _
             "          C.QT_CTER_ATRB,                               " & vbNewLine & _
             "          C.QT_CASA_DECI_ATRB,                          " & vbNewLine & _
             "          '' AS TX_DOMI                                 " & vbNewLine & _
             " FROM     A7.TB_TIPO_MESG A,                            " & vbNewLine & _
             "          A7.TB_TIPO_MESG_ATRB B,                       " & vbNewLine & _
             "          A7.TB_ATRB_MESG C                             " & vbNewLine & _
             " WHERE    A.TP_MESG         = B.TP_MESG                 " & vbNewLine & _
             " AND      B.NO_ATRB_MESG    = C.NO_ATRB_MESG            " & vbNewLine & _
             " AND      A.TP_MESG         = '" & plngTipoMensagem & "'" & vbNewLine & _
             " AND      B.TP_FORM_MESG    = 2                         " & vbNewLine & _
             " ORDER BY A.TP_MESG,                                    " & vbNewLine & _
             "          B.NU_ORDE_AGRU_ATRB"

    strMensagem = fgQueryXMLLerTodos("Mensagem", strSQL, App.EXEName & "." & TypeName(Me))
            
    If strMensagem = vbNullString Then Exit Function
        
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML strMensagem
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        
        If Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_USUA_CADR_OPER" Then
            xmlNode.selectSingleNode("TX_DOMI").Text = fgUsuarioRede
        Else
            'Layout 136
            If plngTipoMensagem = enumTipoMensagemBUS.DespesasBMC Then
                If Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_OPER_ATIV" _
                Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_IDEF_TRAF" _
                Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_VEIC_LEGA" Then
                    xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection("TABELA=A8.TB_MESG_RECB_ENVI_SPB")
                Else
                    xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                End If
            'Layout 249
            ElseIf plngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoInterbancaria Then
                If plngTipoOperacao <> 0 Then
                    'Operacao 231 ou 233
                    If plngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or plngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
                        If Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_REG_OPER_CAMB" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_CNPJ_BASE_IF" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_MOED_ISO" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_MOED_ESTRG" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_TAXA_CAMB" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_MOED_NACIO" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "DT_ENTR_MOED_NACIO" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "DT_ENTR_MOED_ESTR" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "DT_LIQU_OPER" Then
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection("TABELA=CO_REG_OPER_CAMB")
                        Else
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                        End If
                    'Operacao 234
                    ElseIf plngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        If Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "ChACAM" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "DT_OPER_ATIV" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "TP_OPER_CAMB" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_TAXA_CAMB" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_MOED_NACIO" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "VA_MOED_ESTRG" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "DT_LIQU_OPER" Then
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection("TABELA=ChACAM")
                        Else
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                        End If
                    Else
                        xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                    End If
                Else
                    xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                End If
            'Layout 253
            ElseIf plngTipoMensagem = enumTipoMensagemBUS.RegistroOperacaoArbitragem Then
                If plngTipoOperacao <> 0 Then
                    'Operacao 237
                    If plngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
                        If Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_REG_OPER_CAMB" _
                        Or Trim(xmlNode.selectSingleNode("NO_ATRB_MESG").Text) = "CO_REG_OPER_CAMB2" Then
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection("TABELA=CO_REG_OPER_CAMB2")
                        Else
                            xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                        End If
                    Else
                        xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                    End If
                Else
                    xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
                End If
            Else
                xmlNode.selectSingleNode("TX_DOMI").appendChild fgCreateCDATASection(ObterDominioAtributo(xmlNode.selectSingleNode("NO_ATRB_MESG").Text))
            End If
        End If
    Next
    
    LerMensagem = xmlMensagem.xml
    
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificarTipoMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

'' Consultar o dominio válido para o Atributo de um tipo de mensagem.
''
'' Onde o campo da tabela TP_CNTR_DOMI é o tipo de controle do dominio.
''
'' -Dominio Interno = 1
'' ObterDominioAtributo
''
'' -Dominio SPB = 2
'' Metodo para consultar : flObterDominioSPB
''
'' -Dominio Tabela = 3
'' LerTodosDominioTabela
Public Function ObterDominioAtributo(ByVal pstrNomeAtributo As String) As String

Dim strSQL                                  As String
Dim strDominio                              As String
Dim xmlDominio                              As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    strSQL = " SELECT NO_ATRB,      " & _
             "        CO_DOMI,      " & _
             "        TP_CNTR_DOMI, " & _
             "        DE_DOMI       " & _
             "   FROM A8.TB_CTRL_DOMI " & _
             "  WHERE NO_ATRB = '" & Trim(pstrNomeAtributo) & "'" & _
             " ORDER BY TP_CNTR_DOMI "
             
    strDominio = fgQueryXMLLerTodos("DominioAtributo", strSQL, "DominioAtributo")
        
    If Trim(strDominio) = vbNullString Then
        ObterDominioAtributo = vbNullString
        Exit Function
    End If
    
    Set xmlDominio = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlDominio.loadXML strDominio
    
    If CLng(xmlDominio.selectSingleNode("//TP_CNTR_DOMI").Text) = enumTipoControleDominio.DominioInterno Then

    ElseIf CLng(xmlDominio.selectSingleNode("//TP_CNTR_DOMI").Text) = enumTipoControleDominio.DominioSPB Then
        strDominio = flObterDominioSPB(Trim(xmlDominio.selectSingleNode("//CO_DOMI").Text))
    ElseIf CLng(xmlDominio.selectSingleNode("//TP_CNTR_DOMI").Text) = enumTipoControleDominio.DominioTabela Then
        strDominio = "TABELA=" & xmlDominio.selectSingleNode("//CO_DOMI").Text
    End If
    
    ObterDominioAtributo = strDominio

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDominioAtributo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar o dominios válidos para o Atributo do catalogo SPB.
Public Function ObterDominioSPB(ByVal pstrNomeTipoTag As String, _
                       Optional ByVal pstrTagRetorno As String = "") As String

    ObterDominioSPB = flObterDominioSPB(pstrNomeTipoTag, pstrTagRetorno)

End Function

'' Consultar o dominios válidos para o Atributo do catalogo SPB.
Private Function flObterDominioSPB(ByVal pstrNomeTipoTag As String, _
                          Optional ByVal pstrTagRetorno As String = "") As String

Dim strSQL                                  As String
Dim strDominio                              As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   A.NO_TIPO_TAG," & vbCrLf & _
             "          TRIM(B.CO_DOMI) AS CO_DOMI,    " & vbCrLf & _
             "          B.DE_DOMI     " & vbCrLf & _
             "   FROM   A8.TB_TIPO_TAG A, " & vbCrLf & _
             "          A8.TB_DOMINIO B " & vbCrLf & _
             "  WHERE   A.SQ_TIPO_TAG = B.SQ_TIPO_TAG " & vbCrLf & _
             "    AND   A.NO_TIPO_TAG = '" & pstrNomeTipoTag & "'" & vbCrLf & _
             "  ORDER BY B.CO_DOMI       " & vbCrLf

    strDominio = fgQueryXMLLerTodos(IIf(pstrTagRetorno <> "", pstrTagRetorno, "DominioAtributo"), strSQL, "DominioAtributo")
            
    flObterDominioSPB = strDominio
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDominioSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
'Consultar os Dominios de um atributo
Public Function LerTodosDominioTabela(ByVal pvntCodigoEmpresa As Long, _
                                      ByVal pstrNomeTabela As String, _
                                      ByVal pstrMensagem As String, _
                             Optional ByVal pvntCodigo As Variant = 0, _
                             Optional ByVal pstrDescricao As String = "") As String

Dim strSQL                                  As String
Dim strDataServidor                         As String
Dim strSegregaDados                         As String
Dim objFuncao                               As A6A7A8.clsA6A7A8Funcoes

Dim strDHInicio                             As String
Dim strDHFim                                As String


On Error GoTo ErrorHandler
                
    strDataServidor = Format(fgDataHoraServidor(enumFormatoDataHora.Data), "YYYYMMDD")
                
    Select Case UCase(Trim(pstrNomeTabela))
    
        Case "PJ.TB_PRODUTO"

            strSQL = "  SELECT CO_PROD AS CODIGO   ," & _
                     "         DE_PROD AS DESCRICAO " & _
                     "    FROM A8.TB_PRODUTO    A,  " & _
                     "         A8.TB_EMPRESA_HO B   " & _
                     "   WHERE A.CO_EMPR_FUSI  = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR       = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
            
            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_PROD = " & pvntCodigo
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_PROD) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If
            
            strSQL = strSQL & " ORDER BY CO_PROD "

        Case "PJ.TB_SEGMENTO"
        
            strSQL = "  SELECT CO_SEGM AS CODIGO   ," & _
                     "         DE_SEGM AS DESCRICAO " & _
                     "    FROM A8.TB_SEGMENTO   A,  " & _
                     "         A8.TB_EMPRESA_HO B   " & _
                     "   WHERE A.CO_EMPR_FUSI = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR       = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                     
            
            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_SEGM = " & pvntCodigo
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_SEGM) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If
            
            strSQL = strSQL & " ORDER BY CO_SEGM "
        
        Case "PJ.TB_EVENTO_FINANCEIRO"
        
            strSQL = "  SELECT CO_EVEN_FINC AS CODIGO,     " & _
                     "         DE_EVEN_FINC AS DESCRICAO   " & _
                     "    FROM A8.TB_EVENTO_FINANCEIRO A,  " & _
                     "         A8.TB_EMPRESA_HO        B   " & _
                     "   WHERE A.CO_EMPR_FUSI = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR      = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                     
            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_EVEN_FINC = " & pvntCodigo
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_EVEN_FINC) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If
            
            strSQL = strSQL & " ORDER BY CO_EVEN_FINC "
        
        Case "PJ.TB_INDEXADOR"
        
             strSQL = "  SELECT CO_INDX AS CODIGO   ," & _
                     "         DE_INDX AS DESCRICAO " & _
                     "    FROM A8.TB_INDEXADOR  A,  " & _
                     "         A8.TB_EMPRESA_HO B   " & _
                     "   WHERE A.CO_EMPR_FUSI = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR       = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
                     
            
            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_INDX = " & pvntCodigo
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_INDX) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If
            
            strSQL = strSQL & " ORDER BY CO_INDX "
        
        Case "PJ.TB_LOCAL_LIQUIDACAO"
    
            strSQL = "  SELECT CO_LOCA_LIQU AS CODIGO,       " & _
                     "         DE_LOCA_LIQU AS DESCRICAO,    " & _
                     "         SG_LOCA_LIQU AS SIGLA         " & _
                     "    FROM A8.TB_LOCAL_LIQUIDACAO    A,  " & _
                     "         A8.TB_EMPRESA_HO          B   " & _
                     "   WHERE A.CO_EMPR_FUSI = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR      = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"

            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_LOCA_LIQU = " & pvntCodigo
            End If

            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_LOCA_LIQU) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If

            strSegregaDados = fgSegregaDados("TB_LOCAL_LIQUIDACAO", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             False, _
                                             False, _
                                             True, _
                                             False, _
                                             False)
                    
            If Trim(strSegregaDados) <> vbNullString Then
                strSQL = strSQL & strSegregaDados
            End If

            strSQL = strSQL & " ORDER BY CO_LOCA_LIQU "

        Case "PJ.TB_TIPO_CONTA"
            
            strSQL = "  SELECT CO_TIPO_CNTA AS CODIGO,       " & _
                     "         DE_TIPO_CNTA AS DESCRICAO     " & _
                     "    FROM A8.TB_TIPO_CONTA    A,  " & _
                     "         A8.TB_EMPRESA_HO    B   " & _
                     "   WHERE A.CO_EMPR_FUSI = B.CO_EMPR_FUSI " & _
                     "     AND B.CO_EMPR      = " & pvntCodigoEmpresa & _
                     "     AND A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
                     "     AND (A.DT_FIM_VIGE IS NULL " & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")"
            
            If CLng(Val(pvntCodigo)) <> 0 Then
                strSQL = strSQL & " AND CO_TIPO_CNTA = " & pvntCodigo
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(DE_TIPO_CNTA) LIKE '%" & UCase(pstrDescricao) & "%'"
            End If
            
            strSQL = strSQL & " ORDER BY CO_TIPO_CNTA "
    
        Case "A8.TB_VEIC_LEGA"
            
            strSQL = "  SELECT  A.CO_VEIC_LEGA AS CODIGO,   " & vbCrLf & _
                     "          A.NO_VEIC_LEGA AS DESCRICAO," & vbCrLf & _
                     "          A.CO_CNTA_CUTD_PADR_SELIC," & vbCrLf & _
                     "          A.SG_SIST, " & vbCrLf & _
                     "          C.NO_SIST, " & vbCrLf & _
                     "          NVL(B.IN_SIST_SITU_CNTG," & enumIndicadorSimNao.Nao & ") AS IN_SIST_SITU_CNTG," & vbCrLf & _
                     "          A.CO_CNPJ_VEIC_LEGA " & vbCrLf & _
                     "    FROM  A8.TB_VEIC_LEGA      A, " & vbCrLf & _
                     "          A8.TB_SITU_CNTG_SIST B, " & vbCrLf & _
                     "          A7.TB_SIST C " & vbCrLf & _
                     "   WHERE  A.SG_SIST = B.SG_SIST (+) " & vbCrLf & _
                     "     AND  A.SG_SIST = C.SG_SIST " & vbCrLf & _
                     "     AND  A.CO_EMPR = C.CO_EMPR" & vbCrLf & _
                     "     AND  A.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & vbCrLf & _
                     "     AND (A.DT_FIM_VIGE  IS NULL " & vbCrLf & _
                     "      OR  A.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & vbCrLf
            
            If Trim$(pvntCodigo) <> vbNullString Then
                strSQL = strSQL & " AND CO_VEIC_LEGA LIKE '%" & pvntCodigo & "%' " & vbCrLf
            End If
            
            If Trim(pstrDescricao) <> "" Then
                strSQL = strSQL & " AND UPPER(NO_VEIC_LEGA) LIKE '%" & UCase(pstrDescricao) & "%'" & vbCrLf
            End If
            
            strSegregaDados = fgSegregaDados("A8.TB_VEIC_LEGA", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             False, _
                                             True, _
                                             False, _
                                             True, _
                                             False)

            If Trim(strSegregaDados) <> vbNullString Then
                strSQL = strSQL & strSegregaDados & vbCrLf
            End If
            
            strSQL = strSQL & " ORDER BY CO_VEIC_LEGA "
                        
        Case "A8.TB_TIPO_ATIV_MERC"
        
            strSQL = "SELECT NU_ATIV_MERC, " & vbCrLf & _
                            "DE_ATIV_MERC " & vbCrLf & _
                       "FROM A8.TB_TIPO_ATIV_MERC " & vbCrLf & _
                     " ORDER BY NU_ATIV_MERC "
        
        Case "A8.TB_MESG_RECB_ENVI_SPB"
            
            Set objFuncao = CreateObject("A6A7A8.clsA6A7A8Funcoes")
            
            strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(objFuncao.AdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 1, enumPaginacao.Anterior)) & "000000")
            strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
            
            strSQL = "SELECT " & _
                     "       A.NU_CTRL_IF ," & _
                     "       A.NU_SEQU_CNTR_REPE ," & _
                     "       A.DH_REGT_MESG_SPB ," & _
                     "       A.NU_CTRL_CAMR AS CONTROLE," & _
                     "       A.CO_VEIC_LEGA ," & _
                     "       B.VA_FINC AS VALOR, " & _
                     "       C.NO_VEIC_LEGA " & _
                     " FROM     A8.TB_MESG_RECB_ENVI_SPB A ," & _
                     "          A8.TB_MESG_RECB_SPB_CNCL B ," & _
                     "          A8.TB_VEIC_LEGA          C" & _
                     " WHERE    A.NU_CTRL_IF = B.NU_CTRL_IF" & _
                     "   AND    A.DH_REGT_MESG_SPB  = B.DH_REGT_MESG_SPB" & _
                     "   AND    A.NU_SEQU_CNTR_REPE = B.NU_SEQU_CNTR_REPE" & _
                     "   AND    A.CO_VEIC_LEGA      = C.CO_VEIC_LEGA" & _
                     "   AND    A.SG_SIST           = C.SG_SIST" & _
                     "   AND    A.TP_BKOF           = C.TP_BKOF" & _
                     "   AND    A.CO_EMPR           = " & pvntCodigoEmpresa & _
                     "   AND    A.CO_ULTI_SITU_PROC = " & enumStatusMensagem.AConciliar & _
                     "   AND    A.CO_MESG_SPB       = 'BMC0112'" & _
                     "   AND    A.DH_REGT_MESG_SPB BETWEEN " & strDHInicio & _
                     "   AND " & strDHFim

            strSegregaDados = fgSegregaDados("TB_MESG_RECB_ENVI_SPB", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             True, _
                                             True, _
                                             False, _
                                             False, _
                                             False)
            
            strSQL = strSQL & strSegregaDados
            strSQL = strSQL & " ORDER BY A.NU_CTRL_IF "
            
        Case "CHACAM"
            
            If pstrMensagem <> "BMC0015" Then
                pstrMensagem = ""
            End If
            
            strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
            strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
        
            strSQL = "SELECT " & _
                     "       A.NU_CTRL_IF ," & _
                     "       A.NU_SEQU_CNTR_REPE ," & _
                     "       A.CO_MESG_SPB ," & _
                     "       A.DH_REGT_MESG_SPB ," & _
                     "       B.CD_ASSO_CAMB ," & _
                     "       B.DT_OPER ," & _
                     "       DECODE(B.IN_OPER_DEBT_CRED, 1, 'V', 2, 'C') AS IN_OPER_DEBT_CRED ," & _
                     "       B.PE_TAXA_NEGO ," & _
                     "       B.VA_FINC ," & _
                     "       B.VA_MOED_ESTR ," & _
                     "       B.DT_LIQU " & _
                     " FROM  A8.TB_MESG_RECB_ENVI_SPB A, " & _
                     "       A8.TB_MESG_RECB_SPB_CNCL B " & _
                     " WHERE A.NU_CTRL_IF = B.NU_CTRL_IF " & _
                     " AND   A.DH_REGT_MESG_SPB  = B.DH_REGT_MESG_SPB " & _
                     " AND   A.NU_SEQU_CNTR_REPE = B.NU_SEQU_CNTR_REPE " & _
                     " AND   A.CO_EMPR           = " & pvntCodigoEmpresa & _
                     " AND   A.CO_ULTI_SITU_PROC = " & enumStatusMensagem.AConciliar & _
                     " AND   A.CO_MESG_SPB       = '" & pstrMensagem & "'"
                     
            strSQL = strSQL & _
                     " AND   B.CD_ASSO_CAMB NOT IN (SELECT C.CD_ASSO_CAMB " & _
                     "                              FROM   A8.TB_OPER_ATIV C " & _
                     "                              WHERE  B.CD_ASSO_CAMB            =  C.CD_ASSO_CAMB " & _
                     "                              AND    TRUNC(A.DH_REGT_MESG_SPB) =  C.DT_OPER_ATIV " & _
                     "                              AND    C.CO_ULTI_SITU_PROC       <> 19) " & _
                     " AND   A.DH_REGT_MESG_SPB BETWEEN " & strDHInicio & " AND " & strDHFim

            strSegregaDados = fgSegregaDados("TB_MESG_RECB_ENVI_SPB", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             True, _
                                             True, _
                                             False, _
                                             False, _
                                             False)
            
            strSQL = strSQL & strSegregaDados
            strSQL = strSQL & " ORDER BY A.NU_CTRL_IF "
            
        Case "CO_REG_OPER_CAMB"
            
            If pstrMensagem = "BMC0015" Then
                pstrMensagem = ""
            End If
            
            strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
            strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
        
            strSQL = "SELECT " & _
                     "       A.NU_CTRL_IF ," & _
                     "       A.NU_SEQU_CNTR_REPE ," & _
                     "       A.CO_MESG_SPB ," & _
                     "       A.DH_REGT_MESG_SPB ," & _
                     "       A.NU_COMD_OPER ," & _
                     "       A.CO_TEXT_XML " & _
                     " FROM  A8.TB_MESG_RECB_ENVI_SPB A " & _
                     " WHERE A.CO_EMPR           = " & pvntCodigoEmpresa & _
                     " AND   A.CO_ULTI_SITU_PROC = " & enumStatusMensagem.EnviadaLegado & _
                     " AND   A.CO_MESG_SPB       = '" & pstrMensagem & "'" & _
                     " AND   A.NU_COMD_OPER NOT IN (SELECT B.NU_COMD_OPER " & _
                     "                              FROM   A8.TB_OPER_ATIV B " & _
                     "                              WHERE  A.NU_COMD_OPER            =  B.NU_COMD_OPER " & _
                     "                              AND    TRUNC(A.DH_REGT_MESG_SPB) =  B.DT_OPER_ATIV " & _
                     "                              AND    B.CO_ULTI_SITU_PROC       <> 19) "
                 
            'Se Mensagem = CAM0006R2 ou CAM0009R2 nao carregar na tela se existir uma CAM0055(Expiracao) com o mesmo NU_COMD_OPER
            If pstrMensagem = "CAM0006R2" Or pstrMensagem = "CAM0009R2" Then
                strSQL = strSQL & _
                     " AND   A.NU_COMD_OPER NOT IN (SELECT C.NU_COMD_OPER " & _
                     "                              FROM   A8.TB_MESG_RECB_ENVI_SPB C " & _
                     "                              WHERE  A.NU_COMD_OPER            =  C.NU_COMD_OPER " & _
                     "                              AND    C.CO_MESG_SPB             =  'CAM0055' " & _
                     "                              AND    TRUNC(A.DH_REGT_MESG_SPB) =  TRUNC(C.DH_REGT_MESG_SPB) " & _
                     "                              AND    C.DH_REGT_MESG_SPB        =  (SELECT MAX(D.DH_REGT_MESG_SPB) " & _
                     "                                                                   FROM   A8.TB_MESG_RECB_ENVI_SPB D " & _
                     "                                                                   WHERE  C.NU_COMD_OPER         = D.NU_COMD_OPER " & _
                     "                                                                   AND TRUNC(C.DH_REGT_MESG_SPB) = TRUNC(D.DH_REGT_MESG_SPB))) "
            End If
                     
            strSQL = strSQL & _
                     " AND   A.DH_REGT_MESG_SPB BETWEEN " & strDHInicio & " AND " & strDHFim

            strSegregaDados = fgSegregaDados("TB_MESG_RECB_ENVI_SPB", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             True, _
                                             True, _
                                             False, _
                                             False, _
                                             False)
            
            strSQL = strSQL & strSegregaDados
            strSQL = strSQL & " ORDER BY A.NU_CTRL_IF "
            
        Case "CO_REG_OPER_CAMB2"
        
            If pstrMensagem <> "CAM0013R2" Then
                pstrMensagem = ""
            End If
        
            strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
            strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
        
            strSQL = "SELECT " & _
                     "       A.NU_CTRL_IF ," & _
                     "       A.NU_SEQU_CNTR_REPE ," & _
                     "       A.CO_MESG_SPB ," & _
                     "       A.DH_REGT_MESG_SPB ," & _
                     "       A.NU_COMD_OPER ," & _
                     "       A.NR_OPER_CAMB_2 " & _
                     " FROM  A8.TB_MESG_RECB_ENVI_SPB A " & _
                     " WHERE A.CO_EMPR           = " & pvntCodigoEmpresa & _
                     " AND   A.CO_ULTI_SITU_PROC = " & enumStatusMensagem.EnviadaLegado & _
                     " AND   A.CO_MESG_SPB       = '" & pstrMensagem & "'" & _
                     " AND   A.NR_OPER_CAMB_2    NOT IN (SELECT B.NR_OPER_CAMB_2 " & _
                     "                                   FROM   A8.TB_OPER_ATIV B " & _
                     "                                   WHERE  A.NR_OPER_CAMB_2          =  B.NR_OPER_CAMB_2 " & _
                     "                                   AND    TRUNC(A.DH_REGT_MESG_SPB) =  B.DT_OPER_ATIV " & _
                     "                                   AND    B.CO_ULTI_SITU_PROC       <> 19) "
                     
            'Se Mensagem = CAM0013R2 nao carregar na tela se existir uma CAM0055(Expiracao) com o mesmo NU_COMD_OPER
            If pstrMensagem = "CAM0013R2" Then
                strSQL = strSQL & _
                     " AND   A.NU_COMD_OPER NOT IN (SELECT C.NU_COMD_OPER " & _
                     "                              FROM   A8.TB_MESG_RECB_ENVI_SPB C " & _
                     "                              WHERE  A.NU_COMD_OPER            =  C.NU_COMD_OPER " & _
                     "                              AND    C.CO_MESG_SPB             =  'CAM0055' " & _
                     "                              AND    TRUNC(A.DH_REGT_MESG_SPB) =  TRUNC(C.DH_REGT_MESG_SPB) " & _
                     "                              AND    C.DH_REGT_MESG_SPB        =  (SELECT MAX(D.DH_REGT_MESG_SPB) " & _
                     "                                                                   FROM   A8.TB_MESG_RECB_ENVI_SPB D " & _
                     "                                                                   WHERE  C.NU_COMD_OPER         = D.NU_COMD_OPER " & _
                     "                                                                   AND TRUNC(C.DH_REGT_MESG_SPB) = TRUNC(D.DH_REGT_MESG_SPB))) "
            End If
                     
            strSQL = strSQL & _
                     " AND   A.DH_REGT_MESG_SPB BETWEEN " & strDHInicio & " AND " & strDHFim
                     
            strSegregaDados = fgSegregaDados("TB_MESG_RECB_ENVI_SPB", _
                                             False, _
                                             "A", _
                                             "A", _
                                             False, _
                                             True, _
                                             True, _
                                             False, _
                                             False, _
                                             False)
            
            strSQL = strSQL & strSegregaDados
            strSQL = strSQL & " ORDER BY A.NU_CTRL_IF "
        
    End Select

    LerTodosDominioTabela = fgQueryXMLLerTodos("DominioTabela", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "FiltroDominioTabela Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar detalhe e status de uma operção segregadas por grupo de usuario.
Public Function ObterDetalheMensagem(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                            Optional ByVal blnFluxoAutomatico As Boolean = False) As String

Dim datDataCorte                            As Date
Dim datInicio                               As Date
Dim datFim                                  As Date

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim strSQL                                  As String
Dim strSqlA8                                As String
Dim strSqlBaseHist                          As String


On Error GoTo ErrorHandler

    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
    datInicio = gstrDataVazia
    datFim = gstrDataVazia

    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
        datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
    End If

    If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
        datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
    End If
    
    If datInicio <> gstrDataVazia Then
        If datInicio < datDataCorte Then
            blnBaseHistorica = True
        End If
    End If
    
    If datFim <> gstrDataVazia Then
        If datFim > datDataCorte Then
            blnBaseA8 = True
        End If
    End If
    
    strSqlA8 = flObterQueryDetalheMensagem(xmlFiltro, blnFluxoAutomatico)
    
    If blnBaseHistorica Then
        strSqlBaseHist = Replace$(strSqlA8, "'A8' as OWNER", "'A8HIST' as OWNER")
        strSqlBaseHist = Replace$(strSqlBaseHist, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
        strSqlBaseHist = Replace$(strSqlBaseHist, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV")
        strSqlBaseHist = Replace$(strSqlBaseHist, "A8.TB_MESG_RECB_SPB_CNCL", "A8HIST.TB_MESG_RECB_SPB_CNCL ")
        If blnBaseA8 Then
            strSQL = strSqlA8 & vbNewLine & " UNION ALL " & vbNewLine & strSqlBaseHist
        Else
            strSQL = strSqlBaseHist
        End If
    Else
        strSQL = strSqlA8
    End If
    
    ObterDetalheMensagem = fgQueryXMLLerTodos("DetalheMensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "ObterDetalheMensagem", 0
End Function

'Monta o SQL para consultar detalhe e status de uma operção segregadas por grupo de usuario.
Private Function flObterQueryDetalheMensagem(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                    Optional ByVal blnFluxoAutomatico As Boolean = False) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String

Dim strWhereSegregacaoPrinc                 As String   'TB_MESG_RECB_ENVI_SPB  (Tipo BackOffice + Local Liquidação
'                         + Filtros selecionados)
Dim strWhereSegregacaoPrincAux              As String   'TB_MESG_RECB_ENVI_SPB  (Grupo Veículo Legal)
Dim strWhereSegregacaoApoio                 As String   'TB_GRUP_USUA_MESG_SPB  (Mensagens associadas ao usuário)
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim strValoresContingencia                  As String

Dim blnAdmGeralDefBMF                       As Boolean
Dim blnSegregaBackOffice                    As Boolean
Dim intTipoBackOffice                       As Integer
Dim blnNegacao                              As Boolean
Dim intTipoLiquidacao                       As Integer

    On Error GoTo ErrorHandler

    'Perfil de Acesso
    blnAdmGeralDefBMF = False
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_PerfilAcesso/*")
        blnAdmGeralDefBMF = True
    Next
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    
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
    
    'Filtro por Canal Operação Interna
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalOperacaoInternaMensagem/*")
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
            strWhereFiltro = strWhereFiltro & " AND      F.CO_CNAL_OPER_INTE  IN (" & strValores & ")"
        Else
            strValores = Replace$(strValores, "CONDIÇÃO", vbNullString)
            strWhereFiltro = strWhereFiltro & " AND " & strValores
        End If
    End If
    
    'Filtro por Número Sequencial da Operação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroSeqOperacao/*")
        If IsNumeric(objDomNode.Text) Then
            strValores = strValores & objDomNode.Text & ", "
        Else
            strValores = "NULL"
            Exit For
        End If
    Next
    
    If Trim(strValores) <> "" Then
        If strValores = "NULL" Then
            strWhereFiltro = strWhereFiltro & " AND     (A.NU_SEQU_OPER_ATIV  = 0 "
            strWhereFiltro = strWhereFiltro & " OR       A.NU_SEQU_OPER_ATIV  IS NULL)"
        Else
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_OPER_ATIV  IN (" & strValores & ")"
        End If
    End If
    
    'Filtro por Grupo Mensagem
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Mensagem/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      SUBSTR(A.CO_MESG_SPB, 1, 7)        IN (" & strValores & ")"
    End If
    
    'Filtro por Código Mensagem
    If Not blnAdmGeralDefBMF Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoMensagem/*")
            strValores = strValores & "'" & objDomNode.Text & "', "
        Next
        
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.CO_MESG_SPB        IN (" & strValores & ")"
        End If
    End If
    
    'Filtro excluindo Códigos de Mensagens
    If Not blnAdmGeralDefBMF Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoMensagemExcluir/*")
            strValores = strValores & "'" & objDomNode.Text & "', "
        Next
        
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.CO_MESG_SPB  NOT   IN (" & strValores & ")"
        End If
    End If
    
    'Filtro por MensagemSPB
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_MensagemSPB/*")
        strValores = strValores & "'%" & objDomNode.Text & "%', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & " AND A.CO_MESG_SPB LIKE (" & strValores & ")"
    End If
    
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    strValoresContingencia = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BMC0015Contingencia/*")
        strValoresContingencia = strValoresContingencia & objDomNode.Text & ", "
    Next
    
    If Trim(strValoresContingencia) <> "" Then
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strValoresContingencia = Left$(strValoresContingencia, Len(strValoresContingencia) - 2)
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND    ( A.CO_ULTI_SITU_PROC  IN (" & strValores & ")" & _
                            " OR       A.CO_MESG_SPB = 'BMC0015' " & _
                            " AND      A.CO_ULTI_SITU_PROC  IN (" & strValoresContingencia & ")) "
        Else
            strValoresContingencia = Left$(strValoresContingencia, Len(strValoresContingencia) - 2)
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.CO_MESG_SPB = 'BMC0015' " & _
                            " AND      A.CO_ULTI_SITU_PROC  IN (" & strValoresContingencia & ") "
        End If
    Else
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ") "
        End If
    End If
    
    'Filtro por Sequencial Conciliacao Operacao Mensagem
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SequencialConciliacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.NU_SEQU_CNCL_OPER_ATIV_MESG  IN (" & strValores & ")"
    End If
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_EMPR      IN (" & strValores & ")"
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      D.CO_GRUP_VEIC_LEGA        IN ('" & strValores & "')"
    End If
    
    'Filtro Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_VEIC_LEGA     IN (" & strValores & ")"
    End If
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_LOCA_LIQU     IN (" & strValores & ")"
    End If
    
    'KIDA - SGC
    'Filtro Grupo Canal Venda
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.TP_CNAL_VEND      IN (" & strValores & ")"
        'Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND  IN (" & strValores & ")")
    End If
            
    'Filtro Indicador de Entrada manual
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_EntradaManual/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.IN_ENTR_MANU     IN (" & strValores & ")"
    End If
    
    'Filtro Tipo Operação Rotina Abertura
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoRotinaAbertura/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      F.CO_OPER_SELIC      IN (" & strValores & ")"
    End If
    
    'Filtro por Data
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data").length > 0 Then
        If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                strWhereFiltro = strWhereFiltro & "AND A.DH_REGT_MESG_SPB between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                      objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DH_REGT_MESG_SPB) = " & xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text & vbNewLine
        End If
    End If
    
    'Filtro por Data de Operação
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataOperacao")
        If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataOperacao/DataFim").length > 0 Then
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      F.DT_OPER           BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                            " AND                                  " & objDomNode.selectSingleNode("DataFim").Text & vbNewLine
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND F.DT_OPER = " & xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_DataOperacao/DataIni").Text & vbNewLine
        End If
    Next objDomNode
    
    'Filtro por Indicador de Débito ou Crédito
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NaturezaMovimento/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      F.IN_OPER_DEBT_CRED      IN (" & strValores & ")"
    End If
    
    'Filtro por ISPB Banco Liquidante Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ISPBBancoLiquidanteContraparte/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      F.CO_ISPB_BANC_LIQU_CNPT      IN (" & strValores & ")"
    End If
    
    'Filtro por Número Sequencial de Controle de Repetição
    If Not blnAdmGeralDefBMF Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ControleRepeticao/*")
            strValores = objDomNode.Text
        Next
        
        If Trim(strValores) <> "" Then
            strWhereFiltro = strWhereFiltro & " AND      A.NU_SEQU_CNTR_REPE      " & strValores
        End If
    End If
    
    'Filtro Tipo Informação LDL - Prévia ou Definitiva ('P' ou 'D')
    If Not blnAdmGeralDefBMF Then
        strValores = vbNullString
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoLDL/*")
            strValores = strValores & "'" & objDomNode.Text & "', "
        Next
        
        If Trim(strValores) <> "" Then
            strValores = Left$(strValores, Len(strValores) - 2)
            strWhereFiltro = strWhereFiltro & " AND      F.TP_INFO_LDL      IN (" & strValores & ")"
        End If
    End If
    
    'Filtro Especial para Administrador Geral Definitiva BMF
    If blnAdmGeralDefBMF Then
        strWhereFiltro = strWhereFiltro & " AND    ((A.CO_MESG_SPB = 'LDL0001' AND A.NU_SEQU_CNTR_REPE > 1 AND F.TP_INFO_LDL = 'D') "
        strWhereFiltro = strWhereFiltro & " OR      (A.CO_MESG_SPB = 'LDL0004' AND A.NU_SEQU_CNTR_REPE = 1 AND F.TP_INFO_LDL = 'P')) "
    End If
    
    'Filtro Número de Comando especial para Despesa CETIP
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LayoutEntrada/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If strValores = "88" Then
            strWhereFiltro = strWhereFiltro & " AND      SUBSTR(A.NU_COMD_OPER, 1, 11) = 'CUSTO CETIP'" & vbCrLf
        End If
    End If
    
    'Filtro por Data Evento Câmbio
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataEventoCambio").length > 0 Then
        If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataEventoCambio/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataEventoCambio")
                strWhereFiltro = strWhereFiltro & "AND A.DT_OPER_CAMB_SISBACEN between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                      objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_OPER_CAMB_SISBACEN) = " & xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_DataEventoCambio/DataIni").Text & vbNewLine

        End If
    End If
    
    'RATS 1176 - Incluir o sistema no filtro de mensagens - 05/07/2012
    'Filtro por Sistema Origem
    strValores = vbNullString
    If Not xmlFiltro.selectSingleNode("//Sistema") Is Nothing Then
    
        strValores = strValores & "'" & xmlFiltro.selectSingleNode("//Sistema").Text & "'"
        If Trim(strValores) <> "" Then
            
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.SG_SIST     = " & strValores
        End If
    End If
    
    'Filtro por Identificador Pessoa
    strValores = vbNullString
    If Not xmlFiltro.selectSingleNode("//IdentificadorPessoa") Is Nothing Then
    
        strValores = strValores & "'" & xmlFiltro.selectSingleNode("//IdentificadorPessoa").Text & "'"
        If Trim(strValores) <> "" Then
            
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND      A.CD_CLIE_SISBACEN     = " & strValores
        End If
    End If
    'Filtro por Situação Mensagem
    strValores = vbNullString
    If Not xmlFiltro.selectSingleNode("//SituacaoMensagem") Is Nothing Then
        strValores = strValores & xmlFiltro.selectSingleNode("//SituacaoMensagem").Text
        If Trim(strValores) <> "" Then
            strWhereFiltro = strWhereFiltro & vbNewLine & _
                            " AND    A.CO_ULTI_SITU_PROC =" & strValores
        End If
    End If
    
    
    '---------------------------------------------------------------------------------------------------------
    
    If Not blnFluxoAutomatico Then
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                blnSegregaBackOffice = True
                If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SegregaBackOffice/*").length <> 0 Then
                    Set objDomNode = xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_SegregaBackOffice/*")
                    blnSegregaBackOffice = CBool(objDomNode.Text)
                End If
                
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, blnSegregaBackOffice, True, _
                                                         False, False, , True)
                
                strWhereSegregacaoPrincAux = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                            False, "A", "D", False, False, False, False, True, False, , True)
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, blnSegregaBackOffice, False, False, True)
            Case 0
                strWhereSegregacaoPrinc = vbNullString
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
            Case Is > 0
                strWhereSegregacaoPrinc = vbNewLine & " AND (A.TP_BKOF = " & intTipoBackOffice & " OR A.TP_BKOF IS NULL)" & vbNewLine
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
        End Select
    End If
    
    strSQL = " SELECT   'A8' as OWNER, " & vbNewLine & _
             "          A.NU_CTRL_IF,                                                       " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                                                 " & vbNewLine & _
             "          A.NU_SEQU_CNTR_REPE,                                                " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                                " & vbNewLine & _
             "          A.TP_BKOF,                                                          " & vbNewLine & _
             "          A.CO_EMPR,                                                          " & vbNewLine & _
             "          A.DH_RECB_ENVI_MESG_SPB,                                            " & vbNewLine & _
             "          A.CO_MESG_SPB,                                                      " & vbNewLine & _
             "          A.NU_COMD_OPER,                                                     " & vbNewLine & _
             "          A.CO_SITU_MESG_SPB,                                                 " & vbNewLine & _
             "          A.CO_TEXT_XML,                                                      " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                                " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                     " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ,                                                " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,                                           " & vbNewLine & _
             "          A.HO_ENVI_MESG_SPB,                                                 " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                     " & vbNewLine & _
             "          A.IN_ENTR_MANU,                                                     " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                     " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          D.NO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          A.SG_SIST,                                                          " & vbNewLine & _
             "          A.NU_CTRL_CAMR,                                                     " & vbNewLine & _
             "          F.IN_OPER_DEBT_CRED    AS CAMPO_IN_OPER_DEBT_CRED,                  " & vbNewLine
            
'ESTÁ IMPLEMENTAÇÃO ESTÁ EM STAND-BY, AGUARDANDO PRIORIZAÇÃO PARA SER IMPLANTADA
'campos incluídos por solicitação dos usuário do Comex, devido projeto Sisbacen
'    strSQL = strSQL & _
'             "          NVL(A.NO_CLIE, E.NO_CLIE) AS NO_CLIE,                               " & vbNewLine & _
'             "          NVL(A.NR_CNPJ_CPF, E.NR_CNPJ_CPF) AS NR_CNPJ_CPF,                   " & vbNewLine & _
'             "          NVL(A.CD_MOED_ISO, E.CD_MOED_ISO) AS CD_MOED_ISO,                   " & vbNewLine & _
'             "          NVL(A.VL_MOED_ESTR, E.VA_MOED_ESTR) AS VA_MOED_ESTR,                " & vbNewLine & _
'             "          NVL(A.NR_OPER_CAMB, E.NR_OPER_CAMB) AS NR_OPER_CAMB,                " & vbNewLine

    strSQL = strSQL & _
             "          DECODE(E.IN_OPER_DEBT_CRED,                                         " & vbNewLine & _
             "                  NULL,                                                       " & vbNewLine
             
    strSQL = strSQL & _
             "                      DECODE(F.IN_OPER_DEBT_CRED,                             " & vbNewLine & _
             "                          NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                                      enumTipoDebitoCredito.Credito & ", 'Crédito'), " & vbNewLine & _
             "                  " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                    enumTipoDebitoCredito.Credito & ", 'Crédito') " & _
                                              "IN_OPER_DEBT_CRED, " & vbNewLine
             
    strSQL = strSQL & _
             "          DECODE(E.IN_ENTR_SAID_RECU_FINC,                                    " & vbNewLine & _
             "                 " & enumTipoEntradaSaida.ENTRADA & ", 'Entrada', " & _
                                   enumTipoEntradaSaida.Saida & ", 'Saída')   " & _
                                   "IN_ENTR_SAID_RECU_FINC,                                 " & vbNewLine & _
             "          E.DT_OPER_ATIV,                                                     " & vbNewLine & _
             "          E.NO_CNPT,                                                          " & vbNewLine & _
             "          E.VA_OPER_ATIV,                                                     " & vbNewLine & _
             "          E.DT_LIQU_OPER_ATIV,                                                " & vbNewLine & _
             "          NVL(E.TP_OPER,L.TP_OPER) TP_OPER,                                   " & vbNewLine

    strSQL = strSQL & _
             "          F.PE_TAXA_NEGO,                                                     " & vbNewLine & _
             "          F.NU_COMD_OPER_ORIG,                                                " & vbNewLine & _
             "          F.NU_ATIV_MERC,                                                     " & vbNewLine & _
             "          F.DT_VENC,                                                          " & vbNewLine & _
             "          F.DT_OPER,                                                          " & vbNewLine & _
             "          F.DT_LIQU,                                                          " & vbNewLine & _
             "          F.DT_OPER_ATIV_RETN,                                                " & vbNewLine & _
             "          F.PU_ATIV_MERC,                                                     " & vbNewLine & _
             "          F.QT_ATIV_MERC,                                                     " & vbNewLine & _
             "          F.CO_CNPJ_CNPT,                                                     " & vbNewLine & _
             "          F.NO_CNPT                 AS NO_CNPT_CNCL,                          " & vbNewLine & _
             "          F.VA_FINC,                                                          " & vbNewLine & _
             "          F.CO_ISPB_BANC_LIQU_CNPT,                                           " & vbNewLine & _
             "          F.CO_AGEN_COTR,                                                     " & vbNewLine & _
             "          F.NU_CC_COTR,                                                       " & vbNewLine & _
             "          F.CO_SISB_COTR,                                                     " & vbNewLine & _
             "          F.CO_PRAC,                                                          " & vbNewLine & _
             "          F.CO_MOED_ESTR,                                                     " & vbNewLine & _
             "          F.VA_MOED_ESTR,                                                     " & vbNewLine & _
             "          F.DT_LIQU_OPER_ATIV_MOED_ESTR,                                      " & vbNewLine

    strSQL = strSQL & _
             "          G.SG_LOCA_LIQU,                                                     " & vbNewLine & _
             "          H.NO_TIPO_LIQU_OPER_ATIV,                                           " & vbNewLine & _
             "          F.TP_INFO_LDL,                                                      " & vbNewLine & _
             "          F.CO_OPER_SELIC,                                                    " & vbNewLine & _
             "          F.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                     " & vbNewLine & _
             "          A.NU_PRTC_MESG_LG,                                                  " & vbNewLine & _
             "          F.TP_PAGT_LDL   AS TP_PGTO_LDL,                                     " & vbNewLine & _
             "          A.CO_PARP_CAMR,                                                     " & vbNewLine & _
             "          I.DS_PARP_CAMR,                                                     " & vbNewLine & _
             "          J.DE_BKOF,                                                          " & vbNewLine & _
             "          A.TP_ACAO_MESG_SPB_EXEC,                                            " & vbNewLine & _
             "          D.ID_PART_CAMR_CETIP,                                               " & vbNewLine & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,                                      " & vbNewLine & _
             "          A.CO_ISPB_PART_CAMR,                                                " & vbNewLine & _
             "          E.TP_CNAL_VEND,                                                     " & vbNewLine & _
             "          F.CD_ASSO_CAMB,                                                     " & vbNewLine & _
             "          A.NR_OPER_CAMB_2,                                                   " & vbNewLine & _
             "          K.NU_SEQU_CNCL_OPER_ATIV_MESG                                       " & vbNewLine


    strSQL = strSQL & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB        A,                                  " & vbNewLine
             
    strSQL = strSQL & _
             "         (SELECT DISTINCT                                                     " & vbNewLine & _
             "                  CO_MESG_SPB                                                 " & vbNewLine & _
             "          FROM    A8.TB_GRUP_USUA_MESG_SPB                                    "

    'Incorpora a Segregação de Dados da Tabela de Apoio
    If Trim(strWhereSegregacaoApoio) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoApoio
    End If

    strSQL = strSQL & vbNewLine & _
             "         )                                B , " & vbNewLine
    
    strSQL = strSQL & _
             "          A8.TB_SITU_PROC                 C,                                  " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 D,                                  " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 E,                                  " & vbNewLine & _
             "          A8.TB_MESG_RECB_SPB_CNCL        F,                                  " & vbNewLine & _
             "          A8.TB_LOCAL_LIQUIDACAO          G,                                  " & vbNewLine & _
             "          A8.TB_TIPO_LIQU_OPER_ATIV       H,                                  " & vbNewLine & _
             "          A8.TB_TIPO_BKOF_PARP_CAMR       I,                                  " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 J,                                  " & vbNewLine
             
    'RATS 1233 - 04/10/2012
    strSQL = strSQL & _
             "          (SELECT MIN(CO.NU_SEQU_OPER_ATIV) NU_SEQU_OPER_ATIV, CO.NU_CTRL_IF, CO.DH_REGT_MESG_SPB, CO.NU_SEQU_CNTR_REPE, CO.NU_SEQU_CNCL_OPER_ATIV_MESG  " & vbNewLine & _
             "           FROM A8.TB_CNCL_OPER_ATIV CO " & vbNewLine & _
             "           GROUP BY CO.NU_CTRL_IF, CO.DH_REGT_MESG_SPB, CO.NU_SEQU_CNTR_REPE, CO.NU_SEQU_CNCL_OPER_ATIV_MESG) K, " & vbNewLine & _
             "          A8.TB_OPER_ATIV                 L " & vbNewLine

    
    'Para as mensagens sem repeticao o campo NU_SEQU_CNTR_REPE é sempre 1
    'Para as mensagens com repeticao , foi adotado que a mensagem mae o campo NU_SEQU_CNTR_REPE é sempre 1
    strSQL = strSQL & _
             " WHERE    SUBSTR(A.CO_MESG_SPB, 1, 7)     = SUBSTR(B.CO_MESG_SPB, 1, 7)       " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC             = C.CO_SITU_PROC                    " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                  = D.CO_VEIC_LEGA        (+)         " & vbNewLine & _
             " AND      A.SG_SIST                       = D.SG_SIST             (+)         " & vbNewLine & _
             " AND      A.NU_SEQU_OPER_ATIV             = E.NU_SEQU_OPER_ATIV   (+)         " & vbNewLine & _
             " AND      A.NU_CTRL_IF                    = F.NU_CTRL_IF          (+)         " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB              = F.DH_REGT_MESG_SPB    (+)         " & vbNewLine & _
             " AND      A.NU_SEQU_CNTR_REPE             = F.NU_SEQU_CNTR_REPE   (+)         " & vbNewLine & _
             " AND      A.CO_LOCA_LIQU                  = G.CO_LOCA_LIQU                    " & vbNewLine & _
             " AND      G.CO_EMPR_FUSI                  = 1                                 " & vbNewLine & _
             " AND      F.TP_LIQU_OPER_ATIV             = H.TP_LIQU_OPER_ATIV   (+)         " & vbNewLine & _
             " AND      A.CO_PARP_CAMR                  = I.CO_PARP_CAMR        (+)         " & vbNewLine & _
             " AND      A.TP_BKOF                       = J.TP_BKOF             (+)         " & vbNewLine & _
             " AND      A.NU_CTRL_IF                    = K.NU_CTRL_IF          (+)         " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB              = K.DH_REGT_MESG_SPB    (+)         " & vbNewLine & _
             " AND      A.NU_SEQU_CNTR_REPE             = K.NU_SEQU_CNTR_REPE   (+)         " & vbNewLine & _
             " AND      K.NU_SEQU_OPER_ATIV             = L.NU_SEQU_OPER_ATIV   (+)         " & vbNewLine
            
    If InStr(1, strWhereFiltro, "A.NU_SEQU_CNTR_REPE") = 0 Then
        strSQL = strSQL & _
                 " AND      A.NU_SEQU_CNTR_REPE             = 1                                 " & vbNewLine
    End If
    
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados da Tabela Principal
    If Trim(strWhereSegregacaoPrinc) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoPrinc
    End If
    
    'Incorpora a Segregação de Dados da Tabela Principal Auxiliar
    If Trim(strWhereSegregacaoPrincAux) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoPrincAux
    End If
    
    'Ordernação customizada pelo chamador
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_OrderBy/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strSQL = strSQL & vbNewLine & " ORDER BY " & strValores
    End If

    flObterQueryDetalheMensagem = strSQL
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterQueryDetalheMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todas as mensagens/operação agrupados por status
Public Function ObterMensagensPorStatus(ByVal xmlFiltro As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereSegregacaoPrinc                 As String   'TB_MESG_RECB_ENVI_SPB  (Tipo BackOffice + Local Liquidação
                                                        '                         + Filtros selecionados)
Dim strWhereSegregacaoPrincAux              As String   'TB_MESG_RECB_ENVI_SPB  (Grupo Veículo Legal)
Dim strWhereSegregacaoApoio                 As String   'TB_GRUP_USUA_MESG_SPB  (Mensagens associadas ao usuário)
Dim strValores                              As String

Dim strTabelaB                              As String
Dim strTabelaBA8                            As String
Dim strTabelaBHist                          As String

Dim blnBaseA8                               As Boolean
Dim blnBaseHist                             As Boolean

Dim datInicio                               As Date
Dim datFim                                  As Date
Dim datDataCorte                            As Date

Dim xmlComplementoWhere                     As MSXML2.DOMDocument40
Dim blnOcultarQuantidades                   As Boolean

Dim intTipoBackOffice                       As Integer

On Error GoTo ErrorHandler
    
    'Verifica se é necessário exibir as quantidades
    blnOcultarQuantidades = (Not _
        xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Quantidade/OcultarQuantidade") Is Nothing)
    
    If Not blnOcultarQuantidades Then
        If xmlFiltro.selectNodes("Repeat_Filtros/*").length > 0 Then
        
            datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
        
            If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
                  datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
            End If
        
            If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
                datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
            End If
        
            blnBaseHist = datInicio < datDataCorte
            blnBaseA8 = datFim >= datDataCorte
        
            'Envia dados para a segregação ----------------------------------------------------------------------------
            Set xmlComplementoWhere = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlComplementoWhere, "", "Repeat_Where", "")
            Call fgAppendNode(xmlComplementoWhere, "Repeat_Where", "Grupo_Where", "")
            
            'Filtro Grupo Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_GRUP_VEIC_LEGA      IN (" & strValores & ")")
                'Inclui o atributo ALIAS para filtro específico na tabela DOMÍNIO (Segregação de Dados)
                Call fgAppendAttribute(xmlComplementoWhere, "Where", "Alias", "B")
            End If
            
            'Filtro por Empresa
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If
            
            'Filtro Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_VEIC_LEGA      IN (" & strValores & ")")
            End If
            
            'RATS 1176 - Considerar Sistema - 05/07/2012
            'Filtro Sistema
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "SG_SIST  IN (" & strValores & ")")
            End If
            
            
            'KIDA - SGC
            'Filtro Grupo Canal Venda
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                'strWhereFiltro = strWhereFiltro & " AND      A.TP_CNAL_VEND      IN (" & strValores & ")"
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND  IN (" & strValores & ")")
            End If
            
            
            'Filtro Local de Liquidação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_LOCA_LIQU      IN (" & strValores & ")")
            End If
            
            'Filtro por Data
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", _
                                    "DH_REGT_MESG_SPB BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                                                    " AND " & objDomNode.selectSingleNode("DataFim").Text)
            Next objDomNode
            
            'Filtro por MensagemSPB
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_MensagemSPB/*")
                strValores = strValores & "'%" & objDomNode.Text & "%', "
            Next

            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_MESG_SPB LIKE (" & strValores & ")")
            End If
            
        End If
        
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, True, True, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrincAux = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                            False, "A", "B", False, False, False, False, True, False, , _
                                                            True)
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, True, False, False, True)
            Case 0
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, False, False, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
            Case Is > 0
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, False, False, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrinc = strWhereSegregacaoPrinc & vbNewLine & _
                                          " AND (A.TP_BKOF = " & intTipoBackOffice & " OR A.TP_BKOF IS NULL)" & vbNewLine
                
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
        End Select
    
        Set xmlComplementoWhere = Nothing

    End If
    
    If blnOcultarQuantidades Then
        strSQL = " SELECT   A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.TP_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.DE_SITU_PROC                                                         " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC     A                                                  " & vbNewLine & _
                 " WHERE    A.TP_SITU_PROC      = " & enumTipoSituacaoSistema.MensagemSPB & "      " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
    Else
        strSQL = " SELECT   A.DE_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.CO_SITU_PROC,                                                        " & vbNewLine & _
                 "          COUNT(B.CO_ULTI_SITU_PROC)          AS NU_QTD                          " & vbNewLine & _
                 " FROM     A8.TB_SITU_PROC                             A,                         " & vbNewLine

        strTabelaBA8 = _
                 "         (SELECT  A.CO_ULTI_SITU_PROC,NU_SEQU_CNTR_REPE                          " & vbNewLine & _
                 "          FROM    A8.TB_MESG_RECB_ENVI_SPB    A,                                 " & vbNewLine & _
                 "                  A8.TB_VEIC_LEGA             B,                                 " & vbNewLine & _
                 "         (SELECT DISTINCT                                                        " & vbNewLine & _
                 "                  CO_MESG_SPB                                                    " & vbNewLine & _
                 "          FROM    A8.TB_GRUP_USUA_MESG_SPB                                       "
                            
        'Incorpora a Segregação de Dados da Tabela de Apoio
        If Trim(strWhereSegregacaoApoio) <> "" Then
           strTabelaBA8 = strTabelaBA8 & vbNewLine & _
                    strWhereSegregacaoApoio
        End If
                            
        strTabelaBA8 = strTabelaBA8 & vbNewLine & _
                 "         )                                    C,                                 " & vbNewLine & _
                 "                  A8.TB_LOCAL_LIQUIDACAO      D                                  " & vbNewLine & _
                 " WHERE    A.CO_VEIC_LEGA              = B.CO_VEIC_LEGA    (+)                    " & vbNewLine & _
                 " AND      A.SG_SIST                   = B.SG_SIST         (+)                    " & vbNewLine & _
                 " AND      SUBSTR(A.CO_MESG_SPB, 1, 7) = SUBSTR(C.CO_MESG_SPB, 1, 7)              " & vbNewLine & _
                 " AND      A.NU_SEQU_CNTR_REPE         = 1                                        " & vbNewLine & _
                 " AND      A.CO_LOCA_LIQU              = D.CO_LOCA_LIQU                           " & vbNewLine & _
                 " AND      D.CO_EMPR_FUSI              = 1                                        " & vbNewLine
                 
        'Incorpora a Segregação de Dados da Tabela Principal
        If Trim(strWhereSegregacaoPrinc) <> "" Then
           strTabelaBA8 = strTabelaBA8 & vbNewLine & _
                    strWhereSegregacaoPrinc
        End If
        
        'Incorpora a Segregação de Dados da Tabela Principal Auxiliar
        If Trim(strWhereSegregacaoPrincAux) <> "" Then
           strTabelaBA8 = strTabelaBA8 & vbNewLine & _
                    strWhereSegregacaoPrincAux
        End If
        
        strTabelaBA8 = strTabelaBA8 & " ) "
        
        If blnBaseHist Then
            strTabelaBHist = Replace$(strTabelaBA8, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
            If blnBaseA8 Then
                strTabelaB = " ( " & vbNewLine & strTabelaBA8 & vbNewLine & " UNION ALL " & vbNewLine & strTabelaBHist & " ) "
            Else
                strTabelaB = strTabelaBHist
            End If
        Else
            strTabelaB = strTabelaBA8
        End If
                
        strSQL = strSQL & vbNewLine & _
                strTabelaB & "                                         B                          " & vbNewLine
        
        strSQL = strSQL & _
                 " WHERE    A.CO_SITU_PROC      = B.CO_ULTI_SITU_PROC    (+)                       " & vbNewLine & _
                 " AND      A.TP_SITU_PROC      = " & enumTipoSituacaoSistema.MensagemSPB & "      " & vbNewLine & _
                 " GROUP BY A.DE_SITU_PROC,                                                        " & vbNewLine & _
                 "          A.CO_SITU_PROC                                                         " & vbNewLine & _
                 " ORDER BY A.DE_SITU_PROC      ASC"
    End If
    
    ObterMensagensPorStatus = fgQueryXMLLerTodos("MensagensPorStatus", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensPorStatus Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todas as mensagens/operação agrupados por grupo de mensagem/operação
Public Function ObterMensagensPorGrupo(ByVal xmlFiltro As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim strWhereSegregacaoPrinc                 As String   'TB_MESG_RECB_ENVI_SPB  (Tipo BackOffice + Local Liquidação
                                                        '                         + Filtros selecionados)
Dim strWhereSegregacaoPrincAux              As String   'TB_MESG_RECB_ENVI_SPB  (Grupo Veículo Legal)
Dim strWhereSegregacaoApoio                 As String   'TB_GRUP_USUA_MESG_SPB  (Mensagens associadas ao usuário)
Dim strValores                              As String

Dim blnBaseA8                               As Boolean
Dim blnBaseHist                             As Boolean

Dim datInicio                               As Date
Dim datFim                                  As Date
Dim datDataCorte                            As Date

Dim strTabelaF                              As String
Dim strTabelaFA8                            As String
Dim strTabelaFHist                          As String

Dim xmlComplementoWhere                     As MSXML2.DOMDocument40
Dim xmlMensagens                            As MSXML2.DOMDocument40
Dim objXMLNode                              As MSXML2.IXMLDOMNode

Dim blnOcultarQuantidades                   As Boolean

Dim intTipoBackOffice                       As Integer
                                                                
On Error GoTo ErrorHandler
    
    'Verifica se é necessário exibir as quantidades
    blnOcultarQuantidades = (Not _
        xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Quantidade/OcultarQuantidade") Is Nothing)
        
    If Not blnOcultarQuantidades Then
        If xmlFiltro.selectNodes("Repeat_Filtros/*").length > 0 Then
        
            datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo
        
            If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
                  datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
            End If
        
            If Not xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
                datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
            End If
        
            blnBaseHist = datInicio < datDataCorte
            blnBaseA8 = datFim >= datDataCorte

        
            'Envia dados para a segregação ----------------------------------------------------------------------------
            Set xmlComplementoWhere = CreateObject("MSXML2.DOMDocument.4.0")
            Call fgAppendNode(xmlComplementoWhere, "", "Repeat_Where", "")
            Call fgAppendNode(xmlComplementoWhere, "Repeat_Where", "Grupo_Where", "")
            
            'Filtro Grupo Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_GRUP_VEIC_LEGA      IN (" & strValores & ")")
                'Inclui o atributo ALIAS para filtro específico na tabela DOMÍNIO (Segregação de Dados)
                Call fgAppendAttribute(xmlComplementoWhere, "Where", "Alias", "B")
            End If
            
            'Filtro por Empresa
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If
            
            'Filtro Veiculo Legal
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_VEIC_LEGA      IN (" & strValores & ")")
            End If
            
            'KIDA - SGC
            'Filtro Grupo Canal Venda
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalVenda/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "TP_CNAL_VEND  IN (" & strValores & ")")
            End If
            
            'RATS 1176 - Considerar Sistema - 05/07/2012
            'Filtro Sistema
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Sistema/*")
                strValores = strValores & "'" & objDomNode.Text & "', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "SG_SIST  IN (" & strValores & ")")
            End If
            
            'Filtro Local de Liquidação
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
                strValores = strValores & objDomNode.Text & ", "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_LOCA_LIQU      IN (" & strValores & ")")
            End If
            
            'Filtro por Data
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", _
                                    "DH_REGT_MESG_SPB BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                                                    " AND " & objDomNode.selectSingleNode("DataFim").Text)
            Next objDomNode
            
            'Filtro por MensagemSPB
            strValores = vbNullString
            For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_MensagemSPB/*")
                strValores = strValores & "'%" & objDomNode.Text & "%', "
            Next
            
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_EMPR      IN (" & strValores & ")")
            End If
            If Trim(strValores) <> "" Then
                strValores = Left$(strValores, Len(strValores) - 2)
                Call fgAppendNode(xmlComplementoWhere, "Grupo_Where", "Where", "CO_MESG_SPB LIKE (" & strValores & ")")
            End If
            
        End If
        
        intTipoBackOffice = 99
        For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
            intTipoBackOffice = Val(objDomNode.Text)
        Next
        
        Select Case intTipoBackOffice
            Case 99
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, True, True, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrincAux = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                            False, "A", "B", False, False, False, False, True, False, , _
                                                            True)
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, True, False, False, True)
            Case 0
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, False, False, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
            Case Is > 0
                strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                         False, "A", "A", False, False, False, False, False, False, _
                                                         xmlComplementoWhere, True)
                
                strWhereSegregacaoPrinc = strWhereSegregacaoPrinc & vbNewLine & _
                                          " AND (A.TP_BKOF = " & intTipoBackOffice & " OR A.TP_BKOF IS NULL)" & vbNewLine
                
                strWhereSegregacaoPrincAux = vbNullString
                
                strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                         False, "", "", True, False, False, False, False, True)
        End Select
    
        Set xmlComplementoWhere = Nothing
    
    End If
    
    If blnOcultarQuantidades Then
        strSQL = " SELECT   A.CO_GRUP,                                                              " & vbNewLine & _
                 "          A.NO_GRUP,                                                              " & vbNewLine & _
                 "          D.CO_MESG,                                                              " & vbNewLine & _
                 "          D.NO_MESG                                                               " & vbNewLine & _
                 " FROM     A8.TB_GRUPO                             A,                              " & vbNewLine & _
                 "          A8.TB_SERVICO                           B,                              " & vbNewLine & _
                 "          A8.TB_EVENTO                            C,                              " & vbNewLine & _
                 "          A8.TB_MENSAGEM                          D                               " & vbNewLine & _
                 " WHERE    A.SQ_GRUP           = B.SQ_GRUP                                         " & vbNewLine & _
                 " AND      B.SQ_SERV           = C.SQ_SERV                                         " & vbNewLine & _
                 " AND      C.SQ_EVEN           = D.SQ_EVEN                                         " & vbNewLine & _
                 " ORDER BY A.CO_GRUP,                                                              " & vbNewLine & _
                 "          D.CO_MESG           ASC"
    
        ObterMensagensPorGrupo = fgQueryXMLLerTodos("MensagensPorGrupo", strSQL, App.EXEName & "." & TypeName(Me))
    
    Else
        strTabelaF = vbNullString
        strTabelaFA8 = vbNullString
        strTabelaFHist = vbNullString
        
        strSQL = " SELECT   A.CO_GRUP,                                                              " & vbNewLine & _
                 "          A.NO_GRUP,                                                              " & vbNewLine & _
                 "          D.CO_MESG,                                                              " & vbNewLine & _
                 "          D.NO_MESG,                                                              " & vbNewLine & _
                 "          COUNT(F.CO_MESG_SPB)                    AS NU_QTD                       " & vbNewLine & _
                 " FROM     A8.TB_GRUPO                             A,                              " & vbNewLine & _
                 "          A8.TB_SERVICO                           B,                              " & vbNewLine & _
                 "          A8.TB_EVENTO                            C,                              " & vbNewLine & _
                 "          A8.TB_MENSAGEM                          D,                              " & vbNewLine
                 
        strSQL = strSQL & _
                 "         (SELECT DISTINCT                                                         " & vbNewLine & _
                 "                  CO_MESG_SPB                                                     " & vbNewLine & _
                 "          FROM    A8.TB_GRUP_USUA_MESG_SPB                                        "
                            
        'Incorpora a Segregação de Dados da Tabela de Apoio
        If Trim(strWhereSegregacaoApoio) <> "" Then
           strSQL = strSQL & vbNewLine & _
                    strWhereSegregacaoApoio
        End If
                            
        strSQL = strSQL & vbNewLine & _
                 "         )                                        E,                              " & vbNewLine
        
        '>>> 1º) Remove o sufixo da mensagem (R1 ou E) devido à tabela << A8.TB_GRUP_USUA_MESG_SPB >>
        '        estar alimentada com apenas mensagens com 7 posições.
        '    2º) A função RPAD foi utilizada devido aos campos do tipo << CO_MESG_SPB >> estarem definidos
        '        como CHAR(9), dessa forma o JOIN pôde ser implementado
        strTabelaFA8 = strTabelaFA8 & _
                 "         (SELECT  RPAD(SUBSTR(A.CO_MESG_SPB, 1, 7), 9) CO_MESG_SPB                " & vbNewLine & _
                 "          FROM    A8.TB_MESG_RECB_ENVI_SPB    A,                                  " & vbNewLine & _
                 "                  A8.TB_VEIC_LEGA             B,                                  " & vbNewLine
                 
        strTabelaFA8 = strTabelaFA8 & _
                 "         (SELECT DISTINCT                                                         " & vbNewLine & _
                 "                  CO_MESG_SPB                                                     " & vbNewLine & _
                 "          FROM    A8.TB_GRUP_USUA_MESG_SPB                                        "
                            
        'Incorpora a Segregação de Dados da Tabela de Apoio
        If Trim(strWhereSegregacaoApoio) <> "" Then
           strTabelaFA8 = strTabelaFA8 & vbNewLine & _
                    strWhereSegregacaoApoio
        End If
                            
        strTabelaFA8 = strTabelaFA8 & vbNewLine & _
                 "         )                                    C,                                  " & vbNewLine & _
                 "                  A8.TB_LOCAL_LIQUIDACAO      D                                   " & vbNewLine & _
                 " WHERE    A.CO_VEIC_LEGA              = B.CO_VEIC_LEGA    (+)                     " & vbNewLine & _
                 " AND      A.SG_SIST                   = B.SG_SIST         (+)                     " & vbNewLine & _
                 " AND      SUBSTR(A.CO_MESG_SPB, 1, 7) = SUBSTR(C.CO_MESG_SPB, 1, 7)               " & vbNewLine & _
                 " AND      A.NU_SEQU_CNTR_REPE         = 1                                         " & vbNewLine & _
                 " AND      A.CO_LOCA_LIQU              = D.CO_LOCA_LIQU                            " & vbNewLine & _
                 " AND      D.CO_EMPR_FUSI              = 1                                         " & vbNewLine

        'Incorpora a Segregação de Dados da Tabela Principal
        If Trim(strWhereSegregacaoPrinc) <> "" Then
           strTabelaFA8 = strTabelaFA8 & vbNewLine & _
                    strWhereSegregacaoPrinc
        End If
        
        'Incorpora a Segregação de Dados da Tabela Principal Auxiliar
        If Trim(strWhereSegregacaoPrincAux) <> "" Then
           strTabelaFA8 = strTabelaFA8 & vbNewLine & _
                    strWhereSegregacaoPrincAux
        End If
                
        strTabelaFA8 = strTabelaFA8 & " ) "
        
        If blnBaseHist Then
            strTabelaFHist = Replace$(strTabelaFA8, "A8.TB_MESG_RECB_ENVI_SPB", "A8HIST.TB_MESG_RECB_ENVI_SPB")
            If blnBaseA8 Then
                strTabelaF = " ( " & vbNewLine & strTabelaFA8 & " UNION ALL " & vbNewLine & strTabelaFHist & vbNewLine & " ) "
            Else
                strTabelaF = strTabelaFHist
            End If
        Else
            strTabelaF = strTabelaFA8
        End If
                
        strSQL = strSQL & vbNewLine & _
                strTabelaF & "                                               F                               " & vbNewLine
        
        strSQL = strSQL & _
                 " WHERE    A.SQ_GRUP           = B.SQ_GRUP                                         " & vbNewLine & _
                 " AND      B.SQ_SERV           = C.SQ_SERV                                         " & vbNewLine & _
                 " AND      C.SQ_EVEN           = D.SQ_EVEN                                         " & vbNewLine & _
                 " AND      D.CO_MESG           = E.CO_MESG_SPB                                     " & vbNewLine & _
                 " AND      E.CO_MESG_SPB       = F.CO_MESG_SPB         (+)                         "
    
        strSQL = strSQL & vbNewLine & _
                 " GROUP BY A.CO_GRUP,                                                              " & vbNewLine & _
                 "          A.NO_GRUP,                                                              " & vbNewLine & _
                 "          D.CO_MESG,                                                              " & vbNewLine & _
                 "          D.NO_MESG                                                               " & vbNewLine & _
                 " ORDER BY A.CO_GRUP,                                                              " & vbNewLine & _
                 "          D.CO_MESG           ASC"
    
        Set xmlMensagens = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlMensagens.loadXML(fgQueryXMLLerTodos("MensagensPorGrupo", strSQL, App.EXEName & "." & TypeName(Me)))
        ObterMensagensPorGrupo = xmlMensagens.xml
        
        Set xmlMensagens = Nothing
        
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensPorGrupo Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obtém detalhes de mensagens enviadas por alguma Câmara de Liquidação
Public Function ObterDetalheMensagemCamara(ByVal xmlFiltro As MSXML2.DOMDocument40, _
                                  Optional ByVal pblnAutomatico As Boolean = False) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomAttrs                             As MSXML2.IXMLDOMNamedNodeMap

Dim strSQL                                  As String

Dim strWhereSegregacaoPrinc                 As String   'TB_MESG_RECB_ENVI_SPB  (Tipo BackOffice + Local Liquidação
                                                        '                         + Filtros selecionados)
Dim strWhereSegregacaoPrincAux              As String   'TB_MESG_RECB_ENVI_SPB  (Grupo Veículo Legal)
Dim strWhereSegregacaoApoio                 As String   'TB_GRUP_USUA_MESG_SPB  (Mensagens associadas ao usuário)
Dim strWhereFiltro                          As String
Dim strValores                              As String
Dim lngContador                             As Long

Dim blnSegregaBackOffice                    As Boolean
Dim blnSegregaGrupoVeiculoLegal             As Boolean
Dim strComplemento                          As String
                                                                
On Error GoTo ErrorHandler
    
    'Captura Filtro ------------------------------------------------------------------------------------------
    
    'Filtro por Chave Primaria da tabela
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Chave").length <> 0 Then
        Set objDomNode = xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Chave").Item(0)
        Set objDomAttrs = objDomNode.attributes

        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.NU_CTRL_IF         = " & IIf(objDomAttrs.getNamedItem("NU_CTRL_IF").Text <> vbNullString, _
                                                            "'" & objDomAttrs.getNamedItem("NU_CTRL_IF").Text & "'", _
                                                            "A.NU_CTRL_IF") & vbNewLine & _
                        " AND      A.DH_REGT_MESG_SPB   = " & IIf(objDomAttrs.getNamedItem("DH_REGT_MESG_SPB").Text <> vbNullString, _
                                                             "" & objDomAttrs.getNamedItem("DH_REGT_MESG_SPB").Text & "", _
                                                            "A.DH_REGT_MESG_SPB") & vbNewLine & _
                        " AND      A.NU_SEQU_CNTR_REPE  = " & IIf(objDomAttrs.getNamedItem("NU_SEQU_CNTR_REPE").Text <> vbNullString, _
                                                             "" & objDomAttrs.getNamedItem("NU_SEQU_CNTR_REPE").Text & "", _
                                                            "A.NU_SEQU_CNTR_REPE")
    End If
    
'    'Verificar e Menssagem esta associado ao grupo de usuario
'    flMenssagemGrupoUsuario strWhereFiltro

    'Filtro pelo trio STATUS/STATUSMAE/NU_SEQU_CNTR_REPE
    lngContador = 0
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_StatusContador/*")
        lngContador = lngContador + 1
        If lngContador = 1 Then
            strWhereFiltro = strWhereFiltro & vbNewLine & "AND ( "
        Else
            strWhereFiltro = strWhereFiltro & vbNewLine & " OR "
        End If
        Set objDomAttrs = objDomNode.attributes
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " (A.CO_ULTI_SITU_PROC = " & IIf(objDomAttrs.getNamedItem("Status").Text = "", _
                                                    "A.CO_ULTI_SITU_PROC", _
                                                    objDomAttrs.getNamedItem("Status").Text) & "" & _
                     " AND G.CO_ULTI_SITU_PROC = " & IIf(objDomAttrs.getNamedItem("StatusMae").Text = "", _
                                                    "G.CO_ULTI_SITU_PROC ", _
                                                    objDomAttrs.getNamedItem("StatusMae").Text) & "" & _
                     " AND A.NU_SEQU_CNTR_REPE  " & IIf(objDomAttrs.getNamedItem("NU_SEQU_CNTR_REPE").Text = "", _
                                                    " = A.NU_SEQU_CNTR_REPE ", _
                                                    objDomAttrs.getNamedItem("NU_SEQU_CNTR_REPE").Text) & " )"
    Next
    If lngContador > 0 Then
        strWhereFiltro = strWhereFiltro & ")"
    End If
    
    'Filtro por Grupo Mensagem
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Mensagem/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_MESG_SPB                  IN (" & strValores & ")"
    End If
    
    'Filtro por Status
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Status/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_ULTI_SITU_PROC  IN (" & strValores & ")"
    End If
    
    'Filtro por CNPJ
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJ/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      D.CO_CNPJ_VEIC_LEGA  IN (" & strValores & ")"
    End If
    
    'Filtro por Numero de controle de repetição
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SequenciaControleRepeticao/*")
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.NU_SEQU_CNTR_REPE "
        
        'Monta o OPERADOR sql de acordo com o Nome da Tag que for passada
        strWhereFiltro = strWhereFiltro & fgDECODE(LCase(objDomNode.baseName), _
                                                        "igual", "=", _
                                                        "menor", "<", _
                                                        "maior", ">", _
                                                        "menorouigual", "<=", _
                                                        "maiorouigual", ">=", _
                                                        "diferente", "<>" _
                                                        )
        
        strWhereFiltro = strWhereFiltro & objDomNode.Text
    Next
    
    'Filtro por Classificacao do Veiculo Legal na BMA
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoTitularBMA/*")
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      D.TP_TITL_BMA "
                        
        'Monta o OPERADOR sql de acordo com o Nome da Tag que for passada
        strWhereFiltro = strWhereFiltro & fgDECODE(LCase(objDomNode.baseName), _
                                                        "igual", "=", _
                                                        "menor", "<", _
                                                        "maior", ">", _
                                                        "menorouigual", "<=", _
                                                        "maiorouigual", ">=", _
                                                        "diferente", "<>" _
                                                        )
        
        strWhereFiltro = strWhereFiltro & "'" & objDomNode.Text & "'"
    Next
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_EMPR      IN (" & strValores & ")"
    End If
    
    'Filtro Grupo Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoVeiculoLegal/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      D.CO_GRUP_VEIC_LEGA        IN ('" & strValores & "')"
    End If
    
    'Filtro Veiculo Legal
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_VeiculoLegal/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_VEIC_LEGA     IN (" & strValores & ")"
    End If
    
    'Filtro Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.CO_LOCA_LIQU     IN (" & strValores & ")"
    End If
    
    'Filtro Indicador de Entrada manual
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_EntradaManual/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.IN_ENTR_MANU     IN (" & strValores & ")"
    End If
    
    'Filtro Numero de Comando
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroComando/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.NU_COMD_OPER     IN (" & strValores & ")"
    End If
    
    'Filtro Número do Contrato SISBACEN
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ContratacaoSISBACEN/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_CNTR_SISB     IN (" & strValores & ")" & vbCrLf
    End If

    'Filtro ChACAM
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ChACAM/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CD_ASSO_CAMB     IN ('" & strValores & "')" & vbCrLf
    End If

    'Filtro Numero Identificacao Negociacao BMC
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentNegocioBMC/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next

    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      E.NR_IDEF_NEGO_BMC     IN (" & strValores & ")"
    End If

    'Filtro Canal SISBACEN Corretora
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CanalSISBACENCorretora/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_SISB_COTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Debito/Credito
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DebitoCredito/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      E.IN_OPER_DEBT_CRED     IN (" & strValores & ")"
    End If
    
    'Filtro Código Praça
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoPraca/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_PRAC     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Código Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CodigoMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Moeda Estrangeira
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorMoedaEstrangeira/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.VA_MOED_ESTR     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Taxa Negociação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TaxaNegociacao/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.PE_TAXA_NEGO     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Valor Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ValorFinanceiro/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.VA_FINC     IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data Liquidação Moeda Estrangeira
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").length <> 0 Then
        Set objDomNode = xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataLiquidacaoMoedaEstr").Item(0)

        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      E.DT_LIQU_OPER_ATIV_MOED_ESTR      = " & objDomNode.Text & ""
    End If

    'Filtro Tipo Informacao Mensagem (Ex: previa, definitiva)
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoInformacao/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      E.TP_INFO_LDL     IN (" & strValores & ")"
    End If
    
    'Filtro Tipo Operação Rotina Abertura
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_TipoOperacaoRotinaAbertura/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_OPER_SELIC      IN (" & strValores & ")"
    End If
    
    'Filtro ISPB Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_ISPBContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_ISPB_BANC_LIQU_CNPT      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Número Controle LTR
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_NumeroControleLTR/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.NU_CTRL_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Identificador Participante Câmara
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamara/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_PARP_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro Identificador Participante Câmara Contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_IdentificadorPartCamaraContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_CNPT_CAMR      IN (" & strValores & ")"
    End If
    
    'Filtro CNPJ contraparte
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_CNPJContraparte/*")
        strValores = strValores & "'" & objDomNode.Text & "', "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      E.CO_CNPJ_CNPT      IN (" & strValores & ")"
    End If
    
    'Filtro por Data da Mensagem
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_Data")
        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      A.DH_REGT_MESG_SPB BETWEEN " & objDomNode.selectSingleNode("DataIni").Text & _
                        " AND " & objDomNode.selectSingleNode("DataFim").Text
    Next objDomNode
    
    'Filtro por Data da Operacao
    If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataOperacao").length <> 0 Then
        Set objDomNode = xmlFiltro.selectNodes("Repeat_Filtros/Grupo_DataOperacao").Item(0)

        strWhereFiltro = strWhereFiltro & vbNewLine & _
                        " AND      E.DT_OPER      = " & objDomNode.Text & ""
    End If

    'Filtro Número de Comando especial para Despesa CETIP
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LayoutEntrada/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        If strValores = "88" Then
            strWhereFiltro = strWhereFiltro & " AND      SUBSTR(A.NU_COMD_OPER, 1, 11) = 'CUSTO CETIP'" & vbCrLf
        End If
    End If

    If pblnAutomatico = False Then
        '---------------------------------------------------------------------------------------------------------
        'Segregação de dados por usuário
        
        blnSegregaBackOffice = True
        If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SegregaBackOffice/*").length <> 0 Then
            Set objDomNode = xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_SegregaBackOffice/*")
            blnSegregaBackOffice = CBool(objDomNode.Text)
        End If
        
        blnSegregaGrupoVeiculoLegal = True
        If xmlFiltro.selectNodes("Repeat_Filtros/Grupo_SegregaGrupoVeiculoLegal/*").length <> 0 Then
            Set objDomNode = xmlFiltro.selectSingleNode("Repeat_Filtros/Grupo_SegregaGrupoVeiculoLegal/*")
            blnSegregaGrupoVeiculoLegal = CBool(objDomNode.Text)
        End If
        
        If pblnAutomatico = False Then
            '1ª TABELA - TB_MESG_RECB_ENVI_SPB
            '            (Tipo BackOffice, Local Liquidação e Filtros Selecionados) ------------------------------------
            strWhereSegregacaoPrinc = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                  False, "A", "A", False, False, blnSegregaBackOffice, True, False, False, , True)
            
            '...(cont.)
            '           (Grupo Veículo Legal) -------------------------------------------------------------------------
            strWhereSegregacaoPrincAux = fgSegregaDados("A8.TB_MESG_RECB_ENVI_SPB", _
                                                 False, "A", "D", False, False, False, False, blnSegregaGrupoVeiculoLegal, False, , True)
            
            '2ª TABELA - TB_GRUP_USUA_MESG_SPB
            '           (Tipo BackOffice, Grupo Usuário) --------------------------------------------------------------
            strWhereSegregacaoApoio = fgSegregaDados("A8.TB_GRUP_USUA_MESG_SPB", _
                                                  False, "", "", True, False, blnSegregaBackOffice, False, False, True)
        End If
    End If

    strSQL = " SELECT   A.NU_CTRL_IF,                                                       " & vbNewLine & _
             "          A.DH_REGT_MESG_SPB,                                                 " & vbNewLine & _
             "          A.NU_SEQU_CNTR_REPE,                                                " & vbNewLine & _
             "          A.NU_SEQU_OPER_ATIV,                                                " & vbNewLine & _
             "          A.TP_BKOF,                                                          " & vbNewLine & _
             "          F.DE_BKOF,                                                          " & vbNewLine & _
             "          A.CO_EMPR,                                                          " & vbNewLine & _
             "          A.DH_RECB_ENVI_MESG_SPB,                                            " & vbNewLine & _
             "          A.CO_MESG_SPB,                                                      " & vbNewLine & _
             "          A.NU_COMD_OPER,                                                     " & vbNewLine & _
             "          A.CO_SITU_MESG_SPB,                                                 " & vbNewLine & _
             "          A.CO_TEXT_XML,                                                      " & vbNewLine & _
             "          A.CO_ULTI_SITU_PROC,                                                " & vbNewLine & _
             "          C.DE_SITU_PROC,                                                     " & vbNewLine & _
             "          A.CO_USUA_ULTI_ATLZ,                                                " & vbNewLine & _
             "          A.CO_ETCA_TRAB_ULTI_ATLZ,                                           " & vbNewLine & _
             "          A.HO_ENVI_MESG_SPB,                                                 " & vbNewLine & _
             "          A.DH_ULTI_ATLZ,                                                     " & vbNewLine & _
             "          A.IN_ENTR_MANU,                                                     " & vbNewLine & _
             "          A.CO_LOCA_LIQU,                                                     " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          D.NO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          D.CO_CNPJ_VEIC_LEGA,                                                " & vbNewLine & _
             "          E.CO_CNTA_CUTD_SELIC_VEIC_LEGA,                                     " & vbNewLine & _
             "          A.SG_SIST,                                                          " & vbNewLine
            
    strSQL = strSQL & _
             "          DECODE(E.IN_OPER_DEBT_CRED,                                         " & vbNewLine & _
             "                  NULL, '', " & enumTipoDebitoCredito.Debito & ", 'Débito', " & _
                                              enumTipoDebitoCredito.Credito & ", 'Crédito') " & _
                                              "IN_OPER_DEBT_CRED, " & vbNewLine & _
             "          E.IN_OPER_DEBT_CRED                  AS CAMPO_IN_OPER_DEBT_CRED,    " & vbNewLine & _
             "          E.PE_TAXA_NEGO,                                                     " & vbNewLine & _
             "          E.DT_OPER,                                                          " & vbNewLine & _
             "          E.DT_VENC,                                                          " & vbNewLine & _
             "          E.NO_CNPT,                                                          " & vbNewLine & _
             "          E.CO_CNPJ_CNPT,                                                     " & vbNewLine & _
             "          E.DT_LIQU,                                                          " & vbNewLine & _
             "          E.NU_COMD_OPER_ORIG,                                                " & vbNewLine & _
             "          E.NU_ATIV_MERC,                                                     " & vbNewLine & _
             "          E.PU_ATIV_MERC,                                                     " & vbNewLine & _
             "          E.QT_ATIV_MERC,                                                     " & vbNewLine & _
             "          E.VA_FINC,                                                          " & vbNewLine & _
             "          E.TP_INFO_LDL,                                                      " & vbNewLine & _
             "          E.TP_NEGO_BMA,                                                      " & vbNewLine & _
             "          E.CO_ISPB_BANC_LIQU_CNPT,                                           " & vbNewLine & _
             "          D.TP_TITL_BMA,                                                      " & vbNewLine & _
             "          E.CO_TITL_CUTD,                                                     " & vbNewLine & _
             "          G.CO_ULTI_SITU_PROC            AS  CO_ULTI_SITU_PROC_MAE,           " & vbNewLine & _
             "          G.NU_SEQU_CNCL_OPER_ATIV_MESG  AS  NU_SEQU_CNCL_OPER_ATIV_MESG_MA,  " & vbNewLine & _
             "          A.IN_CONF_MESG_LTR,                                                 " & vbNewLine
    
    strSQL = strSQL & _
             "          A.NU_SEQU_CNCL_OPER_ATIV_MESG,                                      " & vbNewLine & _
             "          E.CO_CNTA_CUTD_CNPT,                                                " & vbNewLine & _
             "          E.CO_OPER_CETIP,                                                    " & vbNewLine & _
             "          E.DE_ATIV_MERC,                                                     " & vbNewLine & _
             "          E.TP_LIQU_OPER_ATIV,                                                " & vbNewLine & _
             "          H.NO_TIPO_LIQU_OPER_ATIV,                                           " & vbNewLine & _
             "          E.DT_OPER_ATIV_RETN,                                                " & vbNewLine & _
             "          E.PZ_DIAS_RETN_OPER_ATIV,                                           " & vbNewLine & _
             "          E.VA_OPER_ATIV_RETN,                                                " & vbNewLine & _
             "          A.NU_CTRL_CAMR,                                                     " & vbNewLine & _
             "          D.ID_PART_CAMR_CETIP,                                               " & vbNewLine & _
             "          E.CO_CNPT_CAMR,                                                     " & vbNewLine & _
             "          E.TP_PAGT_LDL,                                                      " & vbNewLine & _
             "          E.CO_MOED_ESTR,                                                     " & vbNewLine & _
             "          E.CO_CNTR_SISB,                                                     " & vbNewLine & _
             "          E.CO_ISPB_IF_CNPT,                                                  " & vbNewLine & _
             "          E.CO_PRAC,                                                          " & vbNewLine & _
             "          E.VA_MOED_ESTR,                                                     " & vbNewLine & _
             "          E.DT_LIQU_OPER_ATIV_MOED_ESTR,                                      " & vbNewLine & _
             "          E.CO_SISB_COTR,                                                     " & vbNewLine & _
             "          A.CO_PARP_CAMR,                                                     " & vbNewLine & _
             "          E.NR_IDEF_NEGO_BMC,                                                 " & vbNewLine & _
             "          E.CD_ASSO_CAMB                                                      " & vbNewLine

    strSQL = strSQL & _
             " FROM     A8.TB_MESG_RECB_ENVI_SPB        A,                                  " & vbNewLine
             
    strSQL = strSQL & _
             "         (SELECT DISTINCT                                                     " & vbNewLine & _
             "                  CO_MESG_SPB                                                 " & vbNewLine & _
             "          FROM    A8.TB_GRUP_USUA_MESG_SPB                                    "

    'Incorpora a Segregação de Dados da Tabela de Apoio
    If Trim(strWhereSegregacaoApoio) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoApoio
    End If

    strSQL = strSQL & vbNewLine & _
             "         )                                B , " & vbNewLine
    
    strSQL = strSQL & _
             "          A8.TB_SITU_PROC                 C,                                  " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 D,                                  " & vbNewLine & _
             "          A8.TB_MESG_RECB_SPB_CNCL        E,                                  " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 F,                                  " & vbNewLine & _
             "          A8.TB_MESG_RECB_ENVI_SPB        G  /* p/ pegar dados da msg mae  */," & vbNewLine & _
             "          A8.TB_TIPO_LIQU_OPER_ATIV       H " & vbNewLine

    strSQL = strSQL & _
             " WHERE    SUBSTR(A.CO_MESG_SPB, 1, 7)     = SUBSTR(B.CO_MESG_SPB, 1, 7)       " & vbNewLine & _
             " AND      A.CO_ULTI_SITU_PROC             = C.CO_SITU_PROC                    " & vbNewLine & _
             " AND      A.CO_VEIC_LEGA                  = D.CO_VEIC_LEGA        (+)         " & vbNewLine & _
             " AND      A.SG_SIST                       = D.SG_SIST             (+)         " & vbNewLine & _
             " AND      A.NU_CTRL_IF                    = E.NU_CTRL_IF          (+)         " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB              = E.DH_REGT_MESG_SPB    (+)         " & vbNewLine & _
             " AND      A.NU_SEQU_CNTR_REPE             = E.NU_SEQU_CNTR_REPE   (+)         " & vbNewLine & _
             " AND      A.NU_CTRL_IF                    = G.NU_CTRL_IF          (+)         " & vbNewLine & _
             " AND      A.DH_REGT_MESG_SPB              = G.DH_REGT_MESG_SPB    (+)         " & vbNewLine & _
             " AND      1 /*msg mãe é sempre = 1*/      = G.NU_SEQU_CNTR_REPE   (+)         " & vbNewLine & _
             " AND      A.TP_BKOF                       = F.TP_BKOF             (+)         " & vbNewLine & _
             " AND      E.TP_LIQU_OPER_ATIV             = H.TP_LIQU_OPER_ATIV   (+)         "
            
    'Incorpora o Filtro
    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & vbNewLine & _
                 strWhereFiltro
    End If
    
    'Incorpora a Segregação de Dados da Tabela Principal
    If Trim(strWhereSegregacaoPrinc) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoPrinc
    End If
    
    'Incorpora a Segregação de Dados da Tabela Principal Auxiliar
    If Trim(strWhereSegregacaoPrincAux) <> "" Then
       strSQL = strSQL & vbNewLine & _
                strWhereSegregacaoPrincAux
    End If
    
    'Ordernação customizada pelo chamador
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_OrderBy/*")
        strValores = strValores & "" & objDomNode.Text & ", "
    Next
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strSQL = strSQL & vbNewLine & " ORDER BY " & strValores
    End If

    ObterDetalheMensagemCamara = fgQueryXMLLerTodos("DetalheMensagemCamara", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheMensagemCamara Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificar e Menssagem esta associado ao grupo de usuario
Public Function flMenssagemGrupoUsuario(ByVal pstrWhereFiltro As String) As Boolean

Dim strSQL                                  As String
Dim rsMensagem                              As ADODB.Recordset

On Error GoTo ErrorHandler

    Set rsMensagem = CreateObject("ADODB.Recordset")
 
    strSQL = " SELECT     A.CO_MESG_SPB" & _
             " FROM      A8.TB_MESG_RECB_ENVI_SPB    A," & _
             "           A8.TB_GRUP_USUA_MESG_SPB    B " & _
             " Where     SUBSTR(a.CO_MESG_SPB, 1, 7) = SUBSTR(B.CO_MESG_SPB,1, 7) "

    strSQL = strSQL & pstrWhereFiltro

    Set rsMensagem = fgQuerySQL(strSQL)

    If rsMensagem.EOF Then
        '3130 - Mensagem SPB não associado ao grupo de usuário.
        lngCodigoErroNegocio = 3130
        GoTo ErrorHandler
    End If

    Set rsMensagem = Nothing

    Exit Function
ErrorHandler:

    Set rsMensagem = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMenssagemGrupoUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
