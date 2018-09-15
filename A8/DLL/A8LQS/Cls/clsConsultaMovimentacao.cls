VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConsultaMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
''-------------------------------------------------
' Gerado por Adrian Pretel
' Atualização em:      20-jun-2005
''-------------------------------------------------
''
' Recuperar os registros nas tabelas de movimentação
''
Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

'Recuperar registros das tabelas de movimentação de Operações

Public Function Ler(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim objXMLNode                              As MSXML2.IXMLDOMNode
Dim xmlLeitura                              As MSXML2.DOMDocument40

Dim strSQL                                  As String
Dim blnPrimeiro                             As Boolean
Dim blnBaseHistorica                        As Boolean
Dim blnBaseCorrente                         As Boolean
Dim strWhereSegregacao                      As String
Dim intTipoBackOffice                       As Integer
Dim strNomeCampo                            As String
                                                                
Dim objRS                                   As ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    intTipoBackOffice = 99
    For Each objXMLNode In pxmlFiltros.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(objXMLNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "A", "F", False, False, True, True, True, False)
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND F.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseHistorica = True
    End If
    
    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseCorrente = True
    End If
             
    strSQL = " SELECT C.CO_LOCA_LIQU                            , " & vbCrLf & _
             "        C.DE_LOCA_LIQU                            , " & vbCrLf & _
             "        D.NO_EMPR                                 , " & vbCrLf & _
             "        F.TP_BKOF                                 , " & vbCrLf & _
             "        H.DE_BKOF                                 , " & vbCrLf & _
             "        A.DT_OPER_ATIV                            , " & vbCrLf & _
             "        E.NO_SIST                                 , " & vbCrLf & _
             "        B.NO_TIPO_OPER                            , " & vbCrLf & _
             "        A.CO_ULTI_SITU_PROC                       , " & vbCrLf & _
             "        G.DE_SITU_PROC                            , " & vbCrLf & _
             "        A.CO_SITU_PROC_MESG_SPB_RECB              , " & vbCrLf & _
             "        ''       AS DE_SITU_PROC_MESG_SPB_RECB    , " & vbCrLf & _
             "        COUNT(*) AS QUANTIDADE                      " & vbCrLf
    
    'TABELAS
    strSQL = strSQL & _
             " FROM   A8.TB_OPER_ATIV                A, " & vbCrLf & _
             "        A8.TB_TIPO_OPER                B, " & vbCrLf & _
             "        A8.TB_LOCAL_LIQUIDACAO         C, " & vbCrLf & _
             "        A8.TB_EMPRESA_HO               D, " & vbCrLf & _
             "        A7.TB_SIST                     E, " & vbCrLf & _
             "        A8.TB_VEIC_LEGA                F, " & vbCrLf & _
             "        A8.TB_SITU_PROC                G, " & vbCrLf & _
             "        A8.TB_TIPO_BKOF                H  " & vbCrLf
    
    'CLÁUSULAS
    strSQL = strSQL & _
             " WHERE  A.TP_OPER             = B.TP_OPER                       " & vbCrLf & _
             " AND    A.CO_LOCA_LIQU        = C.CO_LOCA_LIQU                  " & vbCrLf & _
             " AND    A.CO_EMPR             = D.CO_EMPR                       " & vbCrLf & _
             " AND    A.SG_SIST             = E.SG_SIST                       " & vbCrLf & _
             " AND    A.CO_EMPR             = E.CO_EMPR                       " & vbCrLf & _
             " AND    C.CO_EMPR_FUSI        = 2                               " & vbCrLf & _
             " AND    A.CO_VEIC_LEGA        = F.CO_VEIC_LEGA                  " & vbCrLf & _
             " AND    A.SG_SIST             = F.SG_SIST                       " & vbCrLf & _
             " AND    A.CO_ULTI_SITU_PROC   = G.CO_SITU_PROC                  " & vbCrLf & _
             " AND    G.TP_SITU_PROC        = 1                               " & vbCrLf & _
             " AND    F.TP_BKOF             = H.TP_BKOF                       " & vbCrLf

    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then
        
        'Empresa
        blnPrimeiro = True
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.CO_EMPR In (" & objXMLNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", " & objXMLNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next objXMLNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If
        
        'Sigla do Sistema
        blnPrimeiro = True
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.SG_SIST In ('" & objXMLNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & objXMLNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next objXMLNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If
        
        'Código do Grupo do Veículo Legal
        blnPrimeiro = True
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_GrupoVeiculoLegal")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND F.CO_GRUP_VEIC_LEGA In ('" & objXMLNode.selectSingleNode("GrupoVeiculoLegal").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & objXMLNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next objXMLNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If
        
        'Código do Veículo Legal
        blnPrimeiro = True
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_VeiculoLegal")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.CO_VEIC_LEGA In ('" & objXMLNode.selectSingleNode("VeiculoLegal").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & objXMLNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next objXMLNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If

        'Local de Liquidação
        blnPrimeiro = True
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_LocalLiquidacao")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.CO_LOCA_LIQU In ('" & objXMLNode.selectSingleNode("LocalLiquidacao").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & objXMLNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next objXMLNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If

        'Filtro Data
        For Each objXMLNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
            strSQL = strSQL & " AND A.DT_OPER_ATIV between " & objXMLNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                              objXMLNode.selectSingleNode("DataFim").Text
        Next objXMLNode
        
    End If
    
    If strWhereSegregacao <> vbNullString Then
        strSQL = strSQL & strWhereSegregacao
    End If
    
    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSQL = strSQL & " GROUP BY C.CO_LOCA_LIQU, C.DE_LOCA_LIQU, D.NO_EMPR, F.TP_BKOF, H.DE_BKOF, A.DT_OPER_ATIV, E.NO_SIST, B.NO_TIPO_OPER, A.CO_ULTI_SITU_PROC, G.DE_SITU_PROC, A.CO_SITU_PROC_MESG_SPB_RECB"
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(Replace(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV"), "A8.TB_HIST_SITU_ACAO_OPER_ATIV", "A8HIST.TB_HIST_SITU_ACAO_OPER_ATIV")
        End If
    Else
        strSQL = Replace(Replace(strSQL, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV"), "A8.TB_HIST_SITU_ACAO_OPER_ATIV", "A8HIST.TB_HIST_SITU_ACAO_OPER_ATIV")
    End If
        
    'AGRUPAMENTO
    If InStr(1, strSQL, "GROUP BY") = 0 Then
        strSQL = strSQL & " GROUP BY C.CO_LOCA_LIQU, C.DE_LOCA_LIQU, D.NO_EMPR, F.TP_BKOF, H.DE_BKOF, A.DT_OPER_ATIV, E.NO_SIST, B.NO_TIPO_OPER, A.CO_ULTI_SITU_PROC, G.DE_SITU_PROC, A.CO_SITU_PROC_MESG_SPB_RECB"
    End If
    
    If blnBaseHistorica Then
        strSQL = strSQL & " ORDER BY DT_OPER_ATIV, DE_LOCA_LIQU, NO_EMPR, DE_SITU_PROC"
    Else
        strSQL = strSQL & " ORDER BY A.DT_OPER_ATIV, C.DE_LOCA_LIQU, D.NO_EMPR, G.DE_SITU_PROC"
    End If
    
    Ler = fgQueryXMLLerTodos("ConsultaMovimentacao", strSQL, App.EXEName & "." & TypeName(Me))

    If Ler <> vbNullString Then
        Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call xmlLeitura.loadXML(Ler)
        For Each objXMLNode In xmlLeitura.selectNodes("Repeat_ConsultaMovimentacao/*")
            If Val(objXMLNode.selectSingleNode("CO_ULTI_SITU_PROC").Text) = enumStatusOperacao.Rejeitada Then
                If UBound(Split(objXMLNode.selectSingleNode("CO_SITU_PROC_MESG_SPB_RECB").Text, "|")) > 0 Then
                    
                    strNomeCampo = "DE_DOMI"
                    strSQL = " SELECT DE_DOMI        " & vbNewLine & _
                             " FROM   A8.TB_DOMINIO  " & vbNewLine & _
                             " WHERE  CO_DOMI     = '" & Split(objXMLNode.selectSingleNode("CO_SITU_PROC_MESG_SPB_RECB").Text, "|")(1) & "'" & vbNewLine & _
                             " AND    SQ_TIPO_TAG =  " & Val(Split(objXMLNode.selectSingleNode("CO_SITU_PROC_MESG_SPB_RECB").Text, "|")(0))
                
                Else
                
                    strNomeCampo = "DE_ERRO_BACEN"
                    strSQL = " SELECT DE_ERRO_BACEN    " & vbNewLine & _
                             " FROM   A8.TB_ERRO_BACEN " & vbNewLine & _
                             " WHERE  CO_ERRO_BACEN = '" & objXMLNode.selectSingleNode("CO_SITU_PROC_MESG_SPB_RECB").Text & "'"
                
                End If
            
                Set objRS = fgQuerySQL(strSQL)
                
                If Not objRS.EOF Then
                    objXMLNode.selectSingleNode("DE_SITU_PROC_MESG_SPB_RECB").Text = objRS.fields(strNomeCampo)
                Else
                    objXMLNode.selectSingleNode("DE_SITU_PROC_MESG_SPB_RECB").Text = objXMLNode.selectSingleNode("CO_SITU_PROC_MESG_SPB_RECB").Text
                End If
                
                objRS.Close
            End If
        Next
        
        Ler = xmlLeitura.xml
        
        Set xmlLeitura = Nothing
    End If
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar registros das tabelas de movimentação de Operações Rejeitadas no SLCC

Public Function LerRejeitada(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim xmlDomNode        As MSXML2.IXMLDOMNode
Dim strSQL            As String
Dim strSQLA6          As String
Dim strSqlA8          As String
Dim blnPrimeiro       As Boolean
Dim blnBaseHistorica  As Boolean
Dim blnBaseCorrente   As Boolean

On Error GoTo ErrorHandler

    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseHistorica = True
    End If
    
    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseCorrente = True
    End If
    
    'Mensagem de Rejeição do Módulo de Liquidação - A8
    strSqlA8 = " SELECT 'Rejeição A8' AS MODULO            , " & vbCrLf & _
    "                 TRUNC(A.DH_REME_REJE) AS DTREJEICAO, " & vbCrLf & _
    "                 A.TP_MESG_INTE AS TPMENSAGEM       , " & vbCrLf & _
    "                 B.NO_TIPO_MESG AS MENSAGEM         , " & vbCrLf & _
    "                 A.SG_SIST_ORIG_INFO AS TPSISTEMA   , " & vbCrLf & _
    "                 D.NO_SIST AS SISTEMA               , " & vbCrLf & _
    "                 A.CO_EMPR AS TPEMPRESA             , " & vbCrLf & _
    "                 C.NO_EMPR AS EMPRESA               , " & vbCrLf & _
    "                 A.TX_XML_ERRO AS MOTIVO            , " & vbCrLf & _
    "                 COUNT(*) AS QUANTIDADE               "
    
    strSqlA8 = strSqlA8 & _
    "          FROM   A8.TB_REME_REJE  A, " & vbCrLf & _
    "                 A7.TB_TIPO_MESG  B, " & vbCrLf & _
    "                 A8.TB_EMPRESA_HO C, " & vbCrLf & _
    "                 A7.TB_SIST       D  "
    
    strSqlA8 = strSqlA8 & _
    "          WHERE  A.TP_MESG_INTE      = B.TP_MESG " & vbCrLf & _
    "          AND    A.CO_EMPR           = C.CO_EMPR " & vbCrLf & _
    "          AND    A.SG_SIST_ORIG_INFO = D.SG_SIST " & vbCrLf & _
    "          AND    A.CO_EMPR           = D.CO_EMPR "
    
    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then
        
        'Empresa
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSqlA8 = strSqlA8 & "AND A.CO_EMPR In (" & xmlDomNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSqlA8 = strSqlA8 & ", " & xmlDomNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSqlA8 = strSqlA8 & ")"
        End If
        
        'Sigla do Sistema
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSqlA8 = strSqlA8 & "AND A.SG_SIST_ORIG_INFO In ('" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSqlA8 = strSqlA8 & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSqlA8 = strSqlA8 & ")"
        End If
        
        'Filtro Data
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
        
            strSqlA8 = strSqlA8 & " AND TRUNC(A.DH_REME_REJE) between " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                              xmlDomNode.selectSingleNode("DataFim").Text
        Next xmlDomNode
        
    End If
    
    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSqlA8 = strSqlA8 & _
            "           GROUP BY TRUNC(A.DH_REME_REJE), A.TP_MESG_INTE, B.NO_TIPO_MESG, A.SG_SIST_ORIG_INFO, D.NO_SIST, A.CO_EMPR, C.NO_EMPR, A.TX_XML_ERRO "
            strSqlA8 = strSqlA8 & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSqlA8, "A8.TB_REME_REJE", "A8HIST.TB_REME_REJE")
        End If
    Else
        strSqlA8 = Replace(strSqlA8, "A8.TB_REME_REJE", "A8HIST.TB_REME_REJE")
    End If
        
    If InStr(1, strSqlA8, "GROUP BY") = 0 Then
        strSqlA8 = strSqlA8 & _
        "           GROUP BY TRUNC(A.DH_REME_REJE), A.TP_MESG_INTE, B.NO_TIPO_MESG, A.SG_SIST_ORIG_INFO, D.NO_SIST, A.CO_EMPR, C.NO_EMPR, A.TX_XML_ERRO "
    End If

    'Mensagem de Rejeição do Módulo de Sub-Reserva - A6
    strSQLA6 = strSQLA6 & _
    "          SELECT 'Rejeição A6' AS MODULO            , " & vbCrLf & _
    "                 TRUNC(A.DH_REME_REJE) AS DTREJEICAO, " & vbCrLf & _
    "                 A.TP_MESG_INTE AS TPMENSAGEM       , " & vbCrLf & _
    "                 B.NO_TIPO_MESG AS MENSAGEM         , " & vbCrLf & _
    "                 A.SG_SIST_ORIG_INFO AS TPSISTEMA   , " & vbCrLf & _
    "                 D.NO_SIST AS SISTEMA               , " & vbCrLf & _
    "                 A.CO_EMPR AS TPEMPRESA             , " & vbCrLf & _
    "                 C.NO_EMPR AS EMPRESA               , " & vbCrLf & _
    "                 A.TX_XML_ERRO AS MOTIVO            , " & vbCrLf & _
    "                 COUNT(*) AS QUANTIDADE               "
    
    strSQLA6 = strSQLA6 & _
    "          FROM   A6.TB_REME_REJE  A, " & vbCrLf & _
    "                 A7.TB_TIPO_MESG  B, " & vbCrLf & _
    "                 A8.TB_EMPRESA_HO C, " & vbCrLf & _
    "                 A7.TB_SIST       D  "
    
    strSQLA6 = strSQLA6 & _
    "          WHERE  A.TP_MESG_INTE      = B.TP_MESG " & vbCrLf & _
    "          AND    A.CO_EMPR           = C.CO_EMPR " & vbCrLf & _
    "          AND    A.SG_SIST_ORIG_INFO = D.SG_SIST " & vbCrLf & _
    "          AND    A.CO_EMPR           = D.CO_EMPR "
    
    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then
        
        'Empresa
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSQLA6 = strSQLA6 & "AND A.CO_EMPR In (" & xmlDomNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSQLA6 = strSQLA6 & ", " & xmlDomNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLA6 = strSQLA6 & ")"
        End If
        
        'Sigla do Sistema
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSQLA6 = strSQLA6 & "AND A.SG_SIST_ORIG_INFO In ('" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSQLA6 = strSQLA6 & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLA6 = strSQLA6 & ")"
        End If
        
        'Filtro Data
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
        
            strSQLA6 = strSQLA6 & " AND A.DH_REME_REJE between " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                              xmlDomNode.selectSingleNode("DataFim").Text
        Next xmlDomNode
        
    End If
    
    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSQLA6 = strSQLA6 & _
            "           GROUP BY TRUNC(A.DH_REME_REJE), A.TP_MESG_INTE, B.NO_TIPO_MESG, A.SG_SIST_ORIG_INFO, D.NO_SIST, A.CO_EMPR, C.NO_EMPR, A.TX_XML_ERRO "
            strSQLA6 = strSQLA6 & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQLA6, "A6.TB_REME_REJE", "A6HIST.TB_REME_REJE")
        End If
    Else
        strSQLA6 = Replace(strSQLA6, "A6.TB_REME_REJE", "A6HIST.TB_REME_REJE")
    End If

    If InStr(1, strSQLA6, "GROUP BY") = 0 Then
        strSQLA6 = strSQLA6 & _
        "           GROUP BY TRUNC(A.DH_REME_REJE), A.TP_MESG_INTE, B.NO_TIPO_MESG, A.SG_SIST_ORIG_INFO, D.NO_SIST, A.CO_EMPR, C.NO_EMPR, A.TX_XML_ERRO "
    End If

    strSQL = strSqlA8 & " UNION ALL " & strSQLA6
    
    strSQL = strSQL & _
    "           ORDER BY MODULO, DTREJEICAO, TPMENSAGEM, TPSISTEMA, TPEMPRESA "

     LerRejeitada = fgQueryXMLLerTodos("Erro", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar registros das tabelas de Integrações de Conta Corrente e Contabilidade

Public Function LerCC_HA(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim xmlDomNode        As MSXML2.IXMLDOMNode
Dim strSQL            As String
Dim strSQLWhere       As String
Dim strSQLGroup       As String
Dim strSQLOrder       As String
Dim strSQLFinal       As String
Dim strSQLErro        As String
Dim strSQLDisp        As String
Dim strSQLUnion       As String
Dim strSQLPrincipal   As String
Dim blnPrimeiro       As Boolean
Dim blnBaseHistorica  As Boolean
Dim blnBaseCorrente   As Boolean

Dim strWhereSegregacao                      As String
Dim intTipoBackOffice                       As Integer
                                                                
    On Error GoTo ErrorHandler
    
    intTipoBackOffice = 99
    For Each xmlDomNode In pxmlFiltros.selectNodes("Repeat_Filtros/Grupo_BackOfficePerfilGeral/*")
        intTipoBackOffice = Val(xmlDomNode.Text)
    Next
    
    Select Case intTipoBackOffice
        Case 99
            strWhereSegregacao = fgSegregaDados("A8.TB_OPER_ATIV", False, "C", "H", False, False, True, True, True, False)
        Case 0
            strWhereSegregacao = vbNullString
        Case Is > 0
            strWhereSegregacao = vbNewLine & " AND H.TP_BKOF = " & intTipoBackOffice & vbNewLine
    End Select

    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseHistorica = True
    End If
    
    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseCorrente = True
    End If

    'CAMPOS
    strSQL = " SELECT B.DE_SITU_PROC AS STATUS                , " & vbCrLf & _
    "                 TRUNC(A.DH_SITU_LANC_CC) AS DTINTEGRACAO, " & vbCrLf & _
    "                 D.NO_TIPO_OPER AS DESCTPOPERACAO        , " & vbCrLf & _
    "                 C.CO_LOCA_LIQU AS LOCALLIQUIDACAO       , " & vbCrLf & _
    "                 F.DE_LOCA_LIQU AS DESCLOCALIQUIDACAO    , " & vbCrLf & _
    "                 C.CO_EMPR AS TPEMPRESA                  , " & vbCrLf & _
    "                 E.NO_EMPR AS EMPRESA                    , " & vbCrLf & _
    "                 C.SG_SIST AS TPSISTEMA                  , " & vbCrLf & _
    "                 G.NO_SIST AS SISTEMA                    , " & vbCrLf & _
    "                 A.CO_ERRO AS TPERRO                     , " & vbCrLf & _
    "                 A.TX_MESG_ERRO AS ERRO                  , " & vbCrLf & _
    "                 A.TX_JUST_CANC AS JUSTIFICATIVA         , " & vbCrLf & _
    "                 H.TP_BKOF                               , " & vbCrLf & _
    "                 I.DE_BKOF                               , " & vbCrLf & _
    "                 COUNT(*) AS QUANTIDADE                    "
    'TABELAS
    strSQL = strSQL & _
    "          FROM   A8.TB_HIST_SITU_LANC_CC_CNTB A, " & vbCrLf & _
    "                 A8.TB_SITU_PROC              B, " & vbCrLf & _
    "                 A8.TB_OPER_ATIV              C, " & vbCrLf & _
    "                 A8.TB_TIPO_OPER              D, " & vbCrLf & _
    "                 A8.TB_EMPRESA_HO             E, " & vbCrLf & _
    "                 A8.TB_LOCAL_LIQUIDACAO       F, " & vbCrLf & _
    "                 A7.TB_SIST                   G, " & vbCrLf & _
    "                 A8.TB_VEIC_LEGA              H, " & vbCrLf & _
    "                 A8.TB_TIPO_BKOF              I  "
    'CLÁUSULAS
    strSQL = strSQL & _
    "          WHERE  A.CO_SITU_PROC      = B.CO_SITU_PROC      " & vbCrLf & _
    "          AND    A.NU_SEQU_OPER_ATIV = C.NU_SEQU_OPER_ATIV " & vbCrLf & _
    "          AND    C.TP_OPER           = D.TP_OPER           " & vbCrLf & _
    "          AND    C.CO_EMPR           = E.CO_EMPR           " & vbCrLf & _
    "          AND    C.CO_LOCA_LIQU      = F.CO_LOCA_LIQU      " & vbCrLf & _
    "          AND    C.SG_SIST           = G.SG_SIST           " & vbCrLf & _
    "          AND    C.CO_EMPR           = G.CO_EMPR           " & vbCrLf & _
    "          AND    C.CO_VEIC_LEGA      = H.CO_VEIC_LEGA      " & vbCrLf & _
    "          AND    B.TP_SITU_PROC      = 3                   " & vbCrLf & _
    "          AND    F.CO_EMPR_FUSI      = 2                   " & vbCrLf & _
    "          AND    H.TP_BKOF           = I.TP_BKOF           "

    'CLÁUSULA OPERAÇÕES FINALIZADAS COM SUCESSO
    strSQLFinal = " AND A.CO_SITU_PROC = 107 "

    'CLÁUSULA OPERAÇÕES FINALIZADAS COM ERRO
    strSQLErro = " AND A.CO_SITU_PROC = 108                                              " & vbCrLf & _
    "              AND A.NU_SEQU_OPER_ATIV NOT IN (SELECT W.NU_SEQU_OPER_ATIV            " & vbCrLf & _
    "                                              FROM   A8.TB_HIST_SITU_LANC_CC_CNTB W " & vbCrLf & _
    "                                              WHERE  W.CO_SITU_PROC = 107)          "
    
    'CLÁUSULA OPERAÇÕES NÃO FINALIZADAS
    strSQLDisp = " AND A.CO_SITU_PROC NOT IN (107,108)                                   " & vbCrLf & _
    "              AND A.NU_SEQU_OPER_ATIV NOT IN (SELECT W.NU_SEQU_OPER_ATIV            " & vbCrLf & _
    "                                              FROM   A8.TB_HIST_SITU_LANC_CC_CNTB W " & vbCrLf & _
    "                                              WHERE  W.CO_SITU_PROC IN (107,109))   "

    'CLÁUSULA DE UNION ENTRE OS SELECTS
    strSQLUnion = " UNION ALL "

    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then

        'Empresa
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSQLWhere = " AND C.CO_EMPR In (" & xmlDomNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSQLWhere = strSQLWhere & ", " & xmlDomNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLWhere = strSQLWhere & ")"
        End If
        
        'Sigla do Sistema
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSQLWhere = strSQLWhere & "AND C.SG_SIST In ('" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSQLWhere = strSQLWhere & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLWhere = strSQLWhere & ")"
        End If
        
        'Código do Grupo do Veículo Legal
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_GrupoVeiculoLegal")
        
            If blnPrimeiro Then
                strSQLWhere = strSQLWhere & "AND H.CO_GRUP_VEIC_LEGA In ('" & xmlDomNode.selectSingleNode("GrupoVeiculoLegal").Text & "'"
                blnPrimeiro = False
            Else
                strSQLWhere = strSQLWhere & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLWhere = strSQLWhere & ")"
        End If
        
        'Código do Veículo Legal
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_VeiculoLegal")
        
            If blnPrimeiro Then
                strSQLWhere = strSQLWhere & "AND C.CO_VEIC_LEGA In ('" & xmlDomNode.selectSingleNode("VeiculoLegal").Text & "'"
                blnPrimeiro = False
            Else
                strSQLWhere = strSQLWhere & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLWhere = strSQLWhere & ")"
        End If

        'Local de Liquidação
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_LocalLiquidacao")
        
            If blnPrimeiro Then
                strSQLWhere = strSQLWhere & "AND C.CO_LOCA_LIQU In ('" & xmlDomNode.selectSingleNode("LocalLiquidacao").Text & "'"
                blnPrimeiro = False
            Else
                strSQLWhere = strSQLWhere & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQLWhere = strSQLWhere & ")"
        End If
        
        'Filtro Data
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
        
            strSQLWhere = strSQLWhere & " AND TRUNC(A.DH_SITU_LANC_CC) BETWEEN " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                              xmlDomNode.selectSingleNode("DataFim").Text
        Next xmlDomNode
        
        If strWhereSegregacao <> vbNullString Then
            strSQLWhere = strSQLWhere & strWhereSegregacao
        End If
        
    End If
    
    'AGRUPAMENTO
    strSQLGroup = " GROUP BY B.DE_SITU_PROC, TRUNC(A.DH_SITU_LANC_CC), D.NO_TIPO_OPER, C.CO_LOCA_LIQU, F.DE_LOCA_LIQU, C.CO_EMPR, E.NO_EMPR,C.SG_SIST , G.NO_SIST, A.CO_ERRO, A.TX_MESG_ERRO, A.TX_JUST_CANC, H.TP_BKOF, I.DE_BKOF "
    strSQLOrder = " ORDER BY DTINTEGRACAO, DESCTPOPERACAO, LOCALLIQUIDACAO, TPEMPRESA, TPSISTEMA "

    'MONTAR O SELECT DE INTEGRAÇOES

    strSQLPrincipal = strSQL + strSQLFinal + strSQLWhere + strSQLGroup + strSQLUnion _
                    + strSQL + strSQLErro + strSQLWhere + strSQLGroup + strSQLUnion _
                    + strSQL + strSQLDisp + strSQLWhere + strSQLGroup

    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSQLPrincipal = strSQLPrincipal & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(Replace(strSQLPrincipal, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV"), "A8.TB_HIST_SITU_LANC_CC_CNTB", "A8HIST.TB_HIST_SITU_LANC_CC_CNTB")
        End If
    Else
        strSQLPrincipal = Replace(Replace(strSQLPrincipal, "A8.TB_OPER_ATIV", "A8HIST.TB_OPER_ATIV"), "A8.TB_HIST_SITU_LANC_CC_CNTB", "A8HIST.TB_HIST_SITU_LANC_CC_CNTB")
    End If
    
    strSQLPrincipal = strSQLPrincipal + strSQLOrder
    LerCC_HA = fgQueryXMLLerTodos("Erro", strSQLPrincipal, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerCC_HA Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


