VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F19838D02DA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Controle de Remessas para o A6.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
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

' Este método é responsável pela leitura à tabela de controle de remessas enviadas ao A6,
' e pelo retorno à camada intermediária.

Public Function LerTodos(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim blnPrimeiro                             As Boolean
Dim blnCriterio                             As Boolean

On Error GoTo ErrorHandler

    strSQL = "Select     Distinct                                 " & vbNewLine & _
             "           B.CO_EMPR, A.DT_REME, A.TP_MESG,         " & vbNewLine & _
             "           A.SG_SIST, A.QT_REGT_INFO_SIST_ORIG,     " & vbNewLine & _
             "           A.QT_REGT_REJE, A.QT_REGT_RECB,          " & vbNewLine & _
             "           A.DH_INIC_REME_ORIG, A.DH_FIM_REME_ORIG, " & vbNewLine & _
             "           A.DH_INIC_PROC_REME, A.DH_FIM_PROC_REME, " & vbNewLine & _
             "           A.CO_REME_PRMR, A.CO_REME_ULTI,          " & vbNewLine & _
             "           A.VA_TOTL_REME, A.CO_SITU_REME,          " & vbNewLine & _
             "           A.DH_ULTI_ATLZ, B.NO_EMPR,               " & vbNewLine & _
             "           C.NO_SIST, E.NO_TIPO_MESG,               " & vbNewLine & _
             "           A.CO_VEIC_LEGA, D.NO_VEIC_LEGA           " & vbNewLine & _
             "From       A6.TB_CTRL_PROC_REME       A,            " & vbNewLine & _
             "           A8.TB_EMPRESA_HO              B,            " & vbNewLine & _
             "           A7.TB_SIST                 C,            " & vbNewLine & _
             "           A8.TB_VEIC_LEGA            D,            " & vbNewLine & _
             "           A7.TB_TIPO_MESG            E             " & vbNewLine & _
             "Where      A.CO_VEIC_LEGA = D.CO_VEIC_LEGA          " & vbNewLine & _
             "AND        A.SG_SIST      = D.SG_SIST               " & vbNewLine & _
             "AND        D.CO_EMPR      = C.CO_EMPR               " & vbNewLine & _
             "AND        D.SG_SIST      = C.SG_SIST               " & vbNewLine & _
             "AND        A.TP_MESG      = E.TP_MESG               " & vbNewLine & _
             "AND        D.CO_EMPR      = B.CO_EMPR               "

    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then
    
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSQL = strSQL & vbNewLine & _
                    " AND D.CO_EMPR In (" & xmlDomNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", " & xmlDomNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ") "
        End If
        
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSQL = strSQL & vbNewLine & _
                    " AND A.SG_SIST In ('" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ") "
        End If
        
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
        
            strSQL = strSQL & vbNewLine & _
                "AND A.DT_REME between " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                             xmlDomNode.selectSingleNode("DataFim").Text
        
        Next xmlDomNode
        
    End If
    
    strSQL = strSQL & _
             fgSegregaDados("A8.TB_VEIC_LEGA", _
                            False, _
                            "D", _
                            "D", _
                            False, _
                            False, _
                            True, _
                            False, _
                            True, _
                            False)
    
    'Verificação da Base histórica
    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQL, "A6.TB_CTRL_PROC_REME", "A6HIST.TB_CTRL_PROC_REME")
        End If
    Else
        strSQL = Replace(strSQL, "A6.TB_CTRL_PROC_REME", "A6HIST.TB_CTRL_PROC_REME")
    End If

    
    strSQL = strSQL & vbNewLine & _
        " Order By NO_SIST, DT_REME, NO_EMPR"

    LerTodos = fgQueryXMLLerTodos("Sistema", strSQL, App.EXEName & "." & TypeName(Me))

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

' Este método é responsável pela atualização ao cadastro de remessas enviadas ao A6,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlRemessa As DOMDocument40) As Boolean
Dim objMovimentoFuturo                      As A6SubReserva.clsMovimentoFuturo

On Error GoTo ErrorHandler

    If Val(pxmlRemessa.documentElement.selectSingleNode("DT_FIM_REME_ORIG").Text) = 0 Then
        Set objMovimentoFuturo = CreateObject("A6SubReserva.clsMovimentoFuturo")
        objMovimentoFuturo.Excluir pxmlRemessa
        Set objMovimentoFuturo = Nothing

        flExcluir pxmlRemessa
    End If

    If flExisteControleRemessa(pxmlRemessa) Then
        flAlterar pxmlRemessa
    Else
        flIncluir pxmlRemessa
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objMovimentoFuturo = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela verificação da existência de um determinado registro de controle de remessa.

Private Function flExisteControleRemessa(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean

Dim xmlExisteControleRemessa                As MSXML2.DOMDocument40
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & vbCrLf & _
             "   COUNT (CO_VEIC_LEGA) AS CO_VEIC_LEGA " & vbCrLf & _
             " FROM " & vbCrLf & _
             "   A6.TB_CTRL_PROC_REME " & vbCrLf & _
             " WHERE CO_VEIC_LEGA = '" & pxmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text & "'" & vbCrLf & _
             "   AND SG_SIST      = '" & pxmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "'" & vbCrLf & _
             "   AND DT_REME      =  " & fgDtXML_To_Oracle(pxmlRemessa.documentElement.selectSingleNode("DT_FECH_PROC").Text) & vbCrLf & _
             "   AND TP_MESG      =  " & CLng(pxmlRemessa.documentElement.selectSingleNode("TP_MESG").Text)

    Set xmlExisteControleRemessa = CreateObject("MSXML2.DOMDocument.4.0")

    xmlExisteControleRemessa.loadXML fgQueryXMLLer("ExisteControleRemessa", strSQL, App.EXEName & "." & TypeName(Me))
    If xmlExisteControleRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text = "1" Then
        flExisteControleRemessa = True
    ElseIf xmlExisteControleRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text = "0" Then
        flExisteControleRemessa = False
    End If

    Set xmlExisteControleRemessa = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    flExisteControleRemessa = False
    Set xmlExisteControleRemessa = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão de um novo registro de controle de remessa.

Private Function flIncluir(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = "INSERT INTO " & vbCrLf & _
             "   A6.TB_CTRL_PROC_REME (  " & vbCrLf & _
             "   DT_REME,                " & vbCrLf & _
             "   TP_MESG,                " & vbCrLf & _
             "   SG_SIST,                " & vbCrLf & _
             "   CO_VEIC_LEGA,           " & vbCrLf & _
             "   QT_REGT_INFO_SIST_ORIG, " & vbCrLf & _
             "   QT_REGT_REJE,           " & vbCrLf & _
             "   QT_REGT_RECB,           " & vbCrLf & _
             "   DH_INIC_REME_ORIG,      " & vbCrLf & _
             "   DH_FIM_REME_ORIG,       " & vbCrLf & _
             "   CO_REME_PRMR,           " & vbCrLf & _
             "   CO_REME_ULTI,           " & vbCrLf & _
             "   DH_INIC_PROC_REME,      " & vbCrLf & _
             "   DH_FIM_PROC_REME,       " & vbCrLf & _
             "   VA_TOTL_REME,           " & vbCrLf & _
             "   CO_SITU_REME,           " & vbCrLf & _
             "   DH_ULTI_ATLZ            " & vbCrLf & _
             ") VALUES (                 " & vbCrLf & _
             "   " & fgDtXML_To_Oracle(pxmlRemessa.documentElement.selectSingleNode("DT_FECH_PROC").Text) & ", " & vbCrLf & _
             "   " & CLng(pxmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) & ", " & vbCrLf & _
             "  '" & pxmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "', " & vbCrLf & _
             "  '" & pxmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text & "', " & vbCrLf & _
             "   " & CLng(pxmlRemessa.documentElement.selectSingleNode("QT_REGT_INFO_SIST_ORIG").Text) & ", " & vbCrLf

    With pxmlRemessa.documentElement
        If Not .selectSingleNode("QT_REGT_REJE") Is Nothing Then
            strSQL = strSQL & _
                "   " & CLng(.selectSingleNode("QT_REGT_REJE").Text) & ", " & vbCrLf
        Else
            strSQL = strSQL & _
                "   0, " & vbCrLf
        End If
    
        If Not .selectSingleNode("QT_REGT_RECB") Is Nothing Then
            strSQL = strSQL & _
                "   " & CLng(.selectSingleNode("QT_REGT_RECB").Text) & ", " & vbCrLf
        Else
            strSQL = strSQL & _
                "   0, " & vbCrLf
        End If
    
        strSQL = strSQL & _
                "   " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_INIC_REME_ORIG").Text & .selectSingleNode("HO_INIC_REME_ORIG").Text & "00") & ", " & vbCrLf
    
        If Val(.selectSingleNode("DT_FIM_REME_ORIG").Text) <> 0 Then
            strSQL = strSQL & _
                "   " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FIM_REME_ORIG").Text & .selectSingleNode("HO_FIM_REME_ORIG").Text & "00") & ", " & vbCrLf
        Else
            strSQL = strSQL & _
                "   NULL, " & vbCrLf
        End If
    
        strSQL = strSQL & _
                "  '" & .selectSingleNode("CO_REME_PRMR").Text & "', " & vbCrLf & _
                "  '" & .selectSingleNode("CO_REME_ULTI").Text & "', " & vbCrLf & _
                "   SYSDATE, " & vbCrLf & _
                "   NULL, " & vbCrLf
    
        strSQL = strSQL & _
                "   " & fgVlrToDBServer(.selectSingleNode("VA_TOTL_REME").Text) & ", " & vbCrLf & _
                "   " & enumSituacaoRemessa.EmProcessamento & ", " & vbCrLf & _
                "   SYSDATE)"
    End With
    
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela alteração de um registro de controle de remessa, já existente.

Private Function flAlterar(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    With pxmlRemessa.documentElement

        strSQL = "UPDATE " & vbCrLf & _
                 "   A6.TB_CTRL_PROC_REME " & vbCrLf & _
                 "SET " & vbCrLf & _
                 "   DH_FIM_REME_ORIG = " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FIM_REME_ORIG").Text & .selectSingleNode("HO_FIM_REME_ORIG").Text & "00") & ", " & vbCrLf & _
                 "   DH_FIM_PROC_REME = SYSDATE, " & vbCrLf & _
                 "   CO_REME_PRMR     = '" & .selectSingleNode("CO_REME_PRMR").Text & "', " & vbCrLf & _
                 "   CO_REME_ULTI     = '" & .selectSingleNode("CO_REME_ULTI").Text & "', " & vbCrLf & _
                 "   QT_REGT_RECB = (SELECT COUNT(*) FROM A6.TB_MOVI_FUTU_OPER_ATIV " & vbCrLf & _
                 "                   WHERE  SG_SIST      = '" & .selectSingleNode("SG_SIST_ORIG").Text & "' " & vbCrLf & _
                 "                   AND    CO_VEIC_LEGA = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' " & vbCrLf & _
                 "                   AND    DH_REME      BETWEEN " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text & "000000") & vbCrLf & _
                 "                                       AND     " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text & "235959") & vbCrLf
        
        strSQL = strSQL & _
                 "                   AND    CO_REME_MOVI_FUTU BETWEEN '" & .selectSingleNode("CO_REME_PRMR").Text & "' " & vbCrLf & _
                 "                   AND                              '" & .selectSingleNode("CO_REME_ULTI").Text & "'), " & vbCrLf & _
                 "   QT_REGT_REJE = (SELECT COUNT(*) FROM A6.TB_REME_REJE " & vbCrLf & _
                 "                   WHERE  SG_SIST_ORIG_INFO = '" & .selectSingleNode("SG_SIST_ORIG").Text & "' " & vbCrLf & _
                 "                   AND    DH_REME_REJE      BETWEEN " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text & "000000") & vbCrLf & _
                 "                                            AND     " & fgDtHrXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text & "235959") & vbCrLf & _
                 "                   AND    CO_VEIC_LEGA      = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' " & vbCrLf & _
                 "                   AND    CO_EMPR           =  " & .selectSingleNode("CO_EMPR").Text & "), " & vbCrLf & _
                 "   CO_SITU_REME = " & enumSituacaoRemessa.Finalizado & ", " & vbCrLf & _
                 "   DH_ULTI_ATLZ = SYSDATE " & vbCrLf & _
                 "WHERE " & vbCrLf & _
                 "   DT_REME      = " & fgDtXML_To_Oracle(.selectSingleNode("DT_FECH_PROC").Text) & " AND " & vbCrLf & _
                 "   TP_MESG      = " & .selectSingleNode("TP_MESG").Text & " AND " & vbCrLf & _
                 "   SG_SIST      = '" & .selectSingleNode("SG_SIST_ORIG").Text & "' AND " & vbCrLf & _
                 "   CO_VEIC_LEGA = '" & .selectSingleNode("CO_VEIC_LEGA").Text & "'"

    End With

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela exclusão de um registro de controle de remessa.

Private Function flExcluir(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM " & vbCrLf & _
             "       A6.TB_CTRL_PROC_REME " & vbCrLf & _
             " WHERE DT_REME = " & fgDtXML_To_Oracle(pxmlRemessa.documentElement.selectSingleNode("DT_FECH_PROC").Text) & vbCrLf & _
             "   AND TP_MESG = " & CLng(pxmlRemessa.documentElement.selectSingleNode("TP_MESG").Text) & vbCrLf & _
             "   AND SG_SIST = '" & pxmlRemessa.documentElement.selectSingleNode("SG_SIST_ORIG").Text & "'" & vbCrLf & _
             "   AND CO_VEIC_LEGA = '" & pxmlRemessa.documentElement.selectSingleNode("CO_VEIC_LEGA").Text & "'"
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
