VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemessaFinanceiraCAMR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas à Remessas Financeiras CAMR, bem
'' como a manipulação dos dados cadastrados

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

'Montar dinamicamente comando select conforme filtro especificado

Private Function flObterQueryDetalheRemessaCAMR(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String
Dim strWhereFiltro                          As String
Dim strValores                              As String

    On Error GoTo ErrorHandler
    
    'Filtro por Empresa
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_BancoLiquidante/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_EMPR      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro por Data
    If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data").length > 0 Then
        If xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data/DataFim").length > 0 Then
            For Each objDomNode In xmlFiltro.selectNodes("/Repeat_Filtros/Grupo_Data")
                strWhereFiltro = strWhereFiltro & "AND A.DT_EFET_LANC between " & objDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                                                                  objDomNode.selectSingleNode("DataFim").Text
            Next objDomNode
        Else
            strWhereFiltro = strWhereFiltro & _
                    "AND TRUNC(A.DT_EFET_LANC) = " & xmlFiltro.selectSingleNode("//Repeat_Filtros/Grupo_Data/DataIni").Text & vbNewLine
        End If
    End If

    'Filtro por Local de Liquidação
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_LocalLiquidacao/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_LOCA_LIQU      IN (" & strValores & ")" & vbCrLf
    End If
    
    'Filtro Grupo Lançamento Financeiro
    strValores = vbNullString
    For Each objDomNode In xmlFiltro.selectNodes("Repeat_Filtros/Grupo_GrupoLancamentoFinanceiro/*")
        strValores = strValores & objDomNode.Text & ", "
    Next
    
    If Trim(strValores) <> "" Then
        strValores = Left$(strValores, Len(strValores) - 2)
        strWhereFiltro = strWhereFiltro & " AND      A.CO_GRUP_LANC_FINC      IN (" & strValores & ")"
    End If
    
    strSQL = " SELECT   'A8' AS OWNER,                                                      " & vbNewLine & _
             "          A.DT_EFET_LANC,                                                     " & vbNewLine & _
             "          A.CO_GRUP_LANC_FINC,                                                " & vbNewLine & _
             "          A.CO_IDEF_LANC,                                                     " & vbNewLine & _
             "          A.CO_LANC_FINC,                                                     " & vbNewLine & _
             "          A.IN_OPER_DEBT_CRED  AS  CAMPO_IN_OPER_DEBT_CRED,                   " & vbNewLine & _
             "          A.CO_CLIE_QULF,                                                     " & vbNewLine & _
             "          A.CO_COTR            AS  CO_PARP_CAMR,                              " & vbNewLine & _
             "          A.NO_COTR,                                                          " & vbNewLine & _
             "          A.VA_LANC            AS  VA_OPER_ATIV_VLRXML,                       " & vbNewLine & _
             "          A.IN_CLIE_QULF,                                                     " & vbNewLine & _
             "          A.CO_EMPR,                                                          " & vbNewLine & _
             "          A.CO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          A.SG_SIST,                                                          " & vbNewLine & _
             "          A.CO_TEXT_XML,                                                      " & vbNewLine & _
             "          B.NO_VEIC_LEGA,                                                     " & vbNewLine & _
             "          B.TP_BKOF,                                                          " & vbNewLine & _
             "          C.DE_GRUP_LANC_FINC,                                                " & vbNewLine & _
             "          D.DS_PARP_CAMR,                                                     " & vbNewLine & _
             "          E.DE_BKOF                                                           " & vbNewLine

    strSQL = strSQL & _
             " FROM     A8.TB_REME_LIQU_FINC_CAMR       A,                                  " & vbNewLine & _
             "          A8.TB_VEIC_LEGA                 B,                                  " & vbNewLine & _
             "          A8.TB_GRUP_LANC_FINC            C,                                  " & vbNewLine & _
             "          A8.TB_TIPO_BKOF_PARP_CAMR       D,                                  " & vbNewLine & _
             "          A8.TB_TIPO_BKOF                 E                                   " & vbNewLine

    strSQL = strSQL & _
             " WHERE    A.CO_VEIC_LEGA       =  B.CO_VEIC_LEGA                              " & vbNewLine & _
             " AND      A.SG_SIST            =  B.SG_SIST                                   " & vbNewLine & _
             " AND      A.CO_GRUP_LANC_FINC  =  C.CO_GRUP_LANC_FINC                         " & vbNewLine & _
             " AND      A.CO_COTR            =  D.CO_PARP_CAMR                              " & vbNewLine & _
             " AND      B.TP_BKOF            =  E.TP_BKOF                                   " & vbNewLine

    If Trim(strWhereFiltro) <> "" Then
        strSQL = strSQL & strWhereFiltro & vbNewLine
    End If
    
    strWhereSegregacao = fgSegregaDados("A8.TB_REME_LIQU_FINC_CAMR", False, "A", "B", False, False, True, False, True, False)
    
    strSQL = strSQL & strWhereSegregacao & vbNewLine
    
    flObterQueryDetalheRemessaCAMR = strSQL
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterQueryDetalheRemessaCAMR Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    Exit Function
    Resume

End Function

'Obter detalhes de remessas líquidas CAMR conforme filtro especificado

Public Function ObterDetalheRemessaCAMR(ByVal xmlFiltro As MSXML2.DOMDocument40) As String

Dim datInicio                               As Date
Dim datFim                                  As Date
Dim datDataCorte                            As Date

Dim blnChecaBaseHistorica                   As Boolean

Dim blnBaseA8                               As Boolean
Dim blnBaseHistorica                        As Boolean

Dim strSqlA8                                As String
Dim strSQLHist                              As String
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    datInicio = gstrDataVazia
    datFim = gstrDataVazia
    datDataCorte = fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo

    blnChecaBaseHistorica = Not xmlFiltro.selectSingleNode("/Repeat_Filtros/BaseHistorica") Is Nothing

    If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data") Is Nothing Then
    
        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni") Is Nothing Then
            datInicio = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataIni").Text, 10, 8))
        End If

        If Not xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim") Is Nothing Then
            datFim = fgDtXML_To_Date(Mid$(xmlFiltro.selectSingleNode("/Repeat_Filtros/Grupo_Data/DataFim").Text, 10, 8))
        End If

    End If

    If datInicio <> gstrDataVazia Then
        blnBaseHistorica = datInicio < datDataCorte And blnChecaBaseHistorica
    End If

    If datFim <> gstrDataVazia Then
        blnBaseA8 = datFim > datDataCorte
    End If

    strSqlA8 = flObterQueryDetalheRemessaCAMR(xmlFiltro)
    strSQLHist = Replace$(strSqlA8, "A8.TB_REME_LIQU_FINC_CAMR", "A8HIST.TB_REME_LIQU_FINC_CAMR")
    strSQLHist = Replace$(strSQLHist, "'A8' AS OWNER", "'A8HIST' AS OWNER ")

    If blnBaseHistorica And Not blnBaseA8 Then
        strSQL = strSQLHist
    ElseIf blnBaseHistorica And blnBaseA8 Then
        strSQL = strSQLHist & _
                 " UNION ALL " & _
                 strSqlA8
    Else
        strSQL = strSqlA8
    End If
 
    ObterDetalheRemessaCAMR = fgQueryXMLLerTodos("DetalheRemessa", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalheRemessaCAMR Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Incluir registro na tabela A8.TB_REME_LIQU_FINC_CAMR
Public Function Incluir(ByVal pxmlRemessa As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_REME_LIQU_FINC_CAMR (" & vbCrLf & _
             " CO_IDEF_LANC, " & vbCrLf & _
             " DT_EFET_LANC, " & vbCrLf & _
             " ID_ARQU_CAMR, " & vbCrLf & _
             " CO_GRUP_LANC_FINC, " & vbCrLf & _
             " CO_LANC_FINC, " & vbCrLf & _
             " IN_OPER_DEBT_CRED, " & vbCrLf & _
             " CO_CLIE_QULF, " & vbCrLf & _
             " CO_COTR, " & vbCrLf & _
             " NO_COTR, " & vbCrLf & _
             " VA_LANC, " & vbCrLf & _
             " IN_CLIE_QULF, " & vbCrLf & _
             " CO_EMPR, " & vbCrLf & _
             " CO_TEXT_XML, " & vbCrLf & _
             " SG_SIST, " & vbCrLf & _
             " CO_VEIC_LEGA " & vbCrLf & _
             " ) VALUES ( " & vbCrLf

    strSQL = strSQL & _
            "'" & pxmlRemessa.selectSingleNode("//CO_IDEF_LANC").Text & "', " & vbCrLf & _
            fgDtXML_To_Oracle(pxmlRemessa.selectSingleNode("//DT_EFET_LANC").Text) & ", " & vbCrLf & _
            "'" & pxmlRemessa.selectSingleNode("//ID_ARQU_CAMR").Text & "', " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//CO_GRUP_LANC_FINC").Text & ", " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//CO_LANC_FINC").Text & ", " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//IN_OPER_DEBT_CRED").Text & ", " & vbCrLf & _
            IIf(pxmlRemessa.selectSingleNode("//CO_CLIE_QULF").Text = "", "NULL", pxmlRemessa.selectSingleNode("//CO_CLIE_QULF").Text) & ", " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//CO_COTR").Text & ", " & vbCrLf & _
            "'" & pxmlRemessa.selectSingleNode("//NO_COTR").Text & "', " & vbCrLf & _
            fgVlrToDBServer(pxmlRemessa.selectSingleNode("//VA_LANC").Text) & ", " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//IN_CLIE_QULF").Text & ", " & vbCrLf & _
            pxmlRemessa.selectSingleNode("//CO_EMPR").Text & ", " & vbCrLf & _
            "0, " & vbCrLf & _
            "'" & pxmlRemessa.selectSingleNode("//SG_SIST").Text & "', " & vbCrLf & _
            "'" & pxmlRemessa.selectSingleNode("//CO_VEIC_LEGA").Text & "') "


    Call fgExecuteSQL(strSQL)
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Incluir Log registro na tabela A8.TB_REME_LIQU_FINC_CAMR

Public Function IncluirLogRemessa(ByVal pstrIDArquivo As String, _
                                  ByVal pstrNomeArquivo As String) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    'CO_SITU_ARQU_CAMR = 1 EmProcessamento
    '                  = 2 Finalizado
                  
    
    strSQL = " INSERT INTO A8.TB_ARQU_LIQU_FINC_CAMR (" & vbCrLf & _
                " ID_ARQU_CAMR, " & vbCrLf & _
                " NO_ARQU_CAMR, " & vbCrLf & _
                " CO_SITU_PROC, " & vbCrLf & _
                " DH_ULTI_ATLZ, " & vbCrLf & _
                " CD_USUA_ULTI_ATLZ " & vbCrLf & _
                " ) VALUES ( " & vbCrLf

    strSQL = strSQL & _
            "'" & pstrIDArquivo & "', " & vbCrLf & _
            "'" & pstrNomeArquivo & "', " & vbCrLf & _
            "1," & vbCrLf & _
            "SYSDATE," & vbCrLf & _
            "'" & fgUsuarioRede & "')"

    Call fgExecuteSQL(strSQL)
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodosLogRemessa() As String

Dim strSQL                                  As String
Dim strDHInicio                             As String
Dim strDHFim                                As String

On Error GoTo ErrorHandler

    strDHInicio = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000")
    strDHFim = fgDtHrXML_To_Oracle(fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959")
    
    
    strSQL = " SELECT    ID_ARQU_CAMR , " & vbCrLf & _
            "           NO_ARQU_CAMR, " & vbCrLf & _
            "           CO_SITU_PROC, " & vbCrLf & _
            "           DH_ULTI_ATLZ, " & vbCrLf & _
            "           CD_USUA_ULTI_ATLZ" & vbCrLf & _
            "  FROM     A8.TB_ARQU_LIQU_FINC_CAMR" & vbCrLf
            '"  WHERE    DH_ULTI_ATLZ BETWEEN " & strDHInicio & vbCrLf & _
            '"   AND " & strDHFim
    
    LerTodosLogRemessa = fgQueryXMLLerTodos("LerTodosLogRemessa", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosLogRemessa Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Function

