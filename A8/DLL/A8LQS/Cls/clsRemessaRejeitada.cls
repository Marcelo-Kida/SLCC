VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemessaRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F295F74034A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
 
'Recuperar e inserir registros na tabela de remessa rejeitada (TB_REME_REJE)


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

'recuperar um registro da tabela de remessa rejeitada

Public Function Ler(ByVal strSG_SIST_ORIG_INFO As String, _
                    ByVal intTP_MESG_INTE As Integer, _
                    ByVal lngCO_EMPR As Long, _
                    ByVal lngCO_TEXT_XML_REJE As Long, _
                    ByVal strDH_REME_REJE As String, _
                    ByVal strOwner As String) As String

Dim strSQL                                  As String
Dim xmlLer                                  As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    strSQL = " Select    A.TP_MESG_INTE,                 " & vbNewLine & _
             "           A.CO_EMPR,                      " & vbNewLine & _
             "           A.CO_TEXT_XML_REJE,             " & vbNewLine & _
             "           A.CO_TEXT_XML_RETN_SIST_ORIG,   " & vbNewLine & _
             "           A.TX_XML_ERRO,                  " & vbNewLine & _
             "           A.SG_SIST_ORIG_INFO,            " & vbNewLine & _
             "           A.DH_REME_REJE                  " & vbNewLine & _
             " From      " & strOwner & ".TB_REME_REJE A               " & vbNewLine & _
             " Where     A.SG_SIST_ORIG_INFO = '" & strSG_SIST_ORIG_INFO & "'" & vbNewLine & _
             " AND       A.CO_EMPR           =  " & lngCO_EMPR & vbNewLine & _
             " AND       A.TP_MESG_INTE      =  " & intTP_MESG_INTE & vbNewLine & _
             " AND       A.CO_TEXT_XML_REJE  =  " & lngCO_TEXT_XML_REJE & vbNewLine & _
             " AND       A.DH_REME_REJE      =  " & fgDtHrXML_To_Oracle(strDH_REME_REJE)

    Ler = fgQueryXMLLer("Erro", strSQL, App.EXEName & "." & TypeName(Me))

    If Len(Ler) = 0 Then
       '4002 - Registro Inexistente ou excluído por outro usuário.
       lngCodigoErroNegocio = 4002
       GoTo ErrorHandler
    Else
        Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlLer.loadXML(Ler)

        '>>>> Foi modificada a gravação do campo "TX_XML_ERRO" para Base 64
        On Error Resume Next
        
        xmlLer.documentElement.selectSingleNode("TX_XML_ERRO").Text = fgBase64Decode(xmlLer.documentElement.selectSingleNode("TX_XML_ERRO").Text)
        
        On Error GoTo ErrorHandler

        Ler = xmlLer.xml
        Set xmlLer = Nothing

    End If

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro, "SQL: " & strSQL)

End Function

'recuperar a remessa original da tabela A8.TB_TEXT_XML

Public Function ObterXMLErroRemessaRejeitada(ByVal lngCO_TEXT_XML_REJE As Long) As String

On Error GoTo ErrorHandler

    ObterXMLErroRemessaRejeitada = fgSelectVarchar4000(lngCO_TEXT_XML_REJE)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro, "CO_TEXT_XML = " & lngCO_TEXT_XML_REJE)

End Function

'recuperar registros da tabela de remessa rejeitada

Public Function LerTodos(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String


Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim blnPrimeiro                             As Boolean
Dim blnBaseHistorica                        As Boolean
Dim blnBaseCorrente                         As Boolean
    
On Error GoTo ErrorHandler

    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseHistorica = True
    End If
    
    If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseCorrente = True
    End If
             
    strSQL = " Select    'A8' AS OWNER, " & vbCrLf & _
             "           A.TP_MESG_INTE,                      " & vbNewLine & _
             "           A.CO_EMPR,                           " & vbNewLine & _
             "           A.SG_SIST_ORIG_INFO,                 " & vbNewLine & _
             "           A.CO_TEXT_XML_REJE,                  " & vbNewLine & _
             "           A.CO_TEXT_XML_RETN_SIST_ORIG,        " & vbNewLine & _
             "           A.TX_XML_ERRO,                       " & vbNewLine & _
             "           A.DH_REME_REJE,                      " & vbNewLine & _
             "           B.NO_SIST,                           " & vbNewLine & _
             "           C.NO_REDU_EMPR,                      " & vbNewLine & _
             "           D.NO_TIPO_MESG                       " & vbNewLine & _
             " From      A8.TB_REME_REJE A,                   " & vbNewLine & _
             "           A7.TB_SIST B,                        " & vbNewLine & _
             "           A8.TB_EMPRESA_HO C,                  " & vbNewLine & _
             "           A7.TB_TIPO_MESG D                    " & vbNewLine & _
             " Where     LTRIM(RTRIM(A.SG_SIST_ORIG_INFO))  =  LTRIM(RTRIM(B.SG_SIST (+)))     " & vbNewLine & _
             " AND       A.CO_EMPR                          =  B.CO_EMPR (+)    " & vbNewLine & _
             " AND       C.CO_EMPR                          =  A.CO_EMPR               " & vbNewLine & _
             " AND       LTRIM(RTRIM(D.TP_MESG))           =  LTRIM(RTRIM(A.TP_MESG_INTE)) "
            
    If pxmlFiltros.childNodes.Item(0).baseName = "Repeat_Filtros" Then
    
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_BancoLiquidante")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.CO_EMPR In (" & xmlDomNode.selectSingleNode("BancoLiquidante").Text
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", " & xmlDomNode.selectSingleNode("BancoLiquidante").Text
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If
        
        blnPrimeiro = True
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Sistema")
        
            If blnPrimeiro Then
                strSQL = strSQL & "AND A.SG_SIST_ORIG_INFO In ('" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
                blnPrimeiro = False
            Else
                strSQL = strSQL & ", '" & xmlDomNode.selectSingleNode("Sistema").Text & "'"
            End If
            
        Next xmlDomNode
    
        If blnPrimeiro = False Then
           strSQL = strSQL & ")"
        End If
        
        For Each xmlDomNode In pxmlFiltros.documentElement.selectNodes("//Grupo_Data")
            
            If Trim(xmlDomNode.selectSingleNode("DataIni").Text) = Trim(xmlDomNode.selectSingleNode("DataFim").Text) Then
                strSQL = strSQL & " AND TO_DATE(DH_REME_REJE) = " & xmlDomNode.selectSingleNode("DataIni").Text
            Else
                strSQL = strSQL & " AND A.DH_REME_REJE between " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                                      xmlDomNode.selectSingleNode("DataFim").Text
            End If
        Next xmlDomNode
        
    End If
    
    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER"), "A8.TB_REME_REJE", "A8HIST.TB_REME_REJE")
        End If
    Else
        strSQL = Replace(Replace(strSQL, "'A8' AS OWNER", "'A8HIST' AS OWNER"), "A8.TB_REME_REJE", "A8HIST.TB_REME_REJE")
    End If
        
    strSQL = strSQL & " order by DH_REME_REJE desc"
    
    LerTodos = fgQueryXMLLerTodos("Erro", strSQL, App.EXEName & "." & TypeName(Me))

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

'Incluir um registro na tabela de remessa rejeitada


Public Function Incluir(ByVal pxmlRemessaRejeitada As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim lngCodigoTextoXmlReme                   As Long

On Error GoTo ErrorHandler

    With pxmlRemessaRejeitada.documentElement

        lngCodigoTextoXmlReme = fgInsertVarchar4000(.selectSingleNode("TX_XML_REME").Text)

        If .selectSingleNode("TX_XML_ERRO").Text = vbNullString Then
            .selectSingleNode("TX_XML_ERRO").Text = "Sem descrição"
        End If

        '>>>> Foi modificada a gravação do campo "TX_XML_ERRO" para Base 64
        .selectSingleNode("TX_XML_ERRO").Text = fgBase64Encode(.selectSingleNode("TX_XML_ERRO").Text)

        strSQL = " INSERT INTO A8.TB_REME_REJE ( " & vbCrLf & _
                 " SG_SIST_ORIG_INFO,   " & vbCrLf & _
                 " TP_MESG_INTE,        " & vbCrLf & _
                 " CO_EMPR,             " & vbCrLf & _
                 " CO_TEXT_XML_REJE,    " & vbCrLf & _
                 " CO_TEXT_XML_RETN_SIST_ORIG,    " & vbCrLf & _
                 " TX_XML_ERRO,         " & vbCrLf & _
                 " DH_REME_REJE         " & vbCrLf & _
                 " ) VALUES ( " & vbCrLf

        strSQL = strSQL & _
                 IIf(.selectSingleNode("SG_SIST_INFO").Text = vbNullString, " NULL ", " '" & .selectSingleNode("SG_SIST_INFO").Text & "' ") & " , " & vbCrLf & _
                 IIf(.selectSingleNode("TP_MESG").Text = vbNullString, " NULL ", " " & CLng("0" & .selectSingleNode("TP_MESG").Text)) & "  , " & vbCrLf & _
                 IIf(.selectSingleNode("CO_EMPR").Text = vbNullString, " NULL ", " " & CLng("0" & .selectSingleNode("CO_EMPR").Text)) & "  , " & vbCrLf & _
                 lngCodigoTextoXmlReme & ", " & vbCrLf & _
                 lngCodigoTextoXmlReme & ", " & vbCrLf & _
                 IIf(.selectSingleNode("TX_XML_ERRO").Text = vbNullString, " NULL ", " '" & .selectSingleNode("TX_XML_ERRO").Text & "' ") & " , " & vbCrLf & _
                 fgDataHoraServidor_To_Oracle & " )" & vbCrLf
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


