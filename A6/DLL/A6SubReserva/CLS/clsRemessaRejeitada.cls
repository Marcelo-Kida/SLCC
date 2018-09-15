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
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsMovimentoRejeitado\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105A6501BD"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Remessas Rejeitadas pelo A6.

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

' Este método é responsável pela leitura do XML completo da remessa rejeitada pelo A6,
' e pelo retorno à camada intermediária.

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura de detalhes da remessa rejeitada,
' e pelo retorno à camada intermediária.

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
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set xmlLer = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela atualização geral da tabela de remessas rejeitadas pelo A6,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlRemessaRejeitada As MSXML2.DOMDocument40) As Boolean
    On Error GoTo ErrorHandler

    flIncluir pxmlRemessaRejeitada

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela leitura da tabela de remessas rejeitadas pelo A6,
' e pelo retorno à camada intermediária.

Public Function LerTodos(ByVal pxmlFiltros As MSXML2.DOMDocument40) As String
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strSQL                                  As String
Dim blnPrimeiro                             As Boolean

On Error GoTo ErrorHandler

    strSQL = " SELECT    'A6' AS OWNER,                  " & vbNewLine & _
             "           A.TP_MESG_INTE,                 " & vbNewLine & _
             "           A.CO_EMPR,                      " & vbNewLine & _
             "           A.SG_SIST_ORIG_INFO,            " & vbNewLine & _
             "           A.CO_TEXT_XML_REJE,             " & vbNewLine & _
             "           A.TX_XML_ERRO,                  " & vbNewLine & _
             "           A.DH_REME_REJE,                 " & vbNewLine & _
             "           A.CO_VEIC_LEGA,                 " & vbNewLine & _
             "           B.NO_SIST,                      " & vbNewLine & _
             "           C.NO_REDU_EMPR,                 " & vbNewLine & _
             "           D.NO_TIPO_MESG,                 " & vbNewLine & _
             "           E.NO_VEIC_LEGA                  " & vbNewLine & _
             " FROM      A6.TB_REME_REJE A,              " & vbNewLine & _
             "           A7.TB_SIST      B,              " & vbNewLine & _
             "           A8.TB_EMPRESA_HO   C,              " & vbNewLine & _
             "           A7.TB_TIPO_MESG D,              " & vbNewLine & _
             "           A8.TB_VEIC_LEGA E               " & vbNewLine & _
             " WHERE     A.SG_SIST_ORIG_INFO  = B.SG_SIST      (+)" & vbNewLine & _
             " AND       A.CO_EMPR            = B.CO_EMPR      (+)" & vbNewLine & _
             " AND       A.CO_EMPR            = C.CO_EMPR      (+)" & vbNewLine & _
             " AND       A.TP_MESG_INTE       = D.TP_MESG      (+)" & vbNewLine & _
             " AND       A.CO_VEIC_LEGA       = E.CO_VEIC_LEGA (+)" & vbNewLine & _
             " AND       A.SG_SIST_ORIG_INFO  = E.SG_SIST      (+)" & vbNewLine

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
        
            strSQL = strSQL & " AND A.DH_REME_REJE between " & xmlDomNode.selectSingleNode("DataIni").Text & " AND" & vbNewLine & _
                              xmlDomNode.selectSingleNode("DataFim").Text
                              
            If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataFim").Text, 10, 8)) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
                If fgDtXML_To_Date(Mid$(pxmlFiltros.selectSingleNode("//DataIni").Text, 10, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
                    strSQL = strSQL & vbCrLf & _
                             " UNION ALL " & vbCrLf & _
                             Replace(Replace(strSQL, "'A6' AS OWNER", "'A6HIST' AS OWNER"), "A6.TB_REME_REJE", "A6HIST.TB_REME_REJE")
                End If
            Else
                strSQL = Replace(Replace(strSQL, "'A6' AS OWNER", "'A6HIST' AS OWNER"), "A6.TB_REME_REJE", "A6HIST.TB_REME_REJE")
            End If
                              
        Next xmlDomNode
        
    End If
    
    strSQL = strSQL & " ORDER BY SG_SIST_ORIG_INFO, CO_EMPR, DH_REME_REJE"

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

' Este método é responsável pela inclusão de uma nova remessa rejeitada.

Public Function flIncluir(ByVal pxmlRemessaRejeitada As MSXML2.DOMDocument40) As Boolean
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
        
        strSQL = " INSERT INTO A6.TB_REME_REJE ( " & vbCrLf & _
                 " SG_SIST_ORIG_INFO,   " & vbCrLf & _
                 " TP_MESG_INTE,        " & vbCrLf & _
                 " CO_EMPR,             " & vbCrLf & _
                 " CO_TEXT_XML_REJE,    " & vbCrLf & _
                 " TX_XML_ERRO,         " & vbCrLf & _
                 " DH_REME_REJE,        " & vbCrLf & _
                 " CO_VEIC_LEGA         " & vbCrLf & _
                 " ) VALUES ( " & vbCrLf
             
        strSQL = strSQL & _
                 IIf(.selectSingleNode("SG_SIST_INFO").Text = vbNullString, " NULL ", " '" & .selectSingleNode("SG_SIST_INFO").Text & "' ") & " , " & vbCrLf & _
                 IIf(.selectSingleNode("TP_MESG").Text = vbNullString, " NULL ", " " & CLng("0" & .selectSingleNode("TP_MESG").Text)) & "  , " & vbCrLf & _
                 IIf(.selectSingleNode("CO_EMPR").Text = vbNullString, " NULL ", " " & CLng("0" & .selectSingleNode("CO_EMPR").Text)) & "  , " & vbCrLf & _
                 " " & lngCodigoTextoXmlReme & " , " & vbCrLf & _
                 "'" & .selectSingleNode("TX_XML_ERRO").Text & "', " & vbCrLf & _
                 fgDataHoraServidor_To_Oracle & ", " & vbCrLf & _
                 IIf(.selectSingleNode("CO_VEIC_LEGA").Text = vbNullString, "NULL ", " '" & .selectSingleNode("CO_VEIC_LEGA").Text & "' ") & ")"

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
