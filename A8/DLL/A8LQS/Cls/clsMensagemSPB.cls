VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemSPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas às  mensagens do
'' catalogo SPB.
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

'' Consultar todas as tags e tipos de tags de uma mensagem
Public Function LerMensagem(ByVal pstrCodigoMensagem As String)

Dim strSQL                                  As String
Dim strMensagem                             As String
Dim strTipoTag                              As String
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim xmlTipoTag                              As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    strSQL = " SELECT C.NO_MESG,     " & _
             "        A.SQ_MESG,     " & _
             "        A.SQ_TAG,      " & _
             "        A.NU_ORDE_TAG, " & _
             "        A.IN_OBRI,     " & _
             "        A.IN_NIVE_REPE," & _
             "        B.QT_REPE,     " & _
             "        B.NO_TAG,      " & _
             "        B.DE_TAG,      " & _
             "        B.IN_CATG_TAG, " & _
             "        B.TX_DEFA,     " & _
             "        B.SQ_TIPO_TAG  " & _
             "  FROM  A8.TB_TAG_MENSAGEM A, " & _
             "        A8.TB_TAG          B, " & _
             "        A8.TB_MENSAGEM C      " & _
             " WHERE  C.CO_MESG = '" & pstrCodigoMensagem & "'" & _
             "   AND  A.SQ_TAG  = B.SQ_TAG  " & _
             "   AND  A.SQ_MESG = C.SQ_MESG " & _
             "ORDER BY A.NU_ORDE_TAG        "
    
    strMensagem = fgQueryXMLLerTodos("Mensagem", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strMensagem = "" Then Exit Function
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlTipoTag = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML strMensagem
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        strTipoTag = ""
        If Trim(xmlNode.selectSingleNode("SQ_TIPO_TAG").Text) <> vbNullString Then
            strTipoTag = flLerTipoTag(CLng(xmlNode.selectSingleNode("SQ_TIPO_TAG").Text))
        End If
        
        If strTipoTag <> "" Then
            xmlTipoTag.loadXML strTipoTag
            xmlNode.appendChild xmlTipoTag.documentElement
        End If
        
    Next
        
    LerMensagem = xmlMensagem.xml
    
    Set xmlTipoTag = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlMensagem = Nothing
    Set xmlTipoTag = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar Tipo das Tags e seus respectivos dominios
Private Function flLerTipoTag(ByVal plngSeqTipoTag As Long) As String

Dim strSQL                                  As String
Dim strTipoTag                              As String
Dim strDomininioTipoTag                     As String
Dim xmlTipoTag                              As MSXML2.DOMDocument40
Dim xmlDominioTipoTag                       As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
        
    strSQL = "  SELECT SQ_TIPO_TAG, " & _
             "         NO_TIPO_TAG, " & _
             "         DE_TIPO_TAG, " & _
             "         NU_TAMA_TAG, " & _
             "         IN_TIPO_CTER," & _
             "         IN_TAG_SITU, " & _
             "         QT_CASA_DECI," & _
             "         '' AS DOMI   " & _
             "    FROM A8.TB_TIPO_TAG " & _
             "   WHERE SQ_TIPO_TAG = " & plngSeqTipoTag
        
    strTipoTag = fgQueryXMLLer("Grupo_TipoTag", strSQL, App.EXEName & "." & TypeName(Me))
        
    If strTipoTag = "" Then Exit Function
        
    Set xmlTipoTag = CreateObject("MSXML2.DOMDocument.4.0")
        
    xmlTipoTag.loadXML strTipoTag
        
    strSQL = " SELECT /*+ RULE */ SQ_TIPO_TAG, " & _
             "        CO_DOMI, " & _
             "        DE_DOMI " & _
             "   FROM A8.TB_DOMINIO " & _
             "  WHERE SQ_TIPO_TAG = " & plngSeqTipoTag & _
             "  ORDER BY CO_DOMI "

    strDomininioTipoTag = fgQueryXMLLerTodos("TipoTag", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strDomininioTipoTag <> "" Then
        Set xmlDominioTipoTag = CreateObject("MSXML2.DOMDocument.4.0")
            
        xmlDominioTipoTag.loadXML strDomininioTipoTag
            
        strDomininioTipoTag = " " & vbTab
        
        For Each xmlNode In xmlDominioTipoTag.documentElement.childNodes
            With xmlNode
                strDomininioTipoTag = strDomininioTipoTag & Trim(.selectSingleNode("CO_DOMI").Text) & " - " & Trim(.selectSingleNode("DE_DOMI").Text) & vbTab
            End With
        Next
    End If
        
    xmlTipoTag.documentElement.selectSingleNode("DOMI").Text = strDomininioTipoTag
        
    flLerTipoTag = xmlTipoTag.xml
        
    Set xmlDominioTipoTag = Nothing
    Set xmlTipoTag = Nothing
        
            
    Exit Function
ErrorHandler:
    
    Set xmlDominioTipoTag = Nothing
    Set xmlTipoTag = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterSequenciaMensagem(ByVal pstrCodigoMensagem As String) As Long

Dim strSQL                                  As String
Dim rs                                      As ADODB.Recordset

On Error GoTo ErrorHandler

    strSQL = " SELECT   SQ_MESG          " & _
             "   FROM   A8.TB_MENSAGEM   " & _
             "  WHERE   CO_MESG = '" & Trim$(pstrCodigoMensagem) & "'"

    Set rs = fgQuerySQL(strSQL)
    
    If rs.EOF Then
        ObterSequenciaMensagem = 0
    Else
        If IsNull(rs!SQ_MESG) Then
            ObterSequenciaMensagem = 0
        Else
            ObterSequenciaMensagem = rs!SQ_MESG
        End If
    End If
            
    Set rs = Nothing
        
    Exit Function
ErrorHandler:
    
    Set rs = Nothing
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterSequenciaMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


