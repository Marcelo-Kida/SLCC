VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilaErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo reprocessamento de mensagens contidas em filas de erro dos sistemas A6, A7 e A8.
Option Explicit

Implements ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private objMQAX200                          As A6A7A8CA.clsMQAX200
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

'Controlar o reprocessamento de lotes de mensagens de acordo com os MessageIDs informados.
Public Function ReprocessarMensagens(ByVal pstrLoteMensagem As String) As Boolean

Dim xmlLote                                 As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strMessageId                            As String
Dim strFilaErro                             As String
Dim strAcao                                 As String


On Error GoTo ErrorHandler

    Set xmlLote = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLote.loadXML pstrLoteMensagem
    
    For Each objDomNode In xmlLote.documentElement.childNodes
        strMessageId = objDomNode.selectSingleNode("MESSAGE_ID").Text
        strFilaErro = objDomNode.selectSingleNode("NO_FILA_ERRO_MQSE").Text
        
        If Not objDomNode.selectSingleNode("@Acao") Is Nothing Then
            strAcao = objDomNode.selectSingleNode("@Acao").Text
        Else
            strAcao = "Arquivar"
        End If
        
        If strAcao <> "" Then
            flReprocessar strMessageId, strFilaErro, strAcao
        End If
    Next objDomNode
    
    
    Set xmlLote = Nothing

Exit Function
ErrorHandler:

    Set xmlLote = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "ReprocessarMensagens", 0


End Function

'Reprocessar uma mensagem específica através do MessageID informado.
Private Function flReprocessar(ByVal pstrMessageId As String, _
                               ByVal pstrFilaErro As String, _
                               ByVal pstrAcao As String) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim strMensagemErro                         As String
Dim lngMensagemErro                         As Long
Dim xmlMensagemErro                         As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim lngTamanhoMensagem                      As Long
Dim strNomeArquivo                          As String

On Error GoTo ErrorHandler


    Set xmlMensagemErro = CreateObject("MSXML2.DOMDocument.4.0")

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    objMQAX200.MQInit
    objMQAX200.MQOpen pstrFilaErro, enumMQOO_Open.MQOO_INPUT_EXCLUSIVE
    If Not objMQAX200.MQGet(lngMensagemErro, _
                            strMensagemErro, True, , _
                            pstrMessageId) Then
                            
        '9999 - Mensagem não localizada na fila especificada
        lngCodigoErroNegocio = 9999
        GoTo ErrorHandler
    End If
    
    xmlMensagemErro.loadXML strMensagemErro
    objMQAX200.MQClose
    
    If pstrAcao = "Reprocessar" Then
        strMensagem = xmlMensagemErro.documentElement.selectSingleNode("TX_MESG_ORIG").Text
        lngTamanhoMensagem = Len(strMensagem)
        objMQAX200.MQOpen xmlMensagemErro.documentElement.selectSingleNode("NO_FILA_ORIG_MQSE").Text, enumMQOO_Open.MQOO_OUTPUT
        objMQAX200.MQPut strMensagem, lngTamanhoMensagem, Binario
    Else
        strNomeArquivo = Replace(pstrFilaErro, ".", "") & "_LOG" & "_" & Format(Now, "yyyymmddHHmmss") & ".dat"
        fgGravarArquivo strNomeArquivo, strMensagemErro
    End If
    
    objMQAX200.MQEnd
    Set objMQAX200 = Nothing
    Set xmlMensagemErro = Nothing

Exit Function
ErrorHandler:

    objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing
    Set xmlMensagemErro = Nothing

    fgRaiseError App.EXEName, TypeName(Me), "LerMensagens", 0

End Function

'Controlar a leitura de informações das filas de erro dos sistemas A6, A7 e A8.
Public Function LerMensagens(ByVal pstrFiltro As String) As String

Dim xmlFiltros                              As MSXML2.DOMDocument40
Dim xmlMensagens                            As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strFila                                 As String

On Error GoTo ErrorHandler

    Set xmlFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    xmlFiltros.loadXML pstrFiltro
    
    Set xmlMensagens = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlMensagens, "", "Repeat_Mensagem", ""
    
    For Each objDomNode In xmlFiltros.documentElement.selectSingleNode("Repeat_Sistema").childNodes
        Select Case objDomNode.Text
            Case "A6"
                strFila = "A6Q.E.ERRO"
            Case "A7"
                strFila = "A7Q.E.ERRO"
            Case "A8"
                strFila = "A8Q.E.ERRO"
        End Select
        
        flLerMensagensFila strFila, xmlMensagens
        
    Next objDomNode
    
    LerMensagens = xmlMensagens.xml
    
    Set xmlFiltros = Nothing

Exit Function
ErrorHandler:

    Set xmlFiltros = Nothing

   fgRaiseError App.EXEName, TypeName(Me), "LerMensagens", 0
End Function

'Obter informações das mensagens contidas nas filas de erro dos sistemas A6, A7 e A8.
Private Function flLerMensagensFila(ByVal pstrFila As String, _
                                    ByRef pxmlMensagens As MSXML2.DOMDocument40) As Boolean

Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strMessageDescriptor                    As String
Dim strMessageId                            As String
Dim xmlMensagem                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")

    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    objMQAX200.MQInit
    objMQAX200.MQOpen pstrFila, MQOO_BROWSE
    
    While objMQAX200.MQGet(lngBufferLen, _
                           strBuffer, _
                           False, _
                           , _
                           strMessageId, , _
                           strMessageDescriptor)
                         
        If xmlMensagem.loadXML(strBuffer) Then
            If Not xmlMensagem.selectSingleNode("//FILA_ERRO") Is Nothing Then
                fgAppendNode xmlMensagem, "FILA_ERRO", "MESSAGE_ID", strMessageId
                fgAppendNode xmlMensagem, "FILA_ERRO", "NO_FILA_ERRO_MQSE", pstrFila
                fgAppendXML pxmlMensagens, "Repeat_Mensagem", xmlMensagem.xml
            Else
                xmlMensagem.loadXML (flMensagemDesconhecida(strBuffer, strMessageDescriptor))
                fgAppendNode xmlMensagem, "FILA_ERRO", "MESSAGE_ID", strMessageId
                fgAppendNode xmlMensagem, "FILA_ERRO", "NO_FILA_ERRO_MQSE", pstrFila
                fgAppendXML pxmlMensagens, "Repeat_Mensagem", xmlMensagem.xml
            End If
        Else
            xmlMensagem.loadXML (flMensagemDesconhecida(strBuffer, strMessageDescriptor))
            fgAppendNode xmlMensagem, "FILA_ERRO", "MESSAGE_ID", strMessageId
            fgAppendNode xmlMensagem, "FILA_ERRO", "NO_FILA_ERRO_MQSE", pstrFila
            fgAppendXML pxmlMensagens, "Repeat_Mensagem", xmlMensagem.xml
        End If
        
        strMessageId = vbNullString
        strMessageDescriptor = vbNullString
    Wend
    
    objMQAX200.MQEnd

    Set objMQAX200 = Nothing

Exit Function
ErrorHandler:

    objMQAX200.MQEnd

    Set objMQAX200 = Nothing

   fgRaiseError App.EXEName, TypeName(Me), "flLerMensagensFila", 0
End Function

'Enviar mensagens rejeitadas para fila de erro.
Public Function PutFilaErro(ByVal pstrNomeFilaGet As String, _
                            ByVal pstrNomeFilaErro As String, _
                            ByVal pstrMessageIdHex As String, _
                            ByRef pstrLogErro As String, _
                   Optional ByRef pstrCorrelationID As String, _
                   Optional ByVal pblnMesgReje As Boolean) As Boolean
                         
Dim objMQAX200                              As A6A7A8CA.clsMQAX200
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim lngEnumFilaEntrada                      As Long
Dim xmlMD                                   As MSXML2.DOMDocument40
Dim strMD                                   As String

Dim xmlFilaErro                             As MSXML2.DOMDocument40
Dim xmlNodeMesgOrig                         As MSXML2.IXMLDOMNode
Dim xmlNodeMesgErro                         As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set xmlFilaErro = CreateObject("MSXML2.DOMDocument.4.0")
   
    Call fgAppendNode(xmlFilaErro, "", "FILA_ERRO", "")
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "NO_FILA_ORIG_MQSE", pstrNomeFilaGet)
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "DH_MESG_ERRO", Format(Now, "yyyyMMddhhmmss"))
    If Not pblnMesgReje Then
        Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "TX_MESG_ORIG", "")
    Else
        Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "TX_MESG_ORIG", pstrMessageIdHex)
    End If
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "TX_MESG_ERRO", "")
    
    Set xmlNodeMesgOrig = xmlFilaErro.selectSingleNode("//TX_MESG_ORIG")
    Set xmlNodeMesgErro = xmlFilaErro.selectSingleNode("//TX_MESG_ERRO")
    
    xmlNodeMesgErro.appendChild fgCreateCDATASection(pstrLogErro)
    
    Set objMQAX200 = CreateObject("A6A7A8CA.clsMQAX200")
    
    objMQAX200.MQInit
    objMQAX200.MQOpen pstrNomeFilaGet, MQOO_INPUT_SHARED

    If Not pblnMesgReje Then
        If objMQAX200.MQGet(lngBufferLen, _
                            strBuffer, _
                            True, _
                            True, _
                            pstrMessageIdHex, _
                            "", _
                            strMD) Then
            
            
            objMQAX200.MQClose
                        
            xmlNodeMesgOrig.appendChild fgCreateCDATASection(strBuffer)
            
            strBuffer = xmlFilaErro.xml
            
            objMQAX200.MQOpen pstrNomeFilaErro, enumMQOO_Open.MQOO_OUTPUT
            objMQAX200.MQPut strBuffer, Len(strBuffer), enumPutOptions.Binario
            objMQAX200.MQClose
        
        End If
    Else
        'Adrian - demanda de reprocessamento para mensagens rejeitadas pelo A7
        xmlNodeMesgOrig.appendChild fgCreateCDATASection(strBuffer)
        
        strBuffer = xmlFilaErro.xml
        
        objMQAX200.MQOpen pstrNomeFilaErro, enumMQOO_Open.MQOO_OUTPUT
        objMQAX200.MQPut strBuffer, Len(strBuffer), enumPutOptions.Binario
        objMQAX200.MQClose
    End If
    
    objMQAX200.MQEnd
    
    
    'Pikachu - 24/01/2005
    'Demanda - 1025
    Call flEnviarEmailErroNaoPrevisto(xmlFilaErro)
    
    Set objMQAX200 = Nothing
    Set xmlFilaErro = Nothing
    
    Exit Function

ErrorHandler:
    
    objMQAX200.MQEnd
    
    Set objMQAX200 = Nothing
    Set xmlFilaErro = Nothing
    
End Function

'Formatar a mensagem rejeitada acrescentando informações pertinentes a ela.
Private Function flMensagemDesconhecida(ByVal pstrMensagem As String, _
                                        ByVal pstrMessageDescriptor As String) As String
                                        
Dim xmlFilaErro                             As MSXML2.DOMDocument40
Dim xmlMD                                   As MSXML2.DOMDocument40
Dim xmlNodeMesgOrig                         As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
                                            
    Set xmlMD = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlFilaErro = CreateObject("MSXML2.DOMDocument.4.0")
   
    xmlMD.loadXML pstrMessageDescriptor
   
    Call fgAppendNode(xmlFilaErro, "", "FILA_ERRO", "")
    Call fgAppendAttribute(xmlFilaErro, "FILA_ERRO", "MensagemDesconhecida", "1")
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "NO_FILA_ORIG_MQSE", xmlMD.selectSingleNode("//PutApplicationName").Text)
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "DH_MESG_ERRO", xmlMD.selectSingleNode("//PutDateTime").Text)
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "TX_MESG_ORIG", "")
    Call fgAppendNode(xmlFilaErro, "FILA_ERRO", "TX_MESG_ERRO", "Mensagem desconhecida")
    
    Set xmlNodeMesgOrig = xmlFilaErro.selectSingleNode("//TX_MESG_ORIG")
    
    xmlNodeMesgOrig.appendChild fgCreateCDATASection(pstrMensagem)

    flMensagemDesconhecida = xmlFilaErro.xml

    Set xmlFilaErro = Nothing
    
    Exit Function

ErrorHandler:
    
    Set xmlFilaErro = Nothing

End Function

'Formatar a mensagem com erro não previsto acrescentando informações pertinentes a ela.

Private Function flEnviarEmailErroNaoPrevisto(ByRef pxmlmensagemErro As MSXML2.DOMDocument40)

Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strErro                                 As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objMensagem                             As A7Server.clsMensagem
Dim strMensagem                             As String

Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim objRespostaOcorrencia                   As A7Server.clsRespostaOcorrencia

On Error GoTo ErrorHandler
                                                            
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = CreateObject("A7Server.clsMensagem")
                                                            
    strMensagem = objMensagem.ObterPropriedades
    
    xmlMensagem.loadXML strMensagem
            
    xmlMensagem.selectSingleNode("//SG_SIST_DEST").Text = "A8"
    xmlMensagem.selectSingleNode("//CO_EMPR_ORIG").Text = enumCodigoEmpresa.Santander
                
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")

    xmlErro.loadXML pxmlmensagemErro.selectSingleNode("//TX_MESG_ERRO").Text

    strErro = vbCrLf & _
              " " & String(57, "*") & vbCrLf & _
              " Verificar mensagem de erro no módulo A7 (Ferramentas/Reprocessamento de Mensagens)" & vbCrLf & _
              " " & String(57, "*") & vbCrLf

    If Not xmlErro.selectSingleNode("//Number") Is Nothing Then
        strErro = strErro & " Error Number         : " & xmlErro.selectSingleNode("//Number").Text & vbCrLf & _
                            " Error Description    : " & xmlErro.selectSingleNode("//Description").Text & vbCrLf & _
                            " ComputerName         : " & xmlErro.selectSingleNode("//ComputerName").Text & vbCrLf & _
                            " Time                 : " & xmlErro.selectSingleNode("//Time").Text & vbCrLf & _
                            " Mensagem Entrada     : " & vbCrLf & pxmlmensagemErro.selectSingleNode("//TX_MESG_ORIG").Text
    Else
        strErro = strErro & vbCrLf & _
                  " Mensagem Entrada    : " & pxmlmensagemErro.selectSingleNode("//TX_MESG_ORIG").Text & vbCrLf & _
                  " Mensagem Erro       : " & xmlErro.xml & _
                  vbCrLf
    End If
    
    Set xmlNode = xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO")
    xmlNode.appendChild fgCreateCDATASection(strErro)
                                                                
    Set objRespostaOcorrencia = CreateObject("A7Server.clsRespostaOcorrencia")
    Call objRespostaOcorrencia.EnviarRespostaOcorrencia(xmlMensagem, enumOcorrencia.FalhaNaoPrevista)
    Set objRespostaOcorrencia = Nothing
    
    
    Set xmlErro = Nothing
    Set xmlMensagem = Nothing

    
    Exit Function

ErrorHandler:
    
    Set xmlErro = Nothing
    Set xmlMensagem = Nothing
    Set objRespostaOcorrencia = Nothing
    
    Err.Clear


End Function

