VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMQAX200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Objeto responsável pelo controle de acesso ao objeto MQAX200.

Option Explicit

Public Enum enumMQOO_Open
    MQOO_BROWSE = MQ.MQOO_BROWSE
    MQOO_INPUT_AS_Q_DEF = MQ.MQOO_INPUT_AS_Q_DEF
    MQOO_INPUT_SHARED = MQ.MQOO_INPUT_SHARED
    MQOO_INPUT_EXCLUSIVE = MQ.MQOO_INPUT_EXCLUSIVE
    MQOO_OUTPUT = MQ.MQOO_OUTPUT
End Enum

Public Enum enumMQFeedBack
    MQFB_COA = MQ.MQFB_COA
    MQFB_COD = MQ.MQFB_COD
End Enum

Public Enum enumPutOptions
    Binario = 1
    MainFrame = 2
End Enum

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

Private objMQSession                        As MQAX200.MQSession
Private objMQQueueManager                   As MQAX200.MQQueueManager
Private objMQQueue                          As MQAX200.MQQueue
Private objMQGet                            As MQAX200.MQGetMessageOptions

Private objContext                          As COMSVCSLib.ObjectContext
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

'Commit na transação do MQSeries

Public Sub MQCommit()

On Error GoTo ErrHandler

    objMQQueueManager.Commit
    
    Exit Sub
        
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQCommit", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Get da mensagem na fila do MQSeries

Public Function MQGet(ByRef plngBufferLen As Long, _
                      ByRef pstrBuffer As String, _
             Optional ByVal pblnTransacional = True, _
             Optional ByVal pblnBinario = False, _
             Optional ByRef pstrMessageID As String, _
             Optional ByRef pstrCorrelationID As String, _
             Optional ByRef pstrMessageDescriptorXML As String, _
             Optional ByRef plngBackOutCount As Long) As Boolean

Dim objMQMessage                            As MQAX200.MQMessage
Dim lblnBuffer()                            As Byte

On Error GoTo ErrorHandler
   
    pstrBuffer = ""
    
    Set objMQGet = objMQSession.AccessGetMessageOptions
    
    If pblnTransacional Then
        objMQGet.Options = MQGMO_SYNCPOINT
    Else
        If objMQQueue.OpenOptions = MQOO_BROWSE Then
            objMQGet.Options = MQGMO_BROWSE_NEXT
        Else
            objMQGet.Options = MQGMO_NONE
        End If
    End If
    
    Set objMQMessage = objMQSession.AccessMessage
    
    If Trim(pstrCorrelationID) <> vbNullString Then
        If Len(pstrCorrelationID) = 48 Then
            objMQMessage.CorrelationIdHex = pstrCorrelationID
        Else
            objMQMessage.CorrelationId = Left(pstrCorrelationID & Space(24), 24)
        End If
    ElseIf Trim(pstrMessageID) <> vbNullString Then
        If Len(pstrMessageID) = 48 Then
            objMQMessage.MessageIdHex = pstrMessageID
        Else
            objMQMessage.MessageId = Left(pstrMessageID & Space(24), 24)
        End If
    End If
                   
    objMQQueue.Get objMQMessage, objMQGet
    
    If Len(pstrMessageID) = 24 Then
        pstrMessageID = objMQMessage.MessageId
    Else
        pstrMessageID = objMQMessage.MessageIdHex
    End If
    
    If Len(pstrCorrelationID) = 24 Then
        pstrCorrelationID = objMQMessage.CorrelationId
    Else
        pstrCorrelationID = Replace(objMQMessage.CorrelationId, vbNullChar, "")
        
        If Len(pstrCorrelationID) = 0 Then
            pstrCorrelationID = Replace(objMQMessage.CorrelationId, vbNullChar, "")
        Else
            pstrCorrelationID = objMQMessage.CorrelationIdHex
        End If
    End If
        
    objMQMessage.DataOffset = 0
    
    'Le mensagem em binario
    If objMQMessage.MessageLength <> 0 Then
        lblnBuffer = objMQMessage.Read(objMQMessage.MessageLength)
        'Converte binario para string
        pstrBuffer = StrConv(lblnBuffer, vbUnicode)
        plngBufferLen = objMQMessage.MessageLength
        If Not pblnBinario Then
            pstrBuffer = Replace(pstrBuffer, vbNullChar, " ")
        End If
    Else
        pstrBuffer = vbNullString
    End If
    
    plngBackOutCount = objMQMessage.BackoutCount
    
    pstrMessageDescriptorXML = flObjterMessageDescriptor(objMQMessage)
       
    MQGet = True
    
    Exit Function
    
ErrorHandler:

    MQGet = False
    
    Select Case objMQQueue.ReasonCode
    
        Case MQRC_NO_MSG_AVAILABLE, _
             MQRC_ALREADY_CONNECTED, _
             MQRC_ANOTHER_Q_MGR_CONNECTED
            
            objMQQueue.ClearErrorCodes
            Err.Clear
                        
            Exit Function
            
    End Select
            
            
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQGet", lngCodigoErroNegocio, intNumeroSequencialErro, "MQGet NomeFila:" & objMQQueue.Name)

End Function

'Nick recupera ambiente
Public Function isProduction() As Boolean
    isProduction = fgIsProduction
End Function

'Obter Conecção com o gerinciador de filas MQSeries

Public Sub MQInit(Optional ByVal pstrQueueMmanagerName As String)

On Error GoTo ErrHandler

ConectarMQ:
             
    If pstrQueueMmanagerName = vbNullString Then
        fgObterVariaveisAmbiente
        pstrQueueMmanagerName = gstrQMgrName
    End If

    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(pstrQueueMmanagerName)
        
    Exit Sub

ErrHandler:
        
    If Not objMQQueueManager Is Nothing Then
        Select Case objMQQueueManager.ReasonCode
        
            Case MQRC_UNEXPECTED_ERROR, _
                 MQRC_NO_MSG_AVAILABLE
                
                objMQQueueManager.ClearErrorCodes
                Err.Clear
                            
                Set objMQSession = Nothing
                
                GoTo ConectarMQ
                
        End Select
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQInit", lngCodigoErroNegocio, intNumeroSequencialErro, "Queue Manager Name: " & pstrQueueMmanagerName)

End Sub

'Acessar uma fila do MQSeries (Get, Put)

Public Sub MQOpen(ByVal pstrQueueName As String, _
                  ByVal MQOO_Open As enumMQOO_Open)

On Error GoTo ErrHandler
    
    Set objMQQueue = objMQQueueManager.AccessQueue(pstrQueueName, MQOO_Open)
    
    Exit Sub
    
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQOpen", lngCodigoErroNegocio, intNumeroSequencialErro, "Nome Fila:" & pstrQueueName)

End Sub

'Encerrar acessoa a fila do gerenciador de filas MQSeries

Public Sub MQClose()

On Error GoTo ErrHandler
             
    If Not objMQQueue Is Nothing Then
        objMQQueue.Close
    End If
   
    Exit Sub
    
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQClose", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Put da mensagens na fila do MQSeries

Public Sub MQPut(ByVal pstrBuffer As String, _
                 ByRef plngBufferLen As Long, _
                 ByVal pintPutOptions As enumPutOptions, _
        Optional ByVal pstrReplyToQueueName As String = "", _
        Optional ByVal pstrCorrelationID As String = "", _
        Optional ByRef pstrMessageID As String = "", _
        Optional ByRef pstrMessageDescriptorXML As String, _
        Optional ByVal plngExpiry As Long = 0, _
        Optional ByVal plngPriority As Long = 0)
        
Dim objMQMessage                            As MQAX200.MQMessage
Dim objMQPut                                As MQAX200.MQPutMessageOptions
Dim blnBuffer()                             As Byte

On Error GoTo ErrHandler
                
    Set objMQMessage = objMQSession.AccessMessage
    Set objMQPut = objMQSession.AccessPutMessageOptions
    
    objMQPut.Options = MQPMO_SYNCPOINT
    
    objMQMessage.Persistence = MQ.MQPER_PERSISTENT
    
    'set up the MQSeries bit to make sure the server responds correctly
    If pstrCorrelationID <> vbNullString Then
        If Len(pstrCorrelationID) = 48 Then
            objMQMessage.CorrelationIdHex = pstrCorrelationID
        Else
            If pintPutOptions = MainFrame Then
                objMQMessage.CorrelationId = flToEBCDIC(Left(pstrCorrelationID & Space(24), 24))
            Else
                objMQMessage.CorrelationId = Left(pstrCorrelationID & Space(24), 24)
            End If
        End If
    End If
        
    If pintPutOptions = MainFrame Then
        'Formato para o mainframe poder converter
        objMQMessage.Format = MQFMT_STRING
        'Mensagem texto
        objMQMessage.WriteString pstrBuffer
    ElseIf pintPutOptions = enumPutOptions.Binario Then
        'Converte para Binario
        blnBuffer = StrConv(pstrBuffer, vbFromUnicode)
        ' set up the contents of the output message
        objMQMessage.Write blnBuffer
    Else
        'Parâmetro PutOptions incorreto
        lngCodigoErroNegocio = 10
        GoTo ErrHandler
    End If
    
    If pstrReplyToQueueName <> vbNullString Then
        
        If pintPutOptions = enumPutOptions.Binario Then
            If objMQQueue.Name = pstrReplyToQueueName Then
                objMQMessage.MessageType = MQMT_REPORT
            Else
                objMQMessage.MessageType = MQMT_REPLY
            End If
        End If
        
        If pstrReplyToQueueName = "" Then
            objMQMessage.Report = MQRO_COD
        Else
            objMQMessage.ReplyToQueueName = pstrReplyToQueueName
            objMQMessage.Report = MQRO_COA + MQRO_COD + MQRO_EXCEPTION
        End If
        
        objMQMessage.ReplyToQueueManagerName = gstrQMgrName
        
    End If
    
    If plngPriority > 0 Then
        objMQMessage.Priority = plngPriority
    End If
    
    If plngExpiry > 0 Then
        objMQMessage.Expiry = plngExpiry
    End If
    
    plngBufferLen = objMQMessage.MessageLength

    objMQQueue.Put objMQMessage, objMQPut
    
    pstrMessageID = objMQMessage.MessageIdHex
    
    pstrMessageDescriptorXML = flObjterMessageDescriptor(objMQMessage)
    
    Exit Sub
    
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQPut", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub


'Rollback das transações com MQSeries ( Get,Put)

Public Sub MQRollBack()

On Error GoTo ErrHandler

    objMQQueueManager.Backout
    
    Exit Sub
    
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQRollBack", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Fechar a conecção com o gerenciador de filas MQSeries

Public Sub MQEnd()

On Local Error Resume Next
         
    If Not objMQQueueManager Is Nothing Then
        objMQQueueManager.Disconnect
    End If
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    
    Exit Sub
    
ErrHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQEnd", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'Transformar uma string de EBCDIC para ASCCII

Public Function flToASCII(ByVal pstrValue As String) As String

Dim vntEbcdic                               As Variant
Dim intCont                                 As Long
Dim lngBufferlen                            As Long
    
On Error GoTo ErrorHandler
    
    vntEbcdic = Array( _
    &H0, &H1, &H2, &H3, &H9C, &H9, &H86, &H7F, &H97, &H8D, &H8E, &HB, &HC, &HD, &HE, &HF, _
    &H10, &H11, &H12, &H13, &H9D, &H85, &H8, &H87, &H18, &H19, &H92, &H8F, &H1C, &H1D, &H1E, &H1F, _
    &H80, &H81, &H82, &H83, &H84, &HA, &H17, &H1B, &H88, &H89, &H8A, &H8B, &H8C, &H5, &H6, &H7, _
    &H90, &H91, &H16, &H93, &H94, &H95, &H96, &H4, &H98, &H99, &H9A, &H9B, &H14, &H15, &H9E, &H1A, _
    &H20, &HA0, &HA1, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &H5B, &H2E, &H3C, &H28, &H2B, &H21, _
    &H26, &HA9, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF, &HB0, &HB1, &H5D, &H24, &H2A, &H29, &H3B, &H5E, _
    &H2D, &H2F, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, &HB8, &HB9, &H7C, &H2C, &H25, &H5F, &H3E, &H3F, _
    &HBA, &HBB, &HBC, &HBD, &HBE, &HBF, &HC0, &HC1, &HC2, &H60, &H3A, &H23, &H40, &H27, &H3D, &H22, _
    &HC3, &H61, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, _
    &HCA, &H6A, &H6B, &H6C, &H6D, &H6E, &H6F, &H70, &H71, &H72, &HCB, &HCC, &HCD, &HCE, &HCF, &HD0, _
    &HD1, &H7E, &H73, &H74, &H75, &H76, &H77, &H78, &H79, &H7A, &HD2, &HD3, &HD4, &HD5, &HD6, &HD7, _
    &HD8, &HD9, &HDA, &HDB, &HDC, &HDD, &HDE, &HDF, &HE0, &HE1, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, _
    &H7B, &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &HE8, &HE9, &HEA, &HEB, &HEC, &HED, _
    &H7D, &H4A, &H4B, &H4C, &H4D, &H4E, &H4F, &H50, &H51, &H52, &HEE, &HEF, &HF0, &HF1, &HF2, &HF3, _
    &H5C, &H9F, &H53, &H54, &H55, &H56, &H57, &H58, &H59, &H5A, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, _
    &H30, &H31, &H32, &H33, &H34, &H35, &H36, &H37, &H38, &H39, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF)
    
    lngBufferlen = Len(pstrValue)

    For intCont = 1 To lngBufferlen
        Mid$(pstrValue, intCont, 1) = Chr$(vntEbcdic(Asc(Mid$(pstrValue, intCont, 1))))
    Next
    
    flToASCII = pstrValue

    Exit Function
    
ErrorHandler:
   
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flToASCII", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Transformar uma string de ASCCII para EBCDIC

Public Function flToEBCDIC(ByVal strValue As String) As String

Dim vntAscii                                As Variant
Dim intCont                                 As Long
Dim lngBufferlen                            As Long
    
On Error GoTo ErrorHandler

    vntAscii = Array( _
    &H0, &H1, &H2, &H3, &H37, &H2D, &H2E, &H2F, &H16, &H5, &H25, &HB, &HC, &HD, &HE, &HF, _
    &H10, &H11, &H12, &H13, &H3C, &H3D, &H32, &H26, &H18, &H19, &H3F, &H27, &H1C, &H1D, &H1E, &H1F, _
    &H40, &H4F, &H7F, &H7B, &H5B, &H6C, &H50, &H7D, &H4D, &H5D, &H5C, &H4E, &H6B, &H60, &H4B, &H61, _
    &HF0, &HF1, &HF2, &HF3, &HF4, &HF5, &HF6, &HF7, &HF8, &HF9, &H7A, &H5E, &H4C, &H7E, &H6E, &H6F, _
    &H7C, &HC1, &HC2, &HC3, &HC4, &HC5, &HC6, &HC7, &HC8, &HC9, &HD1, &HD2, &HD3, &HD4, &HD5, &HD6, _
    &HD7, &HD8, &HD9, &HE2, &HE3, &HE4, &HE5, &HE6, &HE7, &HE8, &HE9, &H4A, &HE0, &H5A, &H5F, &H6D, _
    &H79, &H81, &H82, &H83, &H84, &H85, &H86, &H87, &H88, &H89, &H91, &H92, &H93, &H94, &H95, &H96, _
    &H97, &H98, &H99, &HA2, &HA3, &HA4, &HA5, &HA6, &HA7, &HA8, &HA9, &HC0, &H6A, &HD0, &HA1, &H7, _
    &H20, &H21, &H22, &H23, &H24, &H15, &H6, &H17, &H28, &H29, &H2A, &H2B, &H2C, &H9, &HA, &H1B, _
    &H30, &H31, &H1A, &H33, &H34, &H35, &H36, &H8, &H38, &H39, &H3A, &H3B, &H4, &H14, &H3E, &HE1, _
    &H41, &H42, &H43, &H44, &H45, &H46, &H47, &H48, &H49, &H51, &H52, &H53, &H54, &H55, &H56, &H57, _
    &H58, &H59, &H62, &H63, &H64, &H65, &H66, &H67, &H68, &H69, &H70, &H71, &H72, &H73, &H74, &H75, _
    &H76, &H77, &H78, &H80, &H8A, &H8B, &H8C, &H8D, &H8E, &H8F, &H90, &H9A, &H9B, &H9C, &H9D, &H9E, _
    &H9F, &HA0, &HAA, &HAB, &HAC, &HAD, &HAE, &HAF, &HB0, &HB1, &HB2, &HB3, &HB4, &HB5, &HB6, &HB7, _
    &HB8, &HB9, &HBA, &HBB, &HBC, &HBD, &HBE, &HBF, &HCA, &HCB, &HCC, &HCD, &HCE, &HCF, &HDA, &HDB, _
    &HDC, &HDD, &HDE, &HDF, &HEA, &HEB, &HEC, &HED, &HEE, &HEF, &HFA, &HFB, &HFC, &HFD, &HFE, &HFF)
    
    lngBufferlen = Len(strValue)

    For intCont = 1 To lngBufferlen
        Mid$(strValue, intCont, 1) = Chr$(vntAscii(Asc(Mid$(strValue, intCont, 1))))
    Next
    
    flToEBCDIC = strValue

    Exit Function
    
ErrorHandler:
   
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flToEBCDIC", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o message Descriptor de uma mensagem (Cabeçalho)

Private Function flObjterMessageDescriptor(ByVal pobjMQMessage As MQMessage) As String
    
Dim objMessageDescriptor                    As MSXML2.DOMDocument40
    
On Error Resume Next
    
    Set objMessageDescriptor = CreateObject("MSXML2.DOMDocument.4.0")
       
    Call fgAppendNode(objMessageDescriptor, "", "MessageDescriptor", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "StrucId", "MD")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "Report", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "MessageType", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "Expiry", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "Feedback", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "MessageIdHex", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "MessageId", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "CorrelationIdHex", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "CorrelationId", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "BackoutCount", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "ReplyToQueueManagerName", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "ReplyToQueueName", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "UserId", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "AccountingTokenHex", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "PutApplicationType", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "PutApplicationName", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "PutDateTime", fgDtHr_To_Xml(Now))
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "ApplicationOriginData", "")
    Call fgAppendNode(objMessageDescriptor, "MessageDescriptor", "GroupIdHex", "")
       
       
    objMessageDescriptor.selectSingleNode("//Report").Text = CStr(pobjMQMessage.Report)
    objMessageDescriptor.selectSingleNode("//MessageType").Text = CStr(pobjMQMessage.MessageType)
    objMessageDescriptor.selectSingleNode("//Expiry").Text = CStr(pobjMQMessage.Expiry)
    objMessageDescriptor.selectSingleNode("//Feedback").Text = CStr(pobjMQMessage.Feedback)
    objMessageDescriptor.selectSingleNode("//MessageIdHex").Text = CStr(pobjMQMessage.MessageIdHex)
    objMessageDescriptor.selectSingleNode("//MessageId").Text = CStr(pobjMQMessage.MessageId)
    objMessageDescriptor.selectSingleNode("//CorrelationIdHex").Text = CStr(pobjMQMessage.CorrelationIdHex)
    objMessageDescriptor.selectSingleNode("//CorrelationId").Text = CStr(pobjMQMessage.CorrelationId)
    objMessageDescriptor.selectSingleNode("//BackoutCount").Text = CStr(pobjMQMessage.BackoutCount)
    objMessageDescriptor.selectSingleNode("//ReplyToQueueManagerName").Text = CStr(pobjMQMessage.ReplyToQueueManagerName)
    objMessageDescriptor.selectSingleNode("//ReplyToQueueName").Text = CStr(pobjMQMessage.ReplyToQueueName)
    objMessageDescriptor.selectSingleNode("//UserId").Text = CStr(pobjMQMessage.UserId)
    objMessageDescriptor.selectSingleNode("//AccountingTokenHex").Text = CStr(pobjMQMessage.AccountingTokenHex)
    objMessageDescriptor.selectSingleNode("//PutApplicationType").Text = CStr(pobjMQMessage.PutApplicationType)
    objMessageDescriptor.selectSingleNode("//PutApplicationName").Text = CStr(pobjMQMessage.PutApplicationName)
    objMessageDescriptor.selectSingleNode("//PutDateTime").Text = fgDtHr_To_Xml(pobjMQMessage.PutDateTime)
    objMessageDescriptor.selectSingleNode("//ApplicationOriginData").Text = CStr(pobjMQMessage.ApplicationOriginData)
    objMessageDescriptor.selectSingleNode("//GroupIdHex").Text = CStr(pobjMQMessage.GroupIdHex)
                 
    flObjterMessageDescriptor = objMessageDescriptor.xml
    
    Set objMessageDescriptor = Nothing
               
    Exit Function
ErrorHandler:
                   
    Err.Clear
                   
End Function


