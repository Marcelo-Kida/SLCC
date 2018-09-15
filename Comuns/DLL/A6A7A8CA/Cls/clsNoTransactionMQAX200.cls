VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNoTransactionMQAX200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : CA
'Classe         : clsNoTransactionMQAX200
'Data Criação   : 30-04-2001 11:00
'Objetivo       : Controle de acesso ao objeto: MQAX200
'Analista       : Marcos Gennari Patroni
'
'Programador    : Marcos Gennari Patroni
'Data           : 01-05-2001 12:58
'
'Teste          :
'Autor          :
'

Option Explicit

Public Enum CacheOptions
    Trigger = 1
    DominioMensagem = 2
    ValidadorMensagem = 3
    CopyFixoMensagemEnvio = 4
    RegraMensagem = 5
    CopyFixoMensagemRecebimento = 6
    SituacaoLancamento = 7
    ConfiguracaoMensagem = 8
End Enum

'Variável utilizada para tratamento de erros
Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

Private objMQSession                        As MQAX200.MQSession
Private objMQQueueManager                   As MQAX200.MQQueueManager
Private objMQQueue                          As MQAX200.MQQueue
Private objMQGet                            As MQAX200.MQGetMessageOptions
Private objMQMessage                        As MQAX200.MQMessage

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

Public Sub MQInit(Optional ByVal psQMgrName As String)
On Error GoTo ErrHandler
             
    If psQMgrName = vbNullString Then
        fgObterVariaveisAmbiente
        psQMgrName = gsQMgrName
    End If

    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(psQMgrName)
        
    Exit Sub
    
ErrHandler:
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQInit", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

Public Sub MQEnd()
'Alteração feita por Adilson em 26/08/2001
'On Error GoTo ErrHandler
On Local Error Resume Next
         
    If Not objMQQueueManager Is Nothing Then
        objMQQueueManager.Disconnect
    End If
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    
    Exit Sub
    
ErrHandler:
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQEnd", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

Public Sub MQPutPKI(ByVal psFila As String, _
                    ByRef psMensagem As String, _
                    ByRef psMessageID As String, _
                    ByRef psCorrelationID As String)
        
On Error GoTo ErrHandler

Dim objMQMessage                            As MQAX200.MQMessage
Dim objMQPut                                As MQAX200.MQPutMessageOptions
Dim lbBinario()                             As Byte
Dim liCount                                 As Integer
Dim lsReplyToQueueManagerName               As String
                
    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_OUTPUT)
    Set objMQMessage = objMQSession.AccessMessage
    Set objMQPut = objMQSession.AccessPutMessageOptions
        
    objMQMessage.Persistence = MQPER_NOT_PERSISTENT
    lbBinario = StrConv(psMensagem, vbFromUnicode)
    objMQMessage.Expiry = 1200
    objMQMessage.Write lbBinario
    objMQPut.Options = MQPMO_NO_SYNCPOINT
    If Len(psCorrelationID) = 48 Then
        objMQMessage.CorrelationIdHex = psCorrelationID
    End If
    objMQQueue.Put objMQMessage, objMQPut
    psMessageID = objMQMessage.MessageIdHex
    Exit Sub
    
ErrHandler:
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQEnd", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

Public Sub MQGetPKI(ByVal psFila As String, _
                    ByRef psMensagem As String, _
                    ByRef psMessageID As String, _
                    ByRef psCorrelationID As String)

On Error GoTo ErrHandler

Dim objMQMessage                            As MQAX200.MQMessage
Dim lbBinario()                             As Byte
    
    psMensagem = vbNullString
    
    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_INPUT_SHARED)
    Set objMQGet = objMQSession.AccessGetMessageOptions
    
    objMQGet.Options = MQGMO_NO_SYNCPOINT + MQGMO_WAIT
    Set objMQMessage = objMQSession.AccessMessage
    objMQMessage.Persistence = MQPER_NOT_PERSISTENT
    If Len(psCorrelationID) = 48 Then
        objMQMessage.CorrelationIdHex = psCorrelationID
    End If
    objMQGet.WaitInterval = 20000
    objMQQueue.Get objMQMessage, objMQGet
    objMQMessage.DataOffset = 0
    
    'Le mensagem em binario
    If objMQMessage.MessageLength <> 0 Then
        lbBinario = objMQMessage.Read(objMQMessage.MessageLength)
        psMensagem = StrConv(lbBinario, vbUnicode)
    Else
        psMensagem = vbNullString
    End If
    psMessageID = objMQMessage.MessageIdHex
    
    Exit Sub
    
ErrHandler:
    Select Case objMQQueue.ReasonCode
        Case MQRC_NO_MSG_AVAILABLE, MQRC_ALREADY_CONNECTED, _
             MQRC_ANOTHER_Q_MGR_CONNECTED
            objMQQueue.ClearErrorCodes
            Err.Clear
            Exit Sub
    End Select
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQEnd", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

Public Sub MQPutCache(ByVal psFila As String, _
                      ByRef psMensagem As String, _
                      ByRef plIndice As Long)

Dim objMQMessage                            As MQAX200.MQMessage
Dim objMQPut                                As MQAX200.MQPutMessageOptions
Dim lbBinario()                             As Byte
Dim liCount                                 As Integer
Dim lsReplyToQueueManagerName               As String
Dim lsMensagem                              As String
Dim b()                                     As Byte
Dim llQuantidade                            As Long
                
On Error GoTo ErrHandler
    
    Set objMQGet = objMQSession.AccessGetMessageOptions
    Set objMQMessage = objMQSession.AccessMessage
    objMQMessage.MessageId = Left(Format(plIndice, String(6, "0")) & psFila & Space(24), 24)
    If psFila = "PKQ.E.CACHE" & CacheOptions.Trigger Then
        Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_INPUT_EXCLUSIVE)
        objMQGet.Options = MQGMO_NO_SYNCPOINT + MQGMO_WAIT
        objMQGet.WaitInterval = 200
        objMQQueue.Get objMQMessage, objMQGet
        If psMensagem = "+1" Or psMensagem = "-1" Then
            objMQMessage.DataOffset = 0
            b = objMQMessage.Read(objMQMessage.MessageLength)
            lsMensagem = StrConv(b, vbUnicode)
            If objMQMessage.MessageLength > 170 Then
                llQuantidade = CLng(Mid(lsMensagem, 170, 6))
                If psMensagem = "+1" Then
                    llQuantidade = llQuantidade + 1
                    Mid(lsMensagem, 170, 6) = Format(llQuantidade, "000000")
                Else
                    llQuantidade = llQuantidade - 1
                    If llQuantidade < 0 Then
                        llQuantidade = 0
                    End If
                    Mid(lsMensagem, 170, 6) = Format(llQuantidade, "000000")
                End If
                psMensagem = lsMensagem
            End If
        End If
    Else
        Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_INPUT_SHARED)
        objMQGet.Options = MQGMO_NO_SYNCPOINT
        objMQQueue.Get objMQMessage, objMQGet
    End If
    objMQQueue.Close
    Set objMQQueue = Nothing
                
    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_OUTPUT)
    Set objMQMessage = objMQSession.AccessMessage
    Set objMQPut = objMQSession.AccessPutMessageOptions
        
    objMQMessage.Persistence = MQPER_NOT_PERSISTENT
    lbBinario = StrConv(psMensagem, vbFromUnicode)
    If psFila = "PKQ.E.CACHE" & CacheOptions.Trigger Then
        objMQMessage.Expiry = 6000
    Else
        objMQMessage.Expiry = 600000
    End If
    objMQMessage.Write lbBinario
    objMQPut.Options = MQPMO_NO_SYNCPOINT
    
    objMQMessage.MessageId = Left(Format(plIndice, String(6, "0")) & psFila & Space(24), 24)
    objMQQueue.Put objMQMessage, objMQPut
    Exit Sub
    
ErrHandler:
    If objMQQueue.ReasonCode = MQRC_NO_MSG_AVAILABLE Then
        objMQQueue.ClearErrorCodes
        Err.Clear
        Resume Next
    End If
    Err.Clear
End Sub

Public Sub MQGetCache(ByVal psFila As String, _
                      ByRef psMensagem As String, _
                      ByRef plIndice As Long)

On Error GoTo ErrHandler

Dim objMQMessage                            As MQAX200.MQMessage
Dim lbBinario()                             As Byte
Dim lbFim                                   As Boolean
    
    psMensagem = vbNullString
    
    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_BROWSE)
    Set objMQGet = objMQSession.AccessGetMessageOptions
    
    objMQGet.Options = MQGMO_BROWSE_FIRST
    Set objMQMessage = objMQSession.AccessMessage
    objMQMessage.MessageId = Left(Format(plIndice, String(6, "0")) & psFila & Space(24), 24)
    objMQQueue.Get objMQMessage, objMQGet
    If Not lbFim Then
        If Day(objMQMessage.PutDateTime) <> Day(Now + CDate("3:0:0")) Then
            psMensagem = vbNullString
            Exit Sub
        End If
    End If
    objMQMessage.DataOffset = 0
    
    'Le mensagem em binario
    If objMQMessage.MessageLength <> 0 Then
        lbBinario = objMQMessage.Read(objMQMessage.MessageLength)
        psMensagem = StrConv(lbBinario, vbUnicode)
    Else
        psMensagem = vbNullString
    End If
    
    Exit Sub
    
ErrHandler:
    Err.Clear
End Sub

Public Sub MQGetLimpaCache(ByVal psFila As String)

On Error GoTo ErrHandler

Dim objMQMessage                            As MQAX200.MQMessage
Dim lbBinario()                             As Byte
    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_INPUT_SHARED)
    Set objMQGet = objMQSession.AccessGetMessageOptions
    
    objMQGet.Options = MQGMO_NO_SYNCPOINT + MQGMO_NO_WAIT
    Do
        Set objMQMessage = objMQSession.AccessMessage
        objMQQueue.Get objMQMessage, objMQGet
        Set objMQMessage = Nothing
    Loop

ErrHandler:
    Err.Clear
End Sub

Public Function ConsultaCache(ByVal peOptions As CacheOptions, _
                              ByVal plIndice As Long) As String
On Error GoTo ErrHandler
    Dim lsResp                              As String
    
    MQGetCache "PKQ.E.CACHE" & peOptions, lsResp, plIndice
    ConsultaCache = lsResp

    Exit Function
ErrHandler:
    'Ignora o erro
    Err.Clear
End Function

Public Sub GravaCache(ByVal peOptions As CacheOptions, _
                      ByVal plIndice As Long, _
                      ByVal psConteudo As String)
On Error GoTo ErrHandler
    MQPutCache "PKQ.E.CACHE" & peOptions, psConteudo, plIndice

    Exit Sub
ErrHandler:
    'Ignora o erro
    Err.Clear
End Sub

Public Sub LimpaCache(ByVal peOptions As CacheOptions)
On Error GoTo ErrHandler
    MQGetLimpaCache "PKQ.E.CACHE" & peOptions
    Exit Sub
ErrHandler:
    'Ignora o erro
    Err.Clear
End Sub

Public Function QuantidadeMensagens(ByVal psNomeFila As String) As Long
On Error GoTo ErrHandler

    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(gsQMgrName)
    Set objMQQueue = objMQQueueManager.AccessQueue(psNomeFila, MQOO_INQUIRE + MQQT_LOCAL)

    QuantidadeMensagens = objMQQueue.CurrentDepth
    objMQQueue.Close

    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing

    Exit Function
ErrHandler:
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QuantidadeMensagemFila", llCodigoErroNegocio, liNumeroSequencialErro, "Nome da Fila: " & psNomeFila)
End Function

Public Sub LeMensagens(ByVal psFila As String, _
                       ByRef psMensagens() As String, _
              Optional ByVal psCorrelationID As String = vbNullString)
    Dim lngReasonCode                       As Long
    Dim strCorrelID                         As String
    Dim b()                                 As Byte
    Dim liI                                 As Integer
    
    On Error GoTo ErrHandler
    Erase psMensagens
    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(gsQMgrName)

    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_BROWSE)
    
    Do
        ReDim Preserve psMensagens(liI)
        Set objMQGet = objMQSession.AccessGetMessageOptions
        objMQGet.Options = MQGMO_BROWSE_NEXT
        Set objMQMessage = objMQSession.AccessMessage
        If Len(psCorrelationID) > 0 Then
            objMQMessage.CorrelationId = psCorrelationID
        End If
        objMQQueue.Get objMQMessage, objMQGet
        b = objMQMessage.Read(objMQMessage.MessageLength)
        If objMQMessage.MessageLength > 0 Then
            psMensagens(liI) = StrConv(b, vbUnicode)
        Else
            psMensagens(liI) = vbNullString
        End If
        liI = liI + 1
    Loop
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    Exit Sub
ErrHandler:
    
    Select Case objMQQueue.ReasonCode
    
        Case MQRC_NO_MSG_AVAILABLE, MQRC_ALREADY_CONNECTED, _
             MQRC_ANOTHER_Q_MGR_CONNECTED
            
            objMQQueue.ClearErrorCodes
            Err.Clear
                        
            Exit Sub
            
    End Select
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QuantidadeMensagemFila", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

Public Sub ObtemCorrelID(ByVal psFila As String, psCorrelID() As String)
    Dim lngReasonCode                       As Long
    Dim strCorrelID                         As String
    Dim liI                                 As Integer
    
    On Error GoTo ErrHandler
    Erase psCorrelID
    Set objMQSession = CreateObject("MQAX200.MQSession")
    Set objMQQueueManager = objMQSession.AccessQueueManager(gsQMgrName)

    Set objMQQueue = objMQQueueManager.AccessQueue(psFila, MQOO_BROWSE)
    
    Do
        Set objMQGet = objMQSession.AccessGetMessageOptions
        objMQGet.Options = MQGMO_BROWSE_NEXT
        Set objMQMessage = objMQSession.AccessMessage
        objMQQueue.Get objMQMessage, objMQGet
        If Len(objMQMessage.CorrelationId) > 0 Then
            If InStr(1, Join(psCorrelID), objMQMessage.CorrelationId) = 0 Then
                ReDim Preserve psCorrelID(liI)
                psCorrelID(liI) = objMQMessage.CorrelationId
                liI = liI + 1
            End If
        End If
    Loop
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    Exit Sub
ErrHandler:
    
    Select Case objMQQueue.ReasonCode
    
        Case MQRC_NO_MSG_AVAILABLE, MQRC_ALREADY_CONNECTED, _
             MQRC_ANOTHER_Q_MGR_CONNECTED
            
            objMQQueue.ClearErrorCodes
            Err.Clear
                        
            Exit Sub
            
    End Select
    
    Set objMQQueue = Nothing
    Set objMQQueueManager = Nothing
    Set objMQSession = Nothing
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObtemCorrelID", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub


Public Sub MQOpen(ByVal psQName As String, ByVal MQOO_Open As enumMQOO_Open)
On Error GoTo ErrHandler
    
    Set objMQQueue = objMQQueueManager.AccessQueue(psQName, MQOO_Open)
    
    Exit Sub
    
ErrHandler:
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQOpen", llCodigoErroNegocio, liNumeroSequencialErro, psQName)
End Sub

Public Sub MQClose()
On Error GoTo ErrHandler
             
    If Not objMQQueue Is Nothing Then
        objMQQueue.Close
    End If
    
    Exit Sub
    
ErrHandler:
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "MQClose", llCodigoErroNegocio, liNumeroSequencialErro)
End Sub

