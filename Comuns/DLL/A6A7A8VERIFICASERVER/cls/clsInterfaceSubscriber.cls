VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterfaceSubsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Controle Transacional de acesso ao MQSeries/Oracle, distribuição das Mensagens aos objetos responsáveis pelo tratamento.

Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Implements COMSVCSLib.ObjectControl

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

'-----------------------------------------------------------------------
'Chamada a rotina de liberação de mensagens agendadas
'-----------------------------------------------------------------------

Public Function LiberarAgendamento(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As Object 'A7Server.clsGerenciadorAgendamento
Dim strRetorno                              As String
Dim strLogErro                              As String

    On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A7Server.clsGerenciadorAgendamento")
    objGerenciador.LiberarAgendamento pstrNomeObjeto
    Set objGerenciador = Nothing
    
    LiberarAgendamento = vbNullString
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "InterfaceSubscriberAgendamento", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de liquidação com o sistema TQ
'-----------------------------------------------------------------------

Public Function LiquidarOperacaoTQ(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As Object 'A7Server.clsGerenciadorAgendamento
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A8LQS.clsLiquidacaoTQ")
    strRetorno = objGerenciador.LiquidarOperacoesTQ
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        flGravaArquivo "LiquidacaoTQ", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "LiquidacaoTQ", strLogErro
    
End Function


'-----------------------------------------------------------------------
'Chamada as rotinas de gerenciamento de recebimento de mensagens MQ Series
'-----------------------------------------------------------------------

Public Function ReceberMensagemMQ(ByVal pstrNomeFila As String, _
                                  ByVal pstrNomeObjeto As String, _
                                  ByVal pstrNomeObjetoApoio As String, _
                                  ByVal pstrNomeFilaErro As String, _
                         Optional ByVal pstrProtocolo As String = "") As String


Dim objFilaErro                             As Object 'A7Server.clsFilaErro
Dim objGerenciador                          As Object
Dim objApoio                                As Object
Dim blnTemMensagemFila                      As Boolean
Dim strLogErro                              As String
Dim strMessageIdHex                         As String
Dim strCorrelationID                        As String
Dim strHistLogErro                          As String
Dim lngBackOutCount                         As Long
Dim strErrQualificado                       As String
Dim blnErroQualificado                      As Boolean
Dim lngContadorMensagem                     As Long

On Error GoTo ErrHandler
        
    blnTemMensagemFila = True
        
    If Trim(pstrNomeObjetoApoio) <> vbNullString Then
        Set objApoio = CreateObject(pstrNomeObjetoApoio)
    End If
        
    Do While blnTemMensagemFila
                
        lngContadorMensagem = lngContadorMensagem + 1
                
        strLogErro = vbNullString
        strMessageIdHex = vbNullString
        strCorrelationID = vbNullString
        strErrQualificado = vbNullString
        lngBackOutCount = 0
        blnErroQualificado = True
        
        Set objGerenciador = CreateObject(pstrNomeObjeto)
             
        If Trim(pstrProtocolo) <> vbNullString Then
            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(pstrNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount, _
                                                                  strCorrelationID, _
                                                                  pstrProtocolo)
        Else
            
            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(pstrNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount, _
                                                                  strCorrelationID)
        End If
        
        Set objGerenciador = Nothing
        
        strHistLogErro = strHistLogErro & strLogErro
        
        If Trim(strLogErro) <> vbNullString Then
            'Verificar se o erro é conhecido
            strErrQualificado = fgGetError(strLogErro)
            
            If strErrQualificado <> "" Then
                 flGravaArquivo pstrNomeFila, _
                               strErrQualificado & vbCrLf & _
                               String(50, "*") & vbCrLf & _
                               strLogErro
                blnErroQualificado = True
            Else
                If lngBackOutCount > 2 Then
                    
                    flGravaArquivo pstrNomeFila, strLogErro
                    
                    'Put Fila Erro
                    Call PutFilaErro(pstrNomeFila, _
                                     pstrNomeFilaErro, _
                                     strMessageIdHex, _
                                     strLogErro, _
                                     strCorrelationID)
                End If
            End If
            
            strLogErro = vbNullString
        End If
        
        If lngContadorMensagem = 30 Then
            blnTemMensagemFila = False
        End If
        
    Loop

    ReceberMensagemMQ = strHistLogErro
    
    Set objApoio = Nothing
    
    Exit Function

ErrHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "InterfaceSubscriber_" & pstrNomeFila, strLogErro
    
End Function

'-----------------------------------------------------------------------
'Gravar arquivo de log de erro em disco
'-----------------------------------------------------------------------

Private Sub flGravaArquivo(ByVal pstrNomeFilaArquivo As String, _
                           ByVal pstrErro As String)

Dim strNomeArquivoLogErro                    As String
Dim intFile                                  As Integer
Dim strTemp                                  As String

On Error GoTo ErrorHandler
    
    strNomeArquivoLogErro = App.Path & "/log/" & pstrNomeFilaArquivo & "_" & Format(Now, "yyyymmddHHmmss") & ".log"
    
    strTemp = String(50, "*") & vbCrLf
    strTemp = strTemp & pstrErro & vbCrLf
    strTemp = strTemp & String(50, "*") & vbCrLf
    
    intFile = FreeFile
    Open strNomeArquivoLogErro For Output As intFile
    Print #intFile, strTemp
    Close intFile
  
    Exit Sub

ErrorHandler:
    
    Close intFile
        
    Err.Clear
        
End Sub

'-----------------------------------------------------------------------
'Put de mensagens de erro nas filas de erro
'-----------------------------------------------------------------------

Public Function PutFilaErro(ByVal pstrNomeFilaGet As String, _
                            ByVal pstrNomeFilaErro As String, _
                            ByVal pstrMessageIdHex As String, _
                            ByRef pstrLogErro As String, _
                            ByRef pstrCorrelationID As String) As Boolean
                         
Dim objMQAX200                              As Object  'A7Server.clsFilaErro
                         
On Error GoTo ErrorHandler
        
    Set objMQAX200 = CreateObject("A7Server.clsFilaErro")
    
    Call objMQAX200.PutFilaErro(pstrNomeFilaGet, _
                                pstrNomeFilaErro, _
                                pstrMessageIdHex, _
                                pstrLogErro, _
                                pstrCorrelationID)
    Set objMQAX200 = Nothing
    
    Exit Function

ErrorHandler:
   
    Set objMQAX200 = Nothing
    
    Err.Clear
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de Processamento Mensagem STR0008R2 - Conti
'-----------------------------------------------------------------------

Public Function ProcessaMensagemSTR0008R2Comti(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As Object 'A8LQS.clsProcessoMensagemSTR
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A8LQS.clsProcessoMensagemSTR")
    strRetorno = objGerenciador.VerificaMensagemSTR0008R2Comti()
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        flGravaArquivo "ProcessaMensagemSTR0008R2Comti", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "ProcessaMensagemSTR0008R2Comti", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de Envio e-mail Relatorio Usuários Inativos - SLCC
'-----------------------------------------------------------------------

Public Function VerificaUsuariosInativos() As String

Dim objGerenciador                          As Object 'A7Server.clsControleUsuarioSistema
Dim strRetorno                              As String
Dim strLogErro                              As String
  
On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject("A7Server.clsControleUsuarioSistema")
    strRetorno = objGerenciador.VerificaUsuarioInativo
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        flGravaArquivo "VerificaUsuariosInativos", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "VerificaUsuariosInativos", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de liquidação futura D+1 (CETIP)
'-----------------------------------------------------------------------

Public Function GerenciarFluxoOperacoesLiquidacaoFutura(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As Object 'A8LQS.CLSLIQUIDACAOFUTURA
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject(pstrNomeObjeto)
    strRetorno = objGerenciador.GerenciarFluxoOperacoesLiquidacaoFutura
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        flGravaArquivo "GerenciarFluxoOperacoesLiquidacaoFutura", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "GerenciarFluxoOperacoesLiquidacaoFutura", strLogErro
    
End Function

'-----------------------------------------------------------------------
'Chamada a rotina de conciliacao de Roda de Dolar
'-----------------------------------------------------------------------

Public Function GerenciarFluxoConciliacaoRodaDolar(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As Object
Dim strRetorno                              As String
Dim strLogErro                              As String

On Error GoTo ErrorHandler
    
    Set objGerenciador = CreateObject(pstrNomeObjeto)
    strRetorno = objGerenciador.GerenciarFluxoConciliacaoRodaDolar
    Set objGerenciador = Nothing
    
    If strRetorno <> vbNullString Then
        flGravaArquivo "GerenciarFluxoConciliacaoRodaDolar", strRetorno
    End If
    
    Exit Function
    
ErrorHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "GerenciarFluxoConciliacaoRodaDolar", strLogErro
    
End Function
