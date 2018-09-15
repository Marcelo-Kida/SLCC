VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterfaceSubscriber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsInterfaceSubscriber
'Data Criação   : 01/07/2003
'Objetivo       : Controle Transacional de acesso ao MQSeries/Oracle, distribuição
'                 das Mensagens aos objetos responsáveis pelo tratamento.
'
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 01/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

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

Public Function LiberarAgendamento(ByVal pstrNomeObjeto As String) As String

Dim objGerenciador                          As A7Server.clsGerenciadorAgendamento
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


Public Function ReceberMensagemMQ(ByVal pstrNomeFila As String, _
                                  ByVal pstrNomeObjeto As String, _
                                  ByVal pstrNomeObjetoApoio As String, _
                                  ByVal pstrNomeFilaErro As String, _
                         Optional ByVal pstrProtocolo As String = "") As String


Dim objFilaErro                             As A7Server.clsFilaErro
Dim objGerenciador                          As Object
Dim objApoio                                As Object
Dim blnTemMensagemFila                      As Boolean
Dim strLogErro                              As String
Dim strMessageIdHex                         As String
Dim strHistLogErro                          As String
Dim lngBackOutCount                         As Long
Dim strErrQualificado                       As String
Dim blnErroQualificado                      As Boolean

On Error GoTo ErrHandler
        
    blnTemMensagemFila = True
        
    If Trim(pstrNomeObjetoApoio) <> vbNullString Then
        Set objApoio = CreateObject(pstrNomeObjetoApoio)
    End If
        
    Do While blnTemMensagemFila
        
        strLogErro = vbNullString
        strMessageIdHex = vbNullString
        
        Set objGerenciador = CreateObject(pstrNomeObjeto)
             
        If Trim(pstrProtocolo) <> vbNullString Then
            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(pstrNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount, _
                                                                  pstrProtocolo)
        Else
            
            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(pstrNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount)
        End If
        
        Set objGerenciador = Nothing
        
        strHistLogErro = strHistLogErro & strLogErro
        
        If Trim(strLogErro) <> vbNullString Then
                                     
            If lngBackOutCount > 2 Then
                If Not blnErroQualificado Then
                    'Put Fila Erro
                    Set objFilaErro = CreateObject("A7Server.clsFilaErro")
                    Call objFilaErro.PutFilaErro(pstrNomeFila, pstrNomeFilaErro, strMessageIdHex, strLogErro)
                    Set objFilaErro = Nothing
                    
                    flGravaArquivo pstrNomeFila, strLogErro
                    
                End If
            End If
            
            strErrQualificado = fgGetError(strLogErro)
            
            If strErrQualificado <> "" Then
               'Para Erros MQSeries E Oracle
                Call flDesativaVerificaServer
                flGravaArquivo pstrNomeFila, strErrQualificado & vbCrLf & "A6A7A8 Verifica Server parado."
                blnErroQualificado = True
                Exit Function
            Else
                blnErroQualificado = False
            End If
            
            strLogErro = vbNullString
        End If
    Loop

    ReceberMensagemMQ = strHistLogErro
    
    Set objApoio = Nothing
    
    Exit Function

ErrHandler:
    
    strLogErro = Err.Description
    
    Set objGerenciador = Nothing
    
    flGravaArquivo "InterfaceSubscriber1_" & pstrNomeFila, strLogErro
    
End Function

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
Private Function flDesativaVerificaServer() As Boolean

On Error GoTo ErrorHandler
    
    Shell "NET STOP A7VerificaServer", vbHide
    
    Exit Function
ErrorHandler:

    Shell "NET STOP A7VerificaServer", vbHide
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

