VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcesso_A7"
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
'Chamada as rotinas de gerenciamento de recebimento de mensagens MQ Series
'-----------------------------------------------------------------------

Public Function ReceberMensagemMQ(ByVal pstrParam As String, ByVal strDataHoraUltiExecucao As String) As String


Dim objGerenciador                          As Object
Dim blnTemMensagemFila                      As Boolean
Dim strLogErro                              As String
Dim strMessageIdHex                         As String
Dim strCorrelationID                        As String
Dim strHistLogErro                          As String
Dim lngBackOutCount                         As Long
Dim strErrQualificado                       As String
Dim blnErroQualificado                      As Boolean
Dim lngContadorMensagem                     As Long
Dim xmlParametro                            As MSXML2.DOMDocument40

Dim strNomeFila                             As String
Dim strNomeFilaErro                         As String
Dim strNomeObjeto                           As String
Dim strProtocolo                            As String


On Error GoTo ErrHandler

    blnTemMensagemFila = True
    strHistLogErro = vbNullString

    If Not fgVerificarBancoDados Then
        blnTemMensagemFila = False
    Else
        If Not fgVerificaMQSeries("QM.SLCC.01") Then
            blnTemMensagemFila = False
        End If
    End If
    
    If blnTemMensagemFila Then
        Set xmlParametro = CreateObject("MSXML2.DOMDocument.4.0")
        xmlParametro.loadXML pstrParam
        
        If Not xmlParametro.selectSingleNode("//Protocolo") Is Nothing Then
            strProtocolo = Trim(xmlParametro.selectSingleNode("//Protocolo").Text)
        Else
            strProtocolo = vbNullString
        End If
        
        strNomeFila = Trim(xmlParametro.selectSingleNode("//Fila").Text)
        strNomeFilaErro = Trim(xmlParametro.selectSingleNode("//FilaErro").Text)
        strNomeObjeto = Trim(xmlParametro.selectSingleNode("//NomeObjeto").Text)
            
        Set xmlParametro = Nothing
    End If
    
    
    Do While blnTemMensagemFila

        lngContadorMensagem = lngContadorMensagem + 1

        strLogErro = vbNullString
        strMessageIdHex = vbNullString
        strCorrelationID = vbNullString
        strErrQualificado = vbNullString
        lngBackOutCount = 0
        blnErroQualificado = True

        Set objGerenciador = CreateObject(strNomeObjeto)

        If strProtocolo <> vbNullString Then

            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(strNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount, _
                                                                  strCorrelationID, _
                                                                  strProtocolo)
        Else

            blnTemMensagemFila = objGerenciador.ReceberMensagemMQ(strNomeFila, _
                                                                  strLogErro, _
                                                                  strMessageIdHex, _
                                                                  lngBackOutCount, _
                                                                  strCorrelationID)
        End If

        Set objGerenciador = Nothing

        strHistLogErro = strHistLogErro & strLogErro

        If Trim(strLogErro) <> vbNullString Then

            'Verificar se o erro é conhecido
            strErrQualificado = fgGetError("QM.SLCC.01", strNomeFila, strLogErro)

            If strErrQualificado <> "" Then

                 fgGravaArquivo strNomeFila, _
                                strErrQualificado & vbCrLf & _
                                String(50, "*") & vbCrLf & _
                                strLogErro

                blnErroQualificado = True
            Else
                If lngBackOutCount > 2 Then

                 fgGravaArquivo strNomeFila, _
                                strErrQualificado & vbCrLf & _
                                String(50, "*") & vbCrLf & _
                                strLogErro


                    'Put Fila Erro
                    Call PutFilaErro(strNomeFila, _
                                     strNomeFilaErro, _
                                     strMessageIdHex, _
                                     strLogErro, _
                                     strCorrelationID)
                End If
            End If

            strLogErro = vbNullString
        End If

        If lngContadorMensagem = 20 Then
            blnTemMensagemFila = False
            lngContadorMensagem = 0
        End If
    Loop

    ReceberMensagemMQ = ""


    Exit Function

ErrHandler:

    strLogErro = Err.Description

    Set objGerenciador = Nothing
    Set xmlParametro = Nothing

    fgGravaArquivo strNomeFila, strLogErro

End Function



