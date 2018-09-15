VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsRemessa\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105AB800C3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo
' tratamento em geral de remessas enviadas ao A6.

Option Explicit

Public objValidaRemessa                      As Object 'A6A8ValidaRemessa.clsValidaRemessa

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

Implements ObjectControl
Private objContext                           As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

' Este método é responsável pelo recebimento de mensagens MQ.

Public Function ReceberMensagemMQ(ByVal pstrQName As String, _
                                  ByRef pstrLogErro As String, _
                                  ByRef pstrMessageIDHexOut As String, _
                                  ByRef plngBackOutCount As Long, _
                                  ByRef pstrCorrelationID As String) As Boolean

Dim objTRemessa                             As A6SubReserva.clsTRemessa
Dim xmlRemessa                              As MSXML2.DOMDocument40
Dim xmlLogErro                              As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strMessageDescriptor                    As String
Dim lngBufferLen                            As Long
Dim strBuffer                               As String
Dim strErros                                As String
Dim blnEnviaParaRemesssaRejeitada           As Boolean

On Error GoTo ErrorHandler

    Set xmlRemessa = CreateObject("MSXML2.DOMDocument.4.0")
    Set objTRemessa = CreateObject("A6SubReserva.clsTRemessa")
    
'    fgGravaArquivo "LOGA6_ReceberMensagemMQ_", vbNullString

    Set objTRemessa.objValidaRemessa = objValidaRemessa
    
    ReceberMensagemMQ = objTRemessa.ReceberMensagemMQ(pstrQName, _
                                                      pstrLogErro, _
                                                      pstrMessageIDHexOut, _
                                                      plngBackOutCount, _
                                                      pstrCorrelationID, _
                                                      xmlRemessa)


    Set objTRemessa = Nothing
        
        
    If pstrLogErro <> vbNullString Then
        
        Set xmlLogErro = CreateObject("MSXML2.DOMDocument.4.0")
        
        'Se Log erro OK
        If xmlLogErro.loadXML(pstrLogErro) Then
            If Not xmlLogErro.selectSingleNode("//Number") Is Nothing Then
                 If fgVerificaErroOracleMQSeries(pstrLogErro) Then
                    blnEnviaParaRemesssaRejeitada = False
                 Else
                    blnEnviaParaRemesssaRejeitada = True
                 End If
            Else
                blnEnviaParaRemesssaRejeitada = False
            End If
        Else
            blnEnviaParaRemesssaRejeitada = False
        End If
        
        If blnEnviaParaRemesssaRejeitada Then
            Set objTRemessa = CreateObject("A6SubReserva.clsTRemessa")
            Call objTRemessa.RemessaRejeitada(pstrQName, pstrLogErro, pstrMessageIDHexOut, plngBackOutCount, xmlRemessa)
            Set objTRemessa = Nothing
        End If
    End If

    Set xmlRemessa = Nothing

    Exit Function
ErrorHandler:
    
    pstrLogErro = Err.Description
    
    Set objTRemessa = Nothing
    Set xmlRemessa = Nothing

End Function



