VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGerenciadorRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nPacote         :\r\nClasse         : clsRemessa\r\nData Criação   : 14/07/2003\r\nObjetivo       :\r\n\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14/07/2003\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F105AB800C3"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"

' Este componente tem como objetivo, agrupar os métodos responsáveis pelo
' recebimento e direcionamento de mensagens MQ.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
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

'Obter Remessa Movimentação Item Caixa utilizando a classe MQSeries

Public Function ControlarRemessa(ByVal pstrQName As String) As String

On Error GoTo ErrHandler

Dim objRemessa                              As A6SubReserva.clsRemessa
Dim objValidaRemessa                        As A6A8ValidaRemessa.clsValidaRemessa

Dim blnReturnCode                           As Boolean
Dim strLogErro                              As String
Dim strMessageIDHex                         As String
Dim strHistLogErro                          As String

Dim lngBackOutCount                         As Long

    'Ativa o ActiveX EXE - Validador da Remessa
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")

    Do While Not blnReturnCode
        Set objRemessa = CreateObject("A6SubReserva.clsRemessa")
        strMessageIDHex = ""
        strLogErro = ""
        lngBackOutCount = 0
        
        blnReturnCode = objRemessa.ReceberMensagemMQ(pstrQName, strLogErro, strMessageIDHex, lngBackOutCount, "")
        
        Set objRemessa = Nothing

        strHistLogErro = strHistLogErro & strLogErro

        If Trim(strLogErro) <> "" Then
            'TNS:No Listener
            If InStr(1, strLogErro, "ORA-12541") Then
                'Encerrar aplicação - sem gravar fila de erro
                GoTo ErrHandler
            End If
            strLogErro = ""
            If lngBackOutCount > 2 Then
                Set objRemessa = CreateObject("A6SubReserva.clsRemessa")
                'objRemessa.PutFilaErro pstrQName, strLogErro, strMessageIDHex
                Set objRemessa = Nothing
            End If
        ElseIf lngBackOutCount > 2 Then
            Set objRemessa = CreateObject("A6SubReserva.clsRemessa")
            'objRemessa.PutFilaErro pstrQName, strLogErro, strMessageIDHex
            Set objRemessa = Nothing
        End If
    Loop

    ControlarRemessa = strHistLogErro & strLogErro

    Set objValidaRemessa = Nothing

    Exit Function

ErrHandler:

    ControlarRemessa = strHistLogErro & strLogErro

    Set objRemessa = Nothing
    'Set objValidaRemessa = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError("A6SubReserva", TypeName(Me), "ControlarRemessa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

