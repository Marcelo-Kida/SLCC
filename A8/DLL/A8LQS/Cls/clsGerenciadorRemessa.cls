VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Description = "Componente     : \r\nClasse         : \r\nData Criação   : \r\nObjetivo       : \r\n\r\nAnalista       : \r\n\r\nProgramador    :\r\nData           :\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EF762CB0027"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsGerenciadorRemessa
'Data Criação   : 14/07/2003
'Objetivo       :
'
'Analista       : Carlos Fortes
'
'Programador    : Carlos Fortes
'Data           : 14/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :


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

Public Function ControlarRemessa(ByVal pstrQName As String) As String

'Obter Remessa Movimentação Item Caixa utilizando a classe MQSeries

On Error GoTo ErrHandler

Dim objRemessa                              As A8LQS.clsRemessa
Dim objValidaRemessa                        As A6A8ValidaRemessa.clsValidaRemessa

Dim blnReturnCode                           As Boolean
Dim strLogErro                              As String
Dim strMessageIDHex                         As String
Dim strHistLogErro                          As String

Dim lngBackOutCount                         As Long

    'Ativa o ActiveX EXE - Validador da Remessa
    Set objValidaRemessa = CreateObject("A6A8ValidaRemessa.clsValidaRemessa")

    Do While Not blnReturnCode
        Set objRemessa = CreateObject("A8LQS.clsRemessa")
        strMessageIDHex = ""
        strLogErro = ""
        lngBackOutCount = 0
        
        blnReturnCode = objRemessa.ReceberMensagemMQ(pstrQName, strLogErro, strMessageIDHex, lngBackOutCount)
        
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
                Set objRemessa = CreateObject("A8LQS.clsRemessa")
                'objRemessa.PutFilaErro pstrQName, strLogErro, strMessageIDHex
                Set objRemessa = Nothing
            End If
        ElseIf lngBackOutCount > 2 Then
            Set objRemessa = CreateObject("A8LQS.clsRemessa")
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
