VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSaldoFechSubReserva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela consulta e atualização
' do Cadastro de Saldo de Fechamento de Veículos Legais.

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements ObjectControl
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

' Este método é responsável pela atualização geral da tabela de saldo de fechamento,
' e pelo retorno à camada intermediária.

Public Function Salvar(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean
On Error GoTo ErrorHandler

    flIncluir pxmlRemessa

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Este método é responsável pela inclusão de um novo saldo de fechamento para determinado veículo legal.

Private Function flIncluir(ByVal pxmlRemessa As MSXML2.DOMDocument40) As Boolean
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    With pxmlRemessa.documentElement

    strSQL = "INSERT INTO " & vbCrLf & _
             "   A6.TB_SALD_FECH_VEIC_LEGA ( " & vbCrLf & _
             "   CO_VEIC_LEGA, " & vbCrLf & _
             "   SG_SIST, " & vbCrLf & _
             "   DT_APUR_SALD, " & vbCrLf & _
             "   DH_ENVI_SALD, " & vbCrLf & _
             "   VA_SALD_FECH, " & vbCrLf & _
             "   DH_ULTI_ATLZ " & vbCrLf & _
             ") VALUES ( " & vbCrLf & _
             "   '" & .selectSingleNode("CO_VEIC_LEGA").Text & "', " & vbCrLf & _
             "   '" & .selectSingleNode("SG_SIST_ORIG").Text & "', " & vbCrLf & _
                fgDtXML_To_Oracle(.selectSingleNode("DT_APUR_SALD").Text) & ", " & vbCrLf & _
                fgDtHrXML_To_Oracle(.selectSingleNode("DT_REME").Text & .selectSingleNode("HO_REME").Text & "00") & ", " & vbCrLf & _
             "   " & fgVlrToDBServer(.selectSingleNode("VA_SALD_FECH").Text) & ", " & vbCrLf & _
             "   SYSDATE) " & vbCrLf
            
    End With

    fgExecuteSQL strSQL

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
