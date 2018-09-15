VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 

Attribute VB_Name = "clsTipoConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável por controlar o acesso à tabela A8.TB_TIPO_CNTA_SELIC.
'' 
'' 

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private strConstruct                        As String

Private objContext                          As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
    
On Error Resume Next
    
    strConstruct = pobjContruct.ConstructString

End Sub

'Verica se a conta é uma conta de cliente 1


Public Function VerificarCliente1(ByVal pstrContaSelicContraparte As String) As Boolean

Dim strSQL                                  As String
Dim rsTipoConta                             As ADODB.Recordset
Dim lngTipoConta                            As Long

    On Error GoTo ErrorHandler

    lngTipoConta = CLng(Mid(Right(pstrContaSelicContraparte, 5), 1, 4))

    strSQL = " Select   CO_TIPO_CNTA_SELIC      " & vbNewLine & _
             " From     A8.TB_TIPO_CNTA_SELIC   " & vbNewLine & _
             " Where    CO_TIPO_CNTA_SELIC          =   " & lngTipoConta & vbNewLine & _
             " And      SUBSTR(TP_CNTA_SELIC,1,3)   =   'CL1' "

    Set rsTipoConta = fgQuerySQL(strSQL)

    If rsTipoConta.EOF Then
        VerificarCliente1 = False
    Else
        VerificarCliente1 = True
    End If

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarCliente1 Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
