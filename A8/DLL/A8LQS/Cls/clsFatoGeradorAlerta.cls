VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsFatoGeradorAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações relativas aos fatos
'' geradores de alerta cadastrados

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

'' Consultar todos os fatos geradores de alerta cadastrados para o processo de
'' verificação
Public Function LerTodos() As String

Dim strSQL                                  As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    CO_FATO_GERA_ALER,     " & vbNewLine & _
             "           DE_FATO_GERA_ALER      " & vbNewLine & _
             " FROM      A8.TB_FATO_GERA_ALER   " & vbNewLine & _
             " ORDER BY  CO_FATO_GERA_ALER ASC  "
    
    LerTodos = fgQueryXMLLerTodos("FatoGeradorAlerta", strSQL, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
