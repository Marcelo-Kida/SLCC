VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 3  'UsesTransaction
END
Attribute VB_Name = "clsLogTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Recuperar registros referentes a Trilha de Auditoria

Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

'recuperar registros da tabela de controle de Log <OWNER>.TB_CTRL_TABE_LOG

Public Function LerTodosControleLog(ByVal pintOwner As Integer) As String
   
Dim objLogTabela                            As A6A7A8.clsLogTabela
    
On Error GoTo ErrorHandler
        
    Set objLogTabela = CreateObject("A6A7A8.clsLogTabela")
    LerTodosControleLog = objLogTabela.LerTodosControleLog(pintOwner)
    Set objLogTabela = Nothing
    
    Exit Function

ErrorHandler:

    Set objLogTabela = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosControleLog Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter os nome das colunas de uma tabela

Public Function ObterColunasTabela(ByVal pstrNomeTabela As String) As String

Dim objLogTabela                            As A6A7A8.clsLogTabela
    
On Error GoTo ErrorHandler
        
    Set objLogTabela = CreateObject("A6A7A8.clsLogTabela")
    ObterColunasTabela = objLogTabela.ObterColunasTabela(pstrNomeTabela)
    Set objLogTabela = Nothing

    Exit Function
ErrorHandler:
    
    Set objLogTabela = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterColunasTabela", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'recuperar os registros das tabelas de Log

Public Function LerTodosLogTabela(ByVal pstrNomeTabelaLog As String, _
                                  ByVal pintOwner As Integer, _
                                  ByVal pdtmDataInicio As String, _
                                  ByVal pdtmDataFim As String) As String

Dim objLogTabela                            As A6A7A8.clsLogTabela
Dim datDataInicio                           As Date
Dim datDataFim                              As Date
    
On Error GoTo ErrorHandler

    If Trim(pdtmDataInicio) <> "" Then
       datDataInicio = CDate(Format(pdtmDataInicio, "DD/MM/YYYY"))
    End If
    If Trim(pdtmDataFim) <> "" Then
       datDataFim = CDate(Format(pdtmDataFim, "DD/MM/YYYY"))
    End If
        
    Set objLogTabela = CreateObject("A6A7A8.clsLogTabela")
    LerTodosLogTabela = objLogTabela.LerTodosLogTabela(pstrNomeTabelaLog, _
                                                       pintOwner, _
                                                       datDataInicio, _
                                                       datDataFim)
    Set objLogTabela = Nothing

    Exit Function
ErrorHandler:
    
    Set objLogTabela = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosLogTabela", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function





