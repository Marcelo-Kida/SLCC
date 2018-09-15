VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExecucaoBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela exibição dos logs gerados pelos processos batch dos sistemas A6, A7 e A8. (Replicação, Expurgo, IntegraçãoCC)
Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

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

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Ler
Public Function Executar(ByVal pXMLFiltro As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pXMLFiltro.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Ler"
            Executar = Ler(pXMLFiltro)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Obter os logs gerados segundo os parâmetros de filtro informados.
Private Function Ler(ByRef pXMLFiltro As MSXML2.DOMDocument40) As String
On Error GoTo ErrorHandler
Dim strSQL                                  As String
Dim xmlFiltro                               As MSXML2.DOMDocument40
Dim blnBaseHistorica                        As Boolean
Dim blnBaseCorrente                         As Boolean
    
    If fgDtXML_To_Date(pXMLFiltro.selectSingleNode("//DT_INIC").Text) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseHistorica = True
    End If
    
    If fgDtXML_To_Date(pXMLFiltro.selectSingleNode("//DT_FIM").Text) >= (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        blnBaseCorrente = True
    End If
    
    strSQL = "SELECT CO_ROTI_BATCH, " & vbCrLf & _
                    "DH_FIM_EXEC, " & vbCrLf & _
                    "IN_EXEC_SUCE," & vbCrLf & _
                    "DE_ERRO_EXEC " & vbCrLf & _
               "FROM A7.TB_HIST_EXEC_ROTI_BATCH " & _
             "WHERE DH_FIM_EXEC BETWEEN " & _
             fgDtHrXML_To_Oracle(pXMLFiltro.selectSingleNode("//DT_INIC").Text & "000000") & " AND " & _
             fgDtHrXML_To_Oracle(pXMLFiltro.selectSingleNode("//DT_FIM").Text & "235959")
    
    If Not pXMLFiltro.selectSingleNode("//CO_ROTI_BATCH") Is Nothing Then
        strSQL = strSQL & vbCrLf & _
                " AND CO_ROTI_BATCH = '" & pXMLFiltro.selectSingleNode("//CO_ROTI_BATCH").Text & "'"
    End If
             
    If Not pXMLFiltro.selectSingleNode("//IN_EXEC_SUCE") Is Nothing Then
        strSQL = strSQL & vbCrLf & _
                " AND IN_EXEC_SUCE = " & pXMLFiltro.selectSingleNode("//IN_EXEC_SUCE").Text
    End If
             
    If blnBaseCorrente Then
        If blnBaseHistorica Then
            strSQL = strSQL & vbCrLf & _
                     " UNION ALL " & vbCrLf & _
                     Replace(strSQL, "A7.TB_HIST_EXEC_ROTI_BATCH", "A7HIST.TB_HIST_EXEC_ROTI_BATCH")
        End If
    Else
        strSQL = Replace(strSQL, "A7.TB_HIST_EXEC_ROTI_BATCH", "A7HIST.TB_HIST_EXEC_ROTI_BATCH")
    End If
             
    Ler = fgQueryXMLLerTodos("Execucao", strSQL, "clsExecucaoBatch")
             
    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
