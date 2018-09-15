VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMensagemRejeitada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela persistência das mensagens rejeitadas pelo sistema A7 e apresentação das rejeições na tela de Mensagens Rejeitadas do sistema A7.
Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

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

'Consultar todas as mensagens rejeitadas pelo sistema A7 de acordo com o período informado.
Public Function LerTodos(ByVal pstrDataHoraDe As String, _
                         ByVal pstrDataHoraAte As String) As String
    
Dim strSQL                                  As String
Dim strOwner                                As String
    
On Error GoTo ErrorHandler

    If fgDtXML_To_Date(Left$(pstrDataHoraDe, 8)) < (fgDataHoraServidor(enumFormatoDataHora.Data) - fgObterQtdDiasExpurgo()) Then
        strOwner = "A7HIST"
    Else
        strOwner = "A7"
    End If

    strSQL = "SELECT '" & strOwner & "' as OWNER, " & vbCrLf & _
             "        A.CO_MESG_MQSE,          " & vbCrLf & _
             "        A.DH_ENTR_FILA_MQSE,     " & vbCrLf & _
             "        A.DH_RECB_MESG,          " & vbCrLf & _
             "        A.NO_ARQU_ENTR_FILA_MQSE," & vbCrLf & _
             "        A.CO_TEXT_XML ,          " & vbCrLf & _
             "        A.TX_DTLH_OCOR_ERRO,     " & vbCrLf & _
             "        B.DE_OCOR_MESG           " & vbCrLf & _
             "   FROM " & strOwner & ".TB_MESG_REJE  A,      " & vbCrLf & _
             "        A7.TB_OCOR_MESG  B       " & vbCrLf & _
             "  WHERE A.CO_OCOR_MESG  = B.CO_OCOR_MESG" & vbCrLf & _
             "    AND A.DH_RECB_MESG BETWEEN          " & _
            fgDtHrXML_To_Oracle(pstrDataHoraDe) & " AND " & _
            fgDtHrXML_To_Oracle(pstrDataHoraAte) & _
             " ORDER BY A.DH_RECB_MESG DESC "
    
    LerTodos = fgQueryXMLLerTodos("MensagemRejeitada", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salvar a mensagem rejeitada para posterior avaliação.
Public Function SalvarMensagemRejeitada(ByVal pstrMensagemRecebida As String, _
                                        ByVal plngCodigoOcorrencia As Long, _
                                        ByVal pstrMesgIDHex As String, _
                                        ByVal pstrNomeAplicativo As String, _
                                        ByVal pdtmDataPut As Date, _
                                        ByVal pstrDetalheOcorrenciaErro As String) As Long
                                    
Dim lngCodCntdAtrbXml                        As Long
Dim strSQL                                   As String
                                    
On Error GoTo ErrorHandler
        
    lngCodCntdAtrbXml = fgInsertVarchar4000(pstrMensagemRecebida)
 
    strSQL = " INSERT INTO A7.TB_MESG_REJE ( " & _
            "  CO_MESG_MQSE,          " & _
            "  CO_OCOR_MESG,          " & _
            "  DH_RECB_MESG,          " & _
            "  CO_TEXT_XML,           " & _
            "  DH_ENTR_FILA_MQSE,     " & _
            "  NO_ARQU_ENTR_FILA_MQSE," & _
            "  TX_DTLH_OCOR_ERRO      " & _
            " ) VALUES (              " & _
            "'" & pstrMesgIDHex & "'," & _
            plngCodigoOcorrencia & " ," & _
            "SYSDATE," & _
            lngCodCntdAtrbXml & "," & _
            fgDtHrXML_To_Oracle(Format(pdtmDataPut, "YYYYMMDDHHmmss")) & "," & _
            "'" & pstrNomeAplicativo & "'," & _
            "'" & Replace(pstrDetalheOcorrenciaErro, "'", "") & "')"
    
    fgExecuteSQL (strSQL)
    
    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarMensagemRejeitada", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter a mensagem rejeitada pelo sistema A7.
Public Function ObterMensagem(ByVal plngCodCntdAtrbXml As Long) As String
    
    
On Error GoTo ErrorHandler
    
    ObterMensagem = fgSelectVarchar4000(plngCodCntdAtrbXml)
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagem Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


