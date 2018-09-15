VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperacaoMensagemInterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas às mensagens
'' internas cadastradas para as operações, promovendo uma associação entre os
'' dados da operação e seu XML gerado, bem como a manipulação dos dados cadastrados

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

'' Promover a inclusão de uma nova mensagem interna referente a uma operação
Public Function Incluir(ByRef pvntSequenciaOperacao As Variant, _
                        ByRef plngTipoMensagem As Long, _
                        ByRef plngTipoSolicitacao As enumTipoSolicitacao, _
                        ByRef xmlMensagem As MSXML2.DOMDocument40, _
               Optional ByVal pintFormatoSaidaMensagem As Integer = 0) As Boolean

Dim strSQL                                  As String
Dim lngCodigoTxtXML                         As Long
Dim strDataRegistroMensagem                 As String

    On Error GoTo ErrorHandler

    lngCodigoTxtXML = fgInsertVarchar4000(xmlMensagem.xml)

    strDataRegistroMensagem = flObterDataGravacao(pvntSequenciaOperacao)

    'Verificar se o tamanho TX_MESG_INTE_XML
    strSQL = "Insert Into A8.TB_OPER_ATIV_MESG_INTE (" & vbCrLf & _
             "  NU_SEQU_OPER_ATIV, " & vbCrLf & _
             "  DH_MESG_INTE, " & vbCrLf & _
             "  TP_MESG_INTE, " & vbCrLf & _
             "  TP_SOLI_MESG_INTE, " & vbCrLf & _
             "  CO_TEXT_XML, " & vbCrLf & _
             "  TP_FORM_MESG_SAID) " & vbCrLf & _
             " Values ( " & vbCrLf & _
             pvntSequenciaOperacao & ", " & vbCrLf & _
             fgDtHrXML_To_Oracle(strDataRegistroMensagem) & ", " & vbCrLf & _
             plngTipoMensagem & ", " & vbCrLf & _
             plngTipoSolicitacao & ", " & vbCrLf & _
             lngCodigoTxtXML & ", " & vbCrLf & _
             pintFormatoSaidaMensagem & ")"

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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos as mensagens internas cadastradas para as operações
Public Function LerTodos(Optional ByVal pvntSequenciaOperacao As Variant)
    
Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = "Select    A.NU_SEQU_OPER_ATIV,                      " & vbNewLine & _
             "          A.DH_MESG_INTE,                           " & vbNewLine & _
             "          A.TP_MESG_INTE,                           " & vbNewLine & _
             "          B.NO_TIPO_MESG,                           " & vbNewLine & _
             "          A.TP_SOLI_MESG_INTE,                      " & vbNewLine & _
             "          A.CO_TEXT_XML                             " & vbNewLine & _
             " From     A8.TB_OPER_ATIV_MESG_INTE   A,            " & vbNewLine & _
             "          A7.TB_TIPO_MESG             B             " & vbNewLine & _
             " Where    A.TP_MESG_INTE      = B.TP_MESG(+)        " & vbNewLine & _
             " And     (A.TP_FORM_MESG_SAID = B.TP_FORM_MESG_SAID " & vbNewLine & _
             " Or       A.TP_FORM_MESG_SAID = 0) "
             
    If pvntSequenciaOperacao <> Empty Then
        If pvntSequenciaOperacao < 0 Then
            strSQL = Replace$(strSQL, "A8.TB_OPER_ATIV_MESG_INTE", "A8HIST.TB_OPER_ATIV_MESG_INTE")
            pvntSequenciaOperacao = pvntSequenciaOperacao * -1
        End If
             
        strSQL = strSQL & vbNewLine & _
             " And      A.NU_SEQU_OPER_ATIV = " & pvntSequenciaOperacao
    End If
    
    LerTodos = fgQueryXMLLerTodos("OperacaoMensagemInterna", strSQL, App.EXEName & "." & TypeName(Me))

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

'Obter a última data de gravação de uma mensagem interna
Private Function flObterDataGravacao(ByRef pvntSequenciaOperacao As Variant) As String

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset
Dim datDataGravacao                         As Date

    On Error GoTo ErrorHandler

    strSQL = " SELECT   MAX(DH_MESG_INTE) AS DH_MESG_INTE " & vbNewLine & _
             " FROM     A8.TB_OPER_ATIV_MESG_INTE     " & vbNewLine & _
             " WHERE    NU_SEQU_OPER_ATIV       =     " & pvntSequenciaOperacao

    Set objRS = fgQuerySQL(strSQL)

    If IsNull(objRS!DH_MESG_INTE) Then
        datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
    Else
        If DateDiff("s", fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux), objRS!DH_MESG_INTE) = 0 Then
            datDataGravacao = DateAdd("s", 1, fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux))
        Else
            datDataGravacao = fgDataHoraServidor(enumFormatoDataHoraAux.DataHoraAux)
        End If
    End If

    flObterDataGravacao = fgDtHr_To_Xml(datDataGravacao)

    objRS.Close
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    Set objRS = Nothing

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDataGravacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

' Verificar a existencia do tipo de solicitação em uma operação
Public Function ExisteTipoSolicitacao(ByVal pvntSequenciaOperacao As Variant, _
                                      ByVal plngTipoSolicitacao As enumTipoSolicitacao, _
                             Optional ByRef pvntCodigoXML As Variant) As Boolean

Dim strSQL                                  As String
Dim objRS                                   As ADODB.Recordset

    On Error GoTo ErrorHandler

    strSQL = "Select    NU_SEQU_OPER_ATIV,        " & vbNewLine & _
             "          CO_TEXT_XML               " & vbNewLine & _
             " From     A8.TB_OPER_ATIV_MESG_INTE " & vbNewLine & _
             " Where    NU_SEQU_OPER_ATIV   =     " & pvntSequenciaOperacao & vbNewLine & _
             " And      TP_SOLI_MESG_INTE   =     " & plngTipoSolicitacao

    Set objRS = fgQuerySQL(strSQL)

    If objRS.RecordCount > 0 Then
        pvntCodigoXML = objRS!CO_TEXT_XML
        ExisteTipoSolicitacao = True
    Else
        ExisteTipoSolicitacao = False
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "ExisteTipoSolicitacao Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function
