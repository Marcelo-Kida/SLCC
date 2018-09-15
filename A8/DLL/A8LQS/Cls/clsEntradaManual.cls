VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntradaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pela gravação de uma mensagem gerada via entrada manual

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

'Obter empresas
Public Function LerTodosEmpresa() As String

Dim objEmpresa                              As A6A7A8.clsEmpresa

On Error GoTo ErrorHandler

    Set objEmpresa = CreateObject("A6A7A8.clsEmpresa")
    LerTodosEmpresa = objEmpresa.LerTodos()
    Set objEmpresa = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objEmpresa = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodosEmpresa Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Gravar os parâmetros padrão da entrada manual
Public Sub GravarParametroEntradaManual(ByVal plngTipoBackOffice As Long, _
                                        ByVal plngCodigoEmpresa As Long, _
                                        ByVal plngTipoOperacao As Long, _
                                        ByVal pstrXMLParm As String)

Dim strSQL                                  As String

    strSQL = "DELETE FROM  A8.TB_PARM_ENTR_MANL " & vbCrLf & _
              "WHERE CO_EMPR = " & plngCodigoEmpresa & vbCrLf & _
               " AND TP_BKOF = " & plngTipoBackOffice & vbCrLf & _
               " AND TP_OPER = " & plngTipoOperacao

    fgExecuteSQL strSQL

    strSQL = "INSERT INTO A8.TB_PARM_ENTR_MANL (CO_EMPR, " & vbCrLf & _
                                               "TP_BKOF, " & vbCrLf & _
                                               "TP_OPER, " & vbCrLf & _
                                               "CO_TEXT_XML, " & vbCrLf & _
                                               "CO_ETCA_TRAB_ULTI_ATLZ, " & vbCrLf & _
                                               "CO_USUA_ULTI_ATLZ, " & vbCrLf & _
                                               "DH_ULTI_ATLZ) " & vbCrLf & _
                    "VALUES (" & plngCodigoEmpresa & ", " & vbCrLf & _
                            plngTipoBackOffice & ", " & vbCrLf & _
                            plngTipoOperacao & ", " & vbCrLf & _
                            fgInsertVarchar4000(pstrXMLParm) & ", " & vbCrLf & _
                            "'" & fgEstacaoTrabalhoUsuario & "', " & vbCrLf & _
                            "'" & fgUsuarioRede & "', " & vbCrLf & _
                            "SYSDATE)"

    fgExecuteSQL strSQL

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "GravarParametroEntradaManual Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Sub

'Obter os parâmetros padrão da entrada manual
Public Function ObterParametroEntradaManual(ByVal plngTipoBackOffice As Long, _
                                            ByVal plngCodigoEmpresa As Long, _
                                            ByVal plngTipoOperacao As Long) As String
On Error GoTo ErrorHandler
Dim strSQL                                  As String
Dim Rs                                      As Recordset

    strSQL = "SELECT CO_TEXT_XML FROM A8.TB_PARM_ENTR_MANL " & vbCrLf & _
              "WHERE CO_EMPR = " & plngCodigoEmpresa & vbCrLf & _
               " AND TP_BKOF = " & plngTipoBackOffice & vbCrLf & _
               " AND TP_OPER = " & plngTipoOperacao

    Set Rs = fgQuerySQL(strSQL)
    
    If Not Rs.EOF Then
        ObterParametroEntradaManual = fgSelectVarchar4000(Rs!CO_TEXT_XML)
    End If
    
    Set Rs = Nothing
    
    Exit Function
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterParametroEntradaManual Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
