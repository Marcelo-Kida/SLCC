VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistSitLancCCContabil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Participações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsLancamentoCCContabil
'Data Criação   : 20/01/2004
'Objetivo       : Armazenar os métodos da tabela TB_HIST_SITU_LANC_CC_CNTB
'
'Analista       : Carlos Fortes
'
'Programador    : Carlos Fortes
'Data           : 20/01/2004
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

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

Public Function Incluir(ByVal pvntSequenciaOperacao As Variant, _
                        ByVal plngTipoLancIntegracao As Long, _
                        ByVal plngSituacaoProcessoIntegracao As Long, _
               Optional ByVal plngCodigoErro As Long = 0, _
               Optional ByVal pstrDescricaoErro As String = vbNullString, _
               Optional ByVal pstrJustificativa As String = vbNullString) As Boolean

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " Insert Into A8.TB_HIST_SITU_LANC_CC_CNTB (" & vbNewLine & _
             "  DH_SITU_LANC_CC, " & vbNewLine & _
             "  NU_SEQU_OPER_ATIV, " & vbNewLine & _
             "  TP_LANC_ITGR, " & vbNewLine & _
             "  CO_SITU_PROC, " & vbNewLine & _
             "  CO_ERRO, " & vbNewLine & _
             "  TX_MESG_ERRO, " & vbNewLine & _
             "  TX_JUST_CANC, " & vbNewLine & _
             "  CO_USUA_ATLZ, " & vbNewLine & _
             "  CO_ETCA_TRAB_ATLZ) "
    strSQL = strSQL & vbNewLine & " Values (" & _
             " Sysdate, " & vbNewLine & _
             pvntSequenciaOperacao & ", " & vbNewLine & _
             plngTipoLancIntegracao & ", " & vbNewLine & _
             plngSituacaoProcessoIntegracao & ", " & vbNewLine & _
             IIf(plngCodigoErro = 0, " Null ", plngCodigoErro) & ", " & vbNewLine & _
             "'" & pstrDescricaoErro & "', " & vbNewLine & _
             "'" & pstrJustificativa & "', " & vbNewLine & _
             "'" & fgUsuarioRede & "', " & vbNewLine & _
             "'" & fgEstacaoTrabalhoUsuario & "')"

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
