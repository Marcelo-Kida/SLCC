VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsInstituicaoISPB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações relativas às instituições
'' ISPB cadastradas. Exemplo, dados gerais de:
'' 	Bancos
'' 	BACEN	e
'' 	Câmaras de Liquidação

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl

Private objContext As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

'' Consultar todas as instituições ISPB cadastradas
Public Function LerTodos(Optional ByVal plngSequencialISPB As Long = 0, _
                         Optional ByVal pstrCodigoISPB As String = "") As String

Dim strSQL                                  As String

On Error GoTo ErrorHandLerTodos

    strSQL = "SELECT    SQ_ISPB, " & vbCrLf & _
             "          CO_ISPB, " & vbCrLf & _
             "          NO_ISPB, " & vbCrLf & _
             "          NU_CNPJ, " & vbCrLf & _
             "          CO_CPEN, " & vbCrLf & _
             "          IN_TIPO_ISPB, " & vbCrLf & _
             "          DT_INIC_VIGE, " & vbCrLf & _
             "          DT_FIM_VIGE, " & vbCrLf & _
             "          ID_USUA_ULTI_ATLZ, " & vbCrLf & _
             "          DH_ULTI_ATLZ " & vbCrLf & _
             "  FROM    A8.TB_INSTITUICAO_SPB " & vbCrLf & _
             " WHERE    DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
             "   AND    (DT_FIM_VIGE >= " & fgDataHoraServidor_To_Oracle & " OR DT_FIM_VIGE IS NULL)" & vbCrLf

    If plngSequencialISPB > 0 Then
        strSQL = strSQL & " AND  SQ_ISPB = " & plngSequencialISPB & vbCrLf
    End If

    If pstrCodigoISPB <> vbNullString Then
        strSQL = strSQL & " AND     CO_ISPB = '" & pstrCodigoISPB & "'" & vbCrLf
    End If

    LerTodos = fgQueryXMLLerTodos("Instituicao_ISPB", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandLerTodos:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

