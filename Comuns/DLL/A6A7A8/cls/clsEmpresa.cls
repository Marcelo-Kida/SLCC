VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pela manutenção dos registros de empresas

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

'Recuperar os registros da tabela A8.TB_EMPRESA_HO

Public Function LerTodos() As String

Dim strSQL                                   As String
    
On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_EMPR, " & _
             "          CO_EMPR_FUSI, " & _
             "          CO_BANC, " & _
             "          NU_CNPJ, " & _
             "          NO_EMPR, " & _
             "          NO_REDU_EMPR, " & _
             "          SQ_ISPB, " & _
             "          CO_GRUP_CPUL, " & _
             "          DT_INIC_VIGE, " & _
             "          DT_FIM_VIGE, " & _
             "          ID_USUA_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ " & _
             "  FROM    A8.TB_EMPRESA_HO " & _
             " WHERE    (DT_FIM_VIGE    IS NULL " & _
             "    OR     DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & _
             "   AND     DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
             " ORDER BY  CO_EMPR ASC"

    LerTodos = fgQueryXMLLerTodos("Empresa", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar as propriedades (colunas) da tabela A8.TB_EMPRESA_HO

Public Function ObterPropriedades() As String

Dim strSQL                                   As String
    
On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_EMPR, " & _
             "          CO_EMPR_FUSI, " & _
             "          CO_BANC, " & _
             "          NU_CNPJ, " & _
             "          NO_EMPR, " & _
             "          NO_REDU_EMPR, " & _
             "          SQ_ISPB, " & _
             "          CO_GRUP_CPUL, " & _
             "          DT_INIC_VIGE, " & _
             "          DT_FIM_VIGE, " & _
             "          ID_USUA_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ " & _
             "  FROM    A8.TB_EMPRESA_HO "
           
    ObterPropriedades = fgPropriedades("Grupo_Empresa", strSQL, "A6A7A8.clsEmpresa")
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executar as operações (Inclusão,Alteração,Exclusão)

Public Function Executar(ByRef pxmlEmpresa As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlEmpresa.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "LerTodos"
            Executar = LerTodos()
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

'Recuperar dados de um empresa

Public Function Ler(ByVal plngCodigoEmpresa As Long) As String

Dim strSQL                                  As String
    
On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_EMPR, " & _
             "          CO_EMPR_FUSI, " & _
             "          CO_BANC, " & _
             "          NU_CNPJ, " & _
             "          NO_EMPR, " & _
             "          NO_REDU_EMPR, " & _
             "          SQ_ISPB, " & _
             "          CO_GRUP_CPUL, " & _
             "          DT_INIC_VIGE, " & _
             "          DT_FIM_VIGE, " & _
             "          ID_USUA_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ, " & _
             "          ID_PART_CAMR_CETIP " & _
             "  FROM    A8.TB_EMPRESA_HO " & _
             " WHERE    (DT_FIM_VIGE    IS NULL " & _
             "    OR     DT_FIM_VIGE  >=  " & fgDataHoraServidor_To_Oracle & ")" & _
             "   AND     DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
             "   AND     CO_EMPR      = " & plngCodigoEmpresa
             
    Ler = fgQueryXMLLer("Grupo_Empresa", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Recuperar código da empresa fusionada de uma empresa

Public Function ObterCodigoEmpresaFusionada(ByVal plngCodigoEmpresa As Long) As Long

Dim rsEmpresa                               As ADODB.Recordset
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT   CO_EMPR_FUSI " & _
             "  FROM    A8.TB_EMPRESA_HO " & _
             " WHERE    (DT_FIM_VIGE    IS NULL " & _
             "    OR     DT_FIM_VIGE  >=  " & fgDataHoraServidor_To_Oracle & ")" & _
             "   AND     DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
             "   AND     CO_EMPR      = " & plngCodigoEmpresa
             
    Set rsEmpresa = fgQuerySQL(strSQL)
    
    If rsEmpresa.EOF Then
        '4076 - Empresa não Cadastrada na Tabela Empresa Sistema PJ
        lngCodigoErroNegocio = 4076
        GoTo ErrorHandler
    Else
        ObterCodigoEmpresaFusionada = rsEmpresa!CO_EMPR_FUSI
    End If
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterCodigoEmpresaFusionada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
