VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHoraLimIntSistLega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 
'' Objeto responsável pelo fornecimento de informações relativas aos horários
'' limite de integração dos sistemas com o SLCC, bem como a manipulação dos dados
'' cadastrados


Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private objContext                          As COMSVCSLib.ObjectContext
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

'Consultar todos os horários limite de integração cadastrados para os sistemas
'Alterado   : Adilson G. Damasceno
'Data       : 10/12/2010
'Solicitação: RATS 1028
'Descrição  : Incluido a verificação de Data Fim Vigência na Tabela TB_EMPRESA_HO
Public Function LerTodos() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT    A.SG_SIST,                   " & vbNewLine & _
             "           A.CO_EMPR,                   " & vbNewLine & _
             "           A.HO_LIMI_ENVI_ITGR,         " & vbNewLine & _
             "           A.CO_USUA_ULTI_ATLZ,         " & vbNewLine & _
             "           A.CO_ETCA_TRAB_ULTI_ATLZ,    " & vbNewLine & _
             "           A.DH_ULTI_ATLZ,              " & vbNewLine & _
             "           B.NO_SIST,                   " & vbNewLine & _
             "           C.NO_REDU_EMPR               " & vbNewLine & _
             " FROM      A8.TB_HORA_LIMI_ITGR_SIST A, " & vbNewLine & _
             "           A7.TB_SIST B,                " & vbNewLine & _
             "           A8.TB_EMPRESA_HO C           " & vbNewLine & _
             " WHERE     B.CO_EMPR = A.CO_EMPR        " & vbNewLine & _
             " AND       B.SG_SIST = A.SG_SIST        " & vbNewLine & _
             " AND       A.CO_EMPR = C.CO_EMPR        " & vbNewLine & _
             " AND      (C.DT_FIM_VIGE    IS NULL " & _
             " OR        C.DT_FIM_VIGE  >= " & fgDataHoraServidor_To_Oracle & ")" & _
             " AND       C.DT_INIC_VIGE <= " & fgDataHoraServidor_To_Oracle & _
             " ORDER BY  A.SG_SIST ASC                "
    
    LerTodos = fgQueryXMLLerTodos("HorarioLimite", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
''      TB_HORA_LIMI_ITGR_SIST
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT    SG_SIST,                  " & vbNewLine & _
             "           CO_EMPR,                  " & vbNewLine & _
             "           HO_LIMI_ENVI_ITGR,        " & vbNewLine & _
             "           CO_USUA_ULTI_ATLZ,        " & vbNewLine & _
             "           CO_ETCA_TRAB_ULTI_ATLZ,   " & vbNewLine & _
             "           DH_ULTI_ATLZ              " & vbNewLine & _
             " FROM      A8.TB_HORA_LIMI_ITGR_SIST "

    ObterPropriedades = fgPropriedades("Grupo_HorarioLimite", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados, tais como:
''      Incluir
''      Alterar
''      Excluir
''      Ler     e
''      LerTodos
Public Function Executar(ByVal pxmlHorarioLimite As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlHorarioLimite.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlHorarioLimite)
        Case "Ler"
            Executar = Ler(pxmlHorarioLimite.documentElement.selectSingleNode("//SG_SIST").Text, _
                           pxmlHorarioLimite.documentElement.selectSingleNode("//CO_EMPR").Text)
        Case "LerTodos"
            Executar = LerTodos
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select
    
    Exit Function
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a inclusão de um novo horário limite de integração para um sistema
Private Function flIncluir(ByRef pxmlHorarioLimite As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A8.TB_HORA_LIMI_ITGR_SIST (" & _
             " SG_SIST, " & _
             " CO_EMPR, " & _
             " HO_LIMI_ENVI_ITGR, " & _
             " CO_USUA_ULTI_ATLZ, " & _
             " CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "  DH_ULTI_ATLZ " & _
             " ) VALUES ( '" & _
             pxmlHorarioLimite.documentElement.selectSingleNode("SG_SIST").Text & "'," & _
             pxmlHorarioLimite.documentElement.selectSingleNode("CO_EMPR").Text & "," & _
             fgDtHrXML_To_Oracle(CStr(pxmlHorarioLimite.documentElement.selectSingleNode("HO_LIMI_ENVI_ITGR").Text)) & "," & _
             "'" & fgUsuarioRede & "'," & _
             "'" & fgEstacaoTrabalhoUsuario & "'," & _
             "SYSDATE)"

    fgExecuteSQL strSQL
     
    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a alteração de um horário limite de integração existente
Private Function flAlterar(ByRef pxmlHorarioLimite As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
            
    strSQL = " UPDATE A8.TB_HORA_LIMI_ITGR_SIST " & _
             " SET    SG_SIST                 = '" & pxmlHorarioLimite.documentElement.selectSingleNode("SG_SIST").Text & "'," & _
             "        CO_EMPR                 = " & pxmlHorarioLimite.documentElement.selectSingleNode("CO_EMPR").Text & "," & _
             "        HO_LIMI_ENVI_ITGR       = " & fgDtHrXML_To_Oracle(CStr(pxmlHorarioLimite.documentElement.selectSingleNode("HO_LIMI_ENVI_ITGR").Text)) & "," & _
             "        CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & _
             "        CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & _
             "        DH_ULTI_ATLZ            =  SYSDATE " & _
             " WHERE  SG_SIST  = '" & pxmlHorarioLimite.documentElement.selectSingleNode("SG_SIST").Text & "'" & _
             " AND    CO_EMPR = " & pxmlHorarioLimite.documentElement.selectSingleNode("CO_EMPR").Text & _
             " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlHorarioLimite.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos os atributos de um horário limite de integração específico,
'' cadastrado para um sistema
Public Function Ler(ByVal pstrSG_SIST As String, _
                    ByVal plngCO_EMPR As Long) As String
Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT    SG_SIST,                  " & vbNewLine & _
             "           CO_EMPR,                  " & vbNewLine & _
             "           HO_LIMI_ENVI_ITGR,        " & vbNewLine & _
             "           CO_USUA_ULTI_ATLZ,        " & vbNewLine & _
             "           CO_ETCA_TRAB_ULTI_ATLZ,   " & vbNewLine & _
             "           DH_ULTI_ATLZ              " & vbNewLine & _
             " FROM      A8.TB_HORA_LIMI_ITGR_SIST " & vbNewLine & _
             " WHERE SG_SIST  = '" & pstrSG_SIST & "'" & _
             " AND   CO_EMPR = " & plngCO_EMPR
    
    Ler = fgQueryXMLLer("Grupo_HorarioLimite", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function
    
ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Controlar o roteamento e execução dos métodos privados de manipulação de dados,
'' tais como:
''      Incluir
''      Alterar e
''      Excluir
Public Function Salvar(ByVal pxmlHorarioLimite As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    If pxmlHorarioLimite.documentElement.selectSingleNode("@Operacao").Text = "Incluir" Then
        flIncluir pxmlHorarioLimite
    ElseIf pxmlHorarioLimite.documentElement.selectSingleNode("@Operacao").Text = "Alterar" Then
        flAlterar pxmlHorarioLimite
    ElseIf pxmlHorarioLimite.documentElement.selectSingleNode("@Operacao").Text = "Excluir" Then
        flExcluir pxmlHorarioLimite
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a exclusão de um horário limite de integração existente
Private Function flExcluir(ByRef pxmlHorarioLimite As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE " & strSQL & _
             " FROM A8.TB_HORA_LIMI_ITGR_SIST " & _
             " WHERE SG_SIST  = '" & pxmlHorarioLimite.documentElement.selectSingleNode("SG_SIST").Text & "'" & _
             " AND   CO_EMPR = " & pxmlHorarioLimite.documentElement.selectSingleNode("CO_EMPR").Text & _
             " AND   HO_LIMI_ENVI_ITGR = " & fgDtHrXML_To_Oracle(pxmlHorarioLimite.documentElement.selectSingleNode("HO_LIMI_ENVI_ITGR").Text) & _
             " AND   DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(pxmlHorarioLimite.documentElement.selectSingleNode("DH_ULTI_ATLZ").Text)

    If fgExecuteSQL(strSQL) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function

'Verificação da grade de horário por sitema e empresa
Public Function VerificarGradeHorario(ByVal pstrSiglaSistema As String, _
                                      ByVal plngCodigoEmpresa As Long) As Boolean

Dim strSQL                                  As String
Dim objHoraLimite                           As ADODB.Recordset
Dim datHoraServidor                         As Date

On Error GoTo ErrorHandler

    strSQL = " SELECT   HO_LIMI_ENVI_ITGR         " & vbNewLine & _
             " FROM     A8.TB_HORA_LIMI_ITGR_SIST " & vbNewLine & _
             " WHERE    SG_SIST  = '" & pstrSiglaSistema & "'" & _
             " AND      CO_EMPR = " & plngCodigoEmpresa

    Set objHoraLimite = fgQuerySQL(strSQL)

    If objHoraLimite.RecordCount = 0 Then
        VerificarGradeHorario = False
        Exit Function
    End If

    datHoraServidor = TimeSerial(Hour(fgDataHoraServidor(enumFormatoDataHora.HORA)), Minute(fgDataHoraServidor(enumFormatoDataHora.HORA)), 0)
    If TimeSerial(Hour(objHoraLimite!HO_LIMI_ENVI_ITGR), Minute(objHoraLimite!HO_LIMI_ENVI_ITGR), 0) < datHoraServidor Then
        VerificarGradeHorario = False
    Else
        VerificarGradeHorario = True
    End If

    Exit Function

ErrorHandler:

    VerificarGradeHorario = False

    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function
