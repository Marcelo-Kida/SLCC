VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParamReprocCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelos métodos de controle do parâmetro de reprocessamento de Conta Corrente.

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

'Inclui uma parametrização de histórico de conta corrente
Public Function flIncluir(ByRef pxmlParametroReprocCC As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim strCodSubTipo                           As String
Dim strFinalidadeTED                        As String
Dim strCodProd                              As String
Dim strHrLimiItgr                           As String

On Error GoTo ErrorHandler

    flConsiste pxmlParametroReprocCC

    strSQL = " INSERT INTO A8.TB_PARM_REPR_CNTA_CRRT (  " & vbNewLine & _
             " CD_EMPR                     ," & vbNewLine & _
             " TP_BKOF                     ," & vbNewLine & _
             " TP_OPER                     ," & vbNewLine & _
             " IN_LANC_DEBT_CRED           ," & vbNewLine & _
             " TP_CNAL_VEND                ," & vbNewLine & _
             " QT_REPR_CNTA_CRRT           ," & vbNewLine & _
             " QT_HORA_INTL_REPR           ," & vbNewLine & _
             " IN_ATIV                     ," & vbNewLine & _
             " CD_USUA_ULTI_ATLZ           ," & vbNewLine & _
             " CD_ETCA_USUA_ULTI_ATLZ      ," & vbNewLine & _
             " DH_ULTI_ATLZ                ," & vbNewLine & _
             " HR_LIMI_ITGR                 " & vbNewLine & _
             " ) VALUES (                " & vbNewLine

    With pxmlParametroReprocCC.documentElement
    
        If .selectSingleNode("HO_LIMI_ITGR").Text <> vbNullString Then
            strHrLimiItgr = "to_date('" & .selectSingleNode("HO_LIMI_ITGR").Text & " ', 'DD/MM/YYYY HH24:MI:SS')"
        Else
            strHrLimiItgr = "NULL"
        End If
    
        strSQL = strSQL & _
                 CLng("0" & .selectSingleNode("CD_EMPR").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("TP_BKOF").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("TP_OPER").Text) & "           , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text) & " , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("TP_CNAL_VEND").Text) & "      , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("QT_REPR_CNTA_CRRT").Text) & " , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("QT_HORA_INTL_REPR").Text) & " , " & vbNewLine & _
                 CLng("0" & .selectSingleNode("IN_ATIV").Text) & "           , " & vbNewLine & _
                 "'" & fgUsuarioRede & "'                                    , " & vbNewLine & _
                 "'" & fgEstacaoTrabalhoUsuario & "'                         , " & vbNewLine & _
                 " SYSDATE                                                   , " & vbNewLine & _
                 strHrLimiItgr & ")"
    End With

    Call fgExecuteSQL(strSQL)

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Altera uma parametrização de histórico de conta corrente
Private Function flAlterar(ByRef pxmlParametroReprocCC As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim strHrLimiItgr                           As String

On Error GoTo ErrorHandler

    flConsiste pxmlParametroReprocCC
    
    With pxmlParametroReprocCC.documentElement
        If .selectSingleNode("HO_LIMI_ITGR").Text <> vbNullString Then
            strHrLimiItgr = "to_date('" & .selectSingleNode("HO_LIMI_ITGR").Text & " ', 'DD/MM/YYYY HH24:MI:SS')"
        Else
            strHrLimiItgr = "NULL"
        End If
        strSQL = " UPDATE A8.TB_PARM_REPR_CNTA_CRRT                                                                  " & vbNewLine & _
                 " SET    QT_REPR_CNTA_CRRT      = " & CLng("0" & .selectSingleNode("QT_REPR_CNTA_CRRT").Text) & " , " & vbNewLine & _
                 "        QT_HORA_INTL_REPR      = " & CLng("0" & .selectSingleNode("QT_HORA_INTL_REPR").Text) & " , " & vbNewLine & _
                 "        IN_ATIV                = " & CLng("0" & .selectSingleNode("IN_ATIV").Text) & " , " & vbNewLine & _
                 "        HR_LIMI_ITGR           = " & strHrLimiItgr & " , " & vbNewLine & _
                 "        CD_USUA_ULTI_ATLZ      = '" & fgUsuarioRede & "' , " & vbNewLine & _
                 "        CD_ETCA_USUA_ULTI_ATLZ = '" & fgEstacaoTrabalhoUsuario & "' , " & vbNewLine & _
                 "        DH_ULTI_ATLZ           =  SYSDATE " & vbNewLine & _
                 " WHERE  CD_EMPR           =  " & .selectSingleNode("CD_EMPR").Text & vbNewLine & _
                 " AND    TP_BKOF           =  " & .selectSingleNode("TP_BKOF").Text & vbNewLine & _
                 " AND    TP_OPER           =  " & .selectSingleNode("TP_OPER").Text & vbNewLine & _
                 " AND    IN_LANC_DEBT_CRED =  " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & vbNewLine & _
                 " AND    DH_ULTI_ATLZ      =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & vbNewLine & _
                 " AND    TP_CNAL_VEND      = '" & .selectSingleNode("TP_CNAL_VEND").Text & "'"
    End With
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Salva uma alteração de parametrização de histórico de conta corrente, seja de inclusão, alteração ou exclusão
Public Function Salvar(ByVal pxmlParametroReprocCC As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    With pxmlParametroReprocCC.documentElement
    
        If .selectSingleNode("@Operacao").Text = "Incluir" Then
            flIncluir pxmlParametroReprocCC
        ElseIf .selectSingleNode("@Operacao").Text = "Alterar" Then
            flAlterar pxmlParametroReprocCC
        ElseIf .selectSingleNode("@Operacao").Text = "Excluir" Then
            flExcluir pxmlParametroReprocCC
        Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
        End If
        
    End With
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Exclui uma parametrização de histórico de conta corrente
Private Function flExcluir(ByRef pxmlParametroReprocCC As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    With pxmlParametroReprocCC.documentElement

        strSQL = " DELETE " & _
                 " FROM A8.TB_PARM_REPR_CNTA_CRRT " & _
                 " WHERE  CD_EMPR           =  " & .selectSingleNode("CD_EMPR").Text & "                    " & vbNewLine & _
                 " AND    TP_BKOF           =  " & .selectSingleNode("TP_BKOF").Text & "                    " & vbNewLine & _
                 " AND    TP_OPER           =  " & .selectSingleNode("TP_OPER").Text & "                    " & vbNewLine & _
                 " AND    IN_LANC_DEBT_CRED =  " & .selectSingleNode("IN_LANC_DEBT_CRED").Text & "          " & vbNewLine & _
                 " AND    DH_ULTI_ATLZ      =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text) & " " & vbNewLine & _
                 " AND    TP_CNAL_VEND      =  " & .selectSingleNode("TP_CNAL_VEND").Text & "                 " & vbNewLine
    End With


    If fgExecuteSQL(strSQL) = 0 Then
        '4002 - Registro Inexistente ou excluído por outro usuário.
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter todas as parametrizações de histórico de conta corrente
Public Function ObterPropriedades() As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler
    
    strSQL = " SELECT CD_EMPR                ," & vbNewLine & _
             " TP_BKOF                       ," & vbNewLine & _
             " TP_OPER                       ," & vbNewLine & _
             " IN_LANC_DEBT_CRED             ," & vbNewLine & _
             " TP_CNAL_VEND                  ," & vbNewLine & _
             " QT_REPR_CNTA_CRRT             ," & vbNewLine & _
             " QT_HORA_INTL_REPR             ," & vbNewLine & _
             " HR_LIMI_ITGR  as HO_LIMI_ITGR ," & vbNewLine & _
             " IN_ATIV                       ," & vbNewLine & _
             " CD_USUA_ULTI_ATLZ             ," & vbNewLine & _
             " CD_ETCA_USUA_ULTI_ATLZ        ," & vbNewLine & _
             " DH_ULTI_ATLZ                   " & vbNewLine & _
             " FROM   A8.TB_PARM_REPR_CNTA_CRRT   "
             
             

    ObterPropriedades = fgPropriedades("Grupo_ParametroReprocCC", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter todas as parametrizações de histórico de conta corrente, utilizando a segregação de acesso
Public Function LerTodos() As String

Dim strSQL                                  As String
Dim strWhereSegregacao                      As String

On Error GoTo ErrorHandler

    strSQL = " SELECT CD_EMPR              ," & vbNewLine & _
             " TP_BKOF                     ," & vbNewLine & _
             " TP_OPER                     ," & vbNewLine & _
             " IN_LANC_DEBT_CRED           ," & vbNewLine & _
             " TP_CNAL_VEND                ," & vbNewLine & _
             " QT_REPR_CNTA_CRRT           ," & vbNewLine & _
             " QT_HORA_INTL_REPR           ," & vbNewLine & _
             " HR_LIMI_ITGR as HO_LIMI_ITGR," & vbNewLine & _
             " IN_ATIV                     ," & vbNewLine & _
             " CD_USUA_ULTI_ATLZ           ," & vbNewLine & _
             " CD_ETCA_USUA_ULTI_ATLZ      ," & vbNewLine & _
             " DH_ULTI_ATLZ                 " & vbNewLine & _
             " FROM   A8.TB_PARM_REPR_CNTA_CRRT   "

        strWhereSegregacao = fgSegregaDados("A8.TB_PARM_REPR_CNTA_CRRT", _
                                            False, _
                                            "A8.TB_PARM_REPR_CNTA_CRRT", _
                                            "A8.TB_PARM_REPR_CNTA_CRRT", _
                                            True, _
                                            False, _
                                            True, _
                                            False, _
                                            False, _
                                            False)
                                            
    strSQL = strSQL & strWhereSegregacao
    
    strSQL = strSQL & " ORDER BY CD_EMPR,TP_OPER,TP_CNAL_VEND,IN_LANC_DEBT_CRED "

    LerTodos = fgQueryXMLLerTodos("ParametroReprocCC", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function Ler(ByVal plngCodigoEmpresa As Long, _
                    ByVal plngTipoBackOffice As Long, _
                    ByVal plngTipoOperacao As Long, _
                    ByVal plngDebitoCredito As Long, _
                    ByVal plngTipoCanalVenda As Long) As String

Dim strSQL                                  As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT CD_EMPR               ," & vbNewLine & _
             " TP_BKOF                      ," & vbNewLine & _
             " TP_OPER                      ," & vbNewLine & _
             " IN_LANC_DEBT_CRED            ," & vbNewLine & _
             " TP_CNAL_VEND                 ," & vbNewLine & _
             " QT_REPR_CNTA_CRRT            ," & vbNewLine & _
             " QT_HORA_INTL_REPR            ," & vbNewLine & _
             " HR_LIMI_ITGR as HO_LIMI_ITGR ," & vbNewLine & _
             " IN_ATIV                      ," & vbNewLine & _
             " CD_USUA_ULTI_ATLZ            ," & vbNewLine & _
             " CD_ETCA_USUA_ULTI_ATLZ       ," & vbNewLine & _
             " DH_ULTI_ATLZ                  " & vbNewLine & _
             " FROM   A8.TB_PARM_REPR_CNTA_CRRT        " & vbNewLine & _
             " WHERE    CD_EMPR             =   " & plngCodigoEmpresa & vbNewLine & _
             " AND      TP_BKOF             =   " & plngTipoBackOffice & vbNewLine & _
             " AND      TP_OPER             =   " & plngTipoOperacao & vbNewLine & _
             " AND      TP_CNAL_VEND        =   " & plngTipoCanalVenda & vbNewLine & _
             " AND      IN_LANC_DEBT_CRED   =   " & plngDebitoCredito & vbNewLine
             
    Ler = fgQueryXMLLer("ParametroReprocCC", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Executa uma tarefa em parametrização de histórico de conta corrente, seja de leitura ou alteração
Public Function Executar(ByVal pxmlParametroReprocCC As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlParametroReprocCC.documentElement.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Excluir", "Alterar"
            Salvar pxmlParametroReprocCC
        Case "Ler"
            With pxmlParametroReprocCC.documentElement
                Executar = Ler(CLng("0" & .selectSingleNode("CD_EMPR").Text), _
                               CLng("0" & .selectSingleNode("TP_BKOF").Text), _
                               CLng("0" & .selectSingleNode("TP_OPER").Text), _
                               CLng("0" & .selectSingleNode("IN_LANC_DEBT_CRED").Text), _
                               CLng("0" & .selectSingleNode("TP_CNAL_VEND").Text) _
                               )
            End With
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

'Valida dados de uma parametrização de histórico de conta corrente
Private Function flConsiste(ByRef pxmlParametroReprocCC As MSXML2.DOMDocument40) As String

On Error GoTo ErrorHandler

    With pxmlParametroReprocCC.documentElement
    End With

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flConsiste Function", lngCodigoErroNegocio, intNumeroSequencialErro)
End Function


