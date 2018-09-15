VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
 
Attribute VB_Name = "clsFatoGeraAlerBkOffice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'' Objeto responsável pelo fornecimento de informações relativas aos fatos
'' geradores de alerta para tipos de BackOffice, bem como a manipulação dos dados
'' cadastrados

Option Explicit

Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Implements COMSVCSLib.ObjectControl
Private objContext                          As COMSVCSLib.ObjectContext

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

'' Controlar o roteamento e execução dos métodos privados de manipulação de dados,
'' tais como:
'' 	Incluir
'' 	Alterar	e
'' 	Excluir
Public Function Executar(ByRef pxmlCadastroAlerta As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlCadastroAlerta.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Alterar"
            Executar = flAlterar(pxmlCadastroAlerta)
        Case "Excluir"
            Executar = flExcluir(pxmlCadastroAlerta)
        Case "Incluir"
            Executar = flIncluir(pxmlCadastroAlerta)
        Case Else
            ' 7 - Operação Inválida
            lngCodigoErroNegocio = 7
            GoTo ErrorHandler
    End Select

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Executar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Retorna apenas se o registro existe
Public Function Existe(ByVal plngFatoGeraAlerta As Long, _
                       ByVal plngTipoBackOffice As Long) As Boolean
                       

Dim strSql                                  As String
Dim rsFatoGeraAlerBkOffice                  As ADODB.Recordset

On Error GoTo ErrorHandler
                       
    strSql = " SELECT CO_FATO_GERA_ALER, TP_BKOF " & vbCrLf & _
             " FROM   A8.TB_FATO_GERA_ALER_BKOF  " & vbCrLf & _
             " WHERE  CO_FATO_GERA_ALER = " & plngFatoGeraAlerta & vbCrLf & _
             " AND    TP_BKOF = " & plngTipoBackOffice
                       
    
    Set rsFatoGeraAlerBkOffice = fgQuerySQL(strSql)
    
    Existe = rsFatoGeraAlerBkOffice.RecordCount > 0
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Existe Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                       
                       
End Function

'' Promover a alteração de um fato gerador de alerta para tipo de BackOffice
'' existente
Private Function flAlterar(ByRef pxmlCadastroAlerta As MSXML2.DOMDocument40) As Boolean

Dim strSql                                  As String

    On Error GoTo ErrorHandler
    
    With pxmlCadastroAlerta.documentElement
        strSql = " UPDATE  A8.TB_FATO_GERA_ALER_BKOF  " & vbNewLine & _
                 " SET     IN_ENVI_EMAIL           =  " & .selectSingleNode("IN_ENVI_EMAIL").Text & "," & vbNewLine & _
                 "         IN_ENVI_EMAIL_ANEX      =  " & .selectSingleNode("IN_ENVI_EMAIL_ANEX").Text & "," & vbNewLine & _
                 "         DE_EMAIL_ENVI_ALER      = '" & .selectSingleNode("DE_EMAIL_ENVI_ALER").Text & "'," & vbNewLine & _
                 "         TX_ASSU_EMAIL_ENVI_ALER = '" & .selectSingleNode("TX_ASSU_EMAIL_ENVI_ALER").Text & "'," & vbNewLine & _
                 "         TX_EMAIL_ENVI_ALER      = '" & .selectSingleNode("TX_EMAIL_ENVI_ALER").Text & "'," & vbNewLine & _
                 "         IN_EMIS_ALER_TELA       = '" & .selectSingleNode("IN_EMIS_ALER_TELA").Text & "'," & vbNewLine & _
                 "         CO_USUA_ULTI_ATLZ       = '" & fgUsuarioRede & "'," & vbNewLine & _
                 "         CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgEstacaoTrabalhoUsuario & "'," & vbNewLine & _
                 "         DH_ULTI_ATLZ            =  SYSDATE " & vbNewLine & _
                 " WHERE   CO_FATO_GERA_ALER       =  " & .selectSingleNode("CO_FATO_GERA_ALER").Text & vbNewLine & _
                 " AND     TP_BKOF                 =  " & .selectSingleNode("TP_BKOF").Text & vbNewLine & _
                 " AND     DH_ULTI_ATLZ            =  " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
    If fgExecuteSQL(strSql) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a exclusão de um fato gerador de alerta para tipo de BackOffice
'' existente
Private Function flExcluir(ByRef pxmlCadastroAlerta As MSXML2.DOMDocument40) As Boolean

Dim strSql                                  As String

    On Error GoTo ErrorHandler
    
    With pxmlCadastroAlerta.documentElement
        strSql = " DELETE FROM A8.TB_FATO_GERA_ALER_GRUP_USUA " & vbNewLine & _
                 " WHERE   CO_FATO_GERA_ALER                = " & .selectSingleNode("CO_FATO_GERA_ALER").Text & vbNewLine & _
                 " AND     TP_BKOF                          = " & .selectSingleNode("TP_BKOF").Text
    End With
    
    Call fgExecuteSQL(strSql)
    
    With pxmlCadastroAlerta.documentElement
        strSql = " DELETE FROM A8.TB_FATO_GERA_ALER_BKOF " & vbNewLine & _
                 " WHERE   CO_FATO_GERA_ALER           = " & .selectSingleNode("CO_FATO_GERA_ALER").Text & vbNewLine & _
                 " AND     TP_BKOF                     = " & .selectSingleNode("TP_BKOF").Text & vbNewLine & _
                 " AND     DH_ULTI_ATLZ                = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
    End With
    
    If fgExecuteSQL(strSql) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a inclusão de um novo fato gerador de alerta para tipo de BackOffice
Private Function flIncluir(ByRef pxmlCadastroAlerta As MSXML2.DOMDocument40) As Boolean

Dim strSql                                  As String
    
    On Error GoTo ErrorHandler
    
    With pxmlCadastroAlerta.documentElement
        strSql = " INSERT INTO A8.TB_FATO_GERA_ALER_BKOF (                      " & vbNewLine & _
                 "      CO_FATO_GERA_ALER,                                      " & vbNewLine & _
                 "      TP_BKOF,                                                " & vbNewLine & _
                 "      IN_ENVI_EMAIL,                                          " & vbNewLine & _
                 "      DE_EMAIL_ENVI_ALER,                                     " & vbNewLine & _
                 "      IN_ENVI_EMAIL_ANEX,                                     " & vbNewLine & _
                 "      TX_ASSU_EMAIL_ENVI_ALER,                                " & vbNewLine & _
                 "      TX_EMAIL_ENVI_ALER,                                     " & vbNewLine & _
                 "      IN_EMIS_ALER_TELA,                                      " & vbNewLine & _
                 "      CO_USUA_ULTI_ATLZ,                                      " & vbNewLine & _
                 "      CO_ETCA_TRAB_ULTI_ATLZ,                                 " & vbNewLine & _
                 "      DH_ULTI_ATLZ                                            " & vbNewLine & _
                 " ) VALUES (                                                   " & vbNewLine & _
                 "  " & .selectSingleNode("CO_FATO_GERA_ALER").Text & ",        " & vbNewLine & _
                 "  " & .selectSingleNode("TP_BKOF").Text & ",                  " & vbNewLine & _
                 "  " & .selectSingleNode("IN_ENVI_EMAIL").Text & ",            " & vbNewLine & _
                 " '" & .selectSingleNode("DE_EMAIL_ENVI_ALER").Text & "',      " & vbNewLine & _
                 "  " & .selectSingleNode("IN_ENVI_EMAIL_ANEX").Text & ",       " & vbNewLine & _
                 " '" & .selectSingleNode("TX_ASSU_EMAIL_ENVI_ALER").Text & "', " & vbNewLine & _
                 " '" & .selectSingleNode("TX_EMAIL_ENVI_ALER").Text & "',      " & vbNewLine & _
                 "  " & .selectSingleNode("IN_EMIS_ALER_TELA").Text & ",        " & vbNewLine & _
                 " '" & fgUsuarioRede & "',                                     " & vbNewLine & _
                 " '" & fgEstacaoTrabalhoUsuario & "',                          " & vbNewLine & _
                 "      SYSDATE)"
    End With
    
    Call fgExecuteSQL(strSql)
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos os atributos de um fato gerador de alerta para tipo de
'' BackOffice específico
Public Function Ler(ByVal pintCodFatoGeradorAlerta As Integer, _
                    ByVal pintTipoBackOffice As enumTipoBackOffice) As String

Dim strSql                                  As String
    
    On Error GoTo ErrorHandler

    strSql = " SELECT    CO_FATO_GERA_ALER,        " & vbNewLine & _
             "           TP_BKOF,                  " & vbNewLine & _
             "           IN_ENVI_EMAIL,            " & vbNewLine & _
             "           IN_ENVI_EMAIL_ANEX,       " & vbNewLine & _
             "           DE_EMAIL_ENVI_ALER,       " & vbNewLine & _
             "           TX_ASSU_EMAIL_ENVI_ALER,  " & vbNewLine & _
             "           TX_EMAIL_ENVI_ALER,       " & vbNewLine & _
             "           IN_EMIS_ALER_TELA,        " & vbNewLine & _
             "           DH_ULTI_ATLZ              " & vbNewLine & _
             " FROM      A8.TB_FATO_GERA_ALER_BKOF " & vbNewLine & _
             " WHERE     CO_FATO_GERA_ALER = " & pintCodFatoGeradorAlerta & vbNewLine & _
             " AND       TP_BKOF           = " & pintTipoBackOffice & vbNewLine

    Ler = fgQueryXMLLer("FatoGeraAlerBkOffice", strSql, App.EXEName & "." & TypeName(Me))

    If Len(Ler) = 0 Then
        '4002-Registro inexistente ou excluído por outro usuário
        lngCodigoErroNegocio = 4002
        GoTo ErrorHandler
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
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar as associações dos grupos de usuários ao tipo de BackOffice e fato
'' gerador de alerta específicados. Para tanto, a rotina utiliza os parâmetros
'' abaixo:
'' 	Código do Fato Gerador de Alerta
'' 	Tipo de BackOffice
Public Function LerGruposUsuariosAlerta(ByVal pintCodFatoGeradorAlerta As Integer, _
                                        ByVal pintTipoBackOffice As enumTipoBackOffice) As String

Dim strSql                                  As String
    
    On Error GoTo ErrorHandler

    strSql = " SELECT    CO_GRUP_USUA                   " & vbNewLine & _
             " FROM      A8.TB_FATO_GERA_ALER_GRUP_USUA " & vbNewLine & _
             " WHERE     CO_FATO_GERA_ALER = " & pintCodFatoGeradorAlerta & vbNewLine & _
             " AND       TP_BKOF           = " & pintTipoBackOffice & vbNewLine

    LerGruposUsuariosAlerta = fgQueryXMLLerTodos("GrupoUsuario", strSql, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerGruposUsuariosAlerta Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Consultar todos os fatos geradores de alerta para tipo de BackOffice cadastrados
Public Function LerTodos(ByVal pintTipoBackOffice As enumTipoBackOffice) As String

Dim strSql                                  As String

    On Error GoTo ErrorHandler

    strSql = " SELECT    A.CO_FATO_GERA_ALER,                           " & vbNewLine & _
             "           A.IN_ENVI_EMAIL,                               " & vbNewLine & _
             "           A.DE_EMAIL_ENVI_ALER,                          " & vbNewLine & _
             "           A.IN_EMIS_ALER_TELA,                           " & vbNewLine & _
             "           B.DE_FATO_GERA_ALER                            " & vbNewLine & _
             " FROM      A8.TB_FATO_GERA_ALER_BKOF A,                   " & vbNewLine & _
             "           A8.TB_FATO_GERA_ALER B                         " & vbNewLine & _
             " WHERE     A.CO_FATO_GERA_ALER = B.CO_FATO_GERA_ALER      " & vbNewLine & _
             " AND       A.TP_BKOF           = " & pintTipoBackOffice & vbNewLine & _
             " ORDER BY  B.DE_FATO_GERA_ALER"

    LerTodos = fgQueryXMLLerTodos("FatoGeraAlerBkOffice", strSql, App.EXEName & "." & TypeName(Me))

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

'' Consultar todos os atributos de um fato gerador de alerta para tipo de
'' BackOffice específico, bem como sua associação com as informações de fato
'' gerador de alerta
Public Function LerDetalhe(ByVal plngCodigoFatoGeraAlerta As Long, _
                           ByVal plngTipoBackOffice As Long) As String

Dim strSql                                  As String

On Error GoTo ErrorHandler

    strSql = " SELECT  A.CO_FATO_GERA_ALER,                    " & vbCrLf & _
             "         A.TP_BKOF,                              " & vbCrLf & _
             "         A.IN_ENVI_EMAIL,                        " & vbCrLf & _
             "         A.IN_ENVI_EMAIL_ANEX,                   " & vbCrLf & _
             "         A.DE_EMAIL_ENVI_ALER,                   " & vbCrLf & _
             "         A.TX_ASSU_EMAIL_ENVI_ALER,              " & vbCrLf & _
             "         A.TX_EMAIL_ENVI_ALER,                   " & vbCrLf & _
             "         A.IN_EMIS_ALER_TELA,                    " & vbCrLf & _
             "         B.DE_FATO_GERA_ALER,                    " & vbCrLf & _
             "         B.TX_EXPL_FATO_GERA_ALER                " & vbCrLf & _
             " FROM    A8.TB_FATO_GERA_ALER_BKOF A,            " & vbCrLf & _
             "         A8.TB_FATO_GERA_ALER B                  " & vbCrLf & _
             " WHERE A.CO_FATO_GERA_ALER = B.CO_FATO_GERA_ALER " & vbCrLf & _
             " AND   A.CO_FATO_GERA_ALER = " & plngCodigoFatoGeraAlerta & vbCrLf & _
             " AND   A.TP_BKOF =  " & plngTipoBackOffice
             
    LerDetalhe = fgQueryXMLLer("FatoGeraAlerBkOffice", strSql, App.EXEName & "." & TypeName(Me))
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If

Exit Function
ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerDetalhe Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
                           

'' Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a
'' estrutura da tabela:
'' 	TB_FATO_GERA_ALER_BKOF
Public Function ObterPropriedades() As String

Dim strSql                                  As String

    On Error GoTo ErrorHandler

    strSql = " SELECT * FROM A8.TB_FATO_GERA_ALER_BKOF"
    ObterPropriedades = fgPropriedades("Grupo_FatoGeraAlerBkOffice", strSql, App.EXEName & "." & TypeName(Me))

    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'' Promover a associação dos grupos de usuários ao tipo de BackOffice, para que as
'' mensagens de alerta possam ser enviadas de acordo com as configurações
Public Function SalvarGruposUsuarios(ByVal pxmlGruposUsuarios As MSXML2.DOMDocument40, _
                                     ByVal pstrOperacao As String) As String

Dim strSql                                  As String
Dim xmlDomNode                              As IXMLDOMNode

    On Error GoTo ErrorHandler
    
    strSql = " DELETE FROM A8.TB_FATO_GERA_ALER_GRUP_USUA " & vbNewLine & _
             " WHERE  CO_FATO_GERA_ALER                 = " & pxmlGruposUsuarios.documentElement.selectSingleNode("//CO_FATO_GERA_ALER").Text & vbNewLine & _
             " AND    TP_BKOF                           = " & pxmlGruposUsuarios.documentElement.selectSingleNode("//TP_BKOF").Text
    
    Call fgExecuteSQL(strSql)
    
    For Each xmlDomNode In pxmlGruposUsuarios.selectNodes("/Repeat_GrupoUsuario/*")
        With xmlDomNode
            strSql = " INSERT INTO A8.TB_FATO_GERA_ALER_GRUP_USUA ( " & vbNewLine & _
                     "      CO_FATO_GERA_ALER,                      " & vbNewLine & _
                     "      TP_BKOF,                                " & vbNewLine & _
                     "      CO_GRUP_USUA                            " & vbNewLine & _
                     " ) VALUES (                                   " & vbNewLine & _
                     "  " & .selectSingleNode("CO_FATO_GERA_ALER").Text & ", " & vbNewLine & _
                     "  " & .selectSingleNode("TP_BKOF").Text & ",           " & vbNewLine & _
                     "  " & .selectSingleNode("CO_GRUP_USUA").Text & ")"
        End With
        
        Call fgExecuteSQL(strSql)
    Next
    
    If Not objContext Is Nothing Then
        objContext.SetComplete
    End If
    
    Exit Function

ErrorHandler:
    If Not objContext Is Nothing Then
        'objContext.SetAbort
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "SalvarGruposUsuarios Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
