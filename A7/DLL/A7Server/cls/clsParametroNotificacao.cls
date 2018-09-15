VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametroNotificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFA1F3601D5"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsParametroNotificacao
'Data Criação   : 04/08/2003
'Objetivo       :

'Analista       : Marcelo Kida
'
'Programador    : Cassiano Nicolosi
'Data           : 04/08/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :


Option Explicit

Implements COMSVCSLib.ObjectControl

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

Public Function Executar(ByRef pobjParamNotific As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pobjParamNotific.selectSingleNode("//*/@Operacao").Text
    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(pobjParamNotific)
        Case "Ler"
            Executar = Ler(pobjParamNotific)
        Case "LerTodos"
            Executar = LerTodos
        Case "Excluir"
            Call flExcluir(pobjParamNotific)
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

Private Function flAlterar(ByRef pobjParamNotific As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " UPDATE A7.TB_PARAMETRO_NOTIFICACAO " & _
             "    SET TP_GRAU_SEVE       = " & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/TP_GRAU_SEVE").Text & ", " & vbCrLf & _
             "        IN_ENVI_EMAIL      = '" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/IN_ENVI_EMAIL").Text & "', " & vbCrLf & _
             "        TX_ENDE_EMAIL_NOTI = '" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/TX_ENDE_EMAIL_NOTI").Text & "', " & vbCrLf & _
             "        CO_USUA_ULTI_ATLZ  = '" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/CO_USUA_ULTI_ATLZ").Text & "', " & vbCrLf & _
             "        DH_ULTI_ATLZ       = SYSDATE " & vbCrLf & _
             " WHERE  CO_OCOR            = " & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/CO_OCOR").Text
        
    Call fgExecuteSQL(strSQL)
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flExcluir(ByRef pobjParamNotific As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " DELETE FROM A7.TB_PARAMETRO_NOTIFICACAO " & _
            " WHERE  CO_OCOR = " & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/CO_OCOR").Text
    
    Call fgExecuteSQL(strSQL)
    

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flIncluir(ByRef pobjParamNotific As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " INSERT INTO A7.TB_PARAMETRO_NOTIFICACAO ( " & vbCrLf & _
            "           CO_OCOR,                        " & vbCrLf & _
            "           TP_GRAU_SEVE,                   " & vbCrLf & _
            "           IN_ENVI_EMAIL,                  " & vbCrLf & _
            "           TX_ENDE_EMAIL_NOTI,             " & vbCrLf & _
            "           CO_USUA_ULTI_ATLZ,              " & vbCrLf & _
            "           DH_ULTI_ATLZ                    " & vbCrLf & _
            " ) VALUES ( " & vbCrLf & _
            "" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/CO_OCOR").Text & ", " & vbCrLf & _
            "" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/TP_GRAU_SEVE").Text & ", " & vbCrLf & _
            "'" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/IN_ENVI_EMAIL").Text & "', " & vbCrLf & _
            "'" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/TX_ENDE_EMAIL_NOTI").Text & "', " & vbCrLf & _
            "'" & pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/CO_USUA_ULTI_ATLZ").Text & "', " & vbCrLf & _
            "SYSDATE)"
    
    Call fgExecuteSQL(strSQL)

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Ler(ByRef pobjParamNotific As MSXML2.DOMDocument40) As String
    
Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    TP_GRAU_SEVE,                 " & vbCrLf & _
            "           IN_ENVI_EMAIL,                " & vbCrLf & _
            "           TX_ENDE_EMAIL_NOTI           " & vbCrLf & _
            " FROM      A7.TB_PARAMETRO_NOTIFICACAO    " & vbCrLf & _
            " WHERE     CO_OCOR = " & pobjParamNotific.selectSingleNode("//CO_OCOR").Text

    Ler = fgQueryXMLLerTodos("ParametroNotificacao", strSQL)
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function LerTodos() As String

Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    A.CO_OCOR,                      " & vbCrLf & _
            "           A.DE_OCOR,                      " & vbCrLf & _
            "           B.TP_GRAU_SEVE,                 " & vbCrLf & _
            "           B.IN_ENVI_EMAIL,                " & vbCrLf & _
            "           B.TX_ENDE_EMAIL_NOTI,           " & vbCrLf & _
            "           B.CO_USUA_ULTI_ATLZ,            " & vbCrLf & _
            "           B.DH_ULTI_ATLZ                  " & vbCrLf & _
            " FROM      A7.TB_OCORRENCIA A,             " & vbCrLf & _
            "           A7.TB_PARAMETRO_NOTIFICACAO B   " & vbCrLf & _
            " WHERE     A.CO_OCOR = B.CO_OCOR           " & vbCrLf & _
            " ORDER BY  A.CO_OCOR ASC                   "

    LerTodos = fgQueryXMLLerTodos("ParametroNotificacao", strSQL)
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterParametroNotificacao(ByRef pobjEvento As DOMDocument40) As String

Dim objRS                                   As ADODB.Recordset
Dim strSQL                                   As String


On Error GoTo ErrorHandler

    strSQL = " SELECT"
    

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterParametroNotificacao Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterPropriedades() As String

    Dim strSQL                               As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT    *                           " & vbCrLf & _
            " FROM      A7.TB_PARAMETRO_NOTIFICACAO "

    ObterPropriedades = fgPropriedades("Grupo_ParametroNotificacao", strSQL, "A7BusServer.clsParametroNotificacao")

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Salvar(ByRef pobjParamNotific As MSXML2.DOMDocument40) As Boolean

    On Error GoTo ErrorHandler
    
    If pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/@Operacao").Text = "Incluir" Then
        Call flIncluir(pobjParamNotific)
    ElseIf pobjParamNotific.selectSingleNode("//Grupo_ParametroNotificacao/@Operacao").Text = "Alterar" Then
        Call flAlterar(pobjParamNotific)
    Else
        ' 7 - Operação Inválida
        lngCodigoErroNegocio = 7
        GoTo ErrorHandler
    End If
        
    Salvar = True
      
    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Salvar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
