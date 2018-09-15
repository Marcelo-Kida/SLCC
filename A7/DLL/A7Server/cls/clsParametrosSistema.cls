VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParametrosSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo fornecimento de informações de parâmetros de sistema, bem como a manipulação dos dados cadastrados
Option Explicit

Private intNumeroSequencialErro              As Integer
Private lngCodigoErroNegocio                 As Long

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

'Consultar todos os sistemas para a posterior seleção de manutenção.
Public Function LerTodos(Optional ByVal pintEmpresa As Integer = 0) As String

Dim objSistema                          As A6A7A8.clsSistema

On Error GoTo ErrorHandler

    Set objSistema = CreateObject("A6A7A8.clsSistema")
    
    LerTodos = objSistema.LerTodos(pintEmpresa, "S")
   
    Set objSistema = Nothing

    Exit Function
ErrorHandler:

    Set objSistema = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura das tabelas TB_ENDE_FILA_MQSE e TB_RSPT_OCOR_SIST
Public Function ObterPropriedades() As String
Dim objDOMSistema                           As MSXML2.DOMDocument40
Dim strSQL                                  As String
Dim strPropriedades                         As String

On Error GoTo ErrorHandler
    
    Set objDOMSistema = CreateObject("MSXML2.DOMDocument.4.0")
        
    Call fgAppendNode(objDOMSistema, "", "Grupo_ParametrosSistema", "")
    
    strSQL = " SELECT SG_SIST_DEST,           " & _
             "        CO_EMPR_DEST,           " & _
             "        CO_USUA_ULTI_ATLZ,      " & _
             "        CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "        DH_ULTI_ATLZ,           " & _
             "        NO_FILA_MQSE            " & _
             "   FROM A7.TB_ENDE_FILA_MQSE    "

    strPropriedades = fgPropriedades("Grupo_EnderecoFilaMqseries", strSQL, "A7Server.clsParametrosSistema")
    
    Call fgAppendXML(objDOMSistema, "Grupo_ParametrosSistema", strPropriedades)

    '----------------------------------------------------------------------
    
    strSQL = " SELECT    SG_SIST,                " & _
             "           CO_EMPR,                " & _
             "           CO_OCOR_MESG,           " & _
             "           CO_USUA_ULTI_ATLZ,      " & _
             "           CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "           DH_ULTI_ATLZ            " & _
            " FROM      A7.TB_RSPT_OCOR_SIST     "

    strPropriedades = fgPropriedades("Grupo_RespostaOcorrenciaSistema", strSQL, "A7Server.clsParametrosSistema")
    
    Call fgAppendXML(objDOMSistema, "Grupo_ParametrosSistema", strPropriedades)

    ObterPropriedades = objDOMSistema.xml
    
    Set objDOMSistema = Nothing
    
    Exit Function

ErrorHandler:
    
    Set objDOMSistema = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  Excluir
' -  LerTodos
' -  Ler
Public Function Executar(ByRef pobjSistema As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pobjSistema.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            Executar = Salvar(pobjSistema)
        Case "Ler"
            Executar = Ler(pobjSistema.selectSingleNode("//SG_SIST_DEST").Text, _
                  CLng(Val(pobjSistema.selectSingleNode("//CO_EMPR_DEST").Text)))
        Case "LerTodos"
            Executar = LerTodos(CLng(Val(pobjSistema.selectSingleNode("//CO_EMPR").Text)))
        Case "Excluir"
            Executar = flExcluir(pobjSistema)
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

'Promover a inclusão de um novo parâmetro para um determinado sistema
Private Function flIncluir(ByRef pobjSistema As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String
Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    strSQL = " INSERT INTO A7.TB_ENDE_FILA_MQSE( " & _
            " SG_SIST_DEST, " & _
            " CO_EMPR_DEST, " & _
            " NO_FILA_MQSE, " & _
            " CO_USUA_ULTI_ATLZ, " & _
            " CO_ETCA_TRAB_ULTI_ATLZ, " & _
            " DH_ULTI_ATLZ " & _
            " ) VALUES (" & _
            "'" & pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'," & _
            CLng(pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text) & "," & _
            "'" & pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text & "'," & _
            "'" & fgObterUsuarioRede & "'," & _
            "'" & fgObterEstacaoTrabalhoUsuario & "'," & _
            "SYSDATE)"

    fgExecuteSQL strSQL
    
    
    strSQL = " DELETE FROM A7.TB_RSPT_OCOR_SIST " & _
             " WHERE  SG_SIST = '" & pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'" & _
             " AND    CO_EMPR = " & pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text

    fgExecuteSQL strSQL
    
    
    For Each objNode In pobjSistema.selectNodes("//Repeat_RespostaOcorrenciaSistema/*")
    
        strSQL = " INSERT INTO A7.TB_RSPT_OCOR_SIST ( " & _
                " CO_OCOR_MESG, " & _
                " SG_SIST, " & _
                " CO_EMPR, " & _
                " CO_USUA_ULTI_ATLZ, " & _
                " CO_ETCA_TRAB_ULTI_ATLZ, " & _
                " DH_ULTI_ATLZ  " & _
                " ) VALUES ( " & _
                CLng(objNode.selectSingleNode("CO_OCOR_MESG").Text) & "," & _
                "'" & objNode.selectSingleNode("SG_SIST").Text & "', " & _
                CLng(objNode.selectSingleNode("CO_EMPR").Text) & "," & _
                "'" & fgObterUsuarioRede & "', " & _
                "'" & fgObterEstacaoTrabalhoUsuario & "', " & _
                "SYSDATE)"
        
        fgExecuteSQL strSQL
    
    Next

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Promover a alteração parâmetro já existente para um sistema
Private Function flAlterar(ByRef pxmlParamSistema As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    With pxmlParamSistema.documentElement
    
        strSQL = " UPDATE A7.TB_ENDE_FILA_MQSE " & _
                 " SET    NO_FILA_MQSE            = '" & .selectSingleNode("Grupo_EnderecoFilaMqseries/NO_FILA_MQSE").Text & "'," & _
                 "        CO_USUA_ULTI_ATLZ       = '" & fgObterUsuarioRede & "'," & _
                 "        CO_ETCA_TRAB_ULTI_ATLZ  = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
                 "        DH_ULTI_ATLZ = SYSDATE" & _
                 " WHERE  SG_SIST_DEST = '" & .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'" & _
                 " AND    CO_EMPR_DEST = " & .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text & _
                 " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("Grupo_EnderecoFilaMqseries/DH_ULTI_ATLZ").Text)
    
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
       
        strSQL = " DELETE FROM A7.TB_RSPT_OCOR_SIST " & _
                 " WHERE  SG_SIST = '" & .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'" & _
                 " AND    CO_EMPR = " & .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text
    
        fgExecuteSQL strSQL
    
    End With
    
    For Each objNode In pxmlParamSistema.selectNodes("//Repeat_RespostaOcorrenciaSistema/*")
    
        strSQL = " INSERT INTO A7.TB_RSPT_OCOR_SIST ( " & _
                " CO_OCOR_MESG, " & _
                " SG_SIST, " & _
                " CO_EMPR, " & _
                " CO_USUA_ULTI_ATLZ, " & _
                " CO_ETCA_TRAB_ULTI_ATLZ, " & _
                " DH_ULTI_ATLZ " & _
                " ) VALUES ( " & _
                CLng(objNode.selectSingleNode("CO_OCOR_MESG").Text) & "," & _
                "'" & objNode.selectSingleNode("SG_SIST").Text & "', " & _
                CLng(objNode.selectSingleNode("CO_EMPR").Text) & "," & _
                "'" & fgObterUsuarioRede & "', " & _
                "'" & fgObterEstacaoTrabalhoUsuario & "', " & _
                "SYSDATE)"
        
        fgExecuteSQL strSQL
    
    Next
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function




'Obter Detalhes do Robo - Nick
Public Function ObterParametroRobo(ByVal pstrSiglaSistema As String, _
                                   ByVal pstrCodMesg As String) As Boolean

Dim strSQL                                  As String
Dim objDomOcorrenciaSistema                 As MSXML2.DOMDocument40
Dim strPropriedades                         As String

On Error GoTo ErrorHandler
    
    Set objDomOcorrenciaSistema = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(objDomOcorrenciaSistema, "", "Grupo_ParametrosSistema", "")
    
    strSQL = " SELECT   SG_SIST, " & _
             "          CD_MESG_SPB, " & _
             "          IN_SITU_PROC_MESG " & _
             " FROM     A7.TB_GERE_MESG_ROBO " & _
             " WHERE    SG_SIST = '" & Trim(pstrSiglaSistema) & "'" & _
             " AND      CD_MESG_SPB = '" & Trim(pstrCodMesg) & "' AND IN_SITU_PROC_MESG = 1"
             
   
    strPropriedades = fgQueryXMLLer("Grupo_ParametrosSistema", strSQL, App.EXEName & "." & TypeName(Me))
        
    If strPropriedades = vbNullString Then
        ObterParametroRobo = False
    Else
        ObterParametroRobo = True
    End If
   
    Set objDomOcorrenciaSistema = Nothing
    Exit Function
ErrorHandler:
    
    Set objDomOcorrenciaSistema = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function #NK#", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


'Consultar todos os atributos de um parâmetro de sistema específico
Public Function Ler(ByVal pstrSiglaSistema As String, _
                    ByVal plngCodigoEmpresa As Long) As String

Dim strSQL                                  As String
Dim objDomOcorrenciaSistema                 As MSXML2.DOMDocument40
Dim strPropriedades                         As String

On Error GoTo ErrorHandler
    
    Set objDomOcorrenciaSistema = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(objDomOcorrenciaSistema, "", "Grupo_ParametrosSistema", "")
    
    strSQL = " SELECT   SG_SIST_DEST, " & _
             "          CO_EMPR_DEST, " & _
             "          NO_FILA_MQSE, " & _
             "          CO_USUA_ULTI_ATLZ, " & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ " & _
             " FROM     A7.TB_ENDE_FILA_MQSE  " & _
             " WHERE    SG_SIST_DEST = '" & Trim(pstrSiglaSistema) & "'" & _
             " AND      CO_EMPR_DEST = " & plngCodigoEmpresa
   
    strPropriedades = fgQueryXMLLer("Grupo_EnderecoFilaMqseries", strSQL, App.EXEName & "." & TypeName(Me))
        
    If strPropriedades = vbNullString Then
        Set objDomOcorrenciaSistema = Nothing
        Exit Function
    End If

    Call fgAppendXML(objDomOcorrenciaSistema, "Grupo_ParametrosSistema", strPropriedades)
    
    strSQL = " SELECT    CO_OCOR_MESG " & _
             " FROM      A7.TB_RSPT_OCOR_SIST " & _
             " WHERE     SG_SIST = '" & Trim(pstrSiglaSistema) & "'" & _
             " AND       CO_EMPR = " & plngCodigoEmpresa
    
    strPropriedades = fgQueryXMLLerTodos("RespostaOcorrenciaSistema", strSQL, App.EXEName & "." & TypeName(Me))
    
    If strPropriedades <> vbNullString Then
        Call fgAppendXML(objDomOcorrenciaSistema, "Grupo_ParametrosSistema", strPropriedades)
    Else
        Call fgAppendNode(objDomOcorrenciaSistema, "Grupo_ParametrosSistema", "Repeat_RespostaOcorrenciaSistema", "")
    End If
    
    Ler = objDomOcorrenciaSistema.xml
    
    Set objDomOcorrenciaSistema = Nothing
    

    Exit Function
ErrorHandler:
    
    Set objDomOcorrenciaSistema = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Rotear o comando de slavar para as rotinas internas flIncluir e flAlterar
Public Function Salvar(ByRef pobjSistema As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
            
    If pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/@Operacao").Text = "Incluir" Then
        flIncluir pobjSistema
    ElseIf pobjSistema.selectSingleNode("//Grupo_EnderecoFilaMqseries/@Operacao").Text = "Alterar" Then
        flAlterar pobjSistema
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

'Promover a exclusão de um parâmetro de sistema existente
Private Function flExcluir(ByRef pxmlParamSistema As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                  As String
Dim objNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    With pxmlParamSistema.documentElement
    
        strSQL = " DELETE FROM A7.TB_ENDE_FILA_MQSE " & _
                 " WHERE  SG_SIST_DEST = '" & .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'" & _
                 " AND    CO_EMPR_DEST = " & .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text & _
                 " AND    DH_ULTI_ATLZ = " & fgDtHrXML_To_Oracle(.selectSingleNode("Grupo_EnderecoFilaMqseries/DH_ULTI_ATLZ").Text)
    
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
       
        strSQL = " DELETE FROM A7.TB_RSPT_OCOR_SIST " & _
                 " WHERE  SG_SIST = '" & .selectSingleNode("Grupo_EnderecoFilaMqseries/SG_SIST_DEST").Text & "'" & _
                 " AND    CO_EMPR = " & .selectSingleNode("Grupo_EnderecoFilaMqseries/CO_EMPR_DEST").Text
    
        fgExecuteSQL strSQL
    
    End With
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


