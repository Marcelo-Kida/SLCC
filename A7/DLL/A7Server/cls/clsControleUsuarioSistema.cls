VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsControleUsuarioSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo fornecimento de informações de associações do tipo mensagem e atributos , bem como a manipulação dos dados cadastrados
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

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela:

Public Function ObterPropriedades() As String

Dim strSQL                                   As String


On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_USUA_ACES,    " & _
             "          SG_SIST,         " & _
             "          DT_ULTI_ACES     " & _
             " FROM A7.TB_CTRL_ACES_SIST_USUA"

    ObterPropriedades = fgPropriedades("Grupo_ControleAcessoSistemaUsuariuo", strSQL, "A7Server.clsControleUsuarioSistema")


    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  LerTodos
Public Function Executar(ByVal pxmlControleAcesso As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlControleAcesso.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir"
            Executar = Incluir(pxmlControleAcesso.selectSingleNode("//CO_USUA_ACES").Text, _
                               pxmlControleAcesso.selectSingleNode("//SG_SIST").Text)
            
        Case "LerTodos"
            Executar = LerTodos(pxmlControleAcesso.selectSingleNode("//CO_USUA_ACES").Text, _
                                pxmlControleAcesso.selectSingleNode("//SG_SIST").Text, _
                                pxmlControleAcesso.selectSingleNode("//DT_ULTI_ACES").Text)
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

'Consultar todas os usuários  de acordo com os parâmetos de filtro informados.

Public Function LerTodos(Optional ByVal pstrUsuario As String = "", _
                         Optional ByVal pstrSistema As String = "", _
                         Optional ByVal pstrDataUltimoAcesso As String = "") As String
    
Dim strSQL                                  As String
Dim blnWhere                                As Boolean
    
On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_USUA_ACES,    " & _
             "          SG_SIST,         " & _
             "          DT_ULTI_ACES     " & _
             " FROM A7.TB_CTRL_ACES_SIST_USUA"

    If pstrUsuario <> "" Then
        strSQL = strSQL & " WHERE UPPER(CO_USUA_ACES) like '%" & UCase(Trim$(pstrUsuario)) & "'" & vbCrLf
        blnWhere = True
    End If
    
    If pstrSistema <> "" Then
        If blnWhere Then
             strSQL = strSQL & " AND SG_SIST = '" & Trim$(pstrSistema) & "'" & vbCrLf
        Else
             strSQL = strSQL & " WHERE SG_SIST = '" & Trim$(pstrSistema) & "'" & vbCrLf
             blnWhere = True
        End If
    End If
    
    If pstrDataUltimoAcesso <> "" Then
    
        If blnWhere Then
            strSQL = strSQL & " AND DT_ULTI_ACES BETWEEN " & _
                                                        fgDtHrXML_To_Oracle(pstrDataUltimoAcesso & "000000") & " AND " & _
                                                        fgDtHrXML_To_Oracle(pstrDataUltimoAcesso & "235959") & vbCrLf
    
        Else
            strSQL = strSQL & " WHERE DT_ULTI_ACES BETWEEN " & _
                                                        fgDtHrXML_To_Oracle(pstrDataUltimoAcesso & "000000") & " AND " & _
                                                        fgDtHrXML_To_Oracle(pstrDataUltimoAcesso & "235959") & vbCrLf
            
        End If
    End If
    
    strSQL = strSQL & " ORDER BY  CO_USUA_ACES"

    LerTodos = fgQueryXMLLerTodos("ControleAcessoUsuario", strSQL, App.EXEName & "." & TypeName(Me))
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Incluir(ByVal pstrUsuario As String, _
                        ByVal pstrSistema As String) As Boolean
                        
Dim strSQL                                  As String

On Error GoTo ErrorHandler
    
    If Len(Trim(pstrUsuario)) > 8 Then
        pstrUsuario = Mid(Trim(pstrUsuario), 1, 8)
    End If
    
    pstrUsuario = UCase(pstrUsuario)
    
    strSQL = " UPDATE A7.TB_CTRL_ACES_SIST_USUA " & _
             "    SET DT_ULTI_ACES = SYSDATE " & _
             "  WHERE CO_USUA_ACES ='" & Trim$(pstrUsuario) & "'" & _
             "    AND SG_SIST      ='" & Trim$(pstrSistema) & "'"

    If fgExecuteSQL(strSQL) = 0 Then
        
        strSQL = " INSERT INTO A7.TB_CTRL_ACES_SIST_USUA (" & _
                 "   CO_USUA_ACES,    " & _
                 "   SG_SIST,         " & _
                 "   DT_ULTI_ACES     " & _
                 " ) VALUES ( " & _
                 "'" & Trim$(pstrUsuario) & "' , " & _
                 "'" & Trim$(pstrSistema) & "' , " & _
                 "SYSDATE)"
        
        Call fgExecuteSQL(strSQL)
        
    End If
        
    Exit Function
ErrorHandler:
    
    If InStr(1, Err.Description, "ORA-00942") > 0 Then
        Err.Clear
        Exit Function
    End If
        
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Incluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function VerificaUsuarioInativo() As String

Dim strSQL                                  As String
Dim rsUsuario                               As ADODB.Recordset
Dim strMesgBody                             As String
Dim strSistema                              As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_USUA_ACES ," & _
             "          SG_SIST      ," & _
             "          DT_ULTI_ACES " & _
             "   FROM   A7.TB_CTRL_ACES_SIST_USUA " & _
             "  WHERE   DT_ULTI_ACES < TO_DATE('" & Format(fgDataHoraServidor(DataAux), "DD/MM/YYYY") & "','DD/MM/YYYY') - 30 " & _
             " ORDER BY CO_USUA_ACES,SG_SIST,DT_ULTI_ACES "

    Set rsUsuario = fgQuerySQL(strSQL)
        
    If rsUsuario.EOF Then
        rsUsuario.Close
        Exit Function
    End If
        
    strMesgBody = " " & String$(70, "*") & vbCrLf
        
    While Not rsUsuario.EOF
        
        Select Case Trim$(rsUsuario!SG_SIST)
            Case "A6": strSistema = "A6 - Sub Reserva"
            Case "A6T": strSistema = "A6 - Trilha Auditoria"
            Case "A7": strSistema = "A7 - BUS" & vbTab
            Case "A7T": strSistema = "A7 - Trilha Auditoria"
            Case "A8": strSistema = "A8 - LQS" & vbTab
            Case "A8T": strSistema = "A8 - Trilha Auditoria"
            Case "A8P": strSistema = "A8 - Processamento"
        End Select
        
        strMesgBody = strMesgBody & " " & fgCompletaString(UCase$(rsUsuario!CO_USUA_ACES), " ", 15, False) & vbTab
        strMesgBody = strMesgBody & fgCompletaString(strSistema, " ", 40, False) & vbTab
        strMesgBody = strMesgBody & fgCompletaString(rsUsuario!DT_ULTI_ACES, " ", 30, False) & vbTab & vbCrLf
                
        rsUsuario.MoveNext
    Wend
        
    strMesgBody = strMesgBody & " " & String$(70, "*")
        
    rsUsuario.Close
        
    Call flEnviarEmailUsuarioInativo(strMesgBody)
            
    Exit Function
ErrorHandler:
    
    VerificaUsuarioInativo = Err.Description

End Function

'Formatar a mensagem com erro não previsto acrescentando informações pertinentes a ela.

Private Function flEnviarEmailUsuarioInativo(ByVal pstrMensagemBody As String)

Dim xmlErro                                 As MSXML2.DOMDocument40
Dim strErro                                 As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim objMensagem                             As A7Server.clsMensagem
Dim strMensagem                             As String

Dim xmlNode                                 As MSXML2.IXMLDOMNode

Dim objRespostaOcorrencia                   As A7Server.clsRespostaOcorrencia

On Error GoTo ErrorHandler
                                                            
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = CreateObject("A7Server.clsMensagem")
                                                            
    strMensagem = objMensagem.ObterPropriedades
    
    xmlMensagem.loadXML strMensagem
                
    xmlMensagem.selectSingleNode("//SG_SIST_DEST").Text = "A8"
    xmlMensagem.selectSingleNode("//CO_EMPR_ORIG").Text = enumCodigoEmpresa.Santander
                
    Set xmlNode = xmlMensagem.documentElement.selectSingleNode("TX_DTLH_OCOR_ERRO")
    xmlNode.appendChild fgCreateCDATASection(pstrMensagemBody)
                
    Set objRespostaOcorrencia = CreateObject("A7Server.clsRespostaOcorrencia")
    Call objRespostaOcorrencia.EnviarRespostaOcorrencia(xmlMensagem, enumOcorrencia.UsuarioInativo)
    Set objRespostaOcorrencia = Nothing
    
    
    Set xmlErro = Nothing
    Set xmlMensagem = Nothing

    
    Exit Function

ErrorHandler:
    
    Set xmlErro = Nothing
    Set xmlMensagem = Nothing
    Set objRespostaOcorrencia = Nothing
    
    Err.Clear


End Function


