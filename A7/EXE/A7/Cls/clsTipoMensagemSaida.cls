VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoMensagemSaida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Pacote         :
'Classe         : clsAtributo
'Data Criação   : 01/07/2003
'Objetivo       :
'
'Analista       : Marcelo Garcia
'
'Programador    : Marcelo Garcia
'Data           : 26/02/2004
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :


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

Public Function LerTodos() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "  SELECT   CO_MESG_SAID,           " & vbNewLine & _
             "           DE_MESG_SAID            " & vbNewLine & _
             "    FROM   A7.TB_TIPO_MESG_SAID    "
    
    strSQL = strSQL & " ORDER BY   CO_MESG_SAID "

    LerTodos = fgQueryXMLLerTodos("TipoMensagemSaida", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT " & _
                " CO_MESG_SAID, " & _
                " DE_MESG_SAID " & _
             " FROM A7.TB_TIPO_MESG_SAID "

    ObterPropriedades = fgPropriedades("Grupo_TipoMensagemSaida", strSQL, "A7Server.clsTipoMensagemSaida")

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Executar(ByVal pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlTipoMensagemSaida.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar", "Excluir"
            Executar = Salvar(pxmlTipoMensagemSaida)
        Case "Ler"
            Executar = Ler(CLng(pxmlTipoMensagemSaida.selectSingleNode("//CO_MESG_SAID").Text))
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
Private Function flIncluir(ByRef pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    With pxmlTipoMensagemSaida.documentElement
        
        strSQL = " INSERT INTO A7.TB_TIPO_MESG_SAID (" & _
                "           CO_MESG_SAID, " & _
                "           DE_MESG_SAID  " & _
                " ) VALUES ( " & _
                CLng(.selectSingleNode("CO_MESG_SAID").Text) & _
                ",'" & .selectSingleNode("DE_MESG_SAID").Text & "')"

        fgExecuteSQL strSQL
    
    End With
    
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function
Private Function flAlterar(ByRef pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As Boolean

Dim strSQL                                   As String

On Error GoTo ErrorHandler
                
    With pxmlTipoMensagemSaida.documentElement
                
        strSQL = " UPDATE A7.TB_TIPO_MESG_SAID " & _
                "  SET    DE_MESG_SAID            = '" & .selectSingleNode("DE_MESG_SAID").Text & "' " & _
                " WHERE   CO_MESG_SAID            =  " & .selectSingleNode("CO_MESG_SAID").Text
    
        If fgExecuteSQL(strSQL) = 0 Then
           '4001 - Registro foi alterado ou excluído por outro usuário.
           lngCodigoErroNegocio = 4001
           GoTo ErrorHandler
        End If
      
    End With
      
    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAlterar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Ler(ByVal plngCodigoMensagemSaida As Long) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_MESG_SAID, " & _
            "           DE_MESG_SAID  " & _
            " FROM      A7.TB_TIPO_MESG_SAID " & _
            " WHERE     CO_MESG_SAID =" & plngCodigoMensagemSaida
    
    Ler = fgQueryXMLLer("Grupo_TipoMensagemSaida", strSQL, App.EXEName & "." & TypeName(Me))
    
    If Len(Ler) = 0 Then
       '4002 - Registro Inexistente ou excluído por outro usuário.
       lngCodigoErroNegocio = 4002
       GoTo ErrorHandler
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function Salvar(ByVal pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    flRegrasNegocio pxmlTipoMensagemSaida
            
    With pxmlTipoMensagemSaida.documentElement
            
        If .selectSingleNode("@Operacao").Text = "Incluir" Then
            flIncluir pxmlTipoMensagemSaida
        ElseIf .selectSingleNode("@Operacao").Text = "Alterar" Then
            flAlterar pxmlTipoMensagemSaida
        ElseIf .selectSingleNode("@Operacao").Text = "Excluir" Then
            flExcluir pxmlTipoMensagemSaida
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

Private Function flRegrasNegocio(ByRef pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As Boolean

On Error GoTo ErrorHandler
    
    With pxmlTipoMensagemSaida.documentElement
        
        If .selectSingleNode("@Operacao").Text = "Incluir" Then
            
            If flVerificaTipoMensagemSaidaCadastrado(CLng(.selectSingleNode("CO_MESG_SAID").Text)) Then
                'Tipo de Mensagem de Saida já está cadastrado
                lngCodigoErroNegocio = 1015
                GoTo ErrorHandler
            End If
            
        ElseIf .selectSingleNode("@Operacao").Text = "Alterar" Then
                   
        
        ElseIf .selectSingleNode("@Operacao").Text = "Excluir" Then
        
        End If
    
    End With

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flRegrasNegocio Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flVerificaTipoMensagemSaidaCadastrado(ByVal plngCodigoTipoMensagemSaida As Long) As Boolean

Dim strSQL                                   As String
Dim objRS                                    As ADODB.Recordset

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  COUNT(DE_MESG_SAID) " & _
             "   FROM  A7.TB_TIPO_MESG_SAID " & _
             "  WHERE  CO_MESG_SAID = " & plngCodigoTipoMensagemSaida
    
    Set objRS = fgQuerySQL(strSQL)
    
    If objRS.fields(0) = 0 Then
        flVerificaTipoMensagemSaidaCadastrado = False
    Else
        flVerificaTipoMensagemSaidaCadastrado = True
    End If
    
    objRS.Close

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flVerificaTipoMensagemSaidaCadastrada Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Function flExcluir(ByRef pxmlTipoMensagemSaida As MSXML2.DOMDocument40) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    With pxmlTipoMensagemSaida.documentElement

        strSQL = " DELETE FROM A7.TB_TIPO_MESG_SAID " & _
                 "  WHERE CO_MESG_SAID = " & .selectSingleNode("CO_MESG_SAID").Text

    End With

    If fgExecuteSQL(strSQL) = 0 Then
       '4001 - Registro foi alterado ou excluído por outro usuário.
       lngCodigoErroNegocio = 4001
       GoTo ErrorHandler
    End If

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flExcluir Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function



