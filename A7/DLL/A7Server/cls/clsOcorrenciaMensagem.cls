VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOcorrenciaMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo fornecimento de informações de ocorrência de mensagem, bem como a manipulação dos dados cadastrados.
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

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  Incluir
' -  Alterar
' -  LerTodos
' -  Ler
Public Function Executar(ByRef pxmlOcorrencia As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

    On Error GoTo ErrorHandler

    strOperacao = pxmlOcorrencia.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Incluir", "Alterar"
            
            Executar = flAlterar(pxmlOcorrencia)
        
        Case "Ler"
            Executar = Ler(pxmlOcorrencia.selectSingleNode("//CO_OCOR_MESG").Text)
        
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

'Promover a alteração de uma ocorrência de mensagem cadastrada.
Private Function flAlterar(ByRef pxmlOcorrencia As MSXML2.DOMDocument40) As String

Dim strSQL                                  As String

On Error GoTo ErrorHandler

    With pxmlOcorrencia.documentElement
            
        strSQL = " UPDATE A7.TB_OCOR_MESG " & _
                 " SET       TP_GRAU_SEVE           =  " & .selectSingleNode("TP_GRAU_SEVE").Text & "," & _
                 "           TX_ENDE_EMAIL_NOTI     = '" & LCase(.selectSingleNode("TX_ENDE_EMAIL_NOTI").Text) & "'," & _
                 "           IN_ENVI_EMAIL          = '" & .selectSingleNode("IN_ENVI_EMAIL ").Text & "'," & _
                 "           CO_USUA_ULTI_ATLZ      = '" & fgObterUsuarioRede & "'," & _
                 "           CO_ETCA_TRAB_ULTI_ATLZ = '" & fgObterEstacaoTrabalhoUsuario & "'," & _
                 "           DH_ULTI_ATLZ           =  SYSDATE " & _
                 " WHERE     CO_OCOR_MESG           = " & .selectSingleNode("CO_OCOR_MESG").Text & _
                 " AND     DH_ULTI_ATLZ              = " & fgDtHrXML_To_Oracle(.selectSingleNode("DH_ULTI_ATLZ").Text)
    
    If fgExecuteSQL(strSQL) = 0 Then
        '4001 - Registro foi alterado ou excluído por outro usuário
        lngCodigoErroNegocio = 4001
        GoTo ErrorHandler
    End If

    
    End With
        
    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Atualizar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todos os atributos de uma ocorrência de mensagem específica.
Public Function Ler(ByVal plngCodigoOcorrenciaMensagem As Long) As String

Dim strSQL                                   As String

    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_OCOR_MESG,           " & _
             "          IN_ENVI_EMAIL,          " & _
             "          DE_OCOR_MESG,           " & _
             "          DE_ABRV_OCOR_MESG,      " & _
             "          TP_GRAU_SEVE,           " & _
             "          TX_ENDE_EMAIL_NOTI,     " & _
             "          CO_USUA_ULTI_ATLZ,      " & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ            " & _
             "  FROM    A7.TB_OCOR_MESG         " & _
             "  WHERE CO_OCOR_MESG  = " & plngCodigoOcorrenciaMensagem
    
    Ler = fgQueryXMLLer("Grupo_OcorrenciaMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function
    
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Consultar todas as ocorrências de mensagem cadastradas.
Public Function LerTodos() As String
    
Dim strSQL                                   As String
    
    On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_OCOR_MESG,           " & _
             "          IN_ENVI_EMAIL,          " & _
             "          DE_OCOR_MESG,           " & _
             "          DE_ABRV_OCOR_MESG,      " & _
             "          TP_GRAU_SEVE,           " & _
             "          TX_ENDE_EMAIL_NOTI,     " & _
             "          CO_USUA_ULTI_ATLZ,      " & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ            " & _
             "  FROM    A7.TB_OCOR_MESG         " & _
             " ORDER BY CO_OCOR_MESG ASC        "
            
    LerTodos = fgQueryXMLLerTodos("OcorrenciaMensagem", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LerTodos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter as configurações de tipo e tamanho dos atributos, a partir de consulta a estrutura da tabela TB_OCOR_MESG.
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = " SELECT   CO_OCOR_MESG,           " & _
             "          IN_ENVI_EMAIL,          " & _
             "          DE_OCOR_MESG,           " & _
             "          DE_ABRV_OCOR_MESG,      " & _
             "          TP_GRAU_SEVE,           " & _
             "          TX_ENDE_EMAIL_NOTI,     " & _
             "          CO_USUA_ULTI_ATLZ,      " & _
             "          CO_ETCA_TRAB_ULTI_ATLZ, " & _
             "          DH_ULTI_ATLZ            " & _
             "  FROM    A7.TB_OCOR_MESG         "

    ObterPropriedades = fgPropriedades("Grupo_OcorrenciaMensagem", strSQL, "A7Server.clsOcorrenciaMensagem")

    Exit Function

ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


