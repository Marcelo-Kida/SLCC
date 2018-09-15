VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoMesgAtributo"
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
' - TB_TIPO_MESG_ATRB
Public Function ObterPropriedades() As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler

    strSQL = "SELECT TP_MESG, " & _
                   " TP_FORM_MESG, " & _
                   " NO_ATRB_MESG, " & _
                   " NU_ORDE_AGRU_ATRB, " & _
                   " IN_ATRB_IDEF_MESG, " & _
                   " IN_OBRI_ATRB " & _
             " FROM A7.TB_TIPO_MESG_ATRB"

    ObterPropriedades = fgPropriedades("Grupo_TipoMensagemAtributo", strSQL, "A7Server.clsTipoMesgAtributo")

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterPropriedades Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Controlar o roteamento e execução dos métodos privados, tais como:
' -  LerTodos
Public Function Executar(ByVal pxmlAtributoMensagem As MSXML2.DOMDocument40) As String

Dim strOperacao                             As String

On Error GoTo ErrorHandler

    strOperacao = pxmlAtributoMensagem.selectSingleNode("//*/@Operacao").Text

    Select Case strOperacao
        Case "Ler"
            Executar = Ler(pxmlAtributoMensagem.selectSingleNode("//NO_ATRB_MESG").Text)
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

'Consultar todos os atributos de um tipo de atributo da mensagem
Public Function Ler(ByVal pstrNomeAtributoMensagem As String) As String

Dim strSQL                                   As String

On Error GoTo ErrorHandler
    
    strSQL = " SELECT  A.NO_ATRB_MESG " & _
             "   FROM  A7.TB_TIPO_MESG_ATRB A ," & _
             "         A7.TB_TIPO_MESG      B  " & _
             "  WHERE  A.TP_MESG = B.TP_MESG" & _
             "    AND  UPPER(A.NO_ATRB_MESG) = '" & UCase(Trim(pstrNomeAtributoMensagem)) & "'" & _
             "    AND  B.DT_FIM_VIGE_MESG IS NULL "
    
    Ler = fgQueryXMLLer("TipoMensagemAtributo", strSQL, App.EXEName & "." & TypeName(Me))

    Exit Function

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Ler Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


