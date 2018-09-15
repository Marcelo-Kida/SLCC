VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCtrlProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Objeto responsável pelo roteamento de informações provenientes da camada de interface (A7) para a camada de realização de casos de uso (A7SERVER), utilizando as classes destino:
'A7Server.clsCtrlProcessamento
Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                 As Long
Private intNumeroSequencialErro              As Integer

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

'Rotear a chamada do método LerTodos para a classe de negócio A7Server.clsCtrlProcessamento
Public Function LerTodos(ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As String

Dim objCtrlProcessamento            As A7Server.clsCtrlProcessamento

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objCtrlProcessamento = CreateObject("A7Server.clsCtrlProcessamento")
    
    LerTodos = objCtrlProcessamento.LerTodos()
    Set objCtrlProcessamento = Nothing

    Exit Function

ErrorHandler:
    
    Set objCtrlProcessamento = Nothing
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
    
End Function

'Controlar o roteamento e execução dos métodos incluir, excluir, ler, lertodos para o objeto informado utilizando as classes clsConsulta e clsTransacao do A7Miu.
Public Function Executar(ByVal pstrPropriedades As String, _
                         ByRef pvntCodErro As Variant, _
                         ByRef pvntMensagemErro As Variant) As Boolean
    
Dim xmlPropriedades                 As MSXML2.DOMDocument40
Dim strOperacao                     As String
Dim objCtrlProcessamento            As A7Server.clsCtrlProcessamento

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")

    pstrPropriedades = Replace(pstrPropriedades, "'", vbNullString)

    xmlPropriedades.preserveWhiteSpace = True

    If Not xmlPropriedades.loadXML(pstrPropriedades) Then
        '3 - XML Inválido.
        lngCodigoErroNegocio = 3
        GoTo ErrorHandler
    End If

    strOperacao = xmlPropriedades.selectSingleNode("//*/@Operacao").Text

    If InStr(1, UCase(strOperacao), "INCLUIR") <> 0 Then
        Set objCtrlProcessamento = CreateObject("A7Server.clsCtrlProcessamento")
        Executar = objCtrlProcessamento.flIncluir(xmlPropriedades)
        Set objCtrlProcessamento = Nothing
    ElseIf InStr(1, UCase(strOperacao), "EXCLUIR") <> 0 Then
        Set objCtrlProcessamento = CreateObject("A7Server.clsCtrlProcessamento")
        Executar = objCtrlProcessamento.flExcluir(xmlPropriedades)
        Set objCtrlProcessamento = Nothing
    Else
        lngCodigoErroNegocio = 4  ' 4 - Operação Inválida
        GoTo ErrorHandler
    End If

    Set xmlPropriedades = Nothing

    Exit Function

ErrorHandler:
    Set objCtrlProcessamento = Nothing

    If xmlPropriedades.parseError.errorCode <> 0 Then
        Err.Number = xmlPropriedades.parseError.errorCode
        Err.Description = xmlPropriedades.parseError.reason
        Err.Source = App.EXEName & "-" & TypeName(Me) & "-" & "Executar Function"
    End If

    Set xmlPropriedades = Nothing
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description
   
End Function

