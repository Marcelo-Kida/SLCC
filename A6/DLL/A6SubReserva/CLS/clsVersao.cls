VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVersao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela informação ao usuário,
' de dados referentes à versão atual dos componentes do sistema.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private ObjectContext                       As COMSVCSLib.ObjectContext
Implements COMSVCSLib.ObjectControl

Private Sub ObjectControl_Activate()
    Set ObjectContext = COMSVCSLib.GetObjectContext()
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Deactivate()
    Set ObjectContext = Nothing
End Sub

' Este método é responsável pela obtenção da data de determinado componente.

Private Function flDataComponente() As String

    On Error GoTo ErrorHandler
    
    flDataComponente = fgDtHr_To_Xml(FileDateTime(App.Path & "\" & App.EXEName & ".dll"))

Exit Function
ErrorHandler:
    
    flDataComponente = fgDtHr_To_Xml(Now)
    
End Function

' Este método é responsável pela composição dos dados a serem exibidos ao usuário.

Private Sub flAdicionaDadosVersao(ByRef xmlVersao As MSXML2.DOMDocument40)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objDomNodePropriedade                   As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler
    
    Set objDomNode = xmlVersao.createElement("Componente")
    
    Set objDomNodePropriedade = xmlVersao.createElement("Title")
    objDomNodePropriedade.Text = App.Title
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Tipo")
    objDomNodePropriedade.Text = fgObterTipoComponente
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Major")
    objDomNodePropriedade.Text = App.Major
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Minor")
    objDomNodePropriedade.Text = App.Minor
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Revision")
    objDomNodePropriedade.Text = App.Revision
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("FileDescription")
    objDomNodePropriedade.Text = App.FileDescription
    objDomNode.appendChild objDomNodePropriedade
    
    Set objDomNodePropriedade = xmlVersao.createElement("Date")
    objDomNodePropriedade.Text = flDataComponente
    objDomNode.appendChild objDomNodePropriedade
    
    xmlVersao.documentElement.appendChild objDomNode
    
    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If
    
Exit Sub
ErrorHandler:

    Set objDomNode = Nothing
    Set objDomNodePropriedade = Nothing

    If Not ObjectContext Is Nothing Then
      'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAdicionaDadosVersao Function", lngCodigoErroNegocio, intNumeroSequencialErro)


End Sub

' Este método é responsável pela verifiação de dependências de componentes.

Private Sub flCarregarDependencias(ByRef xmlVersao As MSXML2.DOMDocument40)

Dim objVersao                               As Object
Dim arrDependencias()                       As String
Dim lngCont                                 As Long

    On Error GoTo ErrorHandler
    
    arrDependencias = Split(fgObterDependecias, ";")
    
    For lngCont = LBound(arrDependencias) To UBound(arrDependencias)
        If xmlVersao.documentElement.selectSingleNode("Componente[Title='" & arrDependencias(lngCont) & "']") Is Nothing Then
            Set objVersao = CreateObject(arrDependencias(lngCont) & ".clsVersao")
            xmlVersao.loadXML objVersao.ObterVersoesComponentes(xmlVersao.xml)
            Set objVersao = Nothing
        End If
    Next lngCont
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If

    Exit Sub
ErrorHandler:

    Set objVersao = Nothing

    If Not ObjectContext Is Nothing Then
      'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flObterDependencias Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

' Este método é responsável pela obtenção das versões dos componentes,
' e pelo retorno à camada intermediária.

Public Function ObterVersoesComponentes(ByVal pstrXMLVersao As String) As String

Dim xmlVersoes                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
    
    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlVersoes.loadXML pstrXMLVersao
    flAdicionaDadosVersao xmlVersoes
    flCarregarDependencias xmlVersoes
    
    ObterVersoesComponentes = xmlVersoes.xml
    
    Set xmlVersoes = Nothing
    
    If Not ObjectContext Is Nothing Then
       ObjectContext.SetComplete
    End If

    Exit Function
ErrorHandler:

    Set xmlVersoes = Nothing

    If Not ObjectContext Is Nothing Then
      'ObjectContext.SetAbort
    End If

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterVersoesComponentes Function", lngCodigoErroNegocio, intNumeroSequencialErro)
                    
End Function

