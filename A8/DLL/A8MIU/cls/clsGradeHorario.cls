VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGradeHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao Cadastro e Controle de Grades de Horário.

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

' Este método aciona a leitura de um registro da grade de horário na camada de negócios,
' e retorna para a camada de interface.

Public Function ObterDetalhesGradeHorario(ByVal pstrGradeHorario As String, _
                                          ByRef pvntCodErro As Variant, _
                                          ByRef pvntMensagemErro As Variant) As String

Dim objGradeHorario                         As A8LQS.clsGradeHorario

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
    ObterDetalhesGradeHorario = objGradeHorario.ObterDetalhesGradeHorario(pstrGradeHorario)
    Set objGradeHorario = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:

    Set objGradeHorario = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterDetalhesGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de um ou mais registros da grade de horário de acordo com o filtro passado como
' parâmetro, na camada de negócios, e retorna para a camada de interface.

Public Function ObterGradeHorario(ByVal xmlDocFiltros As String, _
                                  ByRef pvntCodErro As Variant, _
                                  ByRef pvntMensagemErro As Variant) As String

Dim objMensagem                             As A8LQS.clsMensagem
Dim objGradeHorario                         As A8LQS.clsGradeHorario
Dim xmlDomFiltro                            As MSXML2.DOMDocument40
Dim xmlDomAux                               As MSXML2.DOMDocument40
Dim strCodigoMensagem                       As String
Dim lngCodigoMensagemXML                    As Long
Dim lngLocalLiquidacao                      As Long
Dim lngSequencialISPBDestino                As Long

On Error GoTo ErrorHandler

    pvntCodErro = 0
    pvntMensagemErro = ""

    Set xmlDomFiltro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltro.loadXML(xmlDocFiltros) Then
        '100 - Documento XML Inválido.
        lngCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
    
    strCodigoMensagem = xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_GradeHorario/CodigoMensagem").Text
    
    If Not xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_GradeHorario/CodigoMensagemXML") Is Nothing Then
        lngCodigoMensagemXML = Val(xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_GradeHorario/CodigoMensagemXML").Text)
    End If
    
    If Not xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_GradeHorario/LocalLiquidacao") Is Nothing Then
        lngLocalLiquidacao = Val(xmlDomFiltro.selectSingleNode("Repeat_Filtros/Grupo_GradeHorario/LocalLiquidacao").Text)
    End If
    
    Set xmlDomAux = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMensagem = CreateObject("A8LQS.clsMensagem")
    lngSequencialISPBDestino = objMensagem.ObterSequencialISPBDestino(strCodigoMensagem, _
                                                                      lngCodigoMensagemXML, _
                                                                      lngLocalLiquidacao)
    Set objMensagem = Nothing
    
    Set xmlDomAux = Nothing
    
    Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
    ObterGradeHorario = objGradeHorario.ObterGradeHorario(strCodigoMensagem, lngSequencialISPBDestino)
    
    Set objGradeHorario = Nothing
    Set xmlDomFiltro = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

    Exit Function
    
ErrorHandler:
    Set objGradeHorario = Nothing
    Set xmlDomFiltro = Nothing
    
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterGradeHorario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

' Este método aciona a leitura de mensagens SBP assiadas à grade de horário passada como parâmetro,
' na camada de negócios, e retorna para a camada de interface.

Public Function ObterMensagensSPB(ByVal pstrGradeHorario As String, _
                                  ByRef pvntCodErro As Variant, _
                                  ByRef pvntMensagemErro As Variant) As String

Dim objGradeHorario                         As A8LQS.clsGradeHorario

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objGradeHorario = CreateObject("A8LQS.clsGradeHorario")
    ObterMensagensSPB = objGradeHorario.ObterMensagensSPB(pstrGradeHorario)
    Set objGradeHorario = Nothing
    
    If Not ObjectContext Is Nothing Then
        ObjectContext.SetComplete
    End If

Exit Function
ErrorHandler:
    Set objGradeHorario = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMensagensSPB Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
