VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapaFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa            : Regerbanc - Participações , Negócios e Serviços LTDA\r\nComponente         : MIU\r\nClasse             : clsMapaFiltro\r\nData Criação       : 23/07/2003\r\nObjetivo           : Fornece as configurações para gerar o Mapa de Filtro.\r\nAnalista           : Adilson G. Damasceno/Carlos Fortes/Marcelo Kida\r\n\r\nProgramador        : Cassiano Nicolosi\r\nData               : 23/07/2003\r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F201B160344"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"

' Este componente não está sendo utilizado. Deve ser removido.

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

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

Public Function ObterMapaFiltro(ByRef pxmlMapaFiltro As MSXML2.DOMDocument40, ByRef piSistemaSLCC As Integer, ByVal psNomeFormulario As String) As Boolean

Dim strXML                                   As String
Dim strXPath                                 As String

    On Error GoTo ErrorHandler
    
    If Dir(App.Path & "\" & "Mapa de Filtro.xml") = "" Then
        'Arquivo Mapa de Filtro.xml não localizado
        lngCodigoErroNegocio = 5
        GoTo ErrorHandler
    End If

    If Not pxmlMapaFiltro.Load(App.Path & "\" & "Mapa de Filtro.xml") Then
        Call fgErroLoadXML(pxmlMapaFiltro, App.EXEName, TypeName(Me), "ObterMapaFiltro")
    End If
    
    Select Case piSistemaSLCC
        Case enumSistemaSLCC.BUS
            strXPath = "//MAPA_DE_FILTRO/BUS/" & psNomeFormulario
        Case enumSistemaSLCC.LQS
            strXPath = "//MAPA_DE_FILTRO/LQS/" & psNomeFormulario
        Case enumSistemaSLCC.SBR
            strXPath = "//MAPA_DE_FILTRO/SBR/" & psNomeFormulario
    End Select
    
    If pxmlMapaFiltro.selectSingleNode(strXPath) Is Nothing Then
        'XML do Mapa de Navegação Inválido
        lngCodigoErroNegocio = 6
        GoTo ErrorHandler
    Else
        strXML = pxmlMapaFiltro.selectSingleNode(strXPath).xml
    End If

    If Not pxmlMapaFiltro.loadXML(strXML) Then
        'XML do Mapa de Navegação Inválido
        lngCodigoErroNegocio = 6
        GoTo ErrorHandler
    End If

    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMapaFiltro", lngCodigoErroNegocio, intNumeroSequencialErro, "Formulário: " & psNomeFormulario)

End Function
