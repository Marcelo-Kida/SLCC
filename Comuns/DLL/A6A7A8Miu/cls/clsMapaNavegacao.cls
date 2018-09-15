VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMapaNavegacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Empresa            : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nComponente         : MIU\r\nClasse             : clsMIU\r\nData Criação       : 07/10/2002\r\nObjetivo           : Fornece as configurações para gerar o Mapa de Navegação.\r\nAnalista           : Adilson G. Damasceno/Carlos Fortes/Marcelo Kida\r\n\r\nProgramador        : Carlos Fortes\r\nData               : 07/10/2002\r\n\r\nData Teste         :\r\nAutor              :\r\n\r\nData Alteração     :\r\nAutor              :\r\nObjetivo           :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3DB4018D0326"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MTS Class"
' Este componente tem como objetivo, agrupar os métodos responsáveis pela ligação entre a camada de Interface
' e a camada de Negócios, no que diz respeito ao XML de Mapa de Navegação para configurações iniciais do sistema.
Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio As Long
Attribute lngCodigoErroNegocio.VB_VarDescription = "Variável utilizada para tratamento de erros"
Private intNumeroSequencialErro As Integer

Private objContext    As COMSVCSLib.ObjectContext
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

' Este método carrega, a partir do XML Mapa de Navegação, um outro XML contendo instruções a serem executadas
' na inicialização de formulários na camada de interface.
Public Function ObterMapaNavegacao(ByRef pxmlMapaNavegacao As MSXML2.DOMDocument40, _
                                   ByVal penumSistemaSLCC As enumSistemaSLCC, _
                                   ByVal pstrFuncionalidade As String, _
                                   ByRef pvntCodErro As Variant, _
                                   ByRef pvntMensagemErro As Variant) As Boolean

Dim strXML                                   As String
Dim strXPath                                 As String

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    If Dir(App.Path & "\" & "Mapa de Navegação.xml") = "" Then
        'Arquivo Mapa de Navegação.xml não localizado
        lngCodigoErroNegocio = 5
        GoTo ErrorHandler
    End If

    If Not pxmlMapaNavegacao.Load(App.Path & "\" & "Mapa de Navegação.xml") Then
        Call fgErroLoadXML(pxmlMapaNavegacao, App.EXEName, TypeName(Me), "ObterMapaNavegacao")
    End If
    
    Select Case penumSistemaSLCC
        Case enumSistemaSLCC.BUS
            strXPath = "BUS/" & pstrFuncionalidade
        Case enumSistemaSLCC.LQS
            strXPath = "LQS/" & pstrFuncionalidade
        Case enumSistemaSLCC.SBR
            strXPath = "SBR/" & pstrFuncionalidade
    End Select
    
    If pxmlMapaNavegacao.documentElement.selectSingleNode(strXPath) Is Nothing Then
        'XML do Mapa de Navegação Inválido
        lngCodigoErroNegocio = 6
        GoTo ErrorHandler
    Else
        strXML = pxmlMapaNavegacao.documentElement.selectSingleNode(strXPath).xml
    End If

    If Not pxmlMapaNavegacao.loadXML(strXML) Then
        'XML do Mapa de Navegação Inválido
        lngCodigoErroNegocio = 6
        GoTo ErrorHandler
    End If

Exit Function
ErrorHandler:
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterMapaNavegacao", lngCodigoErroNegocio, intNumeroSequencialErro, "Funcionalidade: " & pstrFuncionalidade)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function

