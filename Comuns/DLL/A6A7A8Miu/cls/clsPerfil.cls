VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Componente responsável pelo controle de acesso e Controle de Alçadas

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

'Verifica permissão de executar um dos módulos do sistema

Public Function VerificarPermissao(ByVal pstrUsuario As String, _
                                   ByVal pstrTransacao As String) As Boolean

Dim objPerfil                               As A6A7A8.clsControleAcesso
Dim blnTemAcesso                            As Boolean

On Error GoTo ErrorHandler
    
    Set objPerfil = CreateObject("A6A7A8.clsControleAcesso")
    
    If objPerfil.VerificarPermissao(pstrUsuario, pstrTransacao) Then
        VerificarPermissao = True
        Call flIncluirUsuarioControleAcessoSistema(pstrUsuario, pstrTransacao)
    End If
    
    Set objPerfil = Nothing

    Exit Function
ErrorHandler:
    
    Set objPerfil = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "VerificarPermissao Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter ID do usuário

Public Function ObterUsuario() As String

Dim objUsuario                              As A6A7A8.clsUsuario

On Error GoTo ErrorHandler
    
    Set objUsuario = CreateObject("A6A7A8.clsUsuario")
    ObterUsuario = objUsuario.ObterUsuarioRede()
    Set objUsuario = Nothing

    Exit Function
ErrorHandler:
    
    Set objUsuario = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Perfil de digitador , validação da alçada

Public Sub DigitarAlcada(ByVal pstrUsuario As String, _
                         ByVal pstrValor As String, _
                Optional ByRef pvntProtocolo As Variant)

Dim objPerfil                               As A6A7A8.clsPerfil

On Error GoTo ErrorHandler
    
    Set objPerfil = CreateObject("A6A7A8.clsPerfil")
        
    objPerfil.DigitarAlcada pstrUsuario, pstrValor, pvntProtocolo
        
    Set objPerfil = Nothing

    Exit Sub
ErrorHandler:
    
    Set objPerfil = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "DigitarAlcada Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Perfil de aprovador , validação da alçada

Public Sub AprovarAlcada(ByVal pstrUsuario As String, _
                         ByVal pstrValor As String, _
                Optional ByVal pvntProtocolo As Variant)

Dim objPerfil                               As A6A7A8.clsPerfil

On Error GoTo ErrorHandler
    
    Set objPerfil = CreateObject("A6A7A8.clsPerfil")
    objPerfil.AprovarAlcada pstrUsuario, pstrValor, pvntProtocolo
    Set objPerfil = Nothing

    Exit Sub
ErrorHandler:
    
    Set objPerfil = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AprovarAlcada Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Obter as funções que o usuário tem acesso

Public Function ObterControleAcesso(ByVal pstrSiglaSistema As String, _
                                    ByRef pvntCodErro As Variant, _
                                    ByRef pvntMensagemErro As Variant)

Dim objControleAcesso                               As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    pvntCodErro = 0
    pvntMensagemErro = ""
    
    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    ObterControleAcesso = objControleAcesso.ObterControleAcesso(pstrSiglaSistema)
    Set objControleAcesso = Nothing

Exit Function
ErrorHandler:
    
    Set objControleAcesso = Nothing
    'Comentado devido ao novo tratamento de erro do SOAP
    'If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    'Call fgRaiseError(App.EXEName, TypeName(Me), "ObterControleAcesso Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    pvntCodErro = Err.Number
    pvntMensagemErro = Err.Description

End Function
        

Private Function flIncluirUsuarioControleAcessoSistema(ByVal pstrUsuario As String, _
                                                       ByVal pstrTransacao As String) As Boolean

Dim objControleUsuario                      As Object
Dim objMIU                                  As Object
Dim xmlPropriedades                         As MSXML2.DOMDocument40
Dim strSistema                              As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set objMIU = CreateObject("A7miu.clsMiu")
    
    Set objControleUsuario = CreateObject("A7Server.clsControleUsuarioSistema")
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlPropriedades.loadXML objControleUsuario.ObterPropriedades
        
    xmlPropriedades.documentElement.selectSingleNode("@Operacao").Text = "Incluir"
    xmlPropriedades.selectSingleNode("//CO_USUA_ACES").Text = pstrUsuario

    Select Case pstrTransacao
        Case "A6CCACESS01"
            strSistema = "A6C"
        Case "A6ACESS01"
            strSistema = "A6"
        Case "A7ACESS01"
            strSistema = "A7"
        Case "A8ACESS01"
            strSistema = "A8"
        Case "A6CCACESS02"
            strSistema = "ACT"
        Case "A6ACESS02"
            strSistema = "A6T"
        Case "A7ACESS02"
            strSistema = "A7T"
        Case "A8ACESS02"
            strSistema = "A8T"
    End Select
    
    xmlPropriedades.selectSingleNode("//SG_SIST").Text = strSistema
    
    Call objMIU.Executar(xmlPropriedades.xml, _
                         vntCodErro, _
                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objControleUsuario = Nothing
    Set objMIU = Nothing
    Set xmlPropriedades = Nothing
    
Exit Function
ErrorHandler:

    Set objControleUsuario = Nothing
    Set objMIU = Nothing
    Set xmlPropriedades = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flIncluirUsuarioControleAcessoSistema Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

