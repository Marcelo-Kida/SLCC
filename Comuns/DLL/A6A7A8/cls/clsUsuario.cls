VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

'Objeto responsável pelo controle de acesso do usuário


Option Explicit


Private intNumeroSequencialErro             As Integer
Private lngCodigoErroNegocio                As Long

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Const MAX_COMPUTERNAME_LENGTH       As Long = 31
Private Const NORMAL_PRIORITY_CLASS = &H20&

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

'Obter o login do usuário

Public Function ObterUsuarioRede() As String

Dim objSecurity                             As COMSVCSLib.SecurityCallContext
Dim objIdentity                             As COMSVCSLib.SecurityIdentity
Dim objActiveDS                             As ActiveDs.WinNTSystemInfo
Dim objControleAcesso                       As A6A7A8.clsControleAcesso
Dim strAccount                              As String
Dim strUsuario                              As String

On Error GoTo ErrorHandler

    If Not ObjectContext Is Nothing Then

        Set objSecurity = GetSecurityCallContext()
        Set objIdentity = objSecurity.Item("OriginalCaller")
        
        '*****************************************************************
        'Ver uma maneira melhor de tratar esta situação
        Err.Clear
        On Error Resume Next

        strAccount = objIdentity.Item("AccountName")
        
        If Err.Number = 5 Then
            strUsuario = "SISTEMA"
        ElseIf Err.Number = 0 Then
            strUsuario = Right(strAccount, Len(strAccount) - InStr(strAccount, "\"))
        Else
            GoTo ErrorHandler
        End If
        
        Err.Clear
        
        On Error GoTo ErrorHandler
        'Ver uma maneira melhor de tratar esta situação
        '*****************************************************************

        Set objIdentity = Nothing
        Set objSecurity = Nothing
    Else
        Set objActiveDS = New ActiveDs.WinNTSystemInfo
        strUsuario = objActiveDS.UserName
        Set objActiveDS = Nothing
    End If
        
    strUsuario = Trim$(strUsuario)
    
    If Len(strUsuario) > 8 Then
        strUsuario = UCase$(Left$(strUsuario, 8))
    End If
    
    If Not ObjectContext Is Nothing Then

        Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
        If Not objControleAcesso.PesquisaUsuario(UCase$(strUsuario), "", 0, 0) Then
            strUsuario = "SISTEMA"
        End If
        Set objControleAcesso = Nothing

    End If
    
    ObterUsuarioRede = strUsuario
    
    Exit Function
ErrorHandler:
    
    Set objSecurity = Nothing
    Set objIdentity = Nothing
    Set objActiveDS = Nothing
        
    ObterUsuarioRede = "SISTEMA"
    
    Err.Clear

End Function

'Adicionar o usuário a lista de usuários logados (Controle de Acesso)

Public Function AdicionaUsuario(ByVal pstrUsuarioRede As String, _
                                ByVal pstrUsuarioSistema As String, _
                                ByVal pstrEstacaoTrabalho As String, _
                                ByVal pblnForcaInclusao As Boolean) As Boolean

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")

    'Adiciona o usuário no recordset
    If Not objControleAcesso.AdicionaUsuario(pstrUsuarioRede, _
                                             pstrUsuarioSistema, _
                                             pstrEstacaoTrabalho, _
                                             pblnForcaInclusao, _
                                             lngCodigoErroNegocio, _
                                             intNumeroSequencialErro) Then
        GoTo ErrorHandler
    End If

    AdicionaUsuario = True
    
    Set objControleAcesso = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AdicionaUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Remover o usuário a lista de usuários logados (Controle de Acesso)

Public Function RemoveUsuario(ByVal pstrUsuarioRede As String) As Boolean
   

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    
    If Not objControleAcesso.RemoveUsuario(pstrUsuarioRede, lngCodigoErroNegocio) Then
        GoTo ErrorHandler
    End If

    RemoveUsuario = True
    
    Set objControleAcesso = Nothing
    
    Exit Function
ErrorHandler:
    
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "RemoveUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Verificar se usuário está logado (Controle de Acesso)

Public Function PesquisaUsuario(ByVal pstrUsuarioRede As String) As String

'Retorna uma string com os dados dos usuários concatenados

Dim objControleAcesso                       As A6A7A8.clsControleAcesso
Dim strRetornoPesquisa                      As String

On Error GoTo ErrorHandler

    If Not ObjectContext Is Nothing Then
        
        Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
        objControleAcesso.PesquisaUsuario pstrUsuarioRede, strRetornoPesquisa, lngCodigoErroNegocio, intNumeroSequencialErro
        Set objControleAcesso = Nothing
        
        If lngCodigoErroNegocio <> 0 Then
            GoTo ErrorHandler
        End If
    Else
        strRetornoPesquisa = ObterUsuarioRede()
    End If
    
    PesquisaUsuario = strRetornoPesquisa

    Exit Function
ErrorHandler:
        
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "PesquisaUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'Listar usuários logados (Controle de Acesso)

Public Function ListarUsuariosAtivos() As String
    
Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler

    Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
    
    ListarUsuariosAtivos = objControleAcesso.ListarUsuarios
    
    Set objControleAcesso = Nothing
    
    Exit Function

ErrorHandler:
    
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ListarUsuariosAtivos Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function


Public Function ObterUsuarioSistema(ByVal pstrUsuarioRede As String) As String

'Retorna uma string com os dados dos usuários concatenados
Dim objControleAcesso                       As A6A7A8.clsControleAcesso
Dim strRetornoPesquisa                      As String

On Error GoTo ErrorHandler

    If Not ObjectContext Is Nothing Then
        
        Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
        objControleAcesso.ObterUsuarioSistema pstrUsuarioRede, strRetornoPesquisa, lngCodigoErroNegocio, intNumeroSequencialErro
        Set objControleAcesso = Nothing
        
        If lngCodigoErroNegocio <> 0 Then
            GoTo ErrorHandler
        End If
    Else
        strRetornoPesquisa = ObterUsuarioRede()
    End If
    
    ObterUsuarioSistema = UCase(strRetornoPesquisa)

    Exit Function
    
ErrorHandler:
        
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterUsuarioSistema Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

'Obter o nome da estação de trabalho de um usuário logado (Controle de Acesso)

Public Function ObterEstacaoTrabalhoUsuario(ByVal pstrUsuarioRede As String, _
                                            ByRef pstrRetornoEstacaoTrabalho As String, _
                                            ByRef plngCodigoErro As Long, _
                                            ByRef pintNumeroSequencialErro As Integer) As Boolean

Dim objControleAcesso                       As A6A7A8.clsControleAcesso

On Error GoTo ErrorHandler
    
    If UCase$(Trim$(pstrUsuarioRede)) = "SISTEMA" Then
        pstrRetornoEstacaoTrabalho = "SERVIDOR"
        ObterEstacaoTrabalhoUsuario = True
    Else
        #If DebugMode = 1 Then
            pstrRetornoEstacaoTrabalho = flObterEstacaoTrabalho
            ObterEstacaoTrabalhoUsuario = True
        #Else
            Set objControleAcesso = CreateObject("A6A7A8.clsControleAcesso")
            ObterEstacaoTrabalhoUsuario = objControleAcesso.ObterEstacaoTrabalhoUsuario(pstrUsuarioRede, pstrRetornoEstacaoTrabalho, plngCodigoErro, pintNumeroSequencialErro)
            Set objControleAcesso = Nothing
            ObterEstacaoTrabalhoUsuario = True
        #End If
    End If

    Exit Function
ErrorHandler:
    
    Set objControleAcesso = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterEstacaoTrabalhoUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro, "Usuário: " & pstrUsuarioRede)
    
End Function

'Obter o nome da estação de trabalho Local

Private Function flObterEstacaoTrabalho() As String

Dim strEstacao                               As String
Dim lngLen                                   As Long

On Error GoTo ErrorHandler
    
    lngLen = MAX_COMPUTERNAME_LENGTH + 1
    strEstacao = String(lngLen, "X")
    
    GetComputerName strEstacao, lngLen
    strEstacao = Left$(strEstacao, lngLen)
    flObterEstacaoTrabalho = strEstacao

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "modStartScreen", "flObterEstacaoTrabalho Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

'LogOn usuario

Public Function LogOn(ByVal pstrUsuarioSistema As String, _
                      ByVal pstrUsuarioRede As String, _
                      ByVal pstrEstacaoTrabalho As String, _
             Optional ByVal pblnForcaLogar As Boolean) As Boolean

Dim lngCodigoRetorno                        As Long

On Error GoTo ErrorHandler

    Select Case vbNullString
        Case Trim(pstrUsuarioSistema)
            'É obrigatório informar o usuário.
            lngCodigoErroNegocio = 19
            GoTo ErrorHandler
    End Select
    
    LogOn = AdicionaUsuario(pstrUsuarioRede, pstrUsuarioSistema, pstrEstacaoTrabalho, pblnForcaLogar)
    
    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "LogOn Function", lngCodigoErroNegocio, intNumeroSequencialErro, IIf(lngCodigoRetorno <> 0, "Erro US: " & CStr(lngCodigoRetorno), vbNullString))

End Function

'Logoff usuario

Public Function Logoff(ByVal pstrUsuarioRede As String) As Boolean

On Error GoTo ErrorHandler
    
    Logoff = RemoveUsuario(pstrUsuarioRede)

    Exit Function
ErrorHandler:

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Logoff Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

Public Function Reconectar(ByVal pstrUsuarioSistema As String, _
                           ByVal pstrUsuarioRede As String, _
                           ByVal pstrEstacaoTrabalho As String) As Boolean

On Error GoTo ErrorHandler

    Reconectar = AdicionaUsuario(pstrUsuarioRede, pstrUsuarioSistema, pstrEstacaoTrabalho, False)

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "Reconectar Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Public Function ObterUsuario() As String

On Error GoTo ErrorHandler
    
    ObterUsuario = ObterUsuarioRede()

    Exit Function
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "ObterUsuario Function", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Function

