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

'Objeto responsável pela validação da alaçada (Sistema LG)

Option Explicit

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Enum enumTipoValidacaoAlcada
    Digitacao = 1
    Aprovacao = 2
End Enum

Private objContext                          As COMSVCSLib.ObjectContext
Private strConstruct                        As String
Private Const lngOK                         As Long = 0

Implements COMSVCSLib.ObjectControl
Implements COMSVCSLib.IObjectConstruct

Private Sub ObjectControl_Deactivate()
    Set objContext = Nothing
End Sub

Private Function ObjectControl_CanBePooled() As Boolean
    ObjectControl_CanBePooled = True
End Function

Private Sub ObjectControl_Activate()
    Set objContext = COMSVCSLib.GetObjectContext()
End Sub

Private Sub IObjectConstruct_Construct(ByVal pobjContruct As Object)
    strConstruct = pobjContruct.ConstructString
End Sub

'Validação da alçada , Perfil digitador e aprovador

Private Sub flValidarAlcada(ByVal pstrUsuario As String, _
                            ByVal pstrValor As String, _
                            ByVal plngTipoValidacao As enumTipoValidacaoAlcada, _
                            ByRef pvntProtocolo As Variant)

Dim objMBS                                  As Object
Dim lngCodigoRetorno                        As Long
Dim vntNomeUsuario                          As Variant
Dim vntCodPrestadorServ                     As Variant
Dim lngEmpresa                              As Long
Dim vntMatricula                            As Variant
Dim strSistema                              As String
Dim strMsgRetorno                           As String
Dim strCompMBS                              As String

On Error GoTo ErrorHandler

    #If DebugMode = 1 Then
        If plngTipoValidacao = enumTipoValidacaoAlcada.Digitacao Then
            pvntProtocolo = Format(Now, "ddhhmmss")
        End If
        Exit Sub
    #End If

    If Len(Trim(pstrUsuario)) > 8 Then
        pstrUsuario = Mid(Trim(pstrUsuario), 1, 8)
    End If

    'Desligar caso ocorra problemas com validação de Alçada / Limite
    If UCase(Trim(strConstruct)) = "REGER" Then
        vntMatricula = 1324
        vntCodPrestadorServ = 523
        
        If plngTipoValidacao = enumTipoValidacaoAlcada.Digitacao Then
            pvntProtocolo = Format(Now, "ddhhmmss")
        End If

    Else
        'buscar matricula do usuário para repassar para o sistema LG
        strCompMBS = fgGetPrivateProfileString("MBS", "Componente", App.Path & "\SLCC.INI")
        Set objMBS = CreateObject(strCompMBS)
        lngCodigoRetorno = objMBS.recuperaUsuario(pstrUsuario, vntNomeUsuario, vntMatricula, vntCodPrestadorServ)
        Set objMBS = Nothing
    End If
    
    If (lngCodigoRetorno <> lngOK) Or _
       (vntMatricula <= 0) Or _
       (vntMatricula = vbNullString) Or _
       (vntCodPrestadorServ <= 0) Then
       
        'Obter Erro MBS
        Call fgObterErroMBS(lngCodigoRetorno, lngCodigoErroNegocio, strMsgRetorno)
        
        Call fgGravaArquivo("Usuário: " & pstrUsuario & vbCrLf & _
                            "Codigo Retorno MBS: " & lngCodigoRetorno & " - " & strMsgRetorno & vbCrLf & _
                            "Protocolo: " & pvntProtocolo & vbCrLf & _
                            "Tipo Validação: " & IIf(plngTipoValidacao = Aprovacao, "Aprovação", "Digitação"))

       
        GoTo ErrorHandler
    End If
    
    'atribuido valores fixos devido o uso de usuarios do dominio si_brasil e não possuir
    'estes atributos valcione 07/04/2002
    lngEmpresa = vntCodPrestadorServ
    strSistema = "A8"
    
    Select Case plngTipoValidacao
        Case Digitacao
            Call flDigitarAlcada(lngCodigoRetorno, _
                                 lngEmpresa, _
                                 vntMatricula, _
                                 pstrValor, _
                                 strSistema, _
                                 pvntProtocolo, _
                                 strMsgRetorno)
        Case Aprovacao
            Call flAprovarAlcada(lngCodigoRetorno, _
                                 pvntProtocolo, _
                                 lngEmpresa, _
                                 vntMatricula, _
                                 strMsgRetorno)
    
    End Select
    
    Exit Sub
ErrorHandler:
    
    Set objMBS = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flValidarAlcadaBanespa Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub
            
''Chamada para o comonente do LG (LGAT0AXXa.clsLGAT0A24) - Validação de limites e Alçada
            
Private Sub flAprovarAlcada(ByRef plngRetorno As Long, _
                            ByVal pvntProtocolo As Variant, _
                            ByVal plngEmpresa As Long, _
                            ByVal pvntMatricula As Long, _
                            ByRef pstrMsgRetorno As String)

Dim objLG                                   As Object
Dim lngCodigoRetorno                        As Long

On Error GoTo ErrorHandler
    
    Set objLG = CreateObject("LGAT0AXXa.clsLGAT0A24")
    
    lngCodigoRetorno = objLG.fiAutOper(plngRetorno, _
                                       pvntProtocolo, _
                                       plngEmpresa, _
                                       pvntMatricula, _
                                       pstrMsgRetorno)
    Set objLG = Nothing
    
    
    Select Case plngRetorno
        Case 0 'lngOK
        
        Case 1  'Usuário não Cadastrado - Sistema LG
            lngCodigoErroNegocio = 23
            GoTo ErrorHandler
        
        Case 2  'Usuário não possui Limite / Alçada
            lngCodigoErroNegocio = 24
            GoTo ErrorHandler
        
        Case 3  'Usuário não possui Limite / Alçada - Sistema LG
            lngCodigoErroNegocio = 25
            GoTo ErrorHandler
        
        Case 4  'Usuário não possui Limite / Alçada - Sistema LG
            lngCodigoErroNegocio = 25
            GoTo ErrorHandler
        
        Case 5 'Usuário que digitou operação não pode aprovar a mesma
            lngCodigoErroNegocio = 26
            GoTo ErrorHandler
        
        Case 6  'Tamanho de Commarea Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 27
            GoTo ErrorHandler
        
        Case 7  'Código de Protocólo Enviado ao Sistema LG - Inválido.
            lngCodigoErroNegocio = 28
            GoTo ErrorHandler
        
        Case 8  'Código de Empresa Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 29
            GoTo ErrorHandler
        
        Case 9  'Código de Matrícula Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 30
            GoTo ErrorHandler
        
        Case 11 'Número do protocolo inexistente - Sistema LG
            lngCodigoErroNegocio = 31
            GoTo ErrorHandler
            
        Case 13 'Transação já aprovada
            lngCodigoErroNegocio = 32
            GoTo ErrorHandler
            
        Case 99 'Problemas na Rotina LGAT024 - Ver aplicação DR73 no CICS.
            lngCodigoErroNegocio = 67
            GoTo ErrorHandler
            
        Case Else   'Problema na rotina de validação no Sistema LG
            lngCodigoErroNegocio = 33
            GoTo ErrorHandler
    
    End Select
                                                        
    Exit Sub
ErrorHandler:
    
    Set objLG = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flAprovarAlcadaBanespa Sub", lngCodigoErroNegocio, intNumeroSequencialErro, IIf(plngRetorno <> 0, "Erro LG: " & CStr(plngRetorno), vbNullString))

End Sub

'Verificar digitação da alçada

Public Sub DigitarAlcada(ByVal pstrUsuario As String, _
                         ByVal pstrValor As String, _
                Optional ByRef pvntProtocolo As Variant)

Dim intDominio                               As Long

On Error GoTo ErrorHandler

    'Desligar caso ocorra problemas com validação de Alçada / Limite
    If UCase(Trim(strConstruct)) = "OFF" Then
        pvntProtocolo = Format$(Now, "DDHHMMSS")
        Exit Sub
    End If
            
    'Consiste os parâmetros
    Select Case vbNullString
        
        Case Trim(pstrUsuario)
            'É obrigatório informar o usuário
            lngCodigoErroNegocio = 19
            GoTo ErrorHandler
    
        Case Trim(pstrValor)
            
            pstrValor = "0"
    
    End Select
           
    pstrUsuario = UCase$(pstrUsuario)
    pstrValor = fgVlrToDBServer(pstrValor)
    
    Call flValidarAlcada(pstrUsuario, pstrValor, Digitacao, pvntProtocolo)
       
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "DigitarAlcada Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Chamada para o componentes do LG (LGAT0AXXa.clsLGAT0A23) - Validação de limites e Alçada

Private Sub flDigitarAlcada(ByRef plngRetorno As Long, _
                            ByVal plngEmpresa As Long, _
                            ByVal pvntMatricula As Long, _
                            ByVal pstrValor As Variant, _
                            ByVal pstrSistema As String, _
                            ByRef pvntProtocolo As Variant, _
                            ByRef pstrMsgRetorno As String)

Dim objLG                                   As Object
Dim lngCodigoRetorno                        As Long
Dim lngProduto                              As Long
Dim lngTipoUniorg                           As Long
Dim lngCodUniorg                            As Long
Dim strAtributoUsuario                      As String

On Error GoTo ErrorHandler


    'atribuido valores fixos devido o uso de usuarios do dominio si_brasil e não possuir
    'estes atributos valcione 07/04/2002
    lngProduto = 0
    lngTipoUniorg = 0
    lngCodUniorg = 0
    strAtributoUsuario = " "
    
    Set objLG = CreateObject("LGAT0AXXa.clsLGAT0A23")
    
    lngCodigoRetorno = objLG.fiDigOper(plngRetorno, _
                                       plngEmpresa, _
                                       pvntMatricula, _
                                       pstrSistema, _
                                       lngProduto, _
                                       lngTipoUniorg, _
                                       lngCodUniorg, _
                                       strAtributoUsuario, _
                                       pstrValor, _
                                       pvntProtocolo, _
                                       pstrMsgRetorno)
    
    Set objLG = Nothing
                                                        
    Select Case plngRetorno
        Case 0   'lngOK
        
        Case 1   'Usuário não Cadastrado - Sistema LG
            lngCodigoErroNegocio = 23
            GoTo ErrorHandler
        
        Case 2  'Usuário não Autorizado - Sistema LG
            lngCodigoErroNegocio = 34
            GoTo ErrorHandler
        
        Case 6   'Tamanho de Commarea Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 27
            GoTo ErrorHandler
        
        Case 8   'Código de Empresa Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 29
            GoTo ErrorHandler
        
        Case 9   'Código de Matrícula Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 30
            GoTo ErrorHandler
        
        Case 10  'Código do Sistema Enviado ao LG Inválido.
            lngCodigoErroNegocio = 64
            GoTo ErrorHandler
        
        Case 11  'Valor da Operação Enviado ao LG Zerado.
            lngCodigoErroNegocio = 65
            GoTo ErrorHandler
        
        Case 99  'Problemas na Rotina LGAT023 - Ver aplicação DR73 no CICS.
            lngCodigoErroNegocio = 66
            GoTo ErrorHandler
        
        Case Else   'Problema na rotina de validação no sist. LG
            lngCodigoErroNegocio = 33
            GoTo ErrorHandler
    End Select
                                                        
    Exit Sub
ErrorHandler:
    
    Set objLG = Nothing

    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "flDigitarAlcadaBanespa Sub", lngCodigoErroNegocio, intNumeroSequencialErro, IIf(plngRetorno <> 0, "Erro LG: " & CStr(plngRetorno), vbNullString))

End Sub

'Verificação da aprovação da alçada

Public Sub AprovarAlcada(ByVal pstrUsuario As String, _
                         ByVal pstrValor As String, _
                Optional ByVal pvntProtocolo As Variant)

Dim intDominio                               As Long

On Error GoTo ErrorHandler
    
    If Trim$(pvntProtocolo) = vbNullString Then Exit Sub
    
    'Desligar caso ocorra problemas com validação de Alçada / Limite
    If UCase(Trim(strConstruct)) = "OFF" Then
        Exit Sub
    End If
           
    'Consiste os parâmetros
    Select Case vbNullString
        Case Trim(pstrUsuario)
            'É obrigatório informar o usuário
            lngCodigoErroNegocio = 19
            GoTo ErrorHandler
    
        Case Trim(pstrValor)
            pstrValor = "0"
    
    End Select
           
    pstrUsuario = UCase$(pstrUsuario)
    pstrValor = fgVlrToDBServer(pstrValor)
            
    Call flValidarAlcada(pstrUsuario, pstrValor, Aprovacao, pvntProtocolo)
       
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "AprovarAlcada Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

'Chamada para o componentes do LG (LGAT0AXXa.clsLGAT0A26) - Cancelamento de protocolo gerado na digitação

Public Sub CancelarProtocolo(ByVal pstrValor As String, _
                    Optional ByRef pvntProtocolo As Variant)

Dim objLG                                   As Object
Dim lngCodigoRetorno                        As Long
Dim vntCodRetorno                           As Variant
Dim vntMsgRetorno                           As Variant

On Error GoTo ErrorHandler

    'Desligar caso ocorra problemas com validação de Alçada / Limite
    If UCase(Trim(strConstruct)) = "OFF" Then
        Exit Sub
    End If
            
    pstrValor = fgVlrToDBServer(pstrValor)
    
    Set objLG = CreateObject("lgat0axxa.clsLGAT0A26")
    
    lngCodigoRetorno = objLG.fiCanProt(vntCodRetorno, pstrValor, "A8", vntMsgRetorno)
    
    Set objLG = Nothing
    
    Select Case lngCodigoRetorno
        Case 0 'lngOK
        
        Case 1  'Não Existe Operação a ser Cancelda. - Sistema LG
            lngCodigoErroNegocio = 23
            GoTo ErrorHandler
        
        Case 6  'Tamanho de Commarea Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 27
            GoTo ErrorHandler
        
        Case 7  'Código de Protocólo Enviado ao Sistema LG - Inválido.
            lngCodigoErroNegocio = 28
            GoTo ErrorHandler
        
        Case 8  'Código de Empresa Enviado ao Sistema LG - Invalido.
            lngCodigoErroNegocio = 29
            GoTo ErrorHandler
            
        Case Else   'Problema na rotina de cancelamento de protocolo no Sistema LG
            lngCodigoErroNegocio = 33
            GoTo ErrorHandler
    
    End Select

       
    Exit Sub
ErrorHandler:
    
    Set objLG = Nothing
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "CancelarProtocolo Sub", lngCodigoErroNegocio, intNumeroSequencialErro, vntCodRetorno & "-" & vntMsgRetorno)

End Sub




