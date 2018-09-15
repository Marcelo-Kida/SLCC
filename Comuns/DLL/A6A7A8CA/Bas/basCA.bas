Attribute VB_Name = "basCA"
'Funções genéricas e Atalhos para utilização de outros objetos dentro do mesmo Componente

Option Explicit

'Formats
Public Const MQFMT_NONE = "        "
Public Const MQFMT_ADMIN = "MQADMIN "
Public Const MQFMT_CHANNEL_COMPLETED = "MQCHCOM "
Public Const MQFMT_CICS = "MQCICS  "
Public Const MQFMT_COMMAND_1 = "MQCMD1  "
Public Const MQFMT_COMMAND_2 = "MQCMD2  "
Public Const MQFMT_DEAD_LETTER_HEADER = "MQDEAD  "
Public Const MQFMT_DIST_HEADER = "MQHDIST "
Public Const MQFMT_EVENT = "MQEVENT "
Public Const MQFMT_IMS = "MQIMS   "
Public Const MQFMT_IMS_VAR_STRING = "MQIMSVS "
Public Const MQFMT_MD_EXTENSION = "MQHMDE  "
Public Const MQFMT_PCF = "MQPCF   "
Public Const MQFMT_REF_MSG_HEADER = "MQHREF  "
Public Const MQFMT_RF_HEADER = "MQHRF   "
Public Const MQFMT_RF_HEADER_2 = "MQHRF2  "
Public Const MQFMT_STRING = "MQSTR   "
Public Const MQFMT_TRIGGER = "MQTRIG  "
Public Const MQFMT_WORK_INFO_HEADER = "MQHWIH  "
Public Const MQFMT_XMIT_Q_HEADER = "MQXMIT  "

Public Const ADO_CONNECTION_TIMEOUT = 30

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public gstrConexao                          As String
Public gstrConexaoRestore                   As String
Public gstrQMgrName                         As String
Public gstrReplyQMgrName                    As String

'Variável utilizada para tratamento de erros
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Public gblnVariaveisAmbienteOK              As Boolean

'Obter String de conexão

Public Function fgObterVariaveisAmbiente()

Dim blnContinua                              As Boolean
Dim strSenha                                 As String
    
On Error GoTo ErrHandler

    If gblnVariaveisAmbienteOK Then
        Exit Function
    End If
    
    gblnVariaveisAmbienteOK = True
            
    blnContinua = False
    If Len(gstrConexao) = 0 Then
        'A variável de Ambiente (SLCC_Ambiente) define o Path\Name do aquivo UDL
        gstrConexao = flGetEnvironmentVariable("SLCC_Ambiente")
        gstrConexao = flGetPrivateProfileString("OleDB", "Provider", , gstrConexao)
        gstrConexao = "Provider=" & gstrConexao
        If InStr(1, UCase(gstrConexao), "PASSWORD") = 0 Then
            If Len(strSenha) = 0 Then
            End If
            gstrConexao = gstrConexao & ";password=" & strSenha
        End If
    End If
    blnContinua = False
 
    'Variável de Ambiente para obter o QManager, caso não seja encontrado continuar,
    'pois utilizara o QManager default
    blnContinua = True
    gstrQMgrName = flGetEnvironmentVariable("SLCC_QMgrName")
    blnContinua = False
    
    blnContinua = True
    gstrReplyQMgrName = flGetEnvironmentVariable("SLCC_ReplyQMgrName")
    blnContinua = False
    
    blnContinua = True
    gblnLog = False
    'gblnLog = Trim(Replace(flGetEnvironmentVariable("SLCC_LOG"), Chr(0), "")) <> vbNullString
    blnContinua = False
 
    Exit Function

ErrHandler:
    
    If blnContinua Then
        Err.Clear
        lngCodigoErroNegocio = 0
        Resume Next
    End If
    
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basCA", "fgObterVariaveisAmbiente", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

' Função genérica para tratamento de erros ocorridos no sistema.

Public Function fgErroLoadXML(ByRef pxmlDocument As MSXML2.DOMDocument40, ByVal psComponente As String, ByVal psClasse As String, ByVal psMetodo As String) As Variant
    Err.Raise pxmlDocument.parseError.errorCode, psComponente & " - " & psClasse & " - " & psMetodo, pxmlDocument.parseError.reason
End Function

' Retorna variáveis de ambiente.

Public Function flGetEnvironmentVariable(ByVal pstrNomeVarAmbiente As String) As String

Dim lngReturnCode                            As Long
Dim strValorVarAmbiente                      As String

On Error GoTo ErrorHandler
    
    strValorVarAmbiente = String(1000, Chr(0))
        
    lngReturnCode = GetEnvironmentVariable(UCase(pstrNomeVarAmbiente), strValorVarAmbiente, Len(strValorVarAmbiente))
    
    If lngReturnCode <> 0 Then
        flGetEnvironmentVariable = Mid(strValorVarAmbiente, 1, InStr(1, strValorVarAmbiente, Chr(0)) - 1)
    Else
        'Variável de ambiente não cadastrada
        'lngCodigoErroNegocio = 2
        'GoTo ErrorHandler
        lngCodigoErroNegocio = 0
        Err.Raise vbObjectError, "flGetEnvironmentVariable", "Variável de ambiente não cadastrada"
    End If
            
    Exit Function

ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basCA", "flGetEnvironmentVariable", lngCodigoErroNegocio, intNumeroSequencialErro, "VAriável de Ambiente:" & pstrNomeVarAmbiente)
    
End Function

' Retorna profile para string de conexão.

Private Function flGetPrivateProfileString(ByVal pstrSecao As String, _
                                           ByVal pstrChave As String, _
                                  Optional ByVal pstrPathName As String = "App.Path", _
                                  Optional ByVal pstrArqNome As String = "App.EXEName") As String


Dim lngReturnCode                            As Long
Dim strValorKey                              As String

On Error GoTo ErrorHandler
    
    strValorKey = String(1000, Chr(0))
    
    lngReturnCode = GetPrivateProfileString(pstrSecao, pstrChave, "", strValorKey, Len(strValorKey), pstrArqNome)
    
    If lngReturnCode <> 0 Then
        flGetPrivateProfileString = Mid(strValorKey, 1, InStr(1, strValorKey, Chr(0)) - 1)
    Else
        'Arquivo de configuração (.udl) de conexão não encontrado
        'lngCodigoErroNegocio = 14
        'GoTo ErrorHandler
        lngCodigoErroNegocio = 0
        Err.Raise vbObjectError, "flGetPrivateProfileString", "Arquivo de configuração (.udl) de conexão não encontrado"
    End If
        
    Exit Function
    
ErrorHandler:
    If lngCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basCA", "flGetPrivateProfileString", lngCodigoErroNegocio, intNumeroSequencialErro, pstrArqNome)
End Function

' Função genérica para tratamento de erros.

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal pstrClasse As String, _
                        ByVal pstrMetodo As String, _
                        ByRef plngCodigoErroNegocio As Long, _
               Optional ByRef pintNumeroSequencialErro As Integer = 0, _
               Optional ByVal pstrComplemento As String = "")

Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim lngErrNumber                            As Long
Dim strErrSource                            As String
Dim strErrDescription                       As String
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

    objLogErro.RaiseError pstrComponente, _
                          pstrClasse, _
                          pstrMetodo, _
                          plngCodigoErroNegocio, _
                          lngErrNumber, _
                          strErrSource, _
                          strErrDescription, _
                          pintNumeroSequencialErro, _
                          pstrComplemento, _
                          Err
    
    Set objLogErro = Nothing
    
    Err.Raise lngErrNumber, strErrSource, strErrDescription

End Sub

' Trata erro de exclusão de registro de tabela.

Public Function fgTratarErroExclusaoFisica() As Boolean

    'Erro de constraint : Encontrado(s) filho(s) para um registro pai.
    If InStr(1, Err.Description, "ORA-02292") <> 0 Then
        fgTratarErroExclusaoFisica = True
    End If

End Function

Public Function fgIsProduction() As Boolean

    'Nick - verificação de ambiente
    '26/06/2016
    
    If InStr(1, gstrConexao, "ORAPR052") <> 0 Then
        fgIsProduction = True
    End If

End Function


' Trata erro de inclusão de registro com primary key duplicada.
Public Function fgTratarErroInclusaoDuplicada() As Boolean

    'ORA-00001: restrição exclusiva violada
    If InStr(1, Err.Description, "ORA-00001") <> 0 Then
        fgTratarErroInclusaoDuplicada = True
    End If
    
End Function


' Trata erro de inclusão de operação com co_oper_ativ duplicado
Public Function fgTratarInclusaoDuplicadaOperacao() As Boolean
    
    If InStr(1, Err.Description, "Identificador da Operação já existe") <> 0 Then
        fgTratarInclusaoDuplicadaOperacao = True
    End If
    
End Function

