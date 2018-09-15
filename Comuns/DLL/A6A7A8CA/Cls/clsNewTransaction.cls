VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNewTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Componente     : PJPKCA
'Classe         : clsTransacao
'Data Criação   : 01-05-2001 17:13
'Objetivo       : Execução de Querys de Ação (Insert, Update, Delete) através de
'                 sentenças SQL ou Stored Procedures
'Analista       : Adilson Gonçalves Damasceno
'
'Programador    : Adilson Gonçalves Damasceno
'Data           : 01-05-2001 17:13
'
'Teste          :
'Autor          :
'
'Data Alteração : 04-09-2001 10:55
'Autor          : Adilson Gonçalves Damasceno
'Objetivo       : Separar Fisica/Lógicamente as Querys de Consultas da Querys de Ação

Option Explicit

'Variável utilizada para tratamento de erros
Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

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

Public Function ExecuteSQL(ByVal psSQL As String, _
                  Optional ByVal pbConexaoPK As Boolean = False) As Long
On Error GoTo ErrHandler

Dim objTransacao                            As PJPKCA.clsTransacao
    
    Set objTransacao = CreateObject("PJPKCA.clsTransacao")
    ExecuteSQL = objTransacao.ExecuteSQL(psSQL, pbConexaoPK)
    Set objTransacao = Nothing
    
    Exit Function

ErrHandler:
    Set objTransacao = Nothing
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basPJGestaoCaixa", "ExecuteSQL", llCodigoErroNegocio, liNumeroSequencialErro)

End Function

Public Function ExecuteCMD(ByVal psNomeProc As String, _
                           ByVal piPosicaoRetorno As Integer, _
                           ByRef pvParametros() As Variant, _
                  Optional ByVal pbConexaoPK As Boolean = False) As Long

On Error GoTo ErrHandler
        
Dim objTransacao                            As PJPKCA.clsTransacao
    
    Set objTransacao = CreateObject("PJPKCA.clsTransacao")
    ExecuteCMD = objTransacao.ExecuteCMD(psNomeProc, piPosicaoRetorno, pvParametros(), pbConexaoPK)
    Set objTransacao = Nothing
    
    Exit Function
    
ErrHandler:
    
    Set objTransacao = Nothing
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basPJGestaoCaixa", "ExecuteSQL", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Function

Public Function ExecuteCMDMensageria(ByVal psNomeProc As String, _
                                     ByVal piPosicaoRetorno As Integer, _
                                     ByRef pvParametros() As Variant, _
                            Optional ByVal pbConexaoPK As Boolean = False, _
                            Optional ByVal piCommandType As ADODB.CommandTypeEnum = adCmdStoredProc) As Long

On Error GoTo ErrHandler
    
Dim objTransacao                            As PJPKCA.clsTransacao
    
    Set objTransacao = CreateObject("PJPKCA.clsTransacao")
    ExecuteCMDMensageria = objTransacao.ExecuteCMDMensageria(psNomeProc, piPosicaoRetorno, pvParametros(), True, piCommandType)
    Set objTransacao = Nothing
        
    Exit Function
    
ErrHandler:
    Set objTransacao = Nothing
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basPJGestaoCaixa", "ExecuteSQL", llCodigoErroNegocio, liNumeroSequencialErro)

End Function

'Retorna o numero da sequencia da tabela tb_agenda
Public Function fgSequencia(ByVal psTabela As String, _
                   Optional ByVal plEmpresa As Long, _
                   Optional ByVal pbConexaoPK As Boolean = False) As Long

On Error GoTo ErrorHandler

Dim objTransacao                            As PJPKCA.clsTransacao
    
    Set objTransacao = CreateObject("PJPKCA.clsTransacao")
    fgSequencia = objTransacao.fgSequencia(psTabela, plEmpresa, pbConexaoPK)
    Set objTransacao = Nothing
    
    Exit Function
    
ErrorHandler:
    Set objTransacao = Nothing
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "basPJGestaoCaixa", "ExecuteSQL", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Function

