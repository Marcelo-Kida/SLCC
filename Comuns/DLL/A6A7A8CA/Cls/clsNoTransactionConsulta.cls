VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNoTransactionConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : CA
'Classe         : clsNoTransactionConsulta
'Data Criação   : 01-05-2001 17:13
'Objetivo       : Execução de Querys de Consultas
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

'Se piQtdeRegistros = 0, retorna todos os registros selecionados
Public Function QuerySQL(ByVal psSQL As String, _
                Optional ByVal piQtdeRegistros As Integer) As ADODB.Recordset

On Error GoTo ErrHandler

Dim objADODBConnection                        As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset

    Set QuerySQL = Nothing
    
    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")

    objADODBConnection.Mode = adModeRead
    
    objADODBConnection.Open gsConexao

    objADODBConnection.CursorLocation = adUseClient
    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic

    If piQtdeRegistros > 0 Then
        rsDisconnect.MaxRecords = piQtdeRegistros
    End If
    
    Set rsDisconnect = objADODBConnection.Execute(psSQL, , adCmdText)
            
    Set rsDisconnect.ActiveConnection = Nothing
        
    Set QuerySQL = rsDisconnect
    
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    Set objADODBConnection = Nothing
    
    Exit Function

ErrHandler:
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QuerySQL", llCodigoErroNegocio, liNumeroSequencialErro)
End Function

Public Function QueryXML(ByVal psSQL As String, _
                Optional ByVal pbCompressao As Boolean = False) As String
On Error GoTo ErrHandler
Dim objADODBConnection                        As ADODB.Connection
Dim rsDisconnect                            As ADODB.Recordset
Dim objDOMDocument                          As DOMDocument40

    fgObterVariaveisAmbiente
    
    Set objADODBConnection = CreateObject("ADODB.Connection")
    Set rsDisconnect = CreateObject("ADODB.Recordset")
    Set objDOMDocument = New DOMDocument40
    
    objADODBConnection.Open gsConexao
    
    objADODBConnection.CursorLocation = adUseClient
    rsDisconnect.CursorType = adOpenStatic
    rsDisconnect.LockType = adLockBatchOptimistic
    
    Set rsDisconnect = objADODBConnection.Execute(psSQL, , adCmdText)
        
    rsDisconnect.Save objDOMDocument, adPersistXML
        
    QueryXML = objDOMDocument.xml
    Set objDOMDocument = Nothing
    
    If pbCompressao Then
        CompressString QueryXML
    End If
    
    If rsDisconnect.State <> adStateClosed Then
        rsDisconnect.Close
    End If
    Set rsDisconnect = Nothing
    
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    Set objADODBConnection = Nothing

    Exit Function

ErrHandler:
    If objADODBConnection.State <> adStateClosed Then
        objADODBConnection.Close
    End If
    
    QueryXML = vbNullString
    
    Set objDOMDocument = Nothing
    Set rsDisconnect = Nothing
        
    If llCodigoErroNegocio <> 0 And Err.Number = 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, TypeName(Me), "QueryXML", llCodigoErroNegocio, liNumeroSequencialErro)
End Function


