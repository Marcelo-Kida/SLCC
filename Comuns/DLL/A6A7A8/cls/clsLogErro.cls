VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLogErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


'Encapsular chamada para Classe de Log de Erros

Option Explicit

'Raise Error

Public Sub RaiseError(ByVal psComponente As String, ByVal psClasse As String, _
                      ByVal psMetodo As String, ByRef plCodigoErroNegocio As Long, _
             Optional ByRef piNumeroSequencialErro As Integer = 0, _
             Optional ByVal psComplemento As String = "")
                      
Dim objLogErro                              As A6A7A8CA.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String
    
    Set objLogErro = CreateObject("A6A7A8CA.clsLogErro")

    objLogErro.RaiseError psComponente, _
                          psClasse, _
                          psMetodo, _
                          plCodigoErroNegocio, _
                          ErrNumber, _
                          ErrSource, _
                          ErrDescription, _
                          piNumeroSequencialErro, _
                          psComplemento, _
                          Err
    
    Set objLogErro = Nothing
    
    Err.Raise ErrNumber, ErrSource, ErrDescription
    
End Sub
