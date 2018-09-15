Attribute VB_Name = "basMIU"
Attribute VB_Description = "Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA\r\nComponente     : MIU\r\nClasse         : basMIU\r\nData Criação   : 14-10-2002 12:46\r\nObjetivo       : Funções genéricas e Atalhos para utilização de outros objetos\r\n                 dentro do mesmo Componente\r\nAnalista       : Carlos Fortes\r\n\r\nProgramador    : Carlos Fortes\r\nData           : 14-10-2002 12:46\r\n\r\nTeste          :\r\nAutor          :\r\n\r\nData Alteração :\r\nAutor          :\r\nObjetivo       :"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3DB4018D00B0"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"

'Encapsular chamada para Classe de Log de Erros

Option Explicit

'Encapsular chamada para Classe de Log de Erros

Public Sub fgRaiseError(ByVal pstrComponente As String, _
                        ByVal psClasse As String, _
                        ByVal psMetodo As String, _
                        ByRef plCodigoErroNegocio As Long, _
               Optional ByRef piNumeroSequencialErro As Integer = 0, _
               Optional ByVal psComplemento As String = "")

Dim objLogErro                              As A6A7A8.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String

    Set objLogErro = CreateObject("A6A7A8.clsLogErro")

    objLogErro.RaiseError pstrComponente, _
                          psClasse, _
                          psMetodo, _
                          plCodigoErroNegocio, _
                          piNumeroSequencialErro, _
                          psComplemento

    
    Set objLogErro = Nothing
    
    Err.Raise ErrNumber, ErrSource, ErrDescription
    
'**************************************************************************
End Sub

