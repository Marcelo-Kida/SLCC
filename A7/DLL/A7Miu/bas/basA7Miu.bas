Attribute VB_Name = "basA7Miu"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3DB4018D00B0"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Module"
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : BUSMIU
'Classe         : basMIU
'Data Criação   : 14-10-2002 12:46
'Objetivo       : Funções genéricas e Atalhos para utilização de outros objetos
'                 dentro do mesmo Componente
'Analista       : Carlos Fortes
'
'Programador    : Carlos Fortes
'Data           : 14-10-2002 12:46
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

'**************************************************************************
'Encapsular chamada para Classe de Log de Erros
 Public Sub fgRaiseError(ByVal pstrComponente As String, _
                         ByVal pstrClasse As String, _
                         ByVal pstrMetodo As String, _
                         ByRef plngCodigoErroNegocio As Long, _
                Optional ByRef pintNumeroSequencialErro As Integer = 0, _
                Optional ByVal pstrComplemento As String = "")


Dim objLogErro                              As A6A7A8.clsLogErro
Dim ErrNumber                               As Long
Dim ErrSource                               As String
Dim ErrDescription                          As String
    
    Set objLogErro = CreateObject("A6A7A8.clsLogErro")

    objLogErro.RaiseError pstrComponente, _
                          pstrClasse, _
                          pstrMetodo, _
                          plngCodigoErroNegocio, _
                          pintNumeroSequencialErro, _
                          pstrComplemento

    Set objLogErro = Nothing
    
    Err.Raise ErrNumber, ErrSource, ErrDescription
    
'**************************************************************************
End Sub

