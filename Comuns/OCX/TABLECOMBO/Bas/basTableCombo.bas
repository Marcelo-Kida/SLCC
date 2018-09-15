Attribute VB_Name = "basTableCombo"
'Empresa        : Regerbanc
'Componente     : Table Combo
'Classe         :
'Data Criação   : 22/07/203
'Objetivo       :
'Analista       : Carlos Fortes
'
'Programador    : Cassiano Nicolosi
'Data           : 08/08/2003
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Public Function fgErroLoadXML(ByRef objDOMDocument As MSXML2.DOMDocument40, _
                              ByVal psComponente As String, _
                              ByVal psClasse As String, _
                              ByVal psMetodo As String)

    Err.Raise objDOMDocument.parseError.errorCode, psComponente & " - " & psClasse & " - " & psMetodo, objDOMDocument.parseError.reason

End Function
