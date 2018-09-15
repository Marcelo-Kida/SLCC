VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRBLibrary"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Empresa        : Regerbanc - Partticipações , Negócios e Serviços LTDA
'Componente     : RBLibrary
'Classe         : clsRBLibrary
'Data Criação   :
'Objetivo       : Bliblioteca de Funções
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 36/07/2003
'
'Teste          :
'Autor          :
'
Option Explicit

Public Enum enumFormatoDataHora2
    Data = 1
    Hora = 2
    DataHora = 3
End Enum

Private Const gstrDataVazia                  As String = "00:00:00"
Private Const gstrMaskInt                    As String = "##,###,###,###,###,##0;[vbRed](##,###,###,###,###,##0)"
Private Const gstrMaskDec                    As String = "##,###,###,###,###,##0.00;[vbRed](##,###,###,###,###,##0.00)"

Public Function fgAppendNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                             ByVal pstrNodeContext As String, _
                             ByVal pstrNodeName As String, _
                             ByVal pstrNodeValue As String, _
                    Optional ByVal pstrNodeRepetName As String = "") As Boolean

Dim objDomNodeAux                           As MSXML2.IXMLDOMNode
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    If pstrNodeContext = "" Then
        'Se a Tag for o root passar pstrNodeContext = Nome Tag Principal
        Set objDomNodeContext = xmlDocument
    Else
        'Para criar um Grupo (Ex.:Grupo_Usuario) dentro de Uma Repeticao (Ex. Repeat_Usuario)
        'Passar o argumento pNomeRepet = Repet_Usuario
        If pstrNodeRepetName <> "" Then
            Set objDomNodeContext = xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.Item(xmlDocument.selectSingleNode("//" & pstrNodeRepetName).childNodes.length - 1)
        Else
            Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
        End If
    End If
    
    Set objDomNodeAux = xmlDocument.createElement(pstrNodeName)
    objDomNodeAux.Text = pstrNodeValue
    objDomNodeContext.appendChild objDomNodeAux

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

Public Function fgAppendAttribute(ByRef xmlDocument As MSXML2.DOMDocument40, _
                                  ByVal pstrNodeContext As String, _
                                  ByVal pstrNomeAtributo As String, _
                                  ByVal pvntValorAtributo As Variant) As String

Dim xmlAttrib                            As MSXML2.IXMLDOMAttribute
Dim objDomNodeContext                       As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set objDomNodeContext = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeContext)
    
    Set xmlAttrib = xmlDocument.createAttribute(pstrNomeAtributo)
    
    xmlAttrib.Text = pvntValorAtributo
    
    objDomNodeContext.attributes.setNamedItem xmlAttrib

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Description, Err.Description
End Function

Public Sub fgAppendXML(ByRef objMapaNavegacao As MSXML2.DOMDocument40, _
                       ByVal pstrNodeContext As String, _
                       ByVal pstrXMLFilho As String)

Dim xmlFilho                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    'Leitura do XML Filho
    Set xmlFilho = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlFilho.loadXML(pstrXMLFilho) Then
        '100 - Documento XML Inválido.
        'llCodigoErroNegocio = 100
        GoTo ErrorHandler
    End If
       
    'Setar o nivel que deverá entrar o XML Filho
     Set objDomNode = objMapaNavegacao.documentElement.selectSingleNode("//" & pstrNodeContext)

    'Adicionar XML Filho na Saida
    objDomNode.appendChild xmlFilho.childNodes.Item(0)
    
    Set xmlFilho = Nothing

    Exit Sub

ErrorHandler:
    
    Set xmlFilho = Nothing
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd

End Sub
Public Function fgRemoveNode(ByRef xmlDocument As MSXML2.DOMDocument40, _
                             ByVal pstrNodeNameRemove As String) As Boolean

Dim objDomNode                              As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler

    Set objDomNode = xmlDocument.documentElement.selectSingleNode("//" & pstrNodeNameRemove)
    
    xmlDocument.getElementsByTagName(objDomNode.parentNode.nodeName).Item(0).removeChild objDomNode
    
    Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

Public Function fgCreateCDATASection(ByVal psConteudo As String) As MSXML2.IXMLDOMCDATASection

Dim xmlDoc                               As MSXML2.DOMDocument40
Dim objnodeCDATA                            As IXMLDOMCDATASection


On Error GoTo ErrorHandler

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objnodeCDATA = xmlDoc.createCDATASection(psConteudo)

    Set fgCreateCDATASection = objnodeCDATA
    
    Set xmlDoc = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlDoc = Nothing

    Err.Raise Err.Number, Err.Description, Err.DescriptionEnd
    
End Function

Public Function fgDtXML_To_Oracle(ByVal pstrYYYYMMDD As String) As String
On Error GoTo ErrorHandler

    If Len(pstrYYYYMMDD) <> 8 Then
        If Trim(pstrYYYYMMDD) = vbNullString Or _
            pstrYYYYMMDD = gstrDataVazia Then
            fgDtXML_To_Oracle = "Null"
            Exit Function
        Else
            'Parâmetro pstrYYYYMMDD deve ser informado com 8 dígito, no formato "YYYYMMDD"
            Err.Raise vbObjectError + 513, , "Parâmetro pstrYYYYMMDD deve ser informado com 8 dígito, no formato YYYYMMDD"
        End If
    End If

    fgDtXML_To_Oracle = "TO_DATE('" & Format(fgDtXML_To_Date(pstrYYYYMMDD), "DD/MM/YYYY") & "','dd/mm/yyyy')"

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function fgDtHrXML_To_Oracle(ByVal pstrYYYYMMDDHHMMSS As String) As String
On Error GoTo ErrorHandler

    If Len(pstrYYYYMMDDHHMMSS) <> 14 Then
        If Trim(pstrYYYYMMDDHHMMSS) = vbNullString Or _
            pstrYYYYMMDDHHMMSS = gstrDataVazia Then
            fgDtHrXML_To_Oracle = "Null"
            Exit Function
        Else
            'Parâmetro pstrYYYYMMDDHHMMSS deve ser informado com 8 dígito, no formato "YYYYMMDDHHMMSS"
            Err.Raise vbObjectError + 513, , "Parâmetro pstrYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato dd/mm/yyyyy HH:MM:SS"
        End If
    End If
         
    fgDtHrXML_To_Oracle = "TO_DATE('" & Format(fgDtHrStr_To_DateTime(pstrYYYYMMDDHHMMSS), "DD/MM/YYYY HH:mm:ss") & "','dd/mm/yyyy HH24:mi:ss')"

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function fgDtXML_To_Date(ByVal strYYYYMMDD As String) As Date

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long
Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer

On Error GoTo ErrorHandler

    If Len(strYYYYMMDD) <> 8 Then
        If Trim(strYYYYMMDD) = "" Then
            fgDtXML_To_Date = gstrDataVazia
            Exit Function
        Else
            'Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato "yyyyymmdd
            Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDD deve ser informado com 8 dígito, no formato yyyyymmdd"
        End If
    End If

    lngAno = Mid(strYYYYMMDD, 1, 4)
    lngMes = Mid(strYYYYMMDD, 5, 2)
    lngDia = Mid(strYYYYMMDD, 7, 2)

    fgDtXML_To_Date = DateSerial(lngAno, lngMes, lngDia)

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function
Public Function fgDtHrStr_To_DateTime(ByVal strYYYYMMDDHHMMSS As String) As Date

Dim lngDia                                  As Long
Dim lngMes                                  As Long
Dim lngAno                                  As Long
Dim intHora                                 As Integer
Dim intMinuto                               As Integer
Dim intSegundo                              As Integer

On Error GoTo ErrorHandler

    If Len(strYYYYMMDDHHMMSS) <> 14 Then
        If Trim(strYYYYMMDDHHMMSS) = "" Then
            fgDtHrStr_To_DateTime = gstrDataVazia
            Exit Function
        Else
            'Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss
             Err.Raise vbObjectError + 513, App.EXEName & "-fgDtXML_To_Date", "Parâmetro strYYYYMMDDHHMMSS deve ser informado com 14 dígito, no formato yyyyymmddHHmmss"
        End If
    End If

    lngAno = Mid(strYYYYMMDDHHMMSS, 1, 4)
    lngMes = Mid(strYYYYMMDDHHMMSS, 5, 2)
    lngDia = Mid(strYYYYMMDDHHMMSS, 7, 2)

    intHora = Mid(strYYYYMMDDHHMMSS, 9, 2)
    intMinuto = Mid(strYYYYMMDDHHMMSS, 11, 2)
    intSegundo = Mid(strYYYYMMDDHHMMSS, 13, 2)

    fgDtHrStr_To_DateTime = DateSerial(lngAno, lngMes, lngDia) & " " & TimeSerial(intHora, intMinuto, intSegundo)

Exit Function

ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description

End Function

Public Function fgDtString_To_Date(ByVal psDD_MM_YYYY As String) As Date

Dim intDia                                  As Integer
Dim intMes                                  As Integer
Dim intAno                                  As Integer

On Error GoTo ErrorHandler

    If Len(psDD_MM_YYYY) < 10 Then
        'Parâmetro psDD_MM_YYYY deve ser informado com 10 dígito, no formato "dd/mm/yyyyy"
        Err.Raise vbObjectError + 513, , "Parâmetro psDD_MM_YYYY deve ser informado com 10 dígito, no formato dd/mm/yyyyy"
    End If
     
    intDia = Mid(psDD_MM_YYYY, 1, 2)
    intMes = Mid(psDD_MM_YYYY, 4, 2)
    intAno = Mid(psDD_MM_YYYY, 7, 4)
    
    fgDtString_To_Date = DateSerial(intAno, intMes, intDia)
    
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function fgDt_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If
    
    fgDt_To_Xml = Format(pdtmData, "YYYYMMDD")
    
Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function fgDtHr_To_Xml(ByVal pdtmData As Date) As String

On Error GoTo ErrorHandler

    If Not IsDate(pdtmData) Then
        'Parâmetro pdtmData deve ser do tipo Date
        Err.Raise vbObjectError + 513, , "Parâmetro pdtmData deve ser do tipo 'Date'"
    End If

    fgDtHr_To_Xml = Format(pdtmData, "YYYYMMDDHHMMSS")

Exit Function

ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function fgVlrXml_To_Interface(ByVal pvntValor As Variant) As Variant

'Recebe como parâmetro um valor com o separador decimal em qualquer formato.
'Retorna um valor com o separador decimal no formato do Servidor

Dim intPOS                                  As Integer
Dim vntValor                                As Variant
Dim strDecimal                              As String
Dim strThousand                             As String
Dim strNegativeSign                         As String
Dim vntArrayDecimal                         As Variant
Dim vntArrayThousand                        As Variant
Dim vntArrayPonto                           As Variant
Dim vntArrayVirgula                         As Variant
Dim blnValorNegativo                        As Boolean

On Error GoTo ErrorHandler

    If IsEmpty(pvntValor) Then
        fgVlrXml_To_Interface = Format(0, gstrMaskDec)
        Exit Function
    ElseIf Trim(pvntValor) = "" Then
        fgVlrXml_To_Interface = Format(0, gstrMaskDec)
        Exit Function
    End If

    strDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strThousand = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sThousand")
    strNegativeSign = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "sNegativeSign")
    
    If strNegativeSign = "" Then strNegativeSign = "-"

    vntArrayDecimal = Split(pvntValor, strDecimal)
    vntArrayThousand = Split(pvntValor, strThousand)
    
    vntArrayPonto = Split(pvntValor, ".")
    vntArrayVirgula = Split(pvntValor, ",")
    
    If (UBound(vntArrayDecimal) + UBound(vntArrayThousand)) > 1 Then
        'Parametro pvntValor não pode estar formatado
'        llCodigoErroNegocio = 223
'        GoTo ErrorHandler
        pvntValor = Format(pvntValor, "0")
    End If
    
    If strThousand <> "." And strDecimal <> "," Then
        If (UBound(vntArrayPonto) + UBound(vntArrayVirgula)) > 1 Then
            'Parametro pvntValor não pode estar formatado
            Err.Raise vbObjectError + 513, "clsRbLibrary.fgVlrXml_To_Interface", "Parametro pvntValor não pode estar formatado."
        End If
    End If
    
    blnValorNegativo = False

    If InStr(pvntValor, strNegativeSign) > 0 Then
        blnValorNegativo = True
    End If

    If blnValorNegativo Then
        pvntValor = Replace(pvntValor, strNegativeSign, "")
    End If

    intPOS = InStr(1, pvntValor, strDecimal)

    If intPOS > 0 Then
        fgVlrXml_To_Interface = CDec(pvntValor)
    Else
        'Verificar se número tem separador decimal diferente da máquina
        If UBound(vntArrayPonto) Then
            fgVlrXml_To_Interface = CDec(Replace(pvntValor, ".", strDecimal))
        ElseIf UBound(vntArrayVirgula) Then
            fgVlrXml_To_Interface = CDec(Replace(pvntValor, ",", strDecimal))
        Else
            fgVlrXml_To_Interface = CDec(pvntValor)
        End If
        
    End If
    
    If blnValorNegativo Then
        fgVlrXml_To_Interface = fgVlrXml_To_Interface * -1
    End If
    
    fgVlrXml_To_Interface = Format(fgVlrXml_To_Interface, gstrMaskDec)
    
    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function fgVlr_To_Xml(ByVal pvntValor As Variant) As Variant

'Recebe como parâmetro um valor com o separador decimal no formato corrente.
'Retorna um valor com o separador decimal no formato do Servidor (Vírgula)

Dim vntValor                                As Variant
Dim strSeparadorDecimal                     As String

On Error GoTo ErrorHandler
    
    'Ver Valor negativo
    '.00; ""; empty
    
    strSeparadorDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")

    vntValor = CDec(pvntValor)
    vntValor = Replace(vntValor, ".", ",")
    vntValor = Replace(vntValor, strSeparadorDecimal, ",")
    
    fgVlr_To_Xml = vntValor

    Exit Function
ErrorHandler:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function fgDataHoraServidor(ByVal pFormato As enumFormatoDataHora2) As Date
On Error GoTo ErrorHandler

    Select Case pFormato
        Case enumFormatoDataHora2.Data
            fgDataHoraServidor = Date
        Case enumFormatoDataHora2.Hora
            fgDataHoraServidor = Time
        Case enumFormatoDataHora2.DataHora
            fgDataHoraServidor = Now
    End Select

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
