VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmErros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informações sobre o Erro"
   ClientHeight    =   4230
   ClientLeft      =   2490
   ClientTop       =   3000
   ClientWidth     =   6075
   Icon            =   "frmErros.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCopiar 
      Caption         =   "Copiar"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDetalhe 
      Caption         =   "&Detalhes >>"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2505
      Left            =   60
      TabIndex        =   3
      Top             =   1680
      Width           =   5955
      ExtentX         =   10504
      ExtentY         =   4419
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Inside Solid
      X1              =   80
      X2              =   6010
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000004&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      DrawMode        =   4  'Mask Not Pen
      X1              =   60
      X2              =   6000
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmErros.frx":0442
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblDescricaoErro 
      Caption         =   "Descrição do Erro da Aplicação"
      Height          =   675
      Left            =   960
      TabIndex        =   1
      Top             =   320
      Width           =   4890
   End
End
Attribute VB_Name = "frmErros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3EFB2E7C02EE"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
'Objeto responsável pela exibição de erros do sistema.

Option Explicit

Private strArquivoTemp                      As String
Private strErrorDescription                 As String
Private lngErrNumber                        As Long
Private strErrSource                        As String
Private strHTML                             As String

Public Property Let ErrorSouce(ByVal psErrSource As String)
    strErrSource = psErrSource
End Property

Public Property Let ErrorNumber(ByVal plNumeroErro As Long)
    lngErrNumber = plNumeroErro
End Property

Public Property Let ErrorDescription(ByVal psDescricaoErro As String)
    strErrorDescription = psDescricaoErro
End Property

Private Sub cmdCopiar_Click()
    Clipboard.SetText strHTML
End Sub

Private Sub cmdOk_Click()
    Set frmErros = Nothing
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Height = 1900
    
    fgCenterMe Me

    flMostraErro

    Exit Sub

ErrorHandler:
    
    lblDescricaoErro.Caption = "Erro na visualização."

End Sub

'Formatar o erro recebido em formato HTML.
Private Sub flMostraErro()

On Error GoTo ErrorHandler

Dim objDomDocument                          As MSXML2.DOMDocument40

Dim lsHTML                                  As String
Dim liFile                                  As Integer
Dim strRetorno                              As String

    Set objDomDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not objDomDocument.loadXML(strErrorDescription) Then
        Call fgAppendNode(objDomDocument, "", "Erro", "", "")
        Call fgAppendNode(objDomDocument, "Erro", "Grupo_ErrorInfo", "", "")
        Call fgAppendNode(objDomDocument, "Grupo_ErrorInfo", "Number", lngErrNumber)
        Call fgAppendNode(objDomDocument, "Grupo_ErrorInfo", "Description", strErrorDescription)
        Call fgAppendNode(objDomDocument, "Grupo_ErrorInfo", "Source", strErrSource)
    End If
    
    cmdDetalhe.Enabled = True

    If lngErrNumber <> 0 Then
        'Tratamento específico para erros SOAP/HTTP
        strRetorno = fgTrataErroSoapHttp(Val(objDomDocument.documentElement.selectSingleNode("//Erro/Grupo_ErrorInfo/Number").Text), _
                                         objDomDocument.documentElement.selectSingleNode("//Erro/Grupo_ErrorInfo/Description").Text)
        
        lblDescricaoErro.Caption = objDomDocument.documentElement.selectSingleNode("//Erro/Grupo_ErrorInfo/Number").Text & _
                            ": " & strRetorno
                
        lsHTML = "<HTML>" & _
                  "<BODY>" & _
                    "<CENTER>" & _
                        "<TABLE BORDER = 0>" & _
                            flMontaHTML(objDomDocument.xml) & _
                        "</TABLE>" & _
                    "</CENTER>" & _
                 "</BODY>" & _
                 "</HTML>"
        
        strHTML = lsHTML
        strArquivoTemp = ArquivoTemp
        liFile = FreeFile
        Open strArquivoTemp For Output As liFile
        Print #liFile, lsHTML
        Close liFile
        
        Set objDomDocument = Nothing
        
    End If
            
    
    Exit Sub
ErrorHandler:
    Set objDomDocument = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub cmdDetalhe_Click()
Dim liIndex                                 As Integer

    WebBrowser1.Navigate strArquivoTemp
    DoEvents
    
    If cmdDetalhe.Caption = "&Detalhes >>" Then
        cmdDetalhe.Caption = "<< &Detalhes"
        Me.Move Me.Left, Me.Top, Me.Width, 4605
    Else
        cmdDetalhe.Caption = "&Detalhes >>"
        Me.Move Me.Left, Me.Top, Me.Width, 1900
    End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If Dir(strArquivoTemp) <> vbNullString And strArquivoTemp <> vbNullString Then
        Kill strArquivoTemp
    End If
    Set frmErros = Nothing

End Sub

'Criar arquivo temporário para a gravação do arquivo em HTML com o erro.
Private Function ArquivoTemp() As String

    On Error GoTo ErrorHandler
    
    Dim lsTemp                              As String
    Dim llL                                 As Long
    Dim lsDiretorioTemp                     As String
    
    lsDiretorioTemp = String(255, vbNullChar)
    lsDiretorioTemp = Left(lsDiretorioTemp, GetTempPath(255, lsDiretorioTemp))
    
    lsTemp = String(255, vbNullChar)
    llL = GetTempFilename(lsDiretorioTemp, "~html", 0, lsTemp)
    ArquivoTemp = Replace(lsTemp, vbNullChar, vbNullString)
    Kill ArquivoTemp
    DoEvents
    ArquivoTemp = Replace(ArquivoTemp, ".tmp", ".html")
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

'Montar tags HTML a partir do XML de erro.
Private Function flMontaHTML(ByVal psXML As String) As String

Dim objDOMDOC               As MSXML2.DOMDocument40
Dim llCont                       As Integer

On Error GoTo ErrorHandler
        
    Set objDOMDOC = CreateObject("MSXML2.DOMDocument.4.0")
    
    If objDOMDOC.loadXML(psXML) Then
                
        If Not objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ErrorInfo").Item(0) Is Nothing Then
            
            flMontaHTML = flMontaHTML & flMontaCabecalho("Informações do Erro")
            
            For llCont = 0 To objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ErrorInfo/*").length - 1
                        
                flMontaHTML = flMontaHTML & _
                             flMontaTag(objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ErrorInfo/*").Item(llCont).nodeName, _
                             objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ErrorInfo/*").Item(llCont).Text)
            Next
        End If
                
        If Not objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ObjectContext").Item(0) Is Nothing Then
            
            flMontaHTML = flMontaHTML & flMontaCabecalho("Object Context")
            
            For llCont = 0 To objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ObjectContext/*").length - 1
                        
                flMontaHTML = flMontaHTML & _
                             flMontaTag(objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ObjectContext/*").Item(llCont).nodeName, _
                             objDOMDOC.documentElement.selectNodes("//Erro/Grupo_ObjectContext/*").Item(llCont).Text)
            Next
        End If
    
    
        If Not objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem").Item(0) Is Nothing Then
            
            flMontaHTML = flMontaHTML & flMontaCabecalho("Origem")
            
            For llCont = 0 To objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").length - 1
                        
                If Not objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Origem") Is Nothing Then
                        
                    flMontaHTML = flMontaHTML & _
                                 flMontaTag(objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Origem").nodeName, _
                                 objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Origem").Text)
                End If
                
                
                If Not objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Complemento") Is Nothing Then
                    flMontaHTML = flMontaHTML & _
                                 flMontaTag(objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Complemento").nodeName, _
                                 objDOMDOC.documentElement.selectNodes("//Erro/Repet_Origem/*").Item(llCont).selectSingleNode("Complemento").Text)
                End If
            Next
            
        End If
    
    End If
    
    Set objDOMDOC = Nothing
    
    Exit Function
ErrorHandler:
    Set objDOMDOC = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

'Definir cor do nivel do HTML de erro de acordo com sua identação no XML.
Private Function flCorNivel(ByVal piNivel As Integer) As String
    Select Case piNivel
        Case 0
            flCorNivel = "#CEDAEA"
        Case 1
            flCorNivel = "#EEEEEE"
        Case 2
            flCorNivel = "#E6EFD8"
        Case Else
            flCorNivel = "#F3ECD4"
    End Select
End Function

'Montar tag HTML.
Private Function flMontaTag(ByVal NomeTag As String, ByVal ValorTag As String) As String

    flMontaTag = "<TR BGColor=#EEEEEE>" & _
                 "<TD><Font Size=2 Face=Verdana>" & _
                 NomeTag & _
                 "</Font></TD>" & _
                 "<TD><Font Size=2 Face=Verdana>" & _
                 ValorTag & _
                 "</Font></TD>" & _
                 "</TR>"
                 
End Function

'Montar cabeçalho do erro.
Private Function flMontaCabecalho(ByVal psNomeCabecalho As String) As String

    flMontaCabecalho = "<TR>" & _
                            "<TD colspan=2 BGColor=""#BBBBBB"">" & _
                                "<B><FONT COLOR=White SIZE=2 FACE=Verdana>" & _
                                psNomeCabecalho & _
                                "</FONT></B>" & _
                            "</TD>" & _
                        "</TR>"

End Function
