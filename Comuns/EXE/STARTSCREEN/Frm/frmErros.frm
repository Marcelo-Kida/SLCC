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
      Picture         =   "frmErros.frx":030A
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
Option Explicit

Private lsArquivoTemp                       As String
Private mErrorNumber                        As Long
Private mErrorDescription                   As String

Public Property Let ErrorNumber(ByVal NewValue As Long)
    mErrorNumber = NewValue
End Property

Public Property Let ErrorDescription(ByVal NewValue As String)
    mErrorDescription = NewValue
End Property

Private Sub cmdOK_Click()
    Set frmErros = Nothing
    DoEvents
    Unload Me
End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Height = 1900
    CenterMe Me
    
    flMostraErro

    Exit Sub
ErrorHandler:
    lblDescricaoErro.Caption = "Erro na visualização."
End Sub
Private Sub flMostraErro()

On Error GoTo ErrorHandler

Dim lsErrDescription                        As String
Dim lsErrSource                             As String
Dim objDomDocument                          As DOMDocument
Dim lbPossuiXML                             As Boolean
Dim liPosicao                               As Integer
Dim lsHTML                                  As String
Dim lsXMLAux                                As String
Dim liFile                                  As Integer

    If InStr(1, mErrorDescription, "<Description>") <> 0 Then
        lsErrDescription = Mid(mErrorDescription, InStr(1, mErrorDescription, "<Description>") + 13, InStr(1, mErrorDescription, "</Description>") - InStr(1, mErrorDescription, "<Description>") - 13)
        lbPossuiXML = True
    Else
        lsErrDescription = mErrorDescription
        lbPossuiXML = False
    End If

    If mErrorNumber <> 0 Then
        
        lblDescricaoErro.Caption = CStr(mErrorNumber) & ": "
        
        If Len(lsErrDescription) > 70 Then
            liPosicao = InStr(50, lsErrDescription, " ")
            lblDescricaoErro.Caption = lblDescricaoErro.Caption & _
                                       Left(lsErrDescription, liPosicao) & vbCrLf & Mid(lsErrDescription, liPosicao + 1, Len(lsErrDescription))
        Else
            lblDescricaoErro.Caption = lblDescricaoErro.Caption & lsErrDescription
        End If

        If lbPossuiXML Then
            lsErrDescription = mErrorDescription & "</ListaErro>"
            cmdDetalhe.Enabled = True
            
            Set objDomDocument = New DOMDocument
    
            If Not objDomDocument.loadXML(lsErrDescription) Then
                lsXMLAux = Left(lsErrDescription, InStr(lsErrDescription, "<Description>") + 12)
                lsXMLAux = lsXMLAux & Replace(Replace(Mid(lsErrDescription, InStr(lsErrDescription, "<Description>") + 13, InStr(lsErrDescription, "</Description>") - InStr(lsErrDescription, "<Description>") - 13), ">", vbNullString), "<", vbNullString)
                lsXMLAux = lsXMLAux & Mid(lsErrDescription, InStr(lsErrDescription, "</Description>"))
                
                lsHTML = "<Html><Boby><Center><Table border = 0>" & _
                        "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                        "Informações do Erro" & _
                        "</Font></TD></TR>" & _
                        flMontaHTML(lsXMLAux) & _
                        "</Table></Center></Boby></Html>"
            Else
                lsHTML = "<Html><Boby><Center><Table border = 0>" & _
                        "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                        "Informações do Erro" & _
                        "</Font></TD></TR>" & _
                        flMontaHTML(objDomDocument.childNodes(0).xml) & _
                        "</Table></Center></Boby></Html>"
            End If
            
            lsArquivoTemp = ArquivoTemp
            liFile = FreeFile
            Open lsArquivoTemp For Output As liFile
            Print #liFile, lsHTML
            Close liFile
            
            Set objDomDocument = Nothing
        
        Else
            cmdDetalhe.Enabled = False
        End If
        
    End If
    
    Exit Sub
ErrorHandler:
    Set objDomDocument = Nothing
    Err.Raise Err.Number, , Err.Description
End Sub
Private Sub cmdDetalhe_Click()
Dim liIndex                                 As Integer

    WebBrowser1.Navigate lsArquivoTemp
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
    
    If Dir(lsArquivoTemp) <> vbNullString And lsArquivoTemp <> vbNullString Then
        Kill lsArquivoTemp
    End If
    Set frmErros = Nothing

End Sub
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

Private Function flMontaHTML(ByVal psXML As String) As String

Dim objDOMDOC               As DOMDocument
Dim X                       As Integer

    On Error GoTo ErrorHandler
        
    Set objDOMDOC = New DOMDocument
    
    If objDOMDOC.loadXML(psXML) Then

        For X = 0 To objDOMDOC.childNodes(0).childNodes.Length - 1
            If objDOMDOC.childNodes(0).childNodes(X).hasChildNodes Then
                If Left(objDOMDOC.childNodes(0).childNodes(X).childNodes(0).xml, 1) = "<" Then
                    flMontaHTML = flMontaHTML & flMontaHTML(objDOMDOC.childNodes(0).childNodes(X).xml)
                Else
                    flMontaHTML = flMontaHTML & _
                              flMontaTag(objDOMDOC.childNodes(0).childNodes(X).nodeName, objDOMDOC.childNodes(0).childNodes(X).Text)
                End If
            Else
                If Left(objDOMDOC.childNodes(0).childNodes(X).xml, 1) = "<" Then
                    flMontaHTML = flMontaHTML & _
                              flMontaTag(objDOMDOC.childNodes(0).childNodes(X).nodeName, objDOMDOC.childNodes(0).childNodes(X).Text)
                Else
                    flMontaHTML = flMontaHTML & _
                              flMontaTag("Client Error", objDOMDOC.childNodes(0).childNodes(X).Text)
                End If
                
            End If
        Next
    
    End If
    
    Set objDOMDOC = Nothing
    
    Exit Function
ErrorHandler:
    Set objDOMDOC = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

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



