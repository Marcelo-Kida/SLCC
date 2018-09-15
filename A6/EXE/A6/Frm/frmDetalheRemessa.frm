VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDetalheRemessa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhe Remessa"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab sstDetalhe 
      Height          =   6075
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Motivo da Rejeição"
      TabPicture(0)   =   "frmDetalheRemessa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "wbMotivo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Remessa Rejeitada"
      TabPicture(1)   =   "frmDetalheRemessa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "wbRemessa"
      Tab(1).ControlCount=   1
      Begin SHDocVwCtl.WebBrowser wbMotivo 
         Height          =   5655
         Left            =   60
         TabIndex        =   2
         Top             =   360
         Width           =   9015
         ExtentX         =   15901
         ExtentY         =   9975
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
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser wbRemessa 
         Height          =   5655
         Left            =   -74940
         TabIndex        =   3
         Top             =   360
         Width           =   9015
         ExtentX         =   15901
         ExtentY         =   9975
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
         Location        =   ""
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   60
      Top             =   8220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDetalheRemessa.frx":0038
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblBotoes 
      Height          =   330
      Left            =   8100
      TabIndex        =   1
      Top             =   6180
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   582
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDetalheRemessa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a consulta ao detalhe de uma remessa enviada ao A6.

Option Explicit

Private xmlDOMDesc                          As MSXML2.DOMDocument40

Public lngCO_TEXT_XML_REJE                  As Long

Public strXMLErro                           As String
Private strConteudoBrowser                  As String

Private Const MOTIVO_DA_REJEICAO            As Integer = 0
Private Const REMESSA_REJEITADA             As Integer = 1

' Define cores para exibição de acordo com níveis de detalhamento.

Private Function flCorNivel(ByVal pintNivel As Integer) As String

On Error GoTo ErrorHandler

    Select Case pintNivel
        Case 0
            flCorNivel = "#EEEEEE"
        Case 1
            flCorNivel = "#CEDAEA"
        Case 2
            flCorNivel = "#E6EFD8"
        Case Else
            flCorNivel = "#F3ECD4"
    End Select

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flCorNivel", 0

End Function

' Monta HTML para exibição.

Private Function flMontaHTML(ByVal pstrDOMDoc As String, _
                    Optional ByVal pintNivel As Integer, _
                    Optional ByVal pblnMesgErro As Boolean = True) As String

Dim xmlDOMDoc                               As MSXML2.DOMDocument40
Dim intX                                    As Integer
Dim intNivel                                As Integer
Dim strFontColor                            As String
Dim strAux                                  As String
Dim strData                                 As String

On Error GoTo ErrorHandler

    Set xmlDOMDoc = CreateObject("MSXML2.DOMDocument.4.0")
    
    If xmlDOMDoc.loadXML(pstrDOMDoc) Then
        For intX = 0 To xmlDOMDoc.childNodes(0).childNodes.length - 1
            If Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 6) = "Repet_" Then
                flMontaHTML = flMontaHTML & flMontaHTML(xmlDOMDoc.childNodes(0).childNodes(intX).xml, pintNivel)
            ElseIf Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 1) <> "/" Then
               
                If Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 6) = "Grupo_" Then
                    intNivel = pintNivel + 1
                Else
                    intNivel = pintNivel
                End If
                
                If pblnMesgErro Then
                
                    If Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 6) <> "Grupo_" Then
        
                        strFontColor = "Black"
                        flMontaHTML = flMontaHTML & vbCr & _
                                    "<TR BGColor=" & flCorNivel(intNivel) & ">" & vbCr & _
                                    "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"
                        flMontaHTML = flMontaHTML & String(intNivel * 3, Chr(1)) & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName
                        flMontaHTML = flMontaHTML & "</Font></TD>" & vbCr & _
                                    "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"
                    End If
                    
                Else
                
                    strFontColor = "Black"
                    flMontaHTML = flMontaHTML & vbCr & _
                                "<TR BGColor=" & flCorNivel(intNivel) & ">" & vbCr & _
                                "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"
                    
                    'Verifica se a mensagem recebida é proveniente do BUS -> LQS...
                    If Not xmlDOMDoc.selectSingleNode("//MESG") Is Nothing Then
                    'Incluir tratamento para TAGS não fora do dicionário BUS
                    '(TP_MESG, SG_SIST_ORIG, SG_SIST_DEST, CO_EMPR, IN_ENTR_MANU, ...)
                    
                        If xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_ATRB_MESG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/NO_TRAP_ATRB") Is Nothing Then
                            flMontaHTML = flMontaHTML & String(intNivel * 3, Chr(1)) & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName
                        Else
                            flMontaHTML = flMontaHTML & String(intNivel * 3, Chr(1)) & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_ATRB_MESG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/NO_TRAP_ATRB").Text
                        End If
                        
                    '...se não, é uma mensagem SPB
                    Else
                        If xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_TAG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/DE_TAG") Is Nothing Then
                            flMontaHTML = flMontaHTML & String(intNivel * 3, Chr(1)) & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName
                        Else
                            flMontaHTML = flMontaHTML & String(intNivel * 3, Chr(1)) & xmlDOMDesc.selectSingleNode("Repeat_Mensagem/Grupo_Mensagem[NO_TAG='" & xmlDOMDoc.childNodes(0).childNodes(intX).nodeName & "']/DE_TAG").Text
                        End If
                    End If
                    
                    flMontaHTML = flMontaHTML & "</Font></TD>" & vbCr & _
                                "<TD><Font Color=" & strFontColor & " Size=2 Face=Verdana>"
                
                
                
                End If
                
                If Not xmlDOMDoc.childNodes(0).childNodes(intX).childNodes(0) Is Nothing Then
                    If xmlDOMDoc.childNodes(0).childNodes(intX).childNodes(0).hasChildNodes Then
                        flMontaHTML = flMontaHTML & flMontaHTML(xmlDOMDoc.childNodes(0).childNodes(intX).xml, pintNivel + 1)
                    Else
                        If UCase(Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "DTH" Then
                            strData = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                            'YYYYMMDDHHMMSS
                            If IsDate(Mid$(strData, 7, 2) & "/" & Mid$(strData, 5, 2) & "/" & Left$(strData, 4)) Then
                                flMontaHTML = flMontaHTML & FormatDateTime(fgDtHrStr_To_DateTime(Left(strData & String(14, "0"), 14)), vbGeneralDate)
                            Else
                                flMontaHTML = flMontaHTML & strData
                            End If
                        ElseIf UCase(Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 2)) = "DT" Then
                            strData = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                            'YYYYMMDD
                            If IsDate(Mid$(strData, 7, 2) & "/" & Mid$(strData, 5, 2) & "/" & Left$(strData, 4)) Then
                                flMontaHTML = flMontaHTML & FormatDateTime(fgDtHrStr_To_DateTime(Left(strData, 8) & "000000"), vbShortDate)
                            Else
                                flMontaHTML = flMontaHTML & strData
                            End If
                        ElseIf UCase(Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "VLR" Or _
                               UCase(Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "SLD" Then
                            flMontaHTML = flMontaHTML & Format(xmlDOMDoc.childNodes(0).childNodes(intX).Text, "###,###,###,##0.00")
                        ElseIf UCase(Left(xmlDOMDoc.childNodes(0).childNodes(intX).nodeName, 3)) = "QTD" Then
                            flMontaHTML = flMontaHTML & Format(xmlDOMDoc.childNodes(0).childNodes(intX).Text, "###,###,###,##0")
                        Else
                            strAux = xmlDOMDoc.childNodes(0).childNodes(intX).Text
                            strAux = Replace(strAux, "&", "&amp;")
                            strAux = Replace(strAux, "<", "&lt;")
                            strAux = Replace(strAux, ">", "&gt;")
                            flMontaHTML = flMontaHTML & strAux
                        End If
                        
                        flMontaHTML = flMontaHTML & "</Font></TD>" & vbCr & _
                                            "</TR>"
                    End If
                End If
            End If
        Next
    End If
    
    flMontaHTML = Replace(flMontaHTML, Chr(1), "&nbsp")
    
    Exit Function
    
ErrorHandler:
    Call fgRaiseError(App.EXEName, TypeName(Me), "flMontaHTML Function", 0, 0)
    
End Function

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiSBR.Icon
    
    fgCursor True
    Call fgCenterMe(Me)
    
    'Limpa o conteúdo do Browser
    Call flAtualizaConteudoBrowser
    DoEvents
    
    flFormataHTMLMotivoRejeicao
    sstDetalhe.Tab = MOTIVO_DA_REJEICAO
    
    fgCursor
    Exit Sub
    
ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "Form_Load", Me.Caption
    
End Sub

' Monta HTML para exibição de remessa rejeitada.

Private Sub flFormataHTMLRemessaRejeitada()

#If EnableSoap = 1 Then
    Dim objRemessa      As MSSOAPLib30.SoapClient30
#Else
    Dim objRemessa      As A6MIU.clsConsultaRemessaRejeitada
#End If

Dim strMensagemHTML     As String
Dim strHTML             As String
Dim xmlErro             As MSXML2.DOMDocument40
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Call flAtualizaConteudoBrowser
    Set objRemessa = fgCriarObjetoMIU("A6MIU.clsConsultaRemessaRejeitada")
    strMensagemHTML = objRemessa.ObterxmlErroRemessaRejeitada(lngCO_TEXT_XML_REJE, _
                                                              vntCodErro, _
                                                              vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If Len(strMensagemHTML) = 0 Then
       Set objRemessa = Nothing
       frmMural.Caption = Me.Caption
       frmMural.Display = "Não foi encontrado nenhuma Remessa Rejeitada."
       frmMural.Show vbModal
       Exit Sub
    End If
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlErro.loadXML(strMensagemHTML) Then
       Call flAtualizaConteudoBrowser(strMensagemHTML)
       Exit Sub
    End If
    
    strHTML = "<HTML><Body><Center>"
    strHTML = strHTML & "<Table border = 0>" & _
                    "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                    "Remessa Rejeitada" & _
                    "</Font></TD></TR>" & _
                    flMontaHTML(strMensagemHTML)


    strHTML = strHTML & "</Table>"
    strHTML = strHTML & "</Body></HTML>"
    
    Call flAtualizaConteudoBrowser(strHTML)
    
    Set objRemessa = Nothing
    Exit Sub
    
ErrorHandler:
    Set objRemessa = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flFormataHTMLRemessaRejeitada", 0

End Sub

' Monta HTML para exibição do motivo da rejeição da remessa.

Private Sub flFormataHTMLMotivoRejeicao()

Dim strHTML                                 As String
Dim xmlErro                                 As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    flAtualizaConteudoBrowser
    If Len(strXMLErro) = 0 Then
       frmMural.Caption = Me.Caption
       frmMural.Display = "Não foi encontrado nenhum Motivo de Rejeição."
       frmMural.Show vbModal
       Exit Sub
    End If
    
    Set xmlErro = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlErro.loadXML(strXMLErro) Then
        Call flAtualizaConteudoBrowser(strXMLErro)
        Exit Sub
    End If
    
    strHTML = "<HTML><Body><Center>"
    strHTML = strHTML & "<Table border = 0>" & _
                    "<TR><TD colspan=2 BGColor=""#BBBBBB""><Font Color=White Size=2 Face=Verdana>" & _
                    "Motivo da Rejeição" & _
                    "</Font></TD></TR>" & _
                    flMontaHTML(strXMLErro)

    strHTML = strHTML & "</Table>"
    strHTML = strHTML & "</Body></HTML>"
    
    Call flAtualizaConteudoBrowser(strHTML)
    
    Exit Sub
    
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flFormataHTMLMotivoRejeicao", 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmDetalheRemessa = Nothing
    
End Sub

Private Sub sstDetalhe_Click(PreviousTab As Integer)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case sstDetalhe.Tab
    Case MOTIVO_DA_REJEICAO
        flFormataHTMLMotivoRejeicao
         
    Case REMESSA_REJEITADA
        flFormataHTMLRemessaRejeitada
            
    End Select
    
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "sstDetalhe_Click", Me.Caption

End Sub

Private Sub tblBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "Sair"
        Unload Me
    End Select
    
End Sub

Private Sub wbMotivo_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    
On Error Resume Next
    
    pDisp.Document.Body.innerHTML = Replace(strConteudoBrowser, vbNullString, "about:blank")
    
End Sub

' Refresh no HTML.

Private Sub flAtualizaConteudoBrowser(Optional pstrConteudo As String = vbNullString)

On Error GoTo ErrorHandler

    strConteudoBrowser = pstrConteudo
    
    Select Case sstDetalhe.Tab
           Case MOTIVO_DA_REJEICAO
                wbMotivo.Navigate "about:blank"
                
           Case REMESSA_REJEITADA
                wbRemessa.Navigate "about:blank"
    End Select

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flAtualizaConteudoBrowser", 0
    
End Sub

Private Sub wbRemessa_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    
On Error Resume Next
    
    pDisp.Document.Body.innerHTML = Replace(strConteudoBrowser, vbNullString, "about:blank")

End Sub

