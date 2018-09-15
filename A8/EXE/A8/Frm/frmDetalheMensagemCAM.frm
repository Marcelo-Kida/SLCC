VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDetalheMensagemCAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalhe Mensagem CAM"
   ClientHeight    =   8730
   ClientLeft      =   2685
   ClientTop       =   1245
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin SHDocVwCtl.WebBrowser wbDetalhe 
      Height          =   8115
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   6675
      ExtentX         =   11774
      ExtentY         =   14314
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
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   8280
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
            Picture         =   "frmDetalheMensagemCAM.frx":0000
            Key             =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblBotoes 
      Height          =   330
      Left            =   6000
      TabIndex        =   1
      Top             =   8400
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      ButtonWidth     =   1482
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
Attribute VB_Name = "frmDetalheMensagemCAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngCodigoXml                        As Variant
Private strConteudoBrowser                  As String

Public Property Let CodigoXml(ByVal lngCodTextXml As Variant)
    lngCodigoXml = lngCodTextXml
End Property

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True

    Set Me.Icon = mdiLQS.Icon
    
    flCarregaMensagemHTML
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "Form_Load", Me.Caption

End Sub

Private Sub flCarregaMensagemHTML()

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim xmlDomFiltros           As MSXML2.DOMDocument40
Dim strMensagemHTML         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    
    Call flAtualizaConteudoBrowser

    '>>> Formata XML Filtro padrão... --------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sequencial", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Sequencial", _
                                     "Sequencial", lngCodigoXml)
                                     
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Empresa", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_Empresa", _
                                     "Empresa", 0)
                                     
    '>>> -------------------------------------------------------------------------------------------
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    strMensagemHTML = objMensagem.ObterMensagemHTML(xmlDomFiltros.xml, _
                                                    vntCodErro, _
                                                    vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Call flFormataDatas(strMensagemHTML)
    
    'Força o Browser a atualizar a página com o conteúdo obtido
    Call flAtualizaConteudoBrowser(strMensagemHTML)
     
    Set objMensagem = Nothing
    Set xmlDomFiltros = Nothing
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaMensagemHTML", 0

End Sub

Private Sub flAtualizaConteudoBrowser(Optional pstrConteudo As String = vbNullString)

On Error GoTo ErrorHandler

    strConteudoBrowser = pstrConteudo
    wbDetalhe.Navigate "about:blank"

Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flAtualizaConteudoBrowser", 0

End Sub

Private Sub wbDetalhe_DocumentComplete(ByVal pDisp As Object, URL As Variant)

On Error GoTo ErrorHandler

    pDisp.Document.Body.innerHTML = strConteudoBrowser

Exit Sub
ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - wbDetalhe_DocumentComplete"

End Sub

Private Sub flFormataDatas(ByRef strMensagemHTML As String)

Dim lngPosicao                              As Long
Dim strDataRawFormat                        As String
Dim strSaida                                As String
Dim datDataFormatada                        As Date

Const TAG_DATA_HORA                         As String = "|DH|"
Const TAG_DATA_HORA_SIZE                    As Long = 19

Const TAG_DATA                              As String = "|DT|"
Const TAG_DATA_SIZE                         As Long = 13

On Error GoTo ErrorHandler

    Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_SIZE)
            datDataFormatada = fgDtXML_To_Date(Mid$(strDataRawFormat, Len(TAG_DATA) + 1, 8))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0

    Do
        lngPosicao = InStr(strMensagemHTML, TAG_DATA_HORA)
        If lngPosicao > 0 Then
            strSaida = Mid$(strMensagemHTML, 1, lngPosicao - 1)
            strDataRawFormat = Mid$(strMensagemHTML, lngPosicao, TAG_DATA_HORA_SIZE)
            datDataFormatada = fgDtHrStr_To_DateTime(Mid$(strDataRawFormat, Len(TAG_DATA_HORA) + 1, 14))
            strSaida = strSaida & datDataFormatada & Mid$(strMensagemHTML, lngPosicao + TAG_DATA_HORA_SIZE)
            strMensagemHTML = strSaida
        End If
    Loop While lngPosicao <> 0
    
Exit Sub
ErrorHandler:
   fgRaiseError App.EXEName, TypeName(Me), "flFormataDatas", 0

End Sub

Private Sub tblBotoes_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case gstrSair
            Unload Me
    
    End Select

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tblBotoes_ButtonClick"

End Sub
