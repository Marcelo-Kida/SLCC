VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMural 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mural de Alertas e Avisos"
   ClientHeight    =   3555
   ClientLeft      =   2445
   ClientTop       =   4170
   ClientWidth     =   6465
   Icon            =   "frmMural.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6465
   Begin VB.PictureBox pctExibicao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   90
      Picture         =   "frmMural.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   180
      Width           =   480
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   660
      TabIndex        =   0
      Top             =   90
      Width           =   5655
      Begin VB.TextBox txtMural 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   2415
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   5385
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   30
      Top             =   690
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMural.frx":0884
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMural.frx":0CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMural.frx":1128
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMural.frx":157A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4890
      TabIndex        =   3
      Top             =   3060
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   582
      ButtonWidth     =   2487
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgComandos"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fechar Janela"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgComandos 
      Left            =   30
      Top             =   1290
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
            Picture         =   "frmMural.frx":19CC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição de avisos provenientes de consistências feitas nos formulários do sistema.

Option Explicit

Public Enum enumIconeExibicao
    IconExclamation = 1
    IconInformation = 2
    IconCritical = 3
End Enum

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyReturn Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    fgCenterMe Me
    fgCursor
End Sub

'Exibe na tela a mensagem esperada
Public Property Let Display(ByVal vNewValue As String)
    frmMural.txtMural.Text = vNewValue & vbCrLf & frmMural.txtMural.Text
Beep: Beep
End Property

'Define o ícone a ser exibido
Public Property Let IconeExibicao(ByVal intIcone As enumIconeExibicao)
    frmMural.pctExibicao.Picture = imgIcons.ListImages(intIcone).Picture
End Property

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)
    Unload Me
End Sub
