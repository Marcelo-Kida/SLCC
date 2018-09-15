VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmFatoGeradorAlerta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fato Gerador Alerta"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6915
   Begin VB.Frame fraCadastro 
      Height          =   3870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6885
      Begin VB.Frame fraFaqtoGerador 
         Caption         =   "Fato Gerador Alerta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   210
         TabIndex        =   6
         Top             =   2070
         Width           =   4275
         Begin VB.TextBox txtDescricaoFatoGeradorAlerta 
            Height          =   315
            Left            =   90
            TabIndex        =   7
            Top             =   1170
            Width           =   4080
         End
         Begin NumBox.Number numCodigoFatoGeradorAlerta 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   540
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            AutoSelecao     =   0   'False
            AceitaNegativo  =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   90
            TabIndex        =   10
            Top             =   330
            Width           =   495
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   105
            TabIndex        =   9
            Top             =   975
            Width           =   720
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Período de Vigência"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   4620
         TabIndex        =   1
         Top             =   2070
         Width           =   1995
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   2
            Top             =   540
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   19595265
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   1170
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   19595265
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin VB.Label lblDataFimVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   975
            Width           =   240
         End
         Begin VB.Label lblDataInicioVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Início "
            Height          =   195
            Left            =   150
            TabIndex        =   4
            Top             =   330
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lstFatoGeradorAlerta 
         Height          =   1605
         Left            =   210
         TabIndex        =   11
         Top             =   330
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2831
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Data Início "
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Data Fim "
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   30
      Top             =   3900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFatoGeradorAlerta.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3000
      TabIndex        =   12
      Top             =   3930
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
 

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

