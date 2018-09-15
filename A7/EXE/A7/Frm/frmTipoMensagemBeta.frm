VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmTipoMensagemBeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Tipos de Mensagens"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   13530
   Begin VB.Frame Frame1 
      Height          =   2100
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   13485
      Begin MSComctlLib.ListView lstTipoMensagem 
         Height          =   1875
         Left            =   45
         TabIndex        =   1
         Top             =   180
         Width           =   13365
         _ExtentX        =   23574
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   8116
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Natureza"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo de Saída"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Delimitador"
            Object.Width           =   1931
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Prioridade"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Data Início"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Object.Tag             =   "1"
            Text            =   "Data Fim"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Nome Título Mensagem"
            Object.Width           =   7832
         EndProperty
      End
   End
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6945
      Left            =   45
      TabIndex        =   12
      Top             =   2070
      Width           =   13485
      Begin VB.TextBox txtNomeTituloMesg 
         Height          =   315
         Left            =   5200
         MaxLength       =   80
         TabIndex        =   8
         Top             =   1035
         Width           =   4000
      End
      Begin VB.ComboBox cboDelimitador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1035
         Width           =   900
      End
      Begin VB.TextBox txtPrioridade 
         Height          =   315
         Left            =   4155
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1035
         Width           =   840
      End
      Begin VB.ComboBox cboTipoSaida 
         Height          =   315
         Left            =   7005
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   375
         Width           =   2805
      End
      Begin VB.ComboBox cboTipoEvento 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1035
         Width           =   2955
      End
      Begin VB.TextBox txtTipoMensagem 
         Height          =   315
         Left            =   90
         MaxLength       =   9
         TabIndex        =   2
         Top             =   375
         Width           =   930
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   1100
         MaxLength       =   100
         TabIndex        =   3
         Top             =   360
         Width           =   5835
      End
      Begin VB.Frame Frame3 
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
         Height          =   1200
         Left            =   9885
         TabIndex        =   19
         Top             =   135
         Width           =   3525
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   330
            Left            =   180
            TabIndex        =   9
            Top             =   570
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   582
            _Version        =   393216
            Format          =   22675457
            CurrentDate     =   37816
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   330
            Left            =   1785
            TabIndex        =   10
            Top             =   570
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   22675457
            CurrentDate     =   37816
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   210
            TabIndex        =   20
            Top             =   330
            Width           =   510
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1785
            TabIndex        =   21
            Top             =   330
            Width           =   300
         End
      End
      Begin MSComctlLib.ImageList ImgFormatacao 
         Left            =   12645
         Top             =   6435
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":0000
               Key             =   "Node1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":031A
               Key             =   "Parent"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":076C
               Key             =   "Root"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":0BBE
               Key             =   "Node"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":1010
               Key             =   "Rootx"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":132A
               Key             =   "Parentx"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":1644
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":195E
               Key             =   "Del"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":1C78
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":20CA
               Key             =   "Left"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":251C
               Key             =   "Right"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmTipoMensagemBeta.frx":296E
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView trwLayOut 
         Height          =   4695
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   2160
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   8281
         _Version        =   393217
         Indentation     =   617
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImgFormatacao"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbFormatacao 
         Height          =   330
         Left            =   3555
         TabIndex        =   11
         Top             =   1485
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   582
         ButtonWidth     =   1879
         ButtonHeight    =   582
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgFormatacao"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Promote"
               Key             =   "Promote"
               Object.ToolTipText     =   "Promote"
               ImageKey        =   "Left"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Demote"
               Key             =   "Demote"
               Object.ToolTipText     =   "Demote"
               ImageKey        =   "Right"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Subir"
               Key             =   "Up"
               Object.ToolTipText     =   "Subir posição do atributo"
               ImageKey        =   "Up"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Descer"
               Key             =   "Down"
               Object.ToolTipText     =   "Descer posição do atributo"
               ImageKey        =   "Down"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Adicionar"
               Key             =   "Add"
               Object.ToolTipText     =   "Adicionar atributos"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Deletar"
               Key             =   "Del"
               Object.ToolTipText     =   "Deletar atributos"
               ImageKey        =   "Del"
            EndProperty
         EndProperty
         Enabled         =   0   'False
      End
      Begin MSComctlLib.TreeView trwLayOut 
         Height          =   4695
         Index           =   1
         Left            =   4545
         TabIndex        =   23
         Top             =   2160
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   8281
         _Version        =   393217
         Indentation     =   617
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImgFormatacao"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView trwLayOut 
         Height          =   4695
         Index           =   2
         Left            =   9000
         TabIndex        =   24
         Top             =   2160
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   8281
         _Version        =   393217
         Indentation     =   617
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "ImgFormatacao"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Título da Mensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5200
         TabIndex        =   31
         Top             =   810
         Width           =   2565
      End
      Begin VB.Label lblXML 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XML"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   9000
         TabIndex        =   28
         Top             =   1845
         Width           =   4380
      End
      Begin VB.Label lblSTR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "STR / CSV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4545
         TabIndex        =   27
         Top             =   1845
         Width           =   4380
      End
      Begin VB.Label lblId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   90
         TabIndex        =   25
         Top             =   1845
         Width           =   4380
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Prioridade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4170
         TabIndex        =   18
         Top             =   810
         Width           =   870
      End
      Begin VB.Label lblDelimitador 
         AutoSize        =   -1  'True
         Caption         =   "Delimitador"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3105
         TabIndex        =   17
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Saída"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7020
         TabIndex        =   16
         Top             =   150
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Mensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   810
         Width           =   2010
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   150
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1110
         TabIndex        =   14
         Top             =   150
         Width           =   870
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   9990
      TabIndex        =   26
      Top             =   9135
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   582
      ButtonWidth     =   1535
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            Key             =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   8865
      Top             =   9090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":2DC0
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":2ED2
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":31EC
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":353E
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":3650
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":396A
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":3C84
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemBeta.frx":3F9E
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Campos Não Obrigatórios"
      Height          =   195
      Left            =   2925
      TabIndex        =   30
      Top             =   9195
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Campos Obrigatórios"
      Height          =   195
      Left            =   585
      TabIndex        =   29
      Top             =   9200
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   225
      Left            =   2655
      Picture         =   "frmTipoMensagemBeta.frx":42B8
      Top             =   9180
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   225
      Left            =   315
      Picture         =   "frmTipoMensagemBeta.frx":433E
      Top             =   9180
      Width           =   225
   End
End
Attribute VB_Name = "frmTipoMensagemBeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pelo cadastramento e manutenção de tipos de mensagem.
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlTipoMensagem                     As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmTipoMensagem"
Private strEstruturaAtributo                As String

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private WithEvents objInclusaoAtributos     As frmInclusaoAtributo
Attribute objInclusaoAtributos.VB_VarHelpID = -1
Private WithEvents objConfigRepeticao       As frmConfiguraRepeticao
Attribute objConfigRepeticao.VB_VarHelpID = -1
Private objLayOut(2)                        As MSXML2.DOMDocument40
Private intFocu                             As Integer

'Posicionar item no listview de tipos de mensagem.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

    If lstTipoMensagem.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstTipoMensagem.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstTipoMensagem_ItemClick objListItem
           lstTipoMensagem.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimparCampos
    End If

End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = Null
    
    Set objMiu = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.loadXML objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If xmlMapaNavegacao.parseError.errorCode <> 0 Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmTipoMensagem", "flInicializar")
    End If
    
    If xmlTipoMensagem Is Nothing Then
       Set xmlTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
       xmlTipoMensagem.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Tipo_Mensagem").xml
    End If
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Sub

'Carregar listview com os tipos de mensagem cadastrados.
Private Sub flCarregarTipoMensagem()

#If EnableSoap = 1 Then
    Dim objMiu              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu              As A7Miu.clsMIU
#End If

Dim xmlDomTipoMensagem      As MSXML2.DOMDocument40
Dim xmlNode                 As MSXML2.IXMLDOMNode
Dim strLerTodos             As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "LerTodos"
    strLerTodos = objMiu.Executar(xmlMapaNavegacao.selectSingleNode("//Grupo_TipoMensagem").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing

    lstTipoMensagem.ListItems.Clear

    If strLerTodos = "" Then Exit Sub
    
    Set xmlDomTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlDomTipoMensagem.loadXML (strLerTodos)
    
    For Each xmlNode In xmlDomTipoMensagem.selectNodes("//Repeat_TipoMensagem/*")
        With lstTipoMensagem.ListItems.Add(, "EVE" & _
                                            Format(xmlNode.selectSingleNode("TP_FORM_MESG_SAID").Text, "0000") & _
                                            xmlNode.selectSingleNode("TP_MESG").Text, _
                                            xmlNode.selectSingleNode("TP_MESG").Text)
            
            .Tag = xmlNode.selectSingleNode("CO_TEXT_XML").Text
            
            .SubItems(1) = xmlNode.selectSingleNode("NO_TIPO_MESG").Text
            .SubItems(2) = flTipoEventoToSTR(CLng(xmlNode.selectSingleNode("TP_NATZ_MESG").Text))
            .SubItems(3) = flTipoSaidaToSTR(CLng(xmlNode.selectSingleNode("TP_FORM_MESG_SAID").Text))
            .SubItems(4) = xmlNode.selectSingleNode("TP_CTER_DELI").Text
            .SubItems(5) = xmlNode.selectSingleNode("CO_PRIO_FILA_SAID_MESG").Text
            .SubItems(6) = Format(fgDtXML_To_Date(xmlNode.selectSingleNode("DT_INIC_VIGE_MESG").Text), gstrMascaraDataDtp)
            
            If CStr(xmlNode.selectSingleNode("DT_FIM_VIGE_MESG").Text) <> gstrDataVazia Then
                .SubItems(7) = Format(fgDtXML_To_Date(xmlNode.selectSingleNode("DT_FIM_VIGE_MESG").Text), gstrMascaraDataDtp)
            Else
                .SubItems(7) = ""
            End If
            
            .SubItems(8) = xmlNode.selectSingleNode("NO_TITU_MESG").Text
            
        End With
    Next
   
    Set xmlDomTipoMensagem = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set xmlDomTipoMensagem = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarListAtributo", 0

End Sub

Private Sub cboTipoEvento_Click()

    If cboTipoEvento.ListIndex = -1 Then Exit Sub

    If cboTipoEvento.ItemData(cboTipoEvento.ListIndex) = enumNaturezaMensagem.MensagemECO Then
        cboTipoSaida.ListIndex = 0
        cboTipoSaida.Enabled = False
        
        lblId.BackColor = vbMenuBar
        lblId.ForeColor = vbWindowText
        trwLayOut(0).Nodes.Clear
        trwLayOut(0).Enabled = False
        lblSTR.BackColor = vbMenuBar
        lblSTR.ForeColor = vbWindowText
        trwLayOut(1).Nodes.Clear
        trwLayOut(1).Enabled = False
        lblXML.BackColor = vbMenuBar
        lblXML.ForeColor = vbWindowText
        trwLayOut(2).Nodes.Clear
        trwLayOut(2).Enabled = False

        objLayOut(0).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        objLayOut(1).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        objLayOut(2).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"

    ElseIf cboTipoEvento.ItemData(cboTipoEvento.ListIndex) = enumNaturezaMensagem.MensagemConsulta Then
        cboTipoSaida.ListIndex = 0
        cboTipoSaida.Enabled = False
        
        lblId.BackColor = vbWindowBackground
        lblId.ForeColor = vbWindowText
        trwLayOut(0).Nodes.Clear
        trwLayOut(0).Enabled = True
        lblSTR.BackColor = vbMenuBar
        lblSTR.ForeColor = vbWindowText
        trwLayOut(1).Nodes.Clear
        trwLayOut(1).Enabled = False
        lblXML.BackColor = vbMenuBar
        lblXML.ForeColor = vbWindowText
        trwLayOut(2).Nodes.Clear
        trwLayOut(2).Enabled = False

    Else
        If strOperacao <> "Alterar" Then
            cboTipoSaida.Enabled = True
        End If
        
        lblId.BackColor = vbWindowBackground
        lblId.ForeColor = vbWindowText
        trwLayOut(0).Nodes.Clear
        trwLayOut(0).Enabled = True
        lblSTR.BackColor = vbMenuBar
        lblSTR.ForeColor = vbWindowText
        trwLayOut(1).Nodes.Clear
        trwLayOut(1).Enabled = False
        lblXML.BackColor = vbMenuBar
        lblXML.ForeColor = vbWindowText
        trwLayOut(2).Nodes.Clear
        trwLayOut(2).Enabled = False
        
        If cboTipoSaida.ListIndex = -1 Then Exit Sub
        
        Select Case cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
            Case enumTipoSaidaMensagem.SaidaCSV, enumTipoSaidaMensagem.SaidaString
                lblSTR.BackColor = vbWindowBackground
                lblSTR.ForeColor = vbWindowText
                trwLayOut(1).Nodes.Clear
                trwLayOut(1).Enabled = True
            Case enumTipoSaidaMensagem.SaidaXML
                lblXML.BackColor = vbWindowBackground
                lblXML.ForeColor = vbWindowText
                trwLayOut(2).Nodes.Clear
                trwLayOut(2).Enabled = True

            Case enumTipoSaidaMensagem.NaoseAplica
            
            Case Else
                lblSTR.BackColor = vbWindowBackground
                lblSTR.ForeColor = vbWindowText
                trwLayOut(1).Nodes.Clear
                trwLayOut(1).Enabled = True
                lblXML.BackColor = vbWindowBackground
                lblXML.ForeColor = vbWindowText
                trwLayOut(2).Nodes.Clear
                trwLayOut(2).Enabled = True
        End Select
    End If

End Sub

Private Sub cboTipoSaida_Click()

Dim intTipoSaida                            As enumTipoSaidaMensagem

    If cboTipoSaida.ListIndex < 0 Then Exit Sub

    intTipoSaida = cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
    
    'enumTipoSaidaMensagem.NaoseAplica
    lblId.BackColor = vbMenuBar
    lblId.ForeColor = vbWindowText
    trwLayOut(0).Nodes.Clear
    trwLayOut(0).Enabled = False
    lblSTR.BackColor = vbMenuBar
    lblSTR.ForeColor = vbWindowText
    trwLayOut(1).Nodes.Clear
    trwLayOut(1).Enabled = False
    lblXML.BackColor = vbMenuBar
    lblXML.ForeColor = vbWindowText
    trwLayOut(2).Nodes.Clear
    trwLayOut(2).Enabled = False
    objLayOut(0).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
    objLayOut(1).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
    objLayOut(2).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"

    cboDelimitador.ListIndex = -1
    cboDelimitador.Enabled = False
    lblDelimitador.Enabled = False

    If intTipoSaida <> enumTipoSaidaMensagem.NaoseAplica Then
        trwLayOut(0).Enabled = True
        lblId.BackColor = vbWindowBackground
        lblId.ForeColor = vbWindowText
    Else
        Exit Sub
    End If
    
    If intTipoSaida = enumTipoSaidaMensagem.SaidaString Or _
       intTipoSaida = enumTipoSaidaMensagem.SaidaStringXML Or _
       intTipoSaida = enumTipoSaidaMensagem.SaidaCSV Or _
       intTipoSaida = enumTipoSaidaMensagem.SaidaCSVXML Then
        
        trwLayOut(1).Enabled = True
        lblSTR.BackColor = vbWindowBackground
        lblSTR.ForeColor = vbWindowText
        trwLayOut(1).Nodes.Clear
        
        If intTipoSaida = enumTipoSaidaMensagem.SaidaCSV Or _
           intTipoSaida = enumTipoSaidaMensagem.SaidaCSVXML Then
            cboDelimitador.Enabled = True
            lblDelimitador.Enabled = True
        End If
    End If
            
    If intTipoSaida = enumTipoSaidaMensagem.SaidaStringXML Or _
       intTipoSaida = enumTipoSaidaMensagem.SaidaCSVXML Or _
       intTipoSaida = enumTipoSaidaMensagem.SaidaXML Then

        trwLayOut(2).Enabled = True
        lblXML.BackColor = vbWindowBackground
        lblXML.ForeColor = vbWindowText
        trwLayOut(2).Nodes.Clear
    End If
        
    intFocu = -1
    tlbFormatacao.Enabled = False

End Sub

Private Sub dtpDataFimVigencia_Change()
    
    If Not IsNull(dtpDataFimVigencia.Value) Then
        If dtpDataFimVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
            dtpDataFimVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
            dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
        End If
    End If
    If dtpDataInicioVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) And dtpDataInicioVigencia.Enabled Then
        dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
        dtpDataInicioVigencia.MinDate = dtpDataInicioVigencia.Value
    End If

End Sub

Private Sub dtpDataFimVigencia_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            KeyAscii = 0
    End Select

End Sub

Private Sub dtpDataInicioVigencia_Change()
    
    If dtpDataInicioVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
        dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
        dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    End If

    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.Value
    dtpDataFimVigencia.Value = dtpDataInicioVigencia.Value
    dtpDataFimVigencia.Value = Null

End Sub

Private Sub dtpDataInicioVigencia_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            KeyAscii = 0
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
 
    fgCenterMe Me
        
    Me.Icon = mdiBUS.Icon
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    Me.Show
    DoEvents
    
    fgLockWindow Me.hwnd
    flInicializar
    Set objInclusaoAtributos = New frmInclusaoAtributo
    Set objConfigRepeticao = New frmConfiguraRepeticao
    Set objLayOut(0) = New DOMDocument40
    Set objLayOut(1) = New DOMDocument40
    Set objLayOut(2) = New DOMDocument40
    
    flLimparCampos
    
    fgCursor True
    
    flCarregarCboTipoEvento
    flCarregarCboTipoSaida
    flCarregaComboDelimitador
    
    flCarregarTipoMensagem
    
    txtTipoMensagem.SetFocus
    
    fgCursor False
    
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    
    fgCursor False
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - Form_Load")

End Sub

'Limpar campos do formulário.
Private Sub flLimparCampos()
        
    strOperacao = "Incluir"
        
    txtTipoMensagem.Text = vbNullString
    txtTipoMensagem.Enabled = True
    cboTipoSaida.Enabled = True
    cboTipoEvento.Enabled = True
    txtDescricao.Text = vbNullString
    
    cboTipoEvento.ListIndex = -1
    cboTipoSaida.ListIndex = -1
    
    cboDelimitador.ListIndex = -1
    txtPrioridade.Text = vbNullString
    txtNomeTituloMesg.Text = vbNullString
    
    tlbCadastro.Buttons("Excluir").Enabled = False
    
    dtpDataInicioVigencia.Enabled = True
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataInicioVigencia.Value = dtpDataInicioVigencia.MinDate
    
    dtpDataFimVigencia.Enabled = True
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
    dtpDataFimVigencia.Value = Null
    
    lstTipoMensagem.Sorted = False
    
    lblId.BackColor = vbMenuBar
    lblId.ForeColor = vbWindowText
    trwLayOut(0).Nodes.Clear
    trwLayOut(0).Enabled = False
    lblSTR.BackColor = vbMenuBar
    lblSTR.ForeColor = vbWindowText
    trwLayOut(1).Nodes.Clear
    trwLayOut(1).Enabled = False
    lblXML.BackColor = vbMenuBar
    lblXML.ForeColor = vbWindowText
    trwLayOut(2).Nodes.Clear
    trwLayOut(2).Enabled = False
    
    objLayOut(0).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
    objLayOut(1).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
    objLayOut(2).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"

End Sub

'Carregar combo com tipo de eventos.
Private Sub flCarregarCboTipoEvento()

On Error GoTo ErrorHandler
    
    cboTipoEvento.Clear
    cboTipoEvento.AddItem "Envio de dados"
    cboTipoEvento.ItemData(cboTipoEvento.NewIndex) = enumNaturezaMensagem.MensagemEnvio
    cboTipoEvento.AddItem "Consulta"
    cboTipoEvento.ItemData(cboTipoEvento.NewIndex) = enumNaturezaMensagem.MensagemConsulta
    cboTipoEvento.AddItem "Eco"
    cboTipoEvento.ItemData(cboTipoEvento.NewIndex) = enumNaturezaMensagem.MensagemECO
    
    Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flCarregarCboTipoEvento", 0
End Sub

'Carregar combo com tipo de saída de mensagens.
Private Sub flCarregarCboTipoSaida()

On Error GoTo ErrorHandler
        
    cboTipoSaida.Clear
    cboTipoSaida.AddItem "Não se Aplica"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = 0
    cboTipoSaida.AddItem "XML"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = enumTipoSaidaMensagem.SaidaXML
    cboTipoSaida.AddItem "String"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = enumTipoSaidaMensagem.SaidaString
    cboTipoSaida.AddItem "CSV"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = enumTipoSaidaMensagem.SaidaCSV
    cboTipoSaida.AddItem "String + XML"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = enumTipoSaidaMensagem.SaidaStringXML
    cboTipoSaida.AddItem "CSV + XML"
    cboTipoSaida.ItemData(cboTipoSaida.NewIndex) = enumTipoSaidaMensagem.SaidaCSVXML
    
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarCboTipoSaida", 0

End Sub

'Converter o domínio numérico de tipo de saída para literais.
Private Function flTipoSaidaToSTR(lngTipoSaida As Long) As String
    
    Select Case lngTipoSaida
        Case enumTipoSaidaMensagem.SaidaXML
            flTipoSaidaToSTR = "XML"
        Case enumTipoSaidaMensagem.SaidaString
            flTipoSaidaToSTR = "String"
        Case enumTipoSaidaMensagem.SaidaCSV
            flTipoSaidaToSTR = "CSV"
        Case enumTipoSaidaMensagem.SaidaStringXML
            flTipoSaidaToSTR = "String + XML"
        Case enumTipoSaidaMensagem.SaidaCSVXML
            flTipoSaidaToSTR = "CSV + XML"
    End Select

End Function

'Converter as literais de tipo de saída para o domínio numérico.
Private Function flTipoSaidaToEnum(strTipoSaida As String) As Long
    
    Select Case strTipoSaida
        Case "XML"
            flTipoSaidaToEnum = enumTipoSaidaMensagem.SaidaXML
        Case "String"
            flTipoSaidaToEnum = enumTipoSaidaMensagem.SaidaString
        Case "CSV"
            flTipoSaidaToEnum = enumTipoSaidaMensagem.SaidaCSV
        Case "String + XML"
            flTipoSaidaToEnum = enumTipoSaidaMensagem.SaidaStringXML
        Case "CSV + XML"
            flTipoSaidaToEnum = enumTipoSaidaMensagem.SaidaCSVXML
    End Select

End Function

'Converter as literais de tipo de evento para o domínio numérico.
Private Function flTipoEventoToEnum(pstrTipoEvento As String) As Long
    
    Select Case pstrTipoEvento
        Case "Envio de dados"
            flTipoEventoToEnum = enumNaturezaMensagem.MensagemEnvio
        Case "Consulta"
            flTipoEventoToEnum = enumNaturezaMensagem.MensagemConsulta
        Case "Eco"
            flTipoEventoToEnum = enumNaturezaMensagem.MensagemECO
    End Select

End Function

'Converter o domínio numérico de tipo de saída para literais.
Private Function flTipoEventoToSTR(plngTipoEvento As Long) As String

    Select Case plngTipoEvento
        Case enumNaturezaMensagem.MensagemEnvio
            flTipoEventoToSTR = "Envio de dados"
        Case enumNaturezaMensagem.MensagemConsulta
            flTipoEventoToSTR = "Consulta"
        Case enumNaturezaMensagem.MensagemECO
            flTipoEventoToSTR = "Eco"
    End Select

End Function

'Converter o domínio numérico de tipo de dado para literais.
Private Function flTipoDadoToSTR(plngTipoDado As Long) As String
    
    Select Case plngTipoDado
        Case enumTipoDadoAtributo.Alfanumerico
            flTipoDadoToSTR = "Alfanumérico"
        Case enumTipoDadoAtributo.Numerico
            flTipoDadoToSTR = "Numérico"
    End Select

End Function

''Converter as literais de tipo de dado para o domínio numérico.
Private Function flTipoDadoToEnum(pstrTipoDado As String) As Long

    Select Case pstrTipoDado
        Case "Alfanumérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Alfanumerico
        Case "Numérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Numerico
    End Select

End Function

Private Sub Form_Unload(Cancel As Integer)
    
    Set objInclusaoAtributos = Nothing
    Set objConfigRepeticao = Nothing
    Set frmTipoMensagemBeta = Nothing
    Set objLayOut(0) = Nothing
    Set objLayOut(1) = Nothing
    Set objLayOut(2) = Nothing
    
End Sub

Private Sub lstTipoMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    fgClassificarListview lstTipoMensagem, ColumnHeader.Index

End Sub

Private Sub lstTipoMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    fgLockWindow Me.hwnd
    Call flLimparCampos
    strOperacao = "Alterar"
    strKeyItemSelected = Item.Key
    Call flXmlToInterface
    fgLockWindow 0
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiBUS.uctLogErros.MostrarErros Err, "frmTipoMensagem - lstAtributo_ItemClick"
    
    Call flCarregarTipoMensagem
    
    If strOperacao = "Excluir" Then
        flLimparCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
End Sub

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim intAux              As Integer
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    dtpDataInicioVigencia.Enabled = False

    With xmlTipoMensagem
        .selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Ler"
        .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text = lstTipoMensagem.SelectedItem.Text
        .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text = CLng(Mid$(lstTipoMensagem.SelectedItem.Key, 4, 4))
    End With

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlTipoMensagem.loadXML objMiu.Executar(xmlTipoMensagem.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing

    txtTipoMensagem.Enabled = False
    cboTipoSaida.Enabled = False
    cboTipoEvento.Enabled = False

    With xmlTipoMensagem
        txtTipoMensagem.Text = .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text
        txtDescricao.Text = .selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text
        fgSearchItemCombo cboTipoSaida, .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text
        fgSearchItemCombo cboTipoEvento, .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text

        If Trim(.selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text) <> vbNullString Then
            fgSearchItemCombo cboDelimitador, 0, .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text
        Else
            cboDelimitador.ListIndex = -1
        End If

        txtPrioridade.Text = .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text
        txtNomeTituloMesg.Text = .selectSingleNode("//Grupo_TipoMensagem/NO_TITU_MESG").Text
        
        dtpDataInicioVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text)
        dtpDataInicioVigencia.Value = fgDtXML_To_Date(.selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text)

        If dtpDataInicioVigencia.Value > fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
            dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
            dtpDataInicioVigencia.Enabled = True
        End If

        If Trim(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text) <> gstrDataVazia Then
            If fgDtXML_To_Date(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text) < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                dtpDataFimVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text)
                dtpDataInicioVigencia.Enabled = True
            Else
                dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
            End If
            dtpDataFimVigencia.Value = fgDtXML_To_Date(.selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text)
        Else
           dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
           dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
           dtpDataFimVigencia.Value = Null
        End If

        'Inicializa XML com Formato dos tipos
        objLayOut(0).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        
        flMontarMensagem xmlTipoMensagem.documentElement.selectNodes("//Repeat_TipoMensagemAtributo/Grupo_TipoMensagemAtributo[TP_FORM_MESG='" & enumTipoParteSaida.ParteId & "']"), _
                         objLayOut(0)
                
        objLayOut(1).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        flMontarMensagem xmlTipoMensagem.documentElement.selectNodes("//Repeat_TipoMensagemAtributo/Grupo_TipoMensagemAtributo[TP_FORM_MESG='" & enumTipoParteSaida.ParteSTR & "' or TP_FORM_MESG='" & enumTipoParteSaida.ParteCSV & "' ]"), _
                         objLayOut(1)
                         
        objLayOut(2).loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        flMontarMensagem xmlTipoMensagem.documentElement.selectNodes("//Repeat_TipoMensagemAtributo/Grupo_TipoMensagemAtributo[TP_FORM_MESG='" & enumTipoParteSaida.ParteXML & "']"), _
                         objLayOut(2)
        
    End With
        
    For intAux = 0 To 2
        intFocu = intAux
        flLoadTreeFromXML
    Next
    
    intFocu = -1
    tlbFormatacao.Enabled = False
    
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    
    Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flMoveObjetToInterface", 0
    
End Sub

'Montar o layout do tipo de mensagem a partir dos dados recebidos do servidor.
Private Sub flMontarMensagem(ByRef pxmlNodeList As IXMLDOMNodeList, _
                             ByRef pxmlDOMLayout As DOMDocument40)

Dim xmlNodeInclusao                         As IXMLDOMNode
Dim xmlNode                                 As IXMLDOMNode
Dim lngNivel                                As Long
Dim lngX                                    As Long
Dim strNomeNodeInclusao(10)                 As String

On Error GoTo ErrorHandler

    If pxmlNodeList.length = 0 Then Exit Sub
    Set xmlNodeInclusao = pxmlDOMLayout.selectSingleNode("//XML")
    lngNivel = 1
    strNomeNodeInclusao(1) = "XML"
    
    For lngX = 0 To pxmlNodeList.length - 1
        
        Set xmlNode = pxmlNodeList.Item(lngX)
        
        If lngNivel = CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) Then
            'Continuar no mesmo nível

            fgAppendNode pxmlDOMLayout, _
                         strNomeNodeInclusao(lngNivel), _
                         xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                         vbNullString

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "Obrigatorio", _
                              IIf(xmlNode.selectSingleNode("IN_OBRI_ATRB").Text = "1", "True", "False")

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "Expanded", _
                              "True"
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "TP_DADO_ATRB_MESG", _
                              xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "QT_CTER_ATRB", _
                              xmlNode.selectSingleNode("QT_CTER_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "QT_CASA_DECI_ATRB", _
                              xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "QT_REPE", _
                              CLng(xmlNode.selectSingleNode("QT_REPE").Text)

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "IN_ATRB_PRMT_VALO_NEGT", _
                              CLng(xmlNode.selectSingleNode("IN_ATRB_PRMT_VALO_NEGT").Text)

        ElseIf CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) > lngNivel Then
            'Novo Nível, Maior que o anterior --> Incluir como Filho
            'Muda o Pai
            Set xmlNodeInclusao = pxmlNodeList.Item(lngX - 1)
            lngNivel = CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text)
            strNomeNodeInclusao(lngNivel) = xmlNodeInclusao.selectSingleNode("./NO_ATRB_MESG").Text
            lngX = lngX - 1

        ElseIf CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) < lngNivel Then
            'Voltar ao contexto anterior
            
            lngNivel = CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text)
            lngX = lngX - 1

        End If
        
        Set xmlNode = Nothing
    Next

    Set xmlNodeInclusao = Nothing

    Exit Sub
ErrorHandler:
    
    Set xmlNode = Nothing
    Set xmlNodeInclusao = Nothing

    fgRaiseError App.EXEName, "frmTipoMensagem", "flMontarMensagem", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Configurar a quantidade de repetições para um atributo do tipo de mensagem.
Private Sub objConfigRepeticao_QuantidadeRepeticoes(ByVal plngQtdRepeticao As Long)

Dim strNodeFocu                             As String

    fgCursor True
    fgLockWindow Me.hwnd
    
    strNodeFocu = trwLayOut(intFocu).SelectedItem.Key
    objLayOut(intFocu).selectSingleNode("//" & strNodeFocu & "/@QT_REPE").Text = plngQtdRepeticao
    flLoadTreeFromXML
    
    trwLayOut(intFocu).Nodes(strNodeFocu).Selected = True
   
    fgLockWindow 0
    fgCursor

End Sub

'Executar ações recebidas da barra de botões de comando da tela de cadastro de tipos de mensagem, tais como:
' - Salvar tipo de mensagem (inclusão ou alteração);
' - Excluir tipo de mensagem;
' - Limpar dados da tela;
' - Sair da funcionalidade.
Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "Limpar"
            flLimparCampos
            txtTipoMensagem.SetFocus
        Case "Salvar"
            Call flSalvar
        Case "Excluir"
            flExcluir
            flLimparCampos
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select
    
    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
    Exit Sub

ErrorHandler:
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmTipoMensagem - tlbCadastro_ButtonClick"
    
    Call flCarregarTipoMensagem
    
    If strOperacao = "Excluir" Then
        flLimparCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
End Sub

'Executar ações recebidas da barra de botões de formatação do layout do tipo de mensagem
Private Sub tlbFormatacao_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim objNodeInFocus                          As IXMLDOMNode
Dim objNodeAux                              As IXMLDOMNode

    Select Case Button.Key
    
        Case "Demote"
            
            If trwLayOut(intFocu).SelectedItem Is Nothing Then Exit Sub
        
            fgLockWindow Me.hwnd
    
            Set objNodeInFocus = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).cloneNode(True)
            Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).previousSibling
            
            objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode.removeChild objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key)
            
            objNodeAux.appendChild objNodeInFocus
            
            flLoadTreeFromXML
            
            trwLayOut_NodeClick intFocu, trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName)
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).Selected = True
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).EnsureVisible
        
        Case "Promote"
            
            If trwLayOut(intFocu).SelectedItem Is Nothing Then Exit Sub
            
            fgLockWindow Me.hwnd
            
            Set objNodeInFocus = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).cloneNode(True)
            Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode
            
            objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode.removeChild objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key)
            
            If objNodeAux.childNodes(1) Is Nothing Then
                'Remover a quantidade de repetições pois não existem mais filhos
                objNodeAux.selectSingleNode("@QT_REPE").Text = 0
            End If
            
            If objNodeAux.nextSibling Is Nothing Then
                'Incluir no final
                objNodeAux.parentNode.appendChild objNodeInFocus
            Else
                objNodeAux.parentNode.insertBefore objNodeInFocus, objNodeAux.nextSibling
            End If
            
            flLoadTreeFromXML
            
            trwLayOut_NodeClick intFocu, trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName)
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).Selected = True
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).EnsureVisible
            
            
        Case "Up"
            
            If trwLayOut(intFocu).SelectedItem Is Nothing Then Exit Sub
    
            fgLockWindow Me.hwnd
            
            Set objNodeInFocus = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).cloneNode(True)
            Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).previousSibling
            
            objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode.removeChild objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key)
            
            objNodeAux.parentNode.insertBefore objNodeInFocus, objNodeAux
            
            flLoadTreeFromXML
        
            trwLayOut_NodeClick intFocu, trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName)
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).Selected = True
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).EnsureVisible
        
        Case "Down"
            
            If trwLayOut(intFocu).SelectedItem Is Nothing Then Exit Sub
            
            fgLockWindow Me.hwnd
    
            Set objNodeInFocus = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).cloneNode(True)
            Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).nextSibling
            
            objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode.removeChild objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key)
            
            If objNodeAux.nextSibling Is Nothing Then
                'inserir na ultima posição
                objNodeAux.parentNode.appendChild objNodeInFocus
            Else
                objNodeAux.parentNode.insertBefore objNodeInFocus, objNodeAux.nextSibling
            End If
            
            flLoadTreeFromXML
            
            trwLayOut_NodeClick intFocu, trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName)
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).Selected = True
            trwLayOut(intFocu).Nodes(objNodeInFocus.nodeName).EnsureVisible
        
        Case "Add"
            
            objInclusaoAtributos.Show vbModal
            trwLayOut_NodeClick intFocu, trwLayOut(intFocu).SelectedItem
                    
        Case "Del"
            
            If trwLayOut(intFocu).SelectedItem Is Nothing Then Exit Sub
            
            fgLockWindow Me.hwnd
            
            'Define o proximo focu
            If Not objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key & "/preceding-sibling::*") Is Nothing Then
                Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).previousSibling
            ElseIf Not objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).nextSibling Is Nothing Then
                Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).nextSibling
            ElseIf Not objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode Is Nothing Then
                Set objNodeAux = objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode
            End If
            
            objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key).parentNode.removeChild objLayOut(intFocu).selectSingleNode("//" & trwLayOut(intFocu).SelectedItem.Key)
            
            flLoadTreeFromXML
            
            If objNodeAux.nodeName <> "XML" Then
                trwLayOut_NodeClick intFocu, trwLayOut(intFocu).Nodes(objNodeAux.nodeName)
                trwLayOut(intFocu).Nodes(objNodeAux.nodeName).Selected = True
                trwLayOut(intFocu).Nodes(objNodeAux.nodeName).EnsureVisible
            Else
                trwLayOut_NodeClick intFocu, trwLayOut(intFocu).SelectedItem
            End If
        
    End Select
    
    Set objNodeInFocus = Nothing
    Set objNodeAux = Nothing

    fgLockWindow 0

End Sub

Private Sub objInclusaoAtributos_AtributosEscolhidos(ByVal strXMLAtributo As String)

Dim objDOM                                  As DOMDocument40
Dim objNode                                 As IXMLDOMNode
Dim objAddNode                              As IXMLDOMNode
Dim objAttribute                            As IXMLDOMAttribute

Dim strTagErro                              As String

    strTagErro = vbNullString

    Set objDOM = New MSXML2.DOMDocument40

    objDOM.loadXML strXMLAtributo
    
    For Each objNode In objDOM.selectNodes("//XML/*")
    
        If objLayOut(intFocu).selectSingleNode("//" & objNode.nodeName) Is Nothing Then
        
            fgAppendNode objLayOut(intFocu), "XML", objNode.nodeName, vbNullString
            fgAppendAttribute objLayOut(intFocu), objNode.nodeName, "Obrigatorio", "False"
            fgAppendAttribute objLayOut(intFocu), objNode.nodeName, "Expanded", "False"
            
            fgAppendAttribute objLayOut(intFocu), _
                              objNode.nodeName, _
                              "TP_DADO_ATRB_MESG", _
                              objNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text

            fgAppendAttribute objLayOut(intFocu), _
                              objNode.nodeName, _
                              "QT_CTER_ATRB", _
                              objNode.selectSingleNode("@QT_CTER_ATRB").Text
                              
            fgAppendAttribute objLayOut(intFocu), _
                              objNode.nodeName, _
                              "QT_CASA_DECI_ATRB", _
                              objNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text
            
            fgAppendAttribute objLayOut(intFocu), _
                              objNode.nodeName, _
                              "IN_ATRB_PRMT_VALO_NEGT", _
                              objNode.selectSingleNode("@IN_ATRB_PRMT_VALO_NEGT").Text

            fgAppendAttribute objLayOut(intFocu), objNode.nodeName, "QT_REPE", "0"
                
        Else
            'Tag Já incluída
            'Armazenar para retornar erro.
            strTagErro = strTagErro & objNode.nodeName & ", "
        
        End If
        
    Next
    
    Set objDOM = Nothing

    flLoadTreeFromXML
    
    If strTagErro <> vbNullString Then
    
        MsgBox "Os seguintes campos não foram incluídos pois já se encontram na mensagem: " & _
                vbCrLf & Left$(strTagErro, Len(strTagErro) - 2), vbInformation
    
    End If
    
End Sub

'Carregar os TreeViews com lauouts do tipo de mensagem a paritr do xml do tipo de mensagem.
Private Sub flLoadTreeFromXML()

Dim objNode                                 As IXMLDOMNode
    
    trwLayOut(intFocu).Nodes.Clear
    
    For Each objNode In objLayOut(intFocu).selectNodes("//*[name()!='XML']")
        If objNode.parentNode.nodeName = "XML" Then
            'If objNode.hasChildNodes Then
            If Not objNode.childNodes(1) Is Nothing Then
                If CLng(objNode.selectSingleNode("@QT_REPE").Text) > 0 Then
                    trwLayOut(intFocu).Nodes.Add , , objNode.nodeName, objNode.nodeName & " (" & CLng(objNode.selectSingleNode("@QT_REPE").Text) & ")", "Parent"
                Else
                    trwLayOut(intFocu).Nodes.Add , , objNode.nodeName, objNode.nodeName, "Parent"
                End If
            Else
                trwLayOut(intFocu).Nodes.Add , , objNode.nodeName, objNode.nodeName, "Node"
            End If
        Else
            'If objNode.hasChildNodes Then
            If Not objNode.childNodes(1) Is Nothing Then
                If CLng(objNode.selectSingleNode("@QT_REPE").Text) > 0 Then
                    trwLayOut(intFocu).Nodes.Add objNode.parentNode.nodeName, tvwChild, objNode.nodeName, objNode.nodeName & " (" & CLng(objNode.selectSingleNode("@QT_REPE").Text) & ")", "Parent"
                Else
                    trwLayOut(intFocu).Nodes.Add objNode.parentNode.nodeName, tvwChild, objNode.nodeName, objNode.nodeName, "Parent"
                End If
            Else
                trwLayOut(intFocu).Nodes.Add objNode.parentNode.nodeName, tvwChild, objNode.nodeName, objNode.nodeName, "Node"
            End If
        End If
        
        trwLayOut(intFocu).Nodes(objNode.nodeName).Checked = (objNode.selectSingleNode("@Obrigatorio").Text = "True")
        trwLayOut(intFocu).Nodes(objNode.nodeName).Expanded = (objNode.selectSingleNode("@Expanded").Text = "True")
    Next

End Sub

Private Sub trwLayOut_Collapse(Index As Integer, ByVal Node As MSComctlLib.Node)

    objLayOut(Index).selectSingleNode("//" & flGetNodeText(Node.Text) & "/@Expanded").Text = "False"
    trwLayOut_NodeClick Index, Node

End Sub

Private Sub trwLayOut_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
    
    objLayOut(Index).selectSingleNode("//" & flGetNodeText(Node.Text) & "/@Expanded").Text = "True"

End Sub

Private Sub trwLayOut_GotFocus(Index As Integer)
    
    If cboTipoEvento.ListIndex < 0 And cboTipoSaida.ListIndex < 0 Then Exit Sub
    
    intFocu = Index

    tlbFormatacao.Enabled = True
    
    Select Case intFocu
        Case 0
                        
            trwLayOut_NodeClick Index, trwLayOut(Index).SelectedItem
                        
            If trwLayOut(0).Enabled Then
                lblId.BackColor = &H80000001
                lblId.ForeColor = vbWhite
            End If
                        
            If trwLayOut(1).Enabled Then
                lblSTR.BackColor = vbWindowBackground
                lblSTR.ForeColor = vbWindowText
            End If
            
            If trwLayOut(2).Enabled Then
                lblXML.BackColor = vbWindowBackground
                lblXML.ForeColor = vbWindowText
            End If
            
        Case 1
            
            trwLayOut_NodeClick Index, trwLayOut(Index).SelectedItem
            
            If trwLayOut(0).Enabled Then
                lblId.BackColor = vbWindowBackground
                lblId.ForeColor = vbWindowText
            End If
            
            If trwLayOut(1).Enabled Then
                lblSTR.BackColor = &H80000001
                lblSTR.ForeColor = vbWhite
            End If
            
            If trwLayOut(2).Enabled Then
                lblXML.BackColor = vbWindowBackground
                lblXML.ForeColor = vbWindowText
            End If
            
        Case 2
            
            trwLayOut_NodeClick Index, trwLayOut(Index).SelectedItem
            
            If trwLayOut(0).Enabled Then
                lblId.BackColor = vbWindowBackground
                lblId.ForeColor = vbWindowText
            End If
            
            If trwLayOut(1).Enabled Then
                lblSTR.BackColor = vbWindowBackground
                lblSTR.ForeColor = vbWindowText
            End If
            
            If trwLayOut(2).Enabled Then
                lblXML.BackColor = &H80000001
                lblXML.ForeColor = vbWhite
            End If
    End Select
    
End Sub

Private Sub trwLayOut_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    intFocu = Index

    If Button = vbRightButton Then
        If Index > 0 Then
            If Not trwLayOut(Index).SelectedItem Is Nothing Then
                
                If trwLayOut(Index).SelectedItem.children Then

                    If Not (Index = 1 And _
                           (cboTipoSaida.ItemData(cboTipoSaida.ListIndex) = enumTipoSaidaMensagem.SaidaCSV Or _
                            cboTipoSaida.ItemData(cboTipoSaida.ListIndex) = enumTipoSaidaMensagem.SaidaCSVXML)) Then
                       
                        With objConfigRepeticao
                            
                            .Move mdiBUS.Width - mdiBUS.ScaleWidth + Me.Left + trwLayOut(intFocu).Left + X, _
                                  mdiBUS.Height - mdiBUS.ScaleHeight + Me.Top + fraDetalhe.Top + trwLayOut(intFocu).Top + Y
                            .plngQtdRepe = flGetNodeQtd(trwLayOut(Index).SelectedItem.Text)
                            .Show
                        End With
                
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub trwLayOut_NodeCheck(Index As Integer, ByVal Node As MSComctlLib.Node)
    
    If Node.Checked Then
        objLayOut(Index).selectSingleNode("//" & flGetNodeText(Node.Text) & "/@Obrigatorio").Text = "True"
    Else
        objLayOut(Index).selectSingleNode("//" & flGetNodeText(Node.Text) & "/@Obrigatorio").Text = "False"
    End If

End Sub

Private Sub trwLayOut_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)

Dim blnFirstNode                            As Boolean
    
    If Node Is Nothing Then
        tlbFormatacao.Buttons("Promote").Enabled = False
        tlbFormatacao.Buttons("Demote").Enabled = False
        tlbFormatacao.Buttons("Up").Enabled = False
        tlbFormatacao.Buttons("Down").Enabled = False
        tlbFormatacao.Buttons("Del").Enabled = False
        Exit Sub
    End If
    
    With objLayOut(Index).selectSingleNode("//" & Node.Key)

        tlbFormatacao.Buttons("Del").Enabled = True
        
        If .parentNode.firstChild.nodeType = NODE_ELEMENT Then
            blnFirstNode = (.parentNode.firstChild.nodeName = .nodeName)
        Else
            blnFirstNode = (.previousSibling.nodeType <> NODE_ELEMENT)
        End If
        
        If Not blnFirstNode Then
            
            'Permitir a demoção
            tlbFormatacao.Buttons("Demote").Enabled = True
            tlbFormatacao.Buttons("Up").Enabled = True
            
            If .parentNode.lastChild.nodeName = .nodeName Then
                tlbFormatacao.Buttons("Down").Enabled = False
                                
                'Permitir Promoção
                If .parentNode.parentNode.nodeName <> "#document" Then
                    tlbFormatacao.Buttons("Promote").Enabled = True
                Else
                    tlbFormatacao.Buttons("Promote").Enabled = False
                End If
            Else
                tlbFormatacao.Buttons("Promote").Enabled = False
                tlbFormatacao.Buttons("Down").Enabled = True
            End If
        Else
            tlbFormatacao.Buttons("Up").Enabled = False
            
            If .parentNode.lastChild.nodeName = .nodeName Then
                tlbFormatacao.Buttons("Down").Enabled = False
                tlbFormatacao.Buttons("Demote").Enabled = False
                
                If .parentNode.parentNode.nodeName <> "#document" Then
                    'Permitir promoção
                    tlbFormatacao.Buttons("Promote").Enabled = True
                Else
                    tlbFormatacao.Buttons("Promote").Enabled = False
                End If
               
            Else
                tlbFormatacao.Buttons("Down").Enabled = True
                
                'Somente permitir ir para baixo
                tlbFormatacao.Buttons("Promote").Enabled = False
                tlbFormatacao.Buttons("Demote").Enabled = False
            End If
        End If
    End With

End Sub

Private Sub txtPrioridade_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub txtTipoMensagem_Change()

Dim lngPosicao                              As Long
Dim lngTamanho                              As Long

On Error GoTo ErrorHandler
    
    Exit Sub
    
    lngPosicao = txtTipoMensagem.SelStart
    lngTamanho = Len(txtTipoMensagem.Text)

    If Val(txtTipoMensagem.Text) <> 0 Then
        txtTipoMensagem.Text = Val(txtTipoMensagem.Text)
    Else
        txtTipoMensagem.Text = vbNullString
    End If

    If lngTamanho <> Len(txtTipoMensagem.Text) Then
        lngPosicao = lngPosicao - (lngTamanho - Len(txtTipoMensagem.Text))
        If lngPosicao < 0 Then
            lngPosicao = 0
        End If
    End If

    If Len(txtTipoMensagem.Text) >= lngPosicao Then
        txtTipoMensagem.SelStart = lngPosicao
    Else
        txtTipoMensagem.SelStart = Len(txtTipoMensagem.Text)
    End If

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - txtTipoMensagem_Change"
    
End Sub

'Salvar as informações correntes de tipo de mensagem.
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strRetorno          As String
Dim objListItem         As ListItem
Dim strKey              As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
    
    If strRetorno <> "" Then
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
    
    If Not IsNull(dtpDataFimVigencia.Value) Then
        If xmlTipoMensagem.documentElement.selectSingleNode("//DT_FIM_VIGE_MESG").Text <> gstrDataVazia Then
            If fgDtXML_To_Date(xmlTipoMensagem.documentElement.selectSingleNode("//DT_FIM_VIGE_MESG").Text) <> dtpDataFimVigencia.Value Then
                If MsgBox("Deseja desativar o registro a partir do dia " & dtpDataFimVigencia.Value & " ?", vbYesNo, "Atributos Mensagens") = vbNo Then Exit Sub
            End If
        End If
    End If
    
    Call fgCursor(True)
    
    fgLockWindow Me.hwnd
    
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlTipoMensagem.documentElement.xml, _
                         vntCodErro, _
                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
        
    Call fgCursor(False)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
       
    If Not lstTipoMensagem.SelectedItem Is Nothing Then
        If strOperacao = "Incluir" Then
            strKey = "EVE" & _
                      Format(cboTipoSaida.ItemData(cboTipoSaida.ListIndex), "0000") & _
                      Trim(txtTipoMensagem.Text)
        Else
            strKey = lstTipoMensagem.SelectedItem.Key
        End If
    End If
       
    strKeyItemSelected = strKey
    
    flCarregarTipoMensagem

    If strKey <> "" Then
        lstTipoMensagem.ListItems(strKey).EnsureVisible
        lstTipoMensagem.ListItems(strKey).Selected = True
        lstTipoMensagem.HideSelection = False
    End If
            
    strOperacao = "Alterar"
    
    With xmlTipoMensagem
        .selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Ler"
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlTipoMensagem.loadXML objMiu.Executar(xmlTipoMensagem.xml, _
                                            vntCodErro, _
                                            vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    
    fgLockWindow 0
    
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flSalvar", 0
    
End Sub

'Mover valores do formulário para XML para envio ao objeto de negócio.
Private Function flInterfaceToXml() As String

Dim xmlAtributo                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim lngNodePosition                         As Long
Dim intTipoParteSaidaAux                    As enumTipoParteSaida

On Error GoTo ErrorHandler

    With xmlTipoMensagem
        .selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = strOperacao
        .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text = txtTipoMensagem.Text
        .selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text = fgLimpaCaracterEspecial(txtDescricao.Text)
        .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text = cboTipoEvento.ItemData(cboTipoEvento.ListIndex)
        .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG_SAID").Text = cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
        .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text = cboDelimitador.Text
        .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text = fgLimpaCaracterEspecial(txtPrioridade.Text)
        .selectSingleNode("//Grupo_TipoMensagem/NO_TITU_MESG").Text = fgLimpaCaracterEspecial(txtNomeTituloMesg.Text)

        If strOperacao <> "Incluir" Then
            .selectSingleNode("//Grupo_TipoMensagem/CO_TEXT_XML").Text = lstTipoMensagem.SelectedItem.Tag
        End If

        .selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").Text = ""

        .selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").appendChild fgCreateCDATASection(flGerarXSD())
        '.selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").appendChild fgCreateCDATASection(GeraXSD(objLayOut(2)))

        .selectSingleNode("//Grupo_TipoMensagem/DT_INIC_VIGE_MESG").Text = fgDt_To_Xml(dtpDataInicioVigencia.Value)

        If IsNull(dtpDataFimVigencia.Value) Then
            .selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text = ""
        Else
            .selectSingleNode("//Grupo_TipoMensagem/DT_FIM_VIGE_MESG").Text = fgDt_To_Xml(dtpDataFimVigencia.Value)
        End If

    End With

    If strEstruturaAtributo = vbNullString Then
        'Manter a estrutura do atributo pois caso o evento não tenha atributos
        'é necessário guardar a estrutura para a proxima inclusao ou alteração
        'vai passar aqui somente na primeira vez
        strEstruturaAtributo = xmlMapaNavegacao.selectSingleNode("//Repeat_TipoMensagemAtributo/Grupo_TipoMensagemAtributo").xml
    End If

    'Remover todos os nós e evento_atributo
    For Each xmlNode In xmlTipoMensagem.selectNodes("//Repeat_TipoMensagemAtributo/*")
        xmlTipoMensagem.selectSingleNode("//Repeat_TipoMensagemAtributo").removeChild xmlNode
    Next

    If xmlTipoMensagem.selectSingleNode("//Repeat_TipoMensagemAtributo") Is Nothing Then
        Call fgAppendNode(xmlTipoMensagem, "Grupo_TipoMensagem", "Repeat_TipoMensagemAtributo", "")
    End If

    'Monta Atributos
    'Atributos para ID
        
    lngNodePosition = 1
    flObterAtributos objLayOut(0).selectSingleNode("//XML"), _
                     lngNodePosition, _
                     enumTipoParteSaida.ParteId, _
                     1
    
    Select Case cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
        Case enumTipoSaidaMensagem.SaidaCSV, _
             enumTipoSaidaMensagem.SaidaCSVXML
             intTipoParteSaidaAux = ParteCSV
        
        Case enumTipoSaidaMensagem.SaidaString, _
             enumTipoSaidaMensagem.SaidaStringXML
             intTipoParteSaidaAux = ParteSTR
    End Select
    
    'Atributos para String
    lngNodePosition = 1
    flObterAtributos objLayOut(1).selectSingleNode("//XML"), _
                     lngNodePosition, _
                     intTipoParteSaidaAux, _
                     1

    'Atributos para XML
    lngNodePosition = 1
    flObterAtributos objLayOut(2).selectSingleNode("//XML"), _
                     lngNodePosition, _
                     enumTipoParteSaida.ParteXML, _
                     1

    Exit Function
ErrorHandler:

    Set xmlAtributo = Nothing

    fgRaiseError App.EXEName, Me.Name, "flInterfaceToXML", 0

End Function

'Obtém os atributos do tipo de memsagem a partir do xml com layout.
Private Sub flObterAtributos(ByRef pxmlNodeBase As MSXML2.IXMLDOMNode, _
                             ByRef plngPosicaoBase As Long, _
                             ByRef penumTipoParteSaida As enumTipoParteSaida, _
                             ByVal plngNivel As Long)

Dim xmlAtributo                             As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
    
On Error GoTo ErrorHandler
    
    For Each xmlNode In pxmlNodeBase.selectNodes("//" & pxmlNodeBase.nodeName & "/*")

        Set xmlAtributo = CreateObject("MSXML2.DOMDocument.4.0")

        xmlAtributo.loadXML strEstruturaAtributo

        xmlAtributo.selectSingleNode("//TP_MESG").Text = fgLimpaCaracterEspecial(txtTipoMensagem.Text)
        xmlAtributo.selectSingleNode("//TP_FORM_MESG_SAID").Text = cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
        xmlAtributo.selectSingleNode("//NO_ATRB_MESG").Text = xmlNode.nodeName
        xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = penumTipoParteSaida
        xmlAtributo.selectSingleNode("//NU_ORDE_AGRU_ATRB").Text = plngPosicaoBase
        xmlAtributo.selectSingleNode("//IN_OBRI_ATRB").Text = Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True")
        xmlAtributo.selectSingleNode("//NU_NIVE_MESG_ATRB").Text = plngNivel
        xmlAtributo.selectSingleNode("//QT_REPE").Text = CLng(xmlNode.selectSingleNode("@QT_REPE").Text)
        plngPosicaoBase = plngPosicaoBase + 1
        
        Call fgAppendXML(xmlTipoMensagem, "Repeat_TipoMensagemAtributo", xmlAtributo.xml)
        
        Set xmlAtributo = Nothing
        
        If Not xmlNode.childNodes(1) Is Nothing Then
            'Node Tem filho
            Call flObterAtributos(xmlNode, plngPosicaoBase, penumTipoParteSaida, plngNivel + 1)
        End If
        
    Next

    Exit Sub
ErrorHandler:
    
    Set xmlNode = Nothing
    Set xmlAtributo = Nothing
    
    fgRaiseError App.EXEName, "frmTipoMensagem", "flObterAtributos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Function flGetNodeText(ByVal pstrNodeText As String) As String

On Error GoTo ErrorHandler
    
    If InStr(1, pstrNodeText, "(") > 0 Then
        flGetNodeText = Trim$(Mid$(pstrNodeText, 1, InStr(1, pstrNodeText, "(") - 1))
    Else
        flGetNodeText = pstrNodeText
    End If

    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmTipoMensagem", "flGetNodeText", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Function flGetNodeQtd(ByVal pstrNodeText As String) As Long

On Error GoTo ErrorHandler
    
    If InStr(1, pstrNodeText, "(") > 0 Then
        flGetNodeQtd = CLng(Mid$(pstrNodeText, InStr(1, pstrNodeText, "(") + 1, _
                       InStr(1, pstrNodeText, ")") - InStr(1, pstrNodeText, "(") - 1))
    Else
        flGetNodeQtd = 0
    End If
    
    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmTipoMensagem", "flGetNodeQtd", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Validar os valores informados para o tipo de mensagem.
Private Function flValidarCampos() As String

On Error GoTo ErrorHandler
    
    If Len(txtTipoMensagem.Text) = 0 Then
        flValidarCampos = "Informe o código do tipo de mensagem."
        txtTipoMensagem.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescricao) = "" Then
        flValidarCampos = "Informe a descrição do tipo de mensagem."
        txtDescricao.SetFocus
        Exit Function
    End If
            
    If cboTipoEvento.ListIndex < 0 Then
        flValidarCampos = "Informe o tipo da mensagem ."
        cboTipoEvento.SetFocus
        Exit Function
    End If
    
    If cboTipoSaida.ListIndex < 0 Then
        flValidarCampos = "Informe o tipo de saída da mensagem."
        cboTipoSaida.SetFocus
        Exit Function
    End If
    
    If txtPrioridade.Text = vbNullString Then
        flValidarCampos = "Informe a prioridade da mensagem."
        txtPrioridade.SetFocus
        Exit Function
    End If
    
    If txtNomeTituloMesg.Text = vbNullString Then
        flValidarCampos = "Informe o Nome do Título da mensagem."
        txtNomeTituloMesg.SetFocus
        Exit Function
    Else
        If Len(txtNomeTituloMesg) > 80 Then
            flValidarCampos = "Nome do Título da mensagem deve possuir menos de 80 caracteres."
            txtNomeTituloMesg.SetFocus
            Exit Function
        End If
    End If
    
    If cboTipoSaida.ItemData(cboTipoSaida.ListIndex) = enumTipoSaidaMensagem.SaidaCSV Then
        If Trim$(cboDelimitador.Text) = vbNullString Then
            flValidarCampos = "Selecione o caracter delimitador."
            cboDelimitador.SetFocus
            Exit Function
        End If
    End If
        
    If cboTipoEvento.ItemData(cboTipoEvento.ListIndex) <> enumNaturezaMensagem.MensagemECO Then
        If trwLayOut(0).Nodes.Count = 0 Then
            flValidarCampos = "Selecione os Atributos do Identificador."
            Exit Function
        End If
    End If
    
    flValidarCampos = ""
    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmTipoMensagem", "flValidarCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Excluir o tipo de mensagem corrente.
Private Sub flExcluir()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    If MsgBox("Confirma Exclusão ?", vbYesNo, "Tipos de Mensagem") = vbNo Then Exit Sub

    strOperacao = "Excluir"
        
    Call fgCursor(True)
    fgLockWindow Me.hwnd
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMiu = Nothing
    
    Call fgCursor(False)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
   
    Call flInicializar
    Call flCarregarTipoMensagem
    Call flLimparCampos

    fgLockWindow 0

    Exit Sub

ErrorHandler:
    
    fgLockWindow 0
    
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flExcluir", 0
End Sub

'Gerar o XSD(validador) para o tipo de mensagem corrente.
Private Function flGerarXSD() As String

Dim xmlStruct                               As DOMDocument40
Dim xmlType                                 As DOMDocument40
Dim xmlElement                              As IXMLDOMElement
Dim strTipoParteAux                         As String

    Set xmlStruct = New DOMDocument40
    fgAppendNode xmlStruct, "", "xsd:schema", ""
    fgAppendAttribute xmlStruct, "schema", "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
    
    Set xmlType = New DOMDocument40
    fgAppendNode xmlType, "", "Tipos", ""
    fgAppendAttribute xmlType, "Tipos", "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"
    
    'Inicio do XSD
    fgAppendNode xmlStruct, "schema", "xsd:element", ""
    fgAppendAttribute xmlStruct, "schema/element", "name", "Saida"
    fgAppendAttribute xmlStruct, "schema/element", "type", "TipoSaida"
    
    'Incluir os Tipos de Saida da mensagem no ComplexType da Saida
    fgAppendNode xmlStruct, "schema", "xsd:complexType", ""
    fgAppendAttribute xmlStruct, "schema/complexType", "name", "TipoSaida"
    fgAppendNode xmlStruct, "schema/complexType", "xsd:sequence", ""
    
    If trwLayOut(0).Nodes.Count > 0 Then
        'Incluir TipoID
        fgAppendNode xmlStruct, "complexType[@name='TipoSaida']/sequence", "xsd:element", ""
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "name", "SaidaID"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "type", "TipoSaidaID"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "minOccurs", "1"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "maxOccurs", "1"
    End If
    
    If trwLayOut(1).Nodes.Count > 0 Then
        'Verifica se a saida é Str ou Csv
        Select Case cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
            Case enumTipoSaidaMensagem.SaidaString, _
                 enumTipoSaidaMensagem.SaidaStringXML
                'Incluir TipoSTR
                fgAppendNode xmlStruct, "complexType[@name='TipoSaida']/sequence", "xsd:element", ""
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "name", "SaidaSTR"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "type", "TipoSaidaSTR"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "minOccurs", "1"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "maxOccurs", "1"
            Case enumTipoSaidaMensagem.SaidaCSV, _
                 enumTipoSaidaMensagem.SaidaCSVXML
                'Incluir TipoCSV
                fgAppendNode xmlStruct, "complexType[@name='TipoSaida']/sequence", "xsd:element", ""
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "name", "SaidaCSV"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "type", "TipoSaidaCSV"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "minOccurs", "1"
                fgAppendAttribute xmlStruct, _
                                 "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                                 "maxOccurs", "1"
                                 
                'Coloca a definição do Separador
                fgAppendNode xmlType, "Tipos", "xsd:simpleType", vbNullString
                fgAppendAttribute xmlType, "Tipos/simpleType[position()=last()]", _
                                  "name", "TipoCSVSeparador"
                fgAppendNode xmlType, "Tipos/simpleType[@name='TipoCSVSeparador']", _
                             "xsd:restriction", vbNullString
                fgAppendAttribute xmlType, "Tipos/simpleType[@name='TipoCSVSeparador']/restriction", _
                                  "base", "xsd:string"
                fgAppendNode xmlType, "Tipos/simpleType[@name='TipoCSVSeparador']/restriction", _
                             "xsd:pattern", vbNullString
                fgAppendAttribute xmlType, "Tipos/simpleType[@name='TipoCSVSeparador']/restriction/pattern", _
                                  "value", cboDelimitador.Text
        End Select
    End If
        
    If trwLayOut(2).Nodes.Count > 0 Then
        'Incluir TipoXML
        fgAppendNode xmlStruct, "complexType[@name='TipoSaida']/sequence", "xsd:element", ""
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "name", "SaidaXML"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "type", "TipoSaidaXML"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "minOccurs", "1"
        fgAppendAttribute xmlStruct, _
                         "complexType[@name='TipoSaida']/sequence/element[position()=last()]", _
                         "maxOccurs", "1"
    End If
                
    'Obter XSD dos Tipos
    If trwLayOut(0).Nodes.Count > 0 Then
        'Obter XSD da Parte ID
        
        fgAppendNode xmlStruct, "schema", "xsd:complexType", ""
        fgAppendAttribute xmlStruct, "schema/complexType[position()=last()]", "name", "TipoSaidaID"
        fgAppendNode xmlStruct, "schema/complexType[@name='TipoSaidaID']", "xsd:sequence", ""
                
        flSubXSD objLayOut(0).childNodes(0), _
                 xmlStruct, _
                 xmlType, _
                 "TipoSaidaID", _
                 "ID"
        
    End If
    
    If trwLayOut(1).Nodes.Count > 0 Then
        'Obter XSD da parte STR / CSV
        'Verifica se a saida é Str ou Csv
        Select Case cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
            Case enumTipoSaidaMensagem.SaidaString, _
                 enumTipoSaidaMensagem.SaidaStringXML
                'Incluir TipoSTR
                strTipoParteAux = "STR"
            Case enumTipoSaidaMensagem.SaidaCSV, _
                 enumTipoSaidaMensagem.SaidaCSVXML
                strTipoParteAux = "CSV"
        End Select

        fgAppendNode xmlStruct, "schema", "xsd:complexType", ""
        fgAppendAttribute xmlStruct, "schema/complexType[position()=last()]", "name", "TipoSaida" & strTipoParteAux
        fgAppendNode xmlStruct, "schema/complexType[@name='TipoSaida" & strTipoParteAux & "']", "xsd:sequence", ""
                
        flSubXSD objLayOut(1).childNodes(0), _
                 xmlStruct, _
                 xmlType, _
                 "TipoSaida" & strTipoParteAux, _
                 strTipoParteAux

    End If

    If trwLayOut(2).Nodes.Count > 0 Then
        'Obter XSD da Parte XML
        
        fgAppendNode xmlStruct, "schema", "xsd:complexType", ""
        fgAppendAttribute xmlStruct, "schema/complexType[position()=last()]", "name", "TipoSaidaXML"
        fgAppendNode xmlStruct, "schema/complexType[@name='TipoSaidaXML']", "xsd:sequence", ""
                
        flSubXSD objLayOut(2).childNodes(0), _
                 xmlStruct, _
                 xmlType, _
                 "TipoSaidaXML", _
                 "XML"
        
    End If

    'Incluir os Tipos Gerados no XSD
    For Each xmlElement In xmlType.selectNodes("//Tipos/*")
        xmlStruct.childNodes(0).appendChild xmlElement
    Next

    flGerarXSD = xmlStruct.xml
    
    Set xmlType = Nothing
    Set xmlStruct = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlType = Nothing
    Set xmlStruct = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "flMontaXSD Function", 0

End Function

'Sub-função utilizada na geração do XSD do tipo de mensagem.
Private Sub flSubXSD(ByRef pxmlNodeBase As IXMLDOMNode, _
                     ByRef pxmlStruct As DOMDocument40, _
                     ByRef pxmlType As DOMDocument40, _
                     ByVal pstrNomePai As String, _
                     ByRef pstrTipoParte As String)

Dim xmlNode                                 As IXMLDOMNode
Dim strNovoPai                              As String
Dim blnSinal                                As Boolean

On Error GoTo ErrorHandler
    
    For Each xmlNode In pxmlNodeBase.selectNodes("./*")
       
        'Inclui element na sequence da estrutura corrente (pstrNomePai)
        fgAppendNode pxmlStruct, _
                     "complexType[@name='" & pstrNomePai & "']/sequence", _
                     "xsd:element", vbNullString
        fgAppendAttribute pxmlStruct, _
                          "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                          "name", xmlNode.nodeName
        fgAppendAttribute pxmlStruct, _
                          "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                          "type", "Tipo" & pstrTipoParte & xmlNode.nodeName
        fgAppendAttribute pxmlStruct, _
                          "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                          "minOccurs", _
                          Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True")
        
        If CLng(xmlNode.parentNode.selectSingleNode("@QT_REPE").Text) = 0 Then
            fgAppendAttribute pxmlStruct, _
                              "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                              "maxOccurs", 1
        Else
            fgAppendAttribute pxmlStruct, _
                              "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                              "maxOccurs", _
                              CLng(xmlNode.parentNode.selectSingleNode("@QT_REPE").Text)
        End If
        
        If xmlNode.childNodes(1) Is Nothing Then
            'Incluir Tipo quando tipo de node for filho
            fgAppendNode pxmlType, "Tipos", "xsd:simpleType", ""
            fgAppendAttribute pxmlType, "simpleType[position()=last()]", _
                             "name", "Tipo" & pstrTipoParte & xmlNode.nodeName
            fgAppendNode pxmlType, "simpleType[position()=last()]", _
                        "xsd:restriction", ""
            fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction", _
                             "base", "xsd:string"

            'Indicador de sinal negativo
            blnSinal = (CLng(xmlNode.selectSingleNode("@IN_ATRB_PRMT_VALO_NEGT").Text) = enumIndicadorSimNao.sim)
            
            If pstrTipoParte = "STR" Or pstrTipoParte = "ID" Then
                'Usar lenght para tamanho exato
                If CLng(xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text) = enumTipoDadoAtributo.Numerico Then
                    'Valor Numerico
                    
                    fgAppendNode pxmlType, "simpleType[position()=last()]/restriction", _
                                 "xsd:pattern", ""
                    fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/pattern", _
                                 "value", _
                                 IIf(blnSinal, "\-[0-9]{" & (CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) - 1) & "}|", "") & _
                                 "[0-9]{" & CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) & "}"
                    '+ CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)
                    'não acrescentar a quantidade de casas pois o tamanho indica o tamanho total do atributo
                Else
                    'Valor Alfanumerico
                    fgAppendNode pxmlType, "simpleType[position()=last()]/restriction", _
                                 "xsd:length", ""
                    fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/length", _
                                 "value", _
                                 CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) + CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)
                End If
            Else
                If CLng(xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text) = enumTipoDadoAtributo.Numerico Then
                    'Para tipo Numero = Pattern xsd:pattern value="[0-9]{1,4}"
                    fgAppendNode pxmlType, "simpleType[position()=last()]/restriction", _
                             "xsd:pattern", ""
                    
                    If CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text) > 0 Then
                        'Se tag permite casas decimais colocar a "," no pattern
                        fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/pattern", _
                                 "value", _
                                 IIf(blnSinal, "[\-]?", "") & "[0-9]{" & _
                                 Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True") & "," & _
                                 (CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) - CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)) & "}," & _
                                 "[0-9]{1," & CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text) & "}" & _
                                 "|" & IIf(blnSinal, "[\-]?", "") & "[0-9]{" & _
                                 Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True") & "," & _
                                 (CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) - CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)) & "}"
                    Else
                        fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/pattern", _
                                 "value", _
                                 IIf(blnSinal, "[\-]?", "") & "[0-9]{" & _
                                 Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True") & "," & _
                                 CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) & "}"
                    End If
                Else
                    'Para tipo alfa = minLengthe maxlength
                    fgAppendNode pxmlType, "simpleType[position()=last()]/restriction", _
                                 "xsd:minLength", ""
                    fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/minLength", _
                                 "value", _
                                 Abs(xmlNode.selectSingleNode("@Obrigatorio").Text = "True")
                    fgAppendNode pxmlType, "simpleType[position()=last()]/restriction", _
                                 "xsd:maxLength", ""
                    fgAppendAttribute pxmlType, "simpleType[position()=last()]/restriction/maxLength", _
                                 "value", _
                                 CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text) + CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)
                End If
            End If
        
            If pstrTipoParte = "CSV" Then
                If xmlNode.nodeName <> pxmlNodeBase.selectSingleNode("./*[position()=last()]").nodeName Then
                    'Incluir Delimitador (Menos para ultima posição)
                    fgAppendNode pxmlStruct, _
                                 "complexType[@name='" & pstrNomePai & "']/sequence", _
                                 "xsd:element", vbNullString
                    fgAppendAttribute pxmlStruct, _
                                      "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                                      "name", "Separador"
                    fgAppendAttribute pxmlStruct, _
                                      "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                                      "type", "TipoCSVSeparador"
                    fgAppendAttribute pxmlStruct, _
                                      "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                                      "minOccurs", 1
                    fgAppendAttribute pxmlStruct, _
                                      "complexType[@name='" & pstrNomePai & "']/sequence/element[position()=last()]", _
                                      "maxOccurs", 1
                End If
            End If
        Else
            'Node tem Filho --> Incluir complexType para nova estrutura
            
            fgAppendNode pxmlStruct, _
                         "schema", _
                         "xsd:complexType", vbNullString
                         
            strNovoPai = "Tipo" & pstrTipoParte & xmlNode.nodeName
            
            fgAppendAttribute pxmlStruct, _
                              "schema/complexType[position()=last()]", _
                              "name", strNovoPai
            fgAppendNode pxmlStruct, _
                         "complexType[@name='" & strNovoPai & "']", _
                         "xsd:sequence", ""
                        
            'Chama novamente para obter XSD dos filhos
            flSubXSD xmlNode, _
                     pxmlStruct, _
                     pxmlType, _
                     strNovoPai, _
                     pstrTipoParte
            
        End If
    Next
        
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flSubXSD", 0

End Sub

'Proteger chave do tipo de mensagem em operações de alteração.
Private Sub flProtegerChave()

   strOperacao = "Alterar"
   txtTipoMensagem.Enabled = False
   cboTipoSaida.Enabled = False
   cboTipoEvento.Enabled = False
   dtpDataInicioVigencia.Enabled = False
   tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True

End Sub

'Verifica se as datas informadas para o tipo de mensagem estão em um período vigente.
Private Function flRegistroVigente(ByVal pdtmDataServidor As Date, _
                                   ByVal pdtmDataInicio As String, _
                                   ByVal pdtmDataFim As String) As Boolean

On Error GoTo ErrorHandler
    
    If fgDtXML_To_Date(pdtmDataInicio) > pdtmDataServidor Then
        flRegistroVigente = True
        Exit Function
    End If
    
    If pdtmDataFim = gstrDataVazia Then
        flRegistroVigente = False
        Exit Function
    End If
    
    If fgDtXML_To_Date(pdtmDataFim) <= pdtmDataServidor Then
        flRegistroVigente = True
        Exit Function
    End If

    Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flRegistroVigente", 0

End Function

'Carregar combo com tipos de delimitadores (utilizado para CSV)
Private Sub flCarregaComboDelimitador()
    
    cboDelimitador.AddItem ";"
    cboDelimitador.AddItem "|"

End Sub

