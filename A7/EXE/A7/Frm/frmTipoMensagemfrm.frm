VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmTipoMensagem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Tipos de Mensagens"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13530
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   13530
   Begin VB.Frame Frame1 
      Height          =   2100
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   13485
      Begin MSComctlLib.ListView lstTipoMensagem 
         Height          =   1875
         Left            =   75
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
            Text            =   "Tipo Saída"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Natureza"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Tipo de Saída"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Delimitador"
            Object.Width           =   1931
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Prioridade"
            Object.Width           =   1614
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Data Início"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Object.Tag             =   "1"
            Text            =   "Data Fim"
            Object.Width           =   2117
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
      Height          =   7035
      Left            =   45
      TabIndex        =   19
      Top             =   2055
      Width           =   13485
      Begin VB.ComboBox cboTipoMensagemSaida 
         Height          =   315
         Left            =   6255
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   375
         Width           =   3540
      End
      Begin VB.ComboBox cboDelimitador 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5085
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1035
         Width           =   900
      End
      Begin VB.CommandButton cmdToID 
         Caption         =   "id ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7005
         TabIndex        =   12
         Top             =   1710
         Width           =   600
      End
      Begin FPSpread.vaSpread sprID 
         Height          =   1275
         Left            =   7665
         TabIndex        =   18
         Top             =   1710
         Width           =   5730
         _Version        =   196608
         _ExtentX        =   10107
         _ExtentY        =   2249
         _StockProps     =   64
         AutoCalc        =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   5
         MaxRows         =   1
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   11
         SpreadDesigner  =   "frmTipoMensagemfrm.frx":0000
         UserResize      =   0
      End
      Begin MSComCtl2.UpDown updXML 
         Height          =   645
         Left            =   7350
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5565
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1138
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdToXML 
         Caption         =   "xml ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7005
         TabIndex        =   16
         Top             =   5040
         Width           =   600
      End
      Begin FPSpread.vaSpread sprXML 
         Height          =   1875
         Left            =   7665
         TabIndex        =   33
         Top             =   5010
         Width           =   5715
         _Version        =   196608
         _ExtentX        =   10081
         _ExtentY        =   3307
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   3
         MaxRows         =   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   11
         SpreadDesigner  =   "frmTipoMensagemfrm.frx":0396
         UserResize      =   0
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.ListView lstAtributo 
         Height          =   5220
         Left            =   75
         TabIndex        =   11
         Top             =   1695
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   9208
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome Lógico"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nome Físico"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tamanho"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Decimais"
            Object.Width           =   1235
         EndProperty
      End
      Begin VB.TextBox txtPrioridade 
         Height          =   315
         Left            =   6090
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1035
         Width           =   840
      End
      Begin VB.ComboBox cboTipoSaida 
         Height          =   315
         Left            =   2595
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1035
         Width           =   2400
      End
      Begin VB.ComboBox cboTipoEvento 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1035
         Width           =   2415
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
         Top             =   375
         Width           =   5070
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
         TabIndex        =   26
         Top             =   240
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
            Format          =   19660801
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
            Format          =   19660801
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
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   330
            Width           =   300
         End
      End
      Begin MSComCtl2.UpDown updID 
         Height          =   645
         Left            =   7380
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2145
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1138
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdToSTR 
         Caption         =   "str ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6975
         TabIndex        =   14
         Top             =   3090
         Width           =   600
      End
      Begin MSComCtl2.UpDown updSTR 
         Height          =   645
         Left            =   7335
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1138
         _Version        =   393216
         Enabled         =   0   'False
      End
      Begin MSComCtl2.UpDown updCSV 
         Height          =   645
         Left            =   7320
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3510
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   1138
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdToCSV 
         Caption         =   "csv ->"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6975
         TabIndex        =   31
         Top             =   3090
         Visible         =   0   'False
         Width           =   600
      End
      Begin FPSpread.vaSpread sprString 
         Height          =   1875
         Left            =   7680
         TabIndex        =   36
         Top             =   3060
         Width           =   5700
         _Version        =   196608
         _ExtentX        =   10054
         _ExtentY        =   3307
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   5
         MaxRows         =   1
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   11
         SpreadDesigner  =   "frmTipoMensagemfrm.frx":0692
         UserResize      =   0
      End
      Begin FPSpread.vaSpread sprCSV 
         Height          =   1875
         Left            =   7680
         TabIndex        =   35
         Top             =   3060
         Width           =   5700
         _Version        =   196608
         _ExtentX        =   10054
         _ExtentY        =   3307
         _StockProps     =   64
         Enabled         =   0   'False
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   16777215
         MaxCols         =   3
         MaxRows         =   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   11
         SpreadDesigner  =   "frmTipoMensagemfrm.frx":0A1F
         UserResize      =   0
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Mensagem de Saída"
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
         Left            =   6210
         TabIndex        =   37
         Top             =   150
         Width           =   2460
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Atributos Associados"
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
         Left            =   7635
         TabIndex        =   30
         Top             =   1485
         Width           =   1785
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Atributos Disponíveis"
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
         TabIndex        =   29
         Top             =   1470
         Width           =   1830
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
         Left            =   6105
         TabIndex        =   25
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
         Left            =   5040
         TabIndex        =   24
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
         Left            =   2580
         TabIndex        =   23
         Top             =   810
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
         TabIndex        =   22
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   150
         Width           =   870
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   9990
      TabIndex        =   34
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
      Left            =   45
      Top             =   8895
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
            Picture         =   "frmTipoMensagemfrm.frx":0D1B
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":0E2D
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":1147
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":1499
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":15AB
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":18C5
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":1BDF
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoMensagemfrm.frx":1EF9
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTipoMensagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Empresa        : Regerbanc
'Pacote         :
'Classe         : frmTipoMensagem
'Data Criação   : 21/07/2003
'Objetivo       :
'
'Analista       : Marcelo Kida
'
'Programador    : Marcelo Kida
'Data           : 01/07/2003
'
'Teste          :
'Autor          :
'
'Data Alteração : 23/09/2003
'Autor          : Douglas Cavalcante
'Objetivo       : Seguindo os Padrões Definidos foi alterado o Form.
'
'Data Alteração : 26/09/2003
'Autor          : Eder Andrade
'Objetivo       : Impedir que sejam selecionadas datas de vigência inválidas
'
'Data Alteração : 01/10/2003
'Autor          : Marcelo Kida
'Objetivo       : Impedir que sejam associados atributos não vigentes
'
'Data Alteração : 01/10/2003
'Autor          : Marcelo Kida
'Objetivo       : Correção do erro ao tentar remover um atributo associado para o tipo de saida CSV
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlTipoMensagem                     As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmTipoMensagem"
Private strEstruturaAtributo                As String

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Enum enumMoveUpDown
    Up = 1
    Down = 2
End Enum

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


Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

On Error GoTo ErrorHandler
    
    dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = Null
    
    Set objMiu = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.loadXML objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade)
    
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
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Sub

Private Sub flCarregarlistAtributo()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim dtmDataServidor                         As Date

On Error GoTo ErrorHandler
        
    lstAtributo.ListItems.Clear
    lstAtributo.HideSelection = False
        
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_AtributoMensagem/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_AtributoMensagem/IN_VIGE").Text = "N"
    strPropriedades = xmlMapaNavegacao.selectSingleNode("//Grupo_AtributoMensagem").xml
    
    strLerTodos = objMiu.Executar(strPropriedades)
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    dtmDataServidor = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    For Each xmlNode In xmlLerTodos.selectNodes("//Repeat_AtributoMensagem/Grupo_AtributoMensagem")
        
        With lstAtributo.ListItems.Add(, "K" & xmlNode.selectSingleNode("NO_ATRB_MESG").Text, xmlNode.selectSingleNode("NO_TRAP_ATRB").Text)
            
            .SubItems(1) = xmlNode.selectSingleNode("NO_ATRB_MESG").Text
            .SubItems(2) = flTipoDadoToSTR(CLng(xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text))
            .SubItems(3) = xmlNode.selectSingleNode("QT_CTER_ATRB").Text
            .SubItems(4) = xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
            .Tag = xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text
            
            If flRegistroVigente(dtmDataServidor, _
                                 xmlNode.selectSingleNode("DT_INIC_VIGE_ATRB_MESG").Text, _
                                 xmlNode.selectSingleNode("DT_FIM_VIGE_ATRB_MESG").Text) Then
                .ForeColor = vbRed
                .ListSubItems.Item(1).ForeColor = vbRed
                .ListSubItems.Item(2).ForeColor = vbRed
                .ListSubItems.Item(3).ForeColor = vbRed
                .ListSubItems.Item(4).ForeColor = vbRed
            End If
        End With
    Next
    
    lstAtributo.SortKey = 0
    lstAtributo.SortOrder = lvwAscending
    lstAtributo.Sorted = False

    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing

    fgRaiseError App.EXEName, "frmAtributo", "flCarregaListView", 0
   
End Sub

Private Sub flCarregarComboTipoMensagemSaida()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
        
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaNavegacao.selectSingleNode("//Grupo_Propriedades/Grupo_TipoMensagemSaida/@Operacao").Text = "LerTodos"
    
    strPropriedades = xmlMapaNavegacao.selectSingleNode("//Grupo_TipoMensagemSaida").xml
    
    strLerTodos = objMiu.Executar(strPropriedades)
    
    Set objMiu = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    cboTipoMensagemSaida.Clear
    
    For Each xmlNode In xmlLerTodos.selectNodes("//Repeat_TipoMensagemSaida/Grupo_TipoMensagemSaida")
        
        cboTipoMensagemSaida.AddItem xmlNode.selectSingleNode("DE_MESG_SAID").Text
        cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.NewIndex) = CLng(xmlNode.selectSingleNode("CO_MESG_SAID").Text)
    
    Next

    Set xmlLerTodos = Nothing

    Exit Sub
ErrorHandler:

    Set objMiu = Nothing
    Set xmlLerTodos = Nothing

    fgRaiseError App.EXEName, "frmAtributo", "flCarregaListView", 0
   
End Sub


Private Sub flCarregarTipoMensagem()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim xmlDomTipoMensagem                      As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strLerTodos                             As String

On Error GoTo ErrorHandler

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlMapaNavegacao.selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "LerTodos"
    strLerTodos = objMiu.Executar(xmlMapaNavegacao.selectSingleNode("//Grupo_TipoMensagem").xml)
    Set objMiu = Nothing

    lstTipoMensagem.ListItems.Clear

    If strLerTodos = "" Then Exit Sub
    
    Set xmlDomTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlDomTipoMensagem.loadXML (strLerTodos)
    
    For Each xmlNode In xmlDomTipoMensagem.selectNodes("//Repeat_TipoMensagem/*")
        With lstTipoMensagem.ListItems.Add(, "EVE" & _
                                            Format(xmlNode.selectSingleNode("CO_MESG_SAID").Text, "0000") & _
                                            xmlNode.selectSingleNode("TP_MESG").Text, _
                                            xmlNode.selectSingleNode("TP_MESG").Text)
            
            .Tag = xmlNode.selectSingleNode("CO_TEXT_XML").Text
            
            .SubItems(1) = xmlNode.selectSingleNode("NO_TIPO_MESG").Text
            .SubItems(2) = xmlNode.selectSingleNode("DE_MESG_SAID").Text
            .SubItems(3) = flTipoEventoToSTR(CLng(xmlNode.selectSingleNode("TP_NATZ_MESG").Text))
            .SubItems(4) = flTipoSaidaToSTR(CLng(xmlNode.selectSingleNode("TP_FORM_MESG").Text))
            .SubItems(5) = xmlNode.selectSingleNode("TP_CTER_DELI").Text
            .SubItems(6) = xmlNode.selectSingleNode("CO_PRIO_FILA_SAID_MESG").Text
            .SubItems(7) = Format(fgDtXML_To_Date(xmlNode.selectSingleNode("DT_INIC_VIGE_MESG").Text), gstrMascaraDataDtp)
            
            If CStr(xmlNode.selectSingleNode("DT_FIM_VIGE_MESG").Text) <> gstrDataVazia Then
                .SubItems(8) = Format(fgDtXML_To_Date(xmlNode.selectSingleNode("DT_FIM_VIGE_MESG").Text), gstrMascaraDataDtp)
            Else
                .SubItems(8) = ""
            End If
            
        End With
    Next
   
    Set xmlDomTipoMensagem = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set xmlDomTipoMensagem = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarListAtributo", 0

End Sub

Private Sub cboTipoEvento_Click()
    
    If cboTipoEvento.ListIndex = -1 Then Exit Sub

    If cboTipoEvento.ItemData(cboTipoEvento.ListIndex) = enumNaturezaMensagem.MensagemECO Then
        cboTipoSaida.ListIndex = 0
        cboTipoSaida.Enabled = False
        
        cmdToID.Enabled = False
        cmdToXML.Enabled = False
        cmdToCSV.Enabled = False
        cmdToSTR.Enabled = False
        
        updCSV.Enabled = False
        updID.Enabled = False
        updXML.Enabled = False
        updSTR.Enabled = False
        
    ElseIf cboTipoEvento.ItemData(cboTipoEvento.ListIndex) = enumNaturezaMensagem.MensagemConsulta Then
        cboTipoSaida.ListIndex = 0
        cboTipoSaida.Enabled = False
        
        cmdToID.Enabled = True
        cmdToXML.Enabled = False
        cmdToCSV.Enabled = False
        cmdToSTR.Enabled = False
        
        updCSV.Enabled = False
        updID.Enabled = True
        updXML.Enabled = False
        updSTR.Enabled = False
    
    Else
        cboTipoSaida.Enabled = True
        cmdToID.Enabled = True
    
        cmdToID.Enabled = True
        cmdToXML.Enabled = False
        cmdToCSV.Enabled = False
        cmdToSTR.Enabled = False
        
        updCSV.Enabled = False
        updID.Enabled = True
        updXML.Enabled = False
        updSTR.Enabled = False
    End If

End Sub

Private Sub cboTipoSaida_Click()

    If cboTipoSaida.ListIndex < 0 Then Exit Sub
    
    Select Case cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
        
        Case enumTipoSaidaMensagem.NaoseAplica
            
            cmdToID.Enabled = True
            cmdToSTR.Enabled = False
            cmdToCSV.Enabled = False
            cmdToXML.Enabled = False
            
            cmdToID.Visible = True
            cmdToSTR.Visible = True
            cmdToCSV.Visible = False
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = True
            updCSV.Visible = False
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = False
            updCSV.Enabled = False
            updXML.Enabled = False
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = False
            sprCSV.Enabled = False
            sprXML.Enabled = False
            
            lblDelimitador.Enabled = False
            cboDelimitador.Enabled = False
            cboDelimitador.ListIndex = -1
            
            
        Case enumTipoSaidaMensagem.SaidaXML
        
            cmdToID.Enabled = True
            cmdToSTR.Enabled = False
            cmdToCSV.Enabled = False
            cmdToXML.Enabled = True
            
            cmdToID.Visible = True
            cmdToSTR.Visible = True
            cmdToCSV.Visible = False
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = True
            updCSV.Visible = False
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = False
            updCSV.Enabled = False
            updXML.Enabled = True
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = False
            sprCSV.Enabled = False
            sprXML.Enabled = True
            
            lblDelimitador.Enabled = False
            cboDelimitador.Enabled = False
            cboDelimitador.ListIndex = -1
        
        Case enumTipoSaidaMensagem.SaidaString
            
            cmdToID.Enabled = True
            cmdToSTR.Enabled = True
            cmdToCSV.Enabled = False
            cmdToXML.Enabled = False
            
            cmdToID.Visible = True
            cmdToSTR.Visible = True
            cmdToCSV.Visible = False
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = True
            updCSV.Visible = False
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = True
            updCSV.Enabled = False
            updXML.Enabled = False
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = True
            sprCSV.Enabled = False
            sprXML.Enabled = False
            sprString.ZOrder 0
            
            lblDelimitador.Enabled = False
            cboDelimitador.Enabled = False
            cboDelimitador.ListIndex = -1
        
        Case enumTipoSaidaMensagem.SaidaCSV
            
            cmdToID.Enabled = True
            cmdToSTR.Enabled = False
            cmdToCSV.Enabled = True
            cmdToXML.Enabled = False
            
            cmdToID.Visible = True
            cmdToSTR.Visible = False
            cmdToCSV.Visible = True
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = False
            updCSV.Visible = True
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = False
            updCSV.Enabled = True
            updXML.Enabled = False
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = False
            sprCSV.Enabled = True
            sprXML.Enabled = False
            sprCSV.ZOrder 0
            
            lblDelimitador.Enabled = True
            cboDelimitador.Enabled = True
        
        Case enumTipoSaidaMensagem.SaidaStringXML
            
            cmdToID.Enabled = True
            cmdToSTR.Enabled = True
            cmdToCSV.Enabled = False
            cmdToXML.Enabled = True
            
            cmdToID.Visible = True
            cmdToSTR.Visible = True
            cmdToCSV.Visible = False
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = True
            updCSV.Visible = False
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = True
            updCSV.Enabled = False
            updXML.Enabled = True
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = True
            sprCSV.Enabled = False
            sprXML.Enabled = True
            sprString.ZOrder 0
            
            lblDelimitador.Enabled = False
            cboDelimitador.Enabled = False
            cboDelimitador.ListIndex = -1
        
        Case enumTipoSaidaMensagem.SaidaCSVXML
            
            cmdToID.Enabled = True
            cmdToSTR.Enabled = False
            cmdToCSV.Enabled = True
            cmdToXML.Enabled = True
            
            cmdToID.Visible = True
            cmdToSTR.Visible = False
            cmdToCSV.Visible = True
            cmdToXML.Visible = True
            
            updID.Visible = True
            updSTR.Visible = False
            updCSV.Visible = True
            updXML.Visible = True
            
            updID.Enabled = True
            updSTR.Enabled = False
            updCSV.Enabled = True
            updXML.Enabled = True
            
            sprID.MaxRows = 0
            sprXML.MaxRows = 0
            sprString.MaxRows = 0
            sprCSV.MaxRows = 0
            
            sprID.Enabled = True
            sprString.Enabled = False
            sprCSV.Enabled = True
            sprXML.Enabled = True
            sprCSV.ZOrder 0
            
            lblDelimitador.Enabled = True
            cboDelimitador.Enabled = True
    End Select

End Sub

Private Sub cmdToCSV_Click()

Dim objlvwItem                              As MSComctlLib.ListItem

On Error GoTo ErrorHandler
    
    If lstAtributo.SelectedItem Is Nothing Then Exit Sub
    
    For Each objlvwItem In lstAtributo.ListItems
        If objlvwItem.Selected Then
            If objlvwItem.ForeColor <> vbRed Then
                If Not flAtributoSelecionado(Mid(objlvwItem.Key, 2), sprCSV) Then
                    sprCSV.MaxRows = sprCSV.MaxRows + 1
                    sprCSV.SetText 1, sprCSV.MaxRows, objlvwItem.SubItems(1)
                    sprCSV.SetText 3, sprCSV.MaxRows, Mid(objlvwItem.Key, 2)
                End If
            Else
                frmMural.txtMural = "Atributo não vigente , não pode ser associado."
                frmMural.Show
            End If
        End If
    Next
    
    If sprCSV.MaxRows < 5 Then
        sprCSV.Row = sprCSV.MaxRows
    Else
        sprCSV.Row = sprCSV.MaxRows - 4
    End If
    
    sprCSV.Action = ActionGotoCell
    
    sprCSV.Row = sprCSV.MaxRows
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - cmdToCSV_Click")
End Sub

Private Sub cmdToID_Click()

Dim lngInicioAnterior                       As Long
Dim lngTamanhoAnterior                      As Long
Dim objlvwItem                              As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    If lstAtributo.SelectedItem Is Nothing Then Exit Sub
    
    For Each objlvwItem In lstAtributo.ListItems
        If objlvwItem.Selected Then
            If objlvwItem.ForeColor <> vbRed Then
                If Not flAtributoSelecionado(Mid(objlvwItem.Key, 2), sprID) Then
                    sprID.MaxRows = sprID.MaxRows + 1
                    sprID.Row = sprID.MaxRows
                    
                    sprID.Col = 1
                    sprID.Text = objlvwItem.SubItems(1) 'objlvwItem.Text
                    
                    sprID.Col = 3
                    If objlvwItem.SubItems(4) <> "0" Then
                        sprID.Text = objlvwItem.SubItems(3) & "," & objlvwItem.SubItems(4)
                    Else
                        sprID.Text = objlvwItem.SubItems(3)
                    End If
                    
                    sprID.Row = sprID.MaxRows - 1
                    
                    sprID.Col = 2
                    If IsNumeric(sprID.Text) Then
                        lngInicioAnterior = sprID.Text
                    Else
                        lngInicioAnterior = 1
                    End If
                    
                    sprID.Col = 3
                    If IsNumeric(sprID.Text) Then
                         If InStr(sprID.Text, ",") > 0 Then
                            lngTamanhoAnterior = Left(sprID.Text, InStr(sprID.Text, ",") - 1)
                        Else
                            lngTamanhoAnterior = sprID.Text
                        End If
                    Else
                         lngTamanhoAnterior = 0
                    End If
                    
                    sprID.SetText 2, sprID.MaxRows, lngTamanhoAnterior + lngInicioAnterior
                    sprID.SetText 4, sprID.MaxRows, 1
                    sprID.SetText 5, sprID.MaxRows, Mid(objlvwItem.Key, 2)
                End If
            Else
                frmMural.txtMural = "Atributo não vigente , não pode ser associado."
                frmMural.Show
            End If
        End If
    Next
    
    If sprID.MaxRows < 3 Then
        sprID.Row = sprID.MaxRows
    Else
        sprID.Row = sprID.MaxRows - 3
    End If
    
    sprID.Action = ActionGotoCell
    
    Exit Sub
ErrorHandler:
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - cmdToID_Click")

End Sub

Private Sub cmdToSTR_Click()

Dim lngInicioAnterior                        As Long
Dim lngTamanhoAnterior                       As Long
Dim objlvwItem                               As MSComctlLib.ListItem

On Error GoTo ErrorHandler
    
    If lstAtributo.SelectedItem Is Nothing Then Exit Sub
            
    For Each objlvwItem In lstAtributo.ListItems
        
        If objlvwItem.Selected Then
            If objlvwItem.ForeColor <> vbRed Then
                If Not flAtributoSelecionado(Mid(objlvwItem.Key, 2), sprString) Then
                    sprString.MaxRows = sprString.MaxRows + 1
                    sprString.Row = sprString.MaxRows
                    
                    sprString.Col = 1
                    sprString.Text = objlvwItem.SubItems(1) 'objlvwItem.Text
                    
                    sprString.Col = 3
                    If objlvwItem.SubItems(4) <> "0" Then
                        sprString.Text = objlvwItem.SubItems(3) & "," & objlvwItem.SubItems(4)
                    Else
                        sprString.Text = objlvwItem.SubItems(3)
                    End If
                    
                    sprString.Row = sprString.MaxRows - 1
                    
                    sprString.Col = 2
                    If IsNumeric(sprString.Text) Then
                        lngInicioAnterior = sprString.Text
                    Else
                        lngInicioAnterior = 1
                    End If
                    
                    sprString.Col = 3
                    If IsNumeric(sprString.Text) Then
                        If InStr(sprString.Text, ",") > 0 Then
                            lngTamanhoAnterior = Left(sprString.Text, InStr(sprString.Text, ",") - 1)
                        Else
                            lngTamanhoAnterior = sprString.Text
                        End If
                    Else
                         lngTamanhoAnterior = 0
                    End If
                    
                    sprString.SetText 2, sprString.MaxRows, lngTamanhoAnterior + lngInicioAnterior
                    sprString.SetText 5, sprString.MaxRows, Mid(objlvwItem.Key, 2)
                End If
            Else
                frmMural.txtMural = "Atributo não vigente , não pode ser associado."
                frmMural.Show
            End If
        End If
    Next
            
    If sprString.MaxRows < 5 Then
        sprString.Row = sprString.MaxRows
    Else
        sprString.Row = sprString.MaxRows - 4
    End If
    
    sprString.Action = ActionGotoCell
    
    sprString.Row = sprString.MaxRows
            
            
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - cmdToSTR_Click")
End Sub

Private Function flAtributoSelecionado(ByVal strAtributo As String, _
                                       poSpread As vaSpread) As Boolean
Dim lngRow                                   As Long

    If poSpread.Name = "sprXML" Or poSpread.Name = "sprCSV" Then
        poSpread.Col = 3
    Else
        poSpread.Col = 5
    End If
    
    For lngRow = 1 To poSpread.MaxRows
        poSpread.Row = lngRow
        
        If poSpread.Text = strAtributo Then
            flAtributoSelecionado = True
            Exit For
        End If
    Next

End Function

Private Sub cmdToXML_Click()

Dim objlvwItem                              As MSComctlLib.ListItem

On Error GoTo ErrorHandler
    
    If lstAtributo.SelectedItem Is Nothing Then Exit Sub
    
    For Each objlvwItem In lstAtributo.ListItems
        If objlvwItem.Selected Then
            If objlvwItem.ForeColor <> vbRed Then
                If Not flAtributoSelecionado(Mid(objlvwItem.Key, 2), sprXML) Then
                    sprXML.MaxRows = sprXML.MaxRows + 1
                    sprXML.SetText 1, sprXML.MaxRows, objlvwItem.SubItems(1) 'objlvwItem.Text
                    sprXML.SetText 3, sprXML.MaxRows, Mid(objlvwItem.Key, 2)
                End If
            Else
                frmMural.txtMural = "Atributo não vigente , não pode ser associado."
                frmMural.Show
            End If
        End If
    Next
    
    If sprXML.MaxRows < 5 Then
        sprXML.Row = sprXML.MaxRows
    Else
        sprXML.Row = sprXML.MaxRows - 4
    End If
    
    sprXML.Action = ActionGotoCell
    
    sprXML.Row = sprXML.MaxRows
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - cmdToXML_Click")
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
    flLimparCampos
    
    fgCursor True
    
    flCarregarCboTipoEvento
    flCarregarCboTipoSaida
    flCarregaComboDelimitador
    
    flCarregarComboTipoMensagemSaida
    flCarregarlistAtributo
    flCarregarTipoMensagem
    
    txtTipoMensagem.SetFocus
    
    fgCursor False
    
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    
    fgCursor False
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - Form_Load")

End Sub

Private Sub flLimparCampos()
        
    strOperacao = "Incluir"
        
    txtTipoMensagem.Text = vbNullString
    txtTipoMensagem.Enabled = True
    cboTipoMensagemSaida.Enabled = True
    txtDescricao.Text = vbNullString
    
    cboTipoEvento.ListIndex = -1
    cboTipoMensagemSaida.ListIndex = -1
    cboTipoSaida.ListIndex = -1
    
    cboDelimitador.ListIndex = -1
    txtPrioridade.Text = vbNullString
    
    cmdToID.Enabled = False
    cmdToSTR.Enabled = False
    cmdToCSV.Enabled = False
    cmdToXML.Enabled = False
           
    sprID.MaxRows = 0
    sprString.MaxRows = 0
    sprCSV.MaxRows = 0
    sprXML.MaxRows = 0
    
    sprID.Enabled = False
    sprString.Enabled = False
    sprCSV.Enabled = False
    sprXML.Enabled = False
        
    tlbCadastro.Buttons("Excluir").Enabled = False
    
    dtpDataInicioVigencia.Enabled = True
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataInicioVigencia.Value = dtpDataInicioVigencia.MinDate
    
    dtpDataFimVigencia.Enabled = True
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
    dtpDataFimVigencia.Value = Null
    
    lstTipoMensagem.Sorted = False

End Sub

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

Private Function flTipoDadoToSTR(plngTipoDado As Long) As String
    
    Select Case plngTipoDado
        Case enumTipoDadoAtributo.Alfanumerico
            flTipoDadoToSTR = "Alfanumérico"
        Case enumTipoDadoAtributo.Numerico
            flTipoDadoToSTR = "Numérico"
    End Select

End Function

Private Function flTipoDadoToEnum(pstrTipoDado As String) As Long

    Select Case pstrTipoDado
        Case "Alfanumérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Alfanumerico
        Case "Numérico"
            flTipoDadoToEnum = enumTipoDadoAtributo.Numerico
    End Select

End Function

Private Sub Form_Unload(Cancel As Integer)
    Set frmTipoMensagem = Nothing
End Sub

Private Sub lstAtributo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    fgClassificarListview lstAtributo, ColumnHeader.Index
End Sub

Private Sub lstTipoMensagem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    fgClassificarListview lstTipoMensagem, ColumnHeader.Index
End Sub

Private Sub lstTipoMensagem_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    Call flLimparCampos
    strOperacao = "Alterar"
    strKeyItemSelected = Item.Key
    Call flXmlToInterface
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

Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
            
    dtpDataInicioVigencia.Enabled = False
    
    With xmlTipoMensagem
        .selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = "Ler"
        .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text = lstTipoMensagem.SelectedItem.Text
        .selectSingleNode("//Grupo_TipoMensagem/CO_MESG_SAID").Text = CLng(Mid$(lstTipoMensagem.SelectedItem.Key, 4, 4))
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    xmlTipoMensagem.loadXML objMiu.Executar(xmlTipoMensagem.xml)
    Set objMiu = Nothing
       
    txtTipoMensagem.Enabled = False
    cboTipoMensagemSaida.Enabled = False
    
    With xmlTipoMensagem
        txtTipoMensagem.Text = .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text
        txtDescricao.Text = .selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text
        fgSearchItemCombo cboTipoEvento, .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text
        fgSearchItemCombo cboTipoSaida, .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG").Text
        fgSearchItemCombo cboTipoMensagemSaida, .selectSingleNode("//Grupo_TipoMensagem/CO_MESG_SAID").Text
        
        If Trim(.selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text) <> vbNullString Then
            fgSearchItemCombo cboDelimitador, 0, .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text
        Else
            cboDelimitador.ListIndex = -1
        End If
        
        txtPrioridade.Text = .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text
        
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
                
        For Each xmlNode In xmlTipoMensagem.documentElement.selectNodes("//Repeat_TipoMensagemAtributo/Grupo_TipoMensagemAtributo")
            Select Case CLng(xmlNode.selectSingleNode("TP_FORM_MESG").Text)
                Case enumTipoParteSaida.ParteId
                    sprID.Enabled = True
                    sprID.MaxRows = sprID.MaxRows + 1
                    sprID.Row = sprID.MaxRows
                    
                    sprID.SetText 1, sprID.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                    
                    If CLng(xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text) > 0 Then
                        sprID.SetText 3, sprID.MaxRows, xmlNode.selectSingleNode("QT_CTER_ATRB").Text & "," & xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
                    Else
                        sprID.SetText 3, sprID.MaxRows, xmlNode.selectSingleNode("QT_CTER_ATRB").Text
                    End If
                    
                    sprID.Col = 4
                    sprID.Value = xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                    
                    sprID.SetText 5, sprID.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                    
                Case enumTipoParteSaida.ParteSTR
                    sprString.Enabled = True
                    sprString.MaxRows = sprString.MaxRows + 1
                    sprString.Row = sprString.MaxRows
                    
                    sprString.SetText 1, sprString.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                    
                    If CLng(xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text) > 0 Then
                        sprString.SetText 3, sprString.MaxRows, xmlNode.selectSingleNode("QT_CTER_ATRB").Text & "," & xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
                    Else
                        sprString.SetText 3, sprString.MaxRows, xmlNode.selectSingleNode("QT_CTER_ATRB").Text
                    End If
                    
                    sprString.Col = 4
                    sprString.Value = xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                    
                    sprString.SetText 5, sprString.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                
                Case enumTipoParteSaida.ParteCSV
                    sprCSV.Enabled = True
                    sprCSV.MaxRows = sprCSV.MaxRows + 1
                    sprCSV.Row = sprCSV.MaxRows
                    
                    sprCSV.SetText 1, sprCSV.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                    
                    sprCSV.Col = 2
                    sprCSV.Value = xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                    
                    sprCSV.SetText 3, sprCSV.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                
                Case enumTipoParteSaida.ParteXML
                    sprXML.Enabled = True
                    sprXML.MaxRows = sprXML.MaxRows + 1
                    sprXML.Row = sprXML.MaxRows
                    
                    sprXML.SetText 1, sprXML.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                    
                    sprXML.Col = 2
                    sprXML.Value = xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                    
                    sprXML.SetText 3, sprXML.MaxRows, xmlNode.selectSingleNode("NO_ATRB_MESG").Text
            
            End Select
        Next
        
    End With
    
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
        
    flRecalcularPosicoes sprID
    flRecalcularPosicoes sprString
        
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, Me.Name, "flMoveObjetToInterface", 0
End Sub

Private Sub flMoveItem(ByRef poSpread As vaSpread, _
                       ByVal plngLinha As Long, _
                       ByVal plngTipoMovto As enumMoveUpDown)

Dim strValorAtual                            As String
Dim strNovoValor                             As String
Dim intCol                                   As Integer

On Error GoTo ErrorHandler
                       
    If plngTipoMovto = enumMoveUpDown.Up And plngLinha = 1 Then
        Exit Sub
    ElseIf plngTipoMovto = enumMoveUpDown.Down And plngLinha = poSpread.MaxRows Then
        Exit Sub
    ElseIf poSpread.MaxRows = 0 Then
        Exit Sub
    End If
                       
    If plngTipoMovto = enumMoveUpDown.Down Then
        plngLinha = plngLinha + 1
    End If
                     
    With poSpread
        For intCol = 1 To .MaxCols
            .Col = intCol
            .Row = plngLinha
            strValorAtual = .Value
            .Row = plngLinha - 1
            strNovoValor = .Value
            .Value = strValorAtual
            .Row = plngLinha
            .Value = strNovoValor
        Next
                      
        If plngTipoMovto = enumMoveUpDown.Down Then
            .Row = plngLinha
        Else
            .Row = plngLinha - 1
        End If
        .Action = ActionActiveCell
    
    End With
    
    Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flMoveItem", 0
End Sub

Private Sub sprCSV_KeyDown(KeyCode As Integer, Shift As Integer)

    If sprCSV.MaxRows = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        sprCSV.Row = sprCSV.ActiveRow
        sprCSV.Action = ActionDeleteRow
        sprCSV.MaxRows = sprCSV.MaxRows - 1
'        flRecalcularPosicoes sprCSV
        sprCSV.SetFocus
    End If

End Sub

Private Sub sprID_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If sprID.MaxRows = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        sprID.Row = sprID.ActiveRow
        sprID.Action = ActionDeleteRow
        sprID.MaxRows = sprID.MaxRows - 1
        flRecalcularPosicoes sprID
        sprID.SetFocus
    End If
    
End Sub

Private Sub sprString_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If sprString.MaxRows = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        sprString.Row = sprString.ActiveRow
        sprString.Action = ActionDeleteRow
        sprString.MaxRows = sprString.MaxRows - 1
        flRecalcularPosicoes sprString
        sprString.SetFocus
    End If
    
End Sub

Private Sub sprXML_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If sprXML.MaxRows = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        sprXML.Row = sprXML.ActiveRow
        sprXML.Action = ActionDeleteRow
        sprXML.MaxRows = sprXML.MaxRows - 1
        sprXML.SetFocus
    End If
    
End Sub

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

Private Sub txtPrioridade_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub flRecalcularPosicoes(poSpread As vaSpread)

Dim lngInicio                                As Long
Dim lngTamanho                               As Long
Dim lngRow                                   As Long

On Error GoTo ErrorHandler
   
    lngInicio = 1
    
    For lngRow = 1 To poSpread.MaxRows
        With poSpread
            .Row = lngRow
            .Col = 2
            .Text = lngInicio
            .Col = 3
            If InStr(.Text, ",") > 0 Then
                lngTamanho = CLng(Left(.Text, InStr(.Text, ",") - 1))
            Else
                lngTamanho = CLng(.Text)
            End If
            lngInicio = lngInicio + lngTamanho
        End With
    Next
    
    Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flRecalcularPosicoes", 0
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

Private Sub txtTipoMensagem_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub updCSV_DownClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprCSV, sprCSV.ActiveRow, enumMoveUpDown.Down
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updCSV_DownClick")
End Sub

Private Sub updCSV_UpClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprCSV, sprCSV.ActiveRow, enumMoveUpDown.Up
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updCSV_UPClick")
End Sub

Private Sub updID_DownClick()
On Error GoTo ErrorHandler

    flMoveItem sprID, sprID.ActiveRow, enumMoveUpDown.Down
    flRecalcularPosicoes sprID
            
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updID_DownClick")
End Sub

Private Sub updID_UpClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprID, sprID.ActiveRow, enumMoveUpDown.Up
    flRecalcularPosicoes sprID
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updID_UpClick")
End Sub

Private Sub updSTR_DownClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprString, sprString.ActiveRow, enumMoveUpDown.Down
    flRecalcularPosicoes sprString
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updSTR_DownClick")
End Sub

Private Sub updSTR_UpClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprString, sprString.ActiveRow, enumMoveUpDown.Up
    flRecalcularPosicoes sprString
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updSTR_UpClick")
End Sub

Private Sub updXML_DownClick()
On Error GoTo ErrorHandler
    
    flMoveItem sprXML, sprXML.ActiveRow, enumMoveUpDown.Down
    
    Exit Sub
ErrorHandler:
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updXML_DownClick")
End Sub

Private Sub updXML_UpClick()

On Error GoTo ErrorHandler

    flMoveItem sprXML, sprXML.ActiveRow, enumMoveUpDown.Up
    
    Exit Sub
ErrorHandler:
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmTipoMensagem - updXML_UpClick")
End Sub

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu                                  As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                                  As A7Miu.clsMIU
#End If

Dim strRetorno                              As String
Dim objListItem                             As ListItem
Dim strKey                                  As String

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
    Call objMiu.Executar(xmlTipoMensagem.documentElement.xml)
    Set objMiu = Nothing
        
    Call fgCursor(False)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
       
    If Not lstTipoMensagem.SelectedItem Is Nothing Then
        If strOperacao = "Incluir" Then
            strKey = "EVE" & _
                      Format(cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex), "0000") & _
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
    xmlTipoMensagem.loadXML objMiu.Executar(xmlTipoMensagem.xml)
    Set objMiu = Nothing
    
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    
    fgLockWindow 0
    
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "flSalvar", 0
    
End Sub

Private Function flInterfaceToXml() As String
    
Dim xmlAtributo                             As MSXML2.DOMDocument40
Dim lngRow                                  As Long
Dim xmlNode                                 As MSXML2.IXMLDOMNode
    
On Error GoTo ErrorHandler

    With xmlTipoMensagem
        .selectSingleNode("//Grupo_TipoMensagem/@Operacao").Text = strOperacao
        .selectSingleNode("//Grupo_TipoMensagem/TP_MESG").Text = txtTipoMensagem.Text
        .selectSingleNode("//Grupo_TipoMensagem/CO_MESG_SAID").Text = cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex)
        .selectSingleNode("//Grupo_TipoMensagem/NO_TIPO_MESG").Text = fgLimpaCaracterEspecial(txtDescricao.Text)
        .selectSingleNode("//Grupo_TipoMensagem/TP_NATZ_MESG").Text = cboTipoEvento.ItemData(cboTipoEvento.ListIndex)
        .selectSingleNode("//Grupo_TipoMensagem/TP_FORM_MESG").Text = cboTipoSaida.ItemData(cboTipoSaida.ListIndex)
        .selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text = cboDelimitador.Text
        .selectSingleNode("//Grupo_TipoMensagem/CO_PRIO_FILA_SAID_MESG").Text = fgLimpaCaracterEspecial(txtPrioridade.Text)
        
        If strOperacao <> "Incluir" Then
            .selectSingleNode("//Grupo_TipoMensagem/CO_TEXT_XML").Text = lstTipoMensagem.SelectedItem.Tag
        End If
        
        .selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").Text = ""
        
        .selectSingleNode("//Grupo_TipoMensagem/TX_VALID_SAID_MESG").appendChild fgCreateCDATASection(flMontarXSD())
        
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
    For lngRow = 1 To sprID.MaxRows
        
        Set xmlAtributo = CreateObject("MSXML2.DOMDocument.4.0")
        xmlAtributo.loadXML strEstruturaAtributo
        sprID.Row = lngRow
        
        xmlAtributo.selectSingleNode("//TP_MESG").Text = fgLimpaCaracterEspecial(txtTipoMensagem.Text)
        xmlAtributo.selectSingleNode("//CO_MESG_SAID").Text = cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex)
        
        sprID.Col = 5
        xmlAtributo.selectSingleNode("//NO_ATRB_MESG").Text = sprID.Text
        
        xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = enumTipoParteSaida.ParteId
        xmlAtributo.selectSingleNode("//NU_ORDE_AGRU_ATRB").Text = lngRow
        
        sprID.Col = 4
        xmlAtributo.selectSingleNode("//IN_OBRI_ATRB").Text = sprID.Value
        
        Call fgAppendXML(xmlTipoMensagem, "Repeat_TipoMensagemAtributo", xmlAtributo.xml)
        
    Next
    
    'Atributos para String
    For lngRow = 1 To sprString.MaxRows
        
        Set xmlAtributo = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlAtributo.loadXML strEstruturaAtributo
            
        sprString.Row = lngRow
        
        xmlAtributo.selectSingleNode("//TP_MESG").Text = fgLimpaCaracterEspecial(txtTipoMensagem.Text)
        xmlAtributo.selectSingleNode("//CO_MESG_SAID").Text = cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex)

        sprString.Col = 5
        xmlAtributo.selectSingleNode("//NO_ATRB_MESG").Text = sprString.Text
        
        If cmdToSTR.Caption = "str ->" Then
            xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = enumTipoParteSaida.ParteSTR
        Else
            xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = enumTipoParteSaida.ParteCSV
        End If
        
        xmlAtributo.selectSingleNode("//NU_ORDE_AGRU_ATRB").Text = lngRow
        
        sprString.Col = 4
        xmlAtributo.selectSingleNode("//IN_OBRI_ATRB").Text = sprString.Value
    
        Call fgAppendXML(xmlTipoMensagem, "Repeat_TipoMensagemAtributo", xmlAtributo.xml)
    Next
    
    'Atributos para CSV
    For lngRow = 1 To sprCSV.MaxRows
        
        Set xmlAtributo = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlAtributo.loadXML strEstruturaAtributo
    
        sprCSV.Row = lngRow
        
        xmlAtributo.selectSingleNode("//TP_MESG").Text = fgLimpaCaracterEspecial(txtTipoMensagem.Text)
        xmlAtributo.selectSingleNode("//CO_MESG_SAID").Text = cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex)
        sprCSV.Col = 3
        xmlAtributo.selectSingleNode("//NO_ATRB_MESG").Text = sprCSV.Text
        
        xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = enumTipoParteSaida.ParteCSV
        xmlAtributo.selectSingleNode("//NU_ORDE_AGRU_ATRB").Text = lngRow
        
        sprCSV.Col = 2
        xmlAtributo.selectSingleNode("//IN_OBRI_ATRB").Text = sprCSV.Value
    
        'xmlTipoMensagem.selectSingleNode("//Repeat_TipoMensagemAtributo").appendChild xmlAtributo.childNodes(0)
        Call fgAppendXML(xmlTipoMensagem, "Repeat_TipoMensagemAtributo", xmlAtributo.xml)
    Next
    
    'Atributos para XML
    For lngRow = 1 To sprXML.MaxRows
        
        Set xmlAtributo = CreateObject("MSXML2.DOMDocument.4.0")
        
        xmlAtributo.loadXML strEstruturaAtributo
    
        sprXML.Row = lngRow
        
        xmlAtributo.selectSingleNode("//TP_MESG").Text = fgLimpaCaracterEspecial(txtTipoMensagem.Text)
        xmlAtributo.selectSingleNode("//CO_MESG_SAID").Text = cboTipoMensagemSaida.ItemData(cboTipoMensagemSaida.ListIndex)
        sprXML.Col = 3
        xmlAtributo.selectSingleNode("//NO_ATRB_MESG").Text = sprXML.Text
        
        xmlAtributo.selectSingleNode("//TP_FORM_MESG").Text = enumTipoParteSaida.ParteXML
        xmlAtributo.selectSingleNode("//NU_ORDE_AGRU_ATRB").Text = lngRow
        
        sprXML.Col = 2
        xmlAtributo.selectSingleNode("//IN_OBRI_ATRB").Text = sprXML.Value
    
        'xmlTipoMensagem.selectSingleNode("//Repeat_TipoMensagemAtributo").appendChild xmlAtributo.childNodes(0)
        Call fgAppendXML(xmlTipoMensagem, "Repeat_TipoMensagemAtributo", xmlAtributo.xml)
    Next
    
    Set xmlAtributo = Nothing
    
    Exit Function
ErrorHandler:
    
    Set xmlAtributo = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "flInterfaceToXML", 0

End Function

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
    
    If cboTipoMensagemSaida.ListIndex < 0 Then
        flValidarCampos = "Informe o tipo da mensagem de saída."
        cboTipoMensagemSaida.SetFocus
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
    
    If cboTipoSaida.ItemData(cboTipoSaida.ListIndex) = enumTipoSaidaMensagem.SaidaCSV Then
        If Trim$(cboDelimitador.Text) = vbNullString Then
            flValidarCampos = "Selecione o caracter delimitador."
            cboDelimitador.SetFocus
            Exit Function
        End If
    End If
        
    If cboTipoEvento.ItemData(cboTipoEvento.ListIndex) <> enumNaturezaMensagem.MensagemECO Then
        If sprID.MaxRows = 0 Then
            flValidarCampos = "Selecione os Atributos do Identificador."
            Exit Function
        End If
    End If
    
   
    flValidarCampos = ""
    Exit Function

ErrorHandler:
    

    fgRaiseError App.EXEName, "frmTipoMensagem", "flValidarCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Sub flExcluir()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

On Error GoTo ErrorHandler
    
    If MsgBox("Confirma Exclusão ?", vbYesNo, "Tipos de Mensagem") = vbNo Then Exit Sub

    strOperacao = "Excluir"
        
    Call fgCursor(True)
    fgLockWindow Me.hwnd
    Call flInterfaceToXml

    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call objMiu.Executar(xmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem").xml)

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
    
    fgRaiseError App.EXEName, Me.Name, "flExcluir", 0
End Sub

Private Function flMontarXSD() As String

Dim xmlXSD                                  As MSXML2.DOMDocument40
Dim strNomeTag                              As String
Dim blnObrigatorio                          As Boolean
Dim strContexto                             As String
Dim lngRow                                  As Long

On Error GoTo ErrorHandler

    Set xmlXSD = CreateObject("MSXML2.DOMDocument.4.0")

    fgAppendNode xmlXSD, vbNullString, "xsd:schema", vbNullString
    fgAppendAttribute xmlXSD, "schema", "xmlns:xsd", "http://www.w3.org/2001/XMLSchema"

    fgAppendNode xmlXSD, "schema", "xsd:element", vbNullString
    fgAppendAttribute xmlXSD, "schema/element[position() = last()]", "name", "Saida"
    fgAppendAttribute xmlXSD, "schema/element[position() = last()]", "type", "TipoSaida"

    fgAppendNode xmlXSD, "schema", "xsd:complexType", vbNullString
    fgAppendAttribute xmlXSD, "schema/complexType[position() = last()]", "name", "TipoSaida"
    fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaida']", "xsd:sequence", vbNullString

    'Formata ID
    If sprID.MaxRows > 0 Then
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaida']/sequence", "xsd:element", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[@name='TipoSaida']/sequence/element", "name", "SaidaID"
        fgAppendAttribute xmlXSD, "element[@name='SaidaID']", "type", "TipoSaidaID"
        fgAppendAttribute xmlXSD, "element[@name='SaidaID']", "minOccurs", "1"
        fgAppendAttribute xmlXSD, "element[@name='SaidaID']", "maxOccurs", "1"

        fgAppendNode xmlXSD, "schema", "xsd:complexType", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[position() = last()]", "name", "TipoSaidaID"
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaidaID']", "xsd:sequence", vbNullString

        strContexto = "schema/complexType[@name='TipoSaidaID']/sequence"

        With sprID
            For lngRow = 1 To .MaxRows
                .Row = lngRow

                .Col = 5
                strNomeTag = .Text

                .Col = 4
                blnObrigatorio = CBool(.Value)

                fgAppendNode xmlXSD, strContexto, "xsd:element", vbNullString
                
                fgAppendAttribute xmlXSD, strContexto & "/element[position()=last()]", "name", strNomeTag
                fgAppendAttribute xmlXSD, strContexto & "/element[position()=last()]", "type", "TipoID" & strNomeTag
                fgAppendAttribute xmlXSD, strContexto & "/element[position()=last()]", "minOccurs", CStr(Abs(blnObrigatorio))
                fgAppendAttribute xmlXSD, strContexto & "/element[position()=last()]", "maxOccurs", "1"
                fgAppendNode xmlXSD, "schema", "xsd:simpleType", vbNullString
                fgAppendAttribute xmlXSD, "schema/simpleType[position()=last()]", "name", "TipoID" & strNomeTag
                fgAppendNode xmlXSD, "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']", "xsd:restriction", vbNullString

                If CLng(lstAtributo.ListItems("K" & strNomeTag).Tag) = enumTipoDadoAtributo.Alfanumerico Then
                    'Dado Alfanumerico

                    fgAppendAttribute xmlXSD, "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction", "base", "xsd:string"
                    
                    fgAppendNode xmlXSD, "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction", "xsd:length", vbNullString
                    
                    fgAppendAttribute xmlXSD, "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction/length", "value", lstAtributo.ListItems("K" & strNomeTag).SubItems(3)

                Else
                    'Dado Numérico

                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction", _
                                      "base", _
                                      "xsd:string"

                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction", _
                                 "xsd:pattern", _
                                 vbNullString
                    
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoID" & strNomeTag & "']/restriction/pattern", _
                                      "value", _
                                      "[0-9]{" & CStr(Abs(blnObrigatorio)) & "," & lstAtributo.ListItems("K" & strNomeTag).SubItems(3) & "}"

                End If


            Next
        End With
    End If
    
    'Formata String
    If sprString.MaxRows > 0 Then
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaida']/sequence", "xsd:element", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[@name='TipoSaida']/sequence/element[position() = last()]", "name", "SaidaSTR"
        fgAppendAttribute xmlXSD, "element[@name='SaidaSTR']", "type", "TipoSaidaSTR"
        fgAppendAttribute xmlXSD, "element[@name='SaidaSTR']", "minOccurs", "1"
        fgAppendAttribute xmlXSD, "element[@name='SaidaSTR']", "maxOccurs", "1"

        fgAppendNode xmlXSD, "schema", "xsd:complexType", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[position() = last()]", "name", "TipoSaidaSTR"
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaidaSTR']", "xsd:sequence", vbNullString

        strContexto = "schema/complexType[@name='TipoSaidaSTR']/sequence"

        With sprString
            For lngRow = 1 To .MaxRows
                .Row = lngRow

                .Col = 5
                strNomeTag = Replace(.Text, "K", "")

                .Col = 4
                blnObrigatorio = CBool(.Value)

                fgAppendNode xmlXSD, strContexto, "xsd:element", vbNullString
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "name", _
                                  strNomeTag
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "type", _
                                  "TipoSTR" & strNomeTag
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "minOccurs", _
                                  CStr(Abs(blnObrigatorio))
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "maxOccurs", _
                                  "1"

                fgAppendNode xmlXSD, "schema", "xsd:simpleType", vbNullString
                fgAppendAttribute xmlXSD, _
                                  "schema/simpleType[position()=last()]", _
                                  "name", _
                                  "TipoSTR" & strNomeTag

                fgAppendNode xmlXSD, _
                             "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']", _
                             "xsd:restriction", _
                             vbNullString

                If CLng(lstAtributo.ListItems("K" & strNomeTag).Tag) = enumTipoDadoAtributo.Alfanumerico Then
                    'Dado Alfanumerico

                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction", _
                                      "base", _
                                      "xsd:string"
                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction", _
                                 "xsd:length", _
                                 vbNullString
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction/length", _
                                      "value", _
                                      lstAtributo.ListItems("K" & strNomeTag).SubItems(3)

                Else
                    'Dado Numérico

                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction", _
                                      "base", _
                                      "xsd:string"

                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction", _
                                 "xsd:pattern", _
                                 vbNullString
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoSTR" & strNomeTag & "']/restriction/pattern", _
                                      "value", _
                                      "[0-9]{" & lstAtributo.ListItems("K" & strNomeTag).SubItems(3) & "}"

                End If

            Next
        End With

    End If

    'Formata XML
    If sprXML.MaxRows > 0 Then
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaida']/sequence", "xsd:element", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[@name='TipoSaida']/sequence/element[position() = last()]", "name", "SaidaXML"
        fgAppendAttribute xmlXSD, "element[@name='SaidaXML']", "type", "TipoSaidaXML"
        fgAppendAttribute xmlXSD, "element[@name='SaidaXML']", "minOccurs", "1"
        fgAppendAttribute xmlXSD, "element[@name='SaidaXML']", "maxOccurs", "1"

        fgAppendNode xmlXSD, "schema", "xsd:complexType", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[position() = last()]", "name", "TipoSaidaXML"
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaidaXML']", "xsd:sequence", vbNullString

        strContexto = "schema/complexType[@name='TipoSaidaXML']/sequence"

        With sprXML
            For lngRow = 1 To .MaxRows
                .Row = lngRow
                
                .Col = 3
                strNomeTag = .Text
                
                If Mid(strNomeTag, 1, 7) <> "/Grupo_" And Mid(strNomeTag, 1, 6) <> "Grupo_" Then
                
                    .Col = 2
                    blnObrigatorio = CBool(.Value)
    
                    fgAppendNode xmlXSD, strContexto, "xsd:element", vbNullString
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "name", _
                                      strNomeTag
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "type", _
                                      "TipoXML" & strNomeTag
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "minOccurs", _
                                      CStr(Abs(blnObrigatorio))
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "maxOccurs", _
                                      "1"
    
                    fgAppendNode xmlXSD, "schema", "xsd:simpleType", vbNullString
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[position()=last()]", _
                                      "name", _
                                      "TipoXML" & strNomeTag
    
                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']", _
                                 "xsd:restriction", _
                                 vbNullString
    
                    If CLng(lstAtributo.ListItems("K" & strNomeTag).Tag) = enumTipoDadoAtributo.Alfanumerico Then
                        'Dado Alfanumerico
    
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction", _
                                          "base", _
                                          "xsd:string"
                        
                        fgAppendNode xmlXSD, _
                                     "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction", _
                                     "xsd:minLength", _
                                     vbNullString
                        
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction/minLength", _
                                          "value", _
                                          CStr(Abs(blnObrigatorio))
                                          
                        fgAppendNode xmlXSD, _
                                     "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction", _
                                     "xsd:maxLength", _
                                     vbNullString
                        
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction/maxLength", _
                                          "value", _
                                          lstAtributo.ListItems("K" & strNomeTag).SubItems(3)
    
                    Else
                        'Dado Numérico
    
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction", _
                                          "base", _
                                          "xsd:string"
    
                        fgAppendNode xmlXSD, _
                                     "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction", _
                                     "xsd:pattern", _
                                     vbNullString
    
                        If CLng(lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) = 0 Then
                            'Não tem decimais
                            fgAppendAttribute xmlXSD, _
                                              "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction/pattern", _
                                              "value", _
                                              "[0-9]{" & CStr(Abs(blnObrigatorio)) & "," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "}"
                        Else
                            fgAppendAttribute xmlXSD, _
                                              "schema/simpleType[@name='" & "TipoXML" & strNomeTag & "']/restriction/pattern", _
                                              "value", _
                                              "[0-9]{1," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "},[0-9]{1," & lstAtributo.ListItems("K" & strNomeTag).SubItems(4) & "}" & _
                                              "|[0-9]{" & CStr(Abs(blnObrigatorio)) & "," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "}"
                        End If
                        
                    End If
                End If
            Next
        End With
    End If

    'Formata CSV
    If sprCSV.MaxRows > 0 Then
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaida']/sequence", "xsd:element", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[@name='TipoSaida']/sequence/element[position() = last()]", "name", "SaidaCSV"
        fgAppendAttribute xmlXSD, "element[@name='SaidaCSV']", "type", "TipoSaidaCSV"
        fgAppendAttribute xmlXSD, "element[@name='SaidaCSV']", "minOccurs", "1"
        fgAppendAttribute xmlXSD, "element[@name='SaidaCSV']", "maxOccurs", "1"

        fgAppendNode xmlXSD, "schema", "xsd:complexType", vbNullString
        fgAppendAttribute xmlXSD, "schema/complexType[position() = last()]", "name", "TipoSaidaCSV"
        fgAppendNode xmlXSD, "schema/complexType[@name='TipoSaidaCSV']", "xsd:sequence", vbNullString

        strContexto = "schema/complexType[@name='TipoSaidaCSV']/sequence"

        'Coloca a definição do Separador
        fgAppendNode xmlXSD, "schema", "xsd:simpleType", vbNullString
        fgAppendAttribute xmlXSD, _
                          "schema/simpleType[position()=last()]", _
                          "name", _
                          "TipoCSVSeparador"

        fgAppendNode xmlXSD, _
                     "schema/simpleType[@name='TipoCSVSeparador']", _
                     "xsd:restriction", _
                     vbNullString
        
        fgAppendAttribute xmlXSD, _
                           "schema/simpleType[@name='TipoCSVSeparador']/restriction", _
                           "base", _
                           "xsd:string"
        
        fgAppendNode xmlXSD, _
                     "schema/simpleType[@name='TipoCSVSeparador']/restriction", _
                     "xsd:pattern", _
                     vbNullString
        
        fgAppendAttribute xmlXSD, _
                          "schema/simpleType[@name='TipoCSVSeparador']/restriction/pattern", _
                          "value", _
                          cboDelimitador.Text
                          
        With sprCSV
            For lngRow = 1 To .MaxRows
                .Row = lngRow

                .Col = 3
                strNomeTag = .Text

                .Col = 2
                blnObrigatorio = CBool(.Value)

                fgAppendNode xmlXSD, strContexto, "xsd:element", vbNullString
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "name", _
                                  strNomeTag
                
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "type", _
                                  "TipoCSV" & strNomeTag
                
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "minOccurs", _
                                  CStr(Abs(blnObrigatorio))
                
                fgAppendAttribute xmlXSD, _
                                  strContexto & "/element[position()=last()]", _
                                  "maxOccurs", _
                                  "1"

                If lngRow < .MaxRows Then
                    'Incluir Delimitador (Menos para ultima posição)
                    fgAppendNode xmlXSD, strContexto, "xsd:element", vbNullString
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "name", _
                                      "Separador"
                    
                    fgAppendAttribute xmlXSD, _
                                      strContexto & "/element[position()=last()]", _
                                      "type", _
                                      "TipoCSVSeparador"
                End If
                
                'Inclui Tipo Do dado
                fgAppendNode xmlXSD, "schema", "xsd:simpleType", vbNullString
                fgAppendAttribute xmlXSD, _
                                  "schema/simpleType[position()=last()]", _
                                  "name", _
                                  "TipoCSV" & strNomeTag

                fgAppendNode xmlXSD, _
                             "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']", _
                             "xsd:restriction", _
                             vbNullString

                If CLng(lstAtributo.ListItems("K" & strNomeTag).Tag) = enumTipoDadoAtributo.Alfanumerico Then
                    'Dado Alfanumerico
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction", _
                                      "base", _
                                      "xsd:string"
                    
                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction", _
                                 "xsd:minLength", _
                                 vbNullString
                    
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction/minLength", _
                                      "value", _
                                      CStr(Abs(blnObrigatorio))
                                      
                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction", _
                                 "xsd:maxLength", _
                                 vbNullString
                    
                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction/maxLength", _
                                      "value", _
                                      lstAtributo.ListItems("K" & strNomeTag).SubItems(3)

                Else
                    'Dado Numérico

                    fgAppendAttribute xmlXSD, _
                                      "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction", _
                                      "base", _
                                      "xsd:string"

                    fgAppendNode xmlXSD, _
                                 "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction", _
                                 "xsd:pattern", _
                                 vbNullString
                    
                    If CLng(lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) = 0 Then
                        'Não tem decimais
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction/pattern", _
                                          "value", _
                                          "[0-9]{" & CStr(Abs(blnObrigatorio)) & "," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "}"
                    Else
                        fgAppendAttribute xmlXSD, _
                                          "schema/simpleType[@name='" & "TipoCSV" & strNomeTag & "']/restriction/pattern", _
                                          "value", _
                                          "[0-9]{1," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "},[0-9]{1," & lstAtributo.ListItems("K" & strNomeTag).SubItems(4) & "}" & _
                                          "|[0-9]{" & CStr(Abs(blnObrigatorio)) & "," & (lstAtributo.ListItems("K" & strNomeTag).SubItems(3) - lstAtributo.ListItems("K" & strNomeTag).SubItems(4)) & "}"
                    End If

                End If
            Next
        End With
    End If

    flMontarXSD = xmlXSD.xml
    
    Set xmlXSD = Nothing

    Exit Function
ErrorHandler:
    
    Set xmlXSD = Nothing
    
    fgRaiseError App.EXEName, Me.Name, "flMontaXSD Function", 0
    
End Function

Private Sub flProtegerChave()
    
   strOperacao = "Alterar"
   txtTipoMensagem.Enabled = False
   cboTipoMensagemSaida.Enabled = False
   dtpDataInicioVigencia.Enabled = False
   tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True

End Sub

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


Private Sub flCarregaComboDelimitador()
    cboDelimitador.AddItem ";"
    cboDelimitador.AddItem "|"
End Sub
