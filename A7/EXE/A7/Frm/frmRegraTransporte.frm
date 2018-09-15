VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form frmRegraTransporte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Regras de Transporte"
   ClientHeight    =   9465
   ClientLeft      =   75
   ClientTop       =   1410
   ClientWidth     =   15015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   15015
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   3600
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":0112
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":09EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":12C6
            Key             =   "Empresa"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":1BA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":247A
            Key             =   "Sistema"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":2D54
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":362E
            Key             =   "Sistema1"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":3F08
            Key             =   "SistemaDestino"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":4222
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":453C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":4856
            Key             =   "Regra"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":4B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":4E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":51A4
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":54BE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":5810
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":5922
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":5C3C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":5F56
            Key             =   "Evento"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRegraTransporte.frx":6270
            Key             =   "Limpar"
         EndProperty
      EndProperty
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
      Height          =   7020
      Left            =   3375
      TabIndex        =   13
      Top             =   2415
      Width           =   11625
      Begin VB.CheckBox chkIndicadorRegraAtiva 
         Caption         =   "Ativa"
         Height          =   255
         Left            =   7080
         TabIndex        =   30
         Top             =   1400
         Value           =   1  'Checked
         Width           =   800
      End
      Begin VB.TextBox txtTipoFormatoMensagemSaida 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4110
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   345
         Width           =   2835
      End
      Begin VB.TextBox txtFilaDestino 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Top             =   1410
         Width           =   6855
      End
      Begin VB.ComboBox cboDelimitador 
         Height          =   315
         Left            =   7095
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   870
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkIndicadorTraducao 
         Caption         =   "Traduzir"
         Height          =   255
         Left            =   7080
         TabIndex        =   4
         Top             =   360
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.TextBox txtNatureza 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   870
         Width           =   4740
      End
      Begin MSComctlLib.TreeView treSistemaDestino 
         Height          =   4725
         Left            =   8535
         TabIndex        =   10
         Top             =   1770
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   8334
         _Version        =   393217
         Indentation     =   353
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imlIcons"
         Appearance      =   1
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
         Height          =   1170
         Left            =   8265
         TabIndex        =   18
         Top             =   300
         Width           =   3270
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   330
            Left            =   660
            TabIndex        =   8
            Top             =   330
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            Format          =   52625409
            CurrentDate     =   37816
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   330
            Left            =   660
            TabIndex        =   9
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   52625409
            CurrentDate     =   37816
         End
         Begin MSComCtl2.DTPicker dtpHoraInicioVigencia 
            Height          =   330
            Left            =   2160
            TabIndex        =   31
            Top             =   330
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   582
            _Version        =   393216
            Format          =   52625410
            CurrentDate     =   37816
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
            Left            =   285
            TabIndex        =   20
            Top             =   780
            Width           =   300
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
            Left            =   90
            TabIndex        =   19
            Top             =   390
            Width           =   510
         End
      End
      Begin VB.TextBox txtTipoMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   345
         Width           =   3915
      End
      Begin VB.ComboBox cboTipoEntrada 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   870
         Width           =   2040
      End
      Begin FPSpread.vaSpread sprRegra 
         Height          =   4710
         Left            =   90
         TabIndex        =   11
         Top             =   1785
         Width           =   8415
         _Version        =   196608
         _ExtentX        =   14843
         _ExtentY        =   8308
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ArrowsExitEditMode=   -1  'True
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
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
         MaxCols         =   7
         MaxRows         =   1
         ProcessTab      =   -1  'True
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmRegraTransporte.frx":658A
         UserResize      =   0
      End
      Begin MSComctlLib.Toolbar tlbCadastro 
         Height          =   330
         Left            =   8025
         TabIndex        =   24
         Top             =   6600
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "R - Indicador Tag Repetição"
         Height          =   195
         Left            =   2925
         TabIndex        =   29
         Top             =   6680
         Width           =   2025
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Formato de Saída"
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
         Left            =   4110
         TabIndex        =   28
         Top             =   150
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Fila Destino"
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
         TabIndex        =   25
         Top             =   1215
         Width           =   1020
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Campos Não Obrigatórios"
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   6675
         Width           =   1800
      End
      Begin VB.Label Label5 
         BackColor       =   &H00D7B290&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   6645
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sistema Destino:"
         Height          =   195
         Left            =   8565
         TabIndex        =   21
         Top             =   1545
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Mensagem"
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
         TabIndex        =   14
         Top             =   150
         Width           =   1620
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
         Top             =   660
         Width           =   2010
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Entrada"
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
         Left            =   4905
         TabIndex        =   16
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lblDelimitador 
         AutoSize        =   -1  'True
         Caption         =   "Delimitador"
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
         Left            =   7080
         TabIndex        =   17
         Top             =   660
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Regras Cadastradas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3375
      TabIndex        =   12
      Top             =   0
      Width           =   11625
      Begin MSComctlLib.ListView lstRegra 
         Height          =   2130
         Left            =   60
         TabIndex        =   2
         Top             =   210
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   3757
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Tipo de Mensagem"
            Object.Width           =   5503
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Formato Saída"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Natureza da Mensagem"
            Object.Width           =   2619
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Tag             =   "1"
            Text            =   "Tipo de Entrada"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Data Início"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Data Fim"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin MSComctlLib.TreeView treSistema 
      Height          =   9135
      Left            =   45
      TabIndex        =   1
      Top             =   270
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   16113
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Selecione o Sistema de Origem:"
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
      TabIndex        =   0
      Top             =   15
      Width           =   2715
   End
End
Attribute VB_Name = "frmRegraTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pelo cadastramento e manutenção das regras de transporte do sistema A7.

Option Explicit

Private xmlMapaRegra                        As MSXML2.DOMDocument40
Private xmlRegraTransporte                  As MSXML2.DOMDocument40
Private xmlValidaTag                        As MSXML2.DOMDocument40

Private blnTagInvalida                      As Boolean

Private strOperacao                         As String
Private strKeyItemSelected                  As String
Private Const strFuncionalidade             As String = "frmRegraTransporte"

Private strCodigoMensagem                   As String
Private lngTipoFormatoMensagemSaida         As Long
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private WithEvents objInclusaoRegra         As frmInclusaoRegra
Attribute objInclusaoRegra.VB_VarHelpID = -1
Private strDelimitadorCSV                   As String

Private strIndicadorConsulta                As String

Private Sub cboTipoEntrada_Click()
    
    On Error GoTo ErrorHandler

    fgCursor True
    
    If cboTipoEntrada.ListIndex = -1 Then Exit Sub
    
    tlbCadastro.Buttons("Excluir").Enabled = False
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
          
    If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaCSV Then
        cboDelimitador.Enabled = True
        lblDelimitador.Enabled = True
    Else
        cboDelimitador.Enabled = False
        cboDelimitador.ListIndex = -1
        lblDelimitador.Enabled = False
    End If
    
    If strOperacao = "Incluir" Then
        flFormatarSpread strCodigoMensagem, lngTipoFormatoMensagemSaida
    Else
        flFormatarSpread Mid$(lstRegra.SelectedItem.Key, 2, 9), _
                         CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4))
    End If
   
    fgCursor False

    Exit Sub

ErrorHandler:
   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboTipoEntrada_Click"
    
End Sub

'Formatar as colunas do spread de configuração da regra de transporte.
Private Sub flFormatarSpread(pstrCodigoMensagem As String, _
                             plngTipoFormatoMensagemSaida As Long)

Dim xmlTipoMensagem     As MSXML2.DOMDocument40
Dim xmlLayout           As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim lngTamanho          As Long

    On Error GoTo ErrorHandler
        
    If cboTipoEntrada.ListIndex = -1 Or Not CBool(chkIndicadorTraducao.Value) Then Exit Sub
    
    fgLockWindow Me.hwnd
    
    Select Case cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex)
        Case enumTipoEntradaMensagem.EntradaString
            With sprRegra
                .MaxCols = 14
                .Row = 0
                
                .Col = 1
                .Text = "Tp."
                .ColWidth(1) = 3
                
                .Col = 2
                .Text = "Pos. Inicial"
                .ColWidth(2) = 8
                
                .Col = 3
                .Text = "Tamanho"
                .ColWidth(3) = 8
                
                .Col = 4
                .Text = "Campo de Saída"
                .ColWidth(4) = 33
                
                .ColWidth(5) = 0
                .ColWidth(6) = 0
                .ColWidth(7) = 0
                .ColWidth(8) = 0
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0

                .Col = 13
                .Text = "Default"
                .ColWidth(13) = 10
                
                .Col = 14
                .Text = "Obrig."
                .ColWidth(14) = 5
                
            End With
        Case Else
            With sprRegra
                .MaxCols = 14
                .Row = 0
                .Col = 1
                .Text = "Tp."
                .ColWidth(1) = 3
                
                .Col = 2
                .ColWidth(2) = 21
                If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaCSV Then
                    .Text = "Índice da Entrada"
                    'Não Utilizar para CSV e XML
                    .ColWidth(3) = 0
                Else
                    .Text = "Tag de Entrada"
                    'Utilizar para identificar se a tag de origem pertence a uma repeticao
                    .ColWidth(3) = 0
                    .Col = 3
                    .Text = "R"
                End If
                
                .Col = 4
                .Text = "Campo de Saída"
                .ColWidth(4) = 32
                
                .ColWidth(5) = 0
                .ColWidth(6) = 0
                .ColWidth(7) = 0
                .ColWidth(8) = 0
                .ColWidth(9) = 0
                .ColWidth(10) = 0
                .ColWidth(11) = 0
                .ColWidth(12) = 0
            
                .Col = 13
                .Text = "Default"
                .ColWidth(13) = 6
                
                .Col = 14
                .Text = "Obrig."
                .ColWidth(14) = 5
                
            End With
    End Select
    
    Set xmlTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlTipoMensagem, "", "Repeat_ParametrosLeitura", "")
    Call fgAppendNode(xmlTipoMensagem, "Repeat_ParametrosLeitura", "TP_MESG", pstrCodigoMensagem)
    Call fgAppendNode(xmlTipoMensagem, "Repeat_ParametrosLeitura", "TP_FORM_MESG_SAID", plngTipoFormatoMensagemSaida)
    
    Call xmlTipoMensagem.loadXML(fgMIUExecutarGenerico("Ler", "A7Server.clsTipoMensagem", xmlTipoMensagem))
    
    Set xmlLayout = CreateObject("MSXML2.DOMDocument.4.0")
    xmlLayout.loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
        
    flMontarMensagem xmlTipoMensagem.documentElement.selectNodes("//Repeat_TipoMensagemAtributo/*"), _
                     xmlLayout

    sprRegra.MaxRows = 0
    
    flSubFormataSpread xmlLayout.childNodes(0)
    
    strDelimitadorCSV = xmlTipoMensagem.selectSingleNode("//Grupo_TipoMensagem/TP_CTER_DELI").Text
        
    Set xmlTipoMensagem = Nothing
    Set xmlLayout = Nothing
    
    fgLockWindow 0
    
    Exit Sub

ErrorHandler:
    fgLockWindow 0
    
    Set xmlTipoMensagem = Nothing
    Set xmlLayout = Nothing
    
    mdiBUS.uctLogErros.MostrarErros Err, Me.Name & " - flFormatarSpread"

End Sub

'Sub-função utilizada pela flFormatarSpread que formata o spread de acordo com o xml com o layout da mensagem.
Private Sub flSubFormataSpread(ByVal pxmlNodeBase As IXMLDOMNode)

Dim xmlNode                                 As IXMLDOMNode
Dim lngNivel                                As Long
Dim lngTamanho                              As Long
Dim lngRepet                                As Long
Dim lngRepetMax                             As Long

    On Error GoTo ErrorHandler

    For Each xmlNode In pxmlNodeBase.selectNodes("./*")
    
        lngNivel = CLng(xmlNode.selectSingleNode("@NU_NIVE_MESG_ATRB").Text)

        With sprRegra

            .MaxRows = .MaxRows + 1
            .Row = .MaxRows
                    
            .Col = 1
            .Text = flTipoParteSaidaToSTR(CLng(xmlNode.selectSingleNode("@TP_FORM_MESG").Text))
                    
            If CBool(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text) Then
                .BackColor = RGB(255, 255, 255)
                .Col = 1
                .BackColor = RGB(255, 255, 255)
            Else
                .BackColor = RGB(144, 178, 215)
                .Col = 1
                .BackColor = RGB(144, 178, 215)
            End If
            
            .Col = 5
            '.Text = xmlNode.selectSingleNode("NO_ATRB_MESG").Text
            .Text = xmlNode.nodeName
            .Col = 6
            .Text = xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text
            .Col = 7
            .Text = xmlNode.selectSingleNode("@QT_CTER_ATRB").Text
            .Col = 8
            .Text = xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text
            .Col = 9
            .Text = xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text
            .Col = 10
            .Text = xmlNode.selectSingleNode("@TP_FORM_MESG").Text
            .Col = 11
            .Text = xmlNode.selectSingleNode("@NU_NIVE_MESG_ATRB").Text
            .Col = 12
            .Text = xmlNode.selectSingleNode("@QT_REPE").Text
                
            .Col = 2
            Select Case cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex)
                Case enumTipoEntradaMensagem.EntradaXML
                    'String - TagTArget
                    .CellType = CellTypeEdit
                Case Else
                    'numerico - Indice(CSV) ou Inicio(STR)
                    .CellType = CellTypeInteger
                    .TypeIntegerMin = 1
                    .TypeIntegerMax = 9999
                    '.Col = 3
                    '.CellType = CellTypeInteger
                    '.TypeIntegerMin = 1
                    '.TypeIntegerMax = 9999
            End Select
            
            .Col = 14
            .CellType = CellTypeCheckBox
            .TypeCheckCenter = True
            
            If Not xmlNode.selectSingleNode("./*") Is Nothing Then
                'Pai - permitir configuração nesta tag!
                'Alteração feita para atender os casos de repetição dentro de repetição
                'A tag pai é necessária para fechar a procura do tradutor por contexto
                
                'Definir tipo como Pai
                .Col = 6
                .Text = 9 'Tipo Pai (grupo ou repeticao)
                
                .Col = 2
                .CellType = CellTypeEdit
                
                .Col = 3
                If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaString Then
                    'Permitir o Tamanho para o Grupo (blocagem -->Deslocamento)
                    .CellType = CellTypeInteger
                    .TypeIntegerMin = 1
                    .TypeIntegerMax = 9999
                Else
                    .CellType = CellTypeStaticText
                End If
                
                '.BackColor = RGB(215, 178, 144)
                .Col = 4
                If CLng(xmlNode.selectSingleNode("@QT_REPE").Text) > 0 Then
                    .Text = String$((lngNivel - 1) * 5, " ") & xmlNode.nodeName & _
                            " (" & xmlNode.selectSingleNode("@QT_REPE").Text & " Repetições)"
                Else
                    .Text = String$((lngNivel - 1) * 5, " ") & xmlNode.nodeName
                End If
                .CellType = CellTypeStaticText
                .BackColor = flCorNivel(lngNivel + 1)
                .FontBold = True
            Else
                
                .Col = 2
                If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaXML Then
                    'Quando tipo XML --> tipo string - inclusao do nome da tag
                    .CellType = CellTypeEdit
                Else
                    .CellType = CellTypeInteger
                    .TypeIntegerMin = 1
                    .TypeIntegerMax = 9999
                End If
                
                .Col = 3
                If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaXML Then
                    If CLng(xmlNode.selectSingleNode("./@NU_NIVE_MESG_ATRB").Text) > 1 Then
                        .CellType = CellTypeCheckBox
                        .Value = 1
                        .ColWidth(3) = 2
                    Else
                        .CellType = CellTypeStaticText
                    End If
                Else
                    .CellType = CellTypeInteger
                    .TypeIntegerMin = 1
                    .TypeIntegerMax = 9999
                End If
                
                .Col = 4
                '.Text = xmlNode.selectSingleNode("NO_ATRB_MESG").Text
                .Text = String$((lngNivel - 1) * 5, " ") & xmlNode.nodeName & _
                        " [" & flTipoDadoToSTR(CLng(xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text)) & _
                        "(" & xmlNode.selectSingleNode("@QT_CTER_ATRB").Text & _
                        IIf(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text <> "0", "," & xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text, vbNullString) & _
                        ")]"
                 .CellType = CellTypeStaticText
                 .BackColor = flCorNivel(lngNivel)
            End If
                   
            If Not xmlNode.selectSingleNode("./*") Is Nothing Then
                'Tag com filho
                flSubFormataSpread xmlNode
            End If
        End With
    Next
    
    Exit Sub
ErrorHandler:
    
    Call fgRaiseError(App.EXEName, Me.Name, "flSubFormataSpread", lngCodigoErroNegocio)

End Sub

'Definir a cor do nível da mensagem no spread de acordo com a sua identação no xml.
Private Function flCorNivel(ByVal piNivel As Integer) As Long
    
    Select Case piNivel
        Case 1
            flCorNivel = RGB(255, 255, 255)
            'flCorNivel = "#FFFFFF"
        Case 2
            flCorNivel = RGB(238, 238, 238)
            'flCorNivel = "#EEEEEE"
        Case 3
            flCorNivel = RGB(206, 218, 234)
            'flCorNivel = "#CEDAEA"
        Case 4
            flCorNivel = RGB(230, 239, 216)
            'flCorNivel = "#E6EFD8"
        Case Else
            flCorNivel = RGB(243, 236, 212)
            'flCorNivel = "#F3ECD4"
    End Select

End Function

'Converter as literais de tipo de parte de saída para o domínio numérico.
Private Function flTipoParteSaidaToSTR(plTipoParteSaida As Long)
    
    Select Case plTipoParteSaida
        Case enumTipoParteSaida.ParteId
            flTipoParteSaidaToSTR = "Id"
        Case enumTipoParteSaida.ParteSTR
            flTipoParteSaidaToSTR = "Str"
        Case enumTipoParteSaida.ParteXML
            flTipoParteSaidaToSTR = "Xml"
        Case enumTipoParteSaida.ParteCSV
            flTipoParteSaidaToSTR = "Csv"
    End Select
    
End Function

'Converter o domínio numérico de tipo de dado para literais do sistema.
Private Function flTipoDadoToSTR(plTipoDado As Long) As String
    
    Select Case plTipoDado
        Case enumTipoDadoAtributo.Alfanumerico
            flTipoDadoToSTR = "Alfanumérico"
        Case enumTipoDadoAtributo.Numerico
            flTipoDadoToSTR = "Numérico"
    End Select

End Function

'Converter o domínio numérico de tipo de dado para literais da regra de transporte.
Private Function flTipoDadoToFormato(plTipoDado As Long) As String

    Select Case plTipoDado
        Case enumTipoDadoAtributo.Alfanumerico
            flTipoDadoToFormato = "string"
        Case enumTipoDadoAtributo.Numerico
            flTipoDadoToFormato = "number"
        Case 9
            flTipoDadoToFormato = "Grupo"
    End Select
            
End Function

Private Sub chkIndicadorRegraAtiva_Click()

    On Error GoTo ErrorHandler

    fgCursor True

    Call flCarregaListaTipoMensagem
    flHabilitaSalvar

    fgCursor False
    
    Exit Sub

ErrorHandler:
   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - chkIndicadorRegraAtiva_Click"

End Sub

Private Sub chkIndicadorTraducao_Click()

    On Error GoTo ErrorHandler

    fgCursor True
    
    If CBool(chkIndicadorTraducao.Value) Then
        If cboTipoEntrada.ListIndex > -1 Then
            If strOperacao = "Incluir" Then
                flFormatarSpread strCodigoMensagem, lngTipoFormatoMensagemSaida
            Else
                flFormatarSpread Mid$(lstRegra.SelectedItem.Key, 2, 9), _
                                 CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4))
            End If
        End If
    Else
        sprRegra.MaxCols = 2
        sprRegra.MaxRows = 0
    End If

    fgCursor False

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - chkIndicadorTraducao_Click"

End Sub

Private Sub dtpDataFimVigencia_Change()
    
#If EnableSoap = 1 Then
    Dim objRegraTraducao    As MSSOAPLib30.SoapClient30
#Else
    Dim objRegraTraducao    As A7Miu.clsRegraTransporte
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

    On Error GoTo ErrorHandler
    
    With xmlRegraTransporte.documentElement
        
        If .selectSingleNode("Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text = "" Or _
           .selectSingleNode("Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text = gstrDataVazia Then
            
            If Not IsNull(dtpDataFimVigencia.Value) Then
                If dtpDataFimVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                    dtpDataFimVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
                    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
                End If
            End If
        
        Else
            If .selectSingleNode("Grupo_RegraTransporte/IN_CONS").Text = "S" Then Exit Sub
        
            If fgDtXML_To_Date(.selectSingleNode("Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text) <> dtpDataInicioVigencia.Value Then
                                 
                If MsgBox("Deseja reativar a regra?", vbYesNo, "Reativação da Regra") = vbNo Then Exit Sub
                                 
                Set objRegraTraducao = fgCriarObjetoMIU("A7Miu.clsRegraTransporte")
                                    
                If Not objRegraTraducao.ExisteTipoMensagem(.selectSingleNode("Grupo_RegraTransporte/TP_MESG").Text, _
                                                           CLng(.selectSingleNode("Grupo_RegraTransporte/TP_FORM_MESG_SAID").Text), _
                                                           vntCodErro, _
                                                           vntMensagemErro) Then
                    
                    If vntCodErro <> 0 Then
                        GoTo ErrorHandler
                    End If

                    frmMural.txtMural = "Tipo de Mensagem foi excluída ou desativada."
                    frmMural.Show
                    Exit Sub
                End If
                
                Set objRegraTraducao = Nothing
                    
                If .selectSingleNode("Grupo_RegraTransporte/IN_CONS").Text = "N" Then
                    dtpDataFimVigencia.Value = fgDtXML_To_Date(.selectSingleNode("Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text)
                           
                    dtpDataInicioVigencia.Value = dtpDataFimVigencia.Value
                    dtpDataInicioVigencia.MinDate = dtpDataFimVigencia.Value
                    dtpDataInicioVigencia.Enabled = False
                    
                    dtpDataFimVigencia.Value = DateAdd("d", 1, dtpDataInicioVigencia)
                    dtpDataFimVigencia.MinDate = DateAdd("d", 1, dtpDataInicioVigencia)
                    dtpDataFimVigencia.Value = Null
                           
                    .selectSingleNode("Grupo_RegraTransporte/DT_FIM_VIGE_REGR_TRAP").Text = vbNullString
                    
                    strCodigoMensagem = Mid$(lstRegra.SelectedItem.Key, 2, 9)
                    lngTipoFormatoMensagemSaida = CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4))
                    
                    strOperacao = "Incluir"
                End If
            Else
                strOperacao = "Alterar"
            End If
        End If
    
    End With
    
    Exit Sub
ErrorHandler:
    
    Set objRegraTraducao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmRegraTransporte - dtpDataFimVigencia_Change"
    
End Sub

Private Sub dtpDataFimVigencia_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            KeyAscii = 0
    End Select

End Sub

Private Sub dtpDataInicioVigencia_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 8
            KeyAscii = 0
    End Select

End Sub

Private Sub dtpHoraInicioVigencia_KeyPress(KeyAscii As Integer)

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
    
    flLimparCampos
    flCarregarCboTipoEntrada
    
    fgCursor True
    
    flInicializar
    flCarregartreSistemasOrigem
    flCarregaComboDelimitador
    
    fgCursor False
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmRegraTransporte - Form_Load")

End Sub

Private Sub dtpDataInicioVigencia_Change()

    If dtpDataInicioVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
        dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
        dtpDataInicioVigencia.MinDate = dtpDataInicioVigencia.Value
    End If

    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.Value
    dtpDataFimVigencia.Value = dtpDataInicioVigencia.Value
    dtpDataFimVigencia.Value = Null
    
End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A7Miu.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

    On Error GoTo ErrorHandler
    
    Set objMIU = Nothing

    Set xmlMapaRegra = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMIU = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    xmlMapaRegra.loadXML objMIU.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMIU = Nothing
    
    Set xmlRegraTransporte = CreateObject("MSXML2.DOMDocument.4.0")
    xmlRegraTransporte.loadXML xmlMapaRegra.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_RegraTransporteMensagem").xml
    
    Set xmlValidaTag = CreateObject("MSXML2.DOMDocument.4.0")
    
    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    Set xmlValidaTag = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, Me.Name, "flInicializar", lngCodigoErroNegocio)

End Sub

'Carregar o combo com tipo de entrada da mensagem.
Private Sub flCarregarCboTipoEntrada()

    With cboTipoEntrada
        .Clear
        .AddItem "XML"
        .ItemData(.NewIndex) = enumTipoEntradaMensagem.EntradaXML
        .AddItem "String"
        .ItemData(.NewIndex) = enumTipoEntradaMensagem.EntradaString
        'Tipo de entrada CSV removido por não estar sendo utilizado.
        '.AddItem "CSV"
        '.ItemData(.NewIndex) = enumTipoEntradaMensagem.EntradaCSV
    End With

End Sub

'Carregar TreeView com empresas e sistemas de origem.
Private Sub flCarregartreSistemasOrigem()

Dim xmlSistema          As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim strKey              As String
Dim dtmDataServidor     As Date
Dim objNode             As Node
Dim blnVigente          As Boolean

    On Error GoTo ErrorHandler
    
    Set xmlSistema = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlSistema, "", "Repeat_ParametrosLeitura", "")
    Call fgAppendNode(xmlSistema, "Repeat_ParametrosLeitura", "CO_EMPR", "0")
    Call fgAppendNode(xmlSistema, "Repeat_ParametrosLeitura", "TP_VIGE", "S")
    
    Call xmlSistema.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsSistema", xmlSistema))
    
    dtmDataServidor = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    For Each xmlNode In xmlSistema.selectNodes("Repeat_Sistema/*")
        
        On Error Resume Next
            
        strKey = "E" & Format(xmlNode.selectSingleNode("CO_EMPR").Text, "00000")
        
        treSistema.Nodes.Add , , strKey, _
                             xmlNode.selectSingleNode("CO_EMPR").Text & " - " & _
                             xmlNode.selectSingleNode("NO_REDU_EMPR").Text, _
                             "Empresa"
        
        treSistema.Nodes(strKey).Expanded = True
        
        On Error GoTo 0
        
        blnVigente = flRegistroVigente(dtmDataServidor, _
                                       xmlNode.selectSingleNode("DT_INIC_VIGE_SIST").Text, _
                                       xmlNode.selectSingleNode("DT_FIM_VIGE_SIST").Text)
        
        Set objNode = treSistema.Nodes.Add(strKey, _
                                           tvwChild, _
                                           strKey & "S" & xmlNode.selectSingleNode("SG_SIST").Text, _
                                           xmlNode.selectSingleNode("SG_SIST").Text & " - " & xmlNode.selectSingleNode("NO_SIST").Text, _
                                           "Sistema")
        
        If blnVigente Then
            objNode.ForeColor = vbRed
            objNode.Tag = "N"
        Else
            objNode.Tag = "S"
        End If
    
    Next
    
    Exit Sub

ErrorHandler:
    Call fgRaiseError(App.EXEName, Me.Name, "flCarregartreSistemasOrigem", lngCodigoErroNegocio)

End Sub

'Limpar campos do formulário.
Private Sub flLimparCampos()

Dim trwNode                                  As MSComctlLib.Node

    txtTipoMensagem.Text = vbNullString
    txtTipoFormatoMensagemSaida.Text = vbNullString
    txtNatureza.Text = vbNullString
    txtFilaDestino.Text = vbNullString
    
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpHoraInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.HoraAux)
    
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = Null
    
    cboDelimitador.ListIndex = -1
    
    cboDelimitador.Enabled = False
    lblDelimitador.Enabled = False
    
    chkIndicadorTraducao.Value = 0
    chkIndicadorTraducao.Enabled = False
    
    cboTipoEntrada.ListIndex = -1
    cboTipoEntrada.Enabled = False
    treSistemaDestino.Enabled = False
    
    For Each trwNode In treSistemaDestino.Nodes
        trwNode.Checked = False
    Next
    
    flDesabilitaSalvar
    tlbCadastro.Buttons("Sair").Enabled = True
    
    sprRegra.MaxCols = 2
    sprRegra.MaxRows = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlValidaTag = Nothing
    Set frmRegraTransporte = Nothing

End Sub

Private Sub lstRegra_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    On Error GoTo ErrorHandler
    
    fgClassificarListview lstRegra, ColumnHeader.Index
    
    Exit Sub

ErrorHandler:
    mdiBUS.uctLogErros.MostrarErros Err, "frmRegraTraducao - lstRegra_ColumnClick"
    
End Sub

Private Sub lstRegra_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo ErrorHandler
            
    strOperacao = "Alterar"
    
    Call fgCursor(True)
    
    fgLockWindow Me.hwnd
    
    Call flLimparCampos
    Call flXmlToInterface
    
    treSistemaDestino.Enabled = True
    chkIndicadorTraducao.Enabled = True
        
    If strIndicadorConsulta = "S" Then
        tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
        tlbCadastro.Buttons("Salvar").Enabled = False
    Else
        Call flHabilitaSalvar
    End If
                
    Call fgCursor(False)
    
    fgLockWindow 0
   
    Exit Sub

ErrorHandler:
    fgLockWindow 0
    mdiBUS.uctLogErros.MostrarErros Err, "frmRegraTraducao - lstRegra_ItemClick"
    
    If Not treSistema.SelectedItem Is Nothing Then
        If Not treSistema.SelectedItem.Parent Is Nothing Then
            treSistema_NodeClick treSistema.SelectedItem
        End If
    End If
    
    fgCursor False
    
    txtTipoMensagem.SetFocus
    
End Sub

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strKey                                  As String
Dim objTRVNode                              As MSComctlLib.Node

Dim xmlFormato                              As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim xmlNodeFormato                          As MSXML2.IXMLDOMNode

Dim xmlNodeTipoSaida                        As IXMLDOMNode
Dim lngLinhaSpread                          As Long

Dim datDataHoraVigencia                     As Date

    On Error GoTo ErrorHandler
        
    dtpDataInicioVigencia.Enabled = False
    dtpHoraInicioVigencia.Enabled = False
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, "", "Repeat_ParametrosLeitura", "")
    Call fgAppendNode(xmlLeitura, "Repeat_ParametrosLeitura", "TP_MESG", Mid$(lstRegra.SelectedItem.Key, 2, 9))
    Call fgAppendNode(xmlLeitura, "Repeat_ParametrosLeitura", "TP_FORM_MESG_SAID", CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4)))
    Call fgAppendNode(xmlLeitura, "Repeat_ParametrosLeitura", "SG_SIST_ORIG", Mid$(lstRegra.SelectedItem.Key, 15, 3))
    Call fgAppendNode(xmlLeitura, "Repeat_ParametrosLeitura", "CO_EMPR_ORIG", CLng(Mid$(lstRegra.SelectedItem.Key, 18, 5)))
    Call fgAppendNode(xmlLeitura, "Repeat_ParametrosLeitura", "DH_INIC_VIGE_REGR_TRAP", Mid$(lstRegra.SelectedItem.Key, 23, 14))
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("Ler", "A7Server.clsRegraTransporte", xmlLeitura))
    
    'Carrega Informações da Regra
    For Each xmlNode In xmlLeitura.selectNodes("//Grupo_RegraTransporte")
        
        With xmlNode
    
            txtTipoMensagem.Text = lstRegra.SelectedItem.Text & " - " & lstRegra.SelectedItem.SubItems(1)
            txtTipoFormatoMensagemSaida.Text = flTipoSaidaToSTR(CLng(.selectSingleNode("./TP_FORM_MESG_SAID").Text))
            
            txtNatureza.Text = lstRegra.SelectedItem.SubItems(2)
            
            If lstRegra.SelectedItem.SubItems(2) = "ECO" Then
                chkIndicadorTraducao.Value = vbUnchecked
                chkIndicadorTraducao.Enabled = False
                cboTipoEntrada.Enabled = False
            Else
                chkIndicadorTraducao.Enabled = True
                cboTipoEntrada.Enabled = True
            End If
            
            If Trim(.selectSingleNode("./TP_CTER_DELI").Text) <> vbNullString Then
                fgSearchItemCombo cboDelimitador, 0, .selectSingleNode("./TP_CTER_DELI").Text
            End If
            
            chkIndicadorTraducao.Value = Abs(.selectSingleNode("./IN_EXIS_REGR_TRNF").Text = "S")
            
            fgSearchItemCombo cboTipoEntrada, CLng(.selectSingleNode("./TP_FORM_MESG_ENTR").Text)
            
            datDataHoraVigencia = fgDtHrXML_To_Interface(.selectSingleNode("./DH_INIC_VIGE_REGR_TRAP").Text)
            
            dtpDataInicioVigencia.MinDate = Day(datDataHoraVigencia) & "/" & _
                                            Month(datDataHoraVigencia) & "/" & _
                                            Year(datDataHoraVigencia)
            
            dtpDataInicioVigencia.Value = dtpDataInicioVigencia.MinDate
            
            dtpHoraInicioVigencia.Value = Hour(datDataHoraVigencia) & ":" & _
                                          Minute(datDataHoraVigencia) & ":" & _
                                          Second(datDataHoraVigencia)
            
            If dtpDataInicioVigencia.Value > fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
                dtpDataInicioVigencia.Enabled = True
            End If
        
            If Trim(.selectSingleNode("./DT_FIM_VIGE_REGR_TRAP").Text) <> gstrDataVazia Then
                If fgDtXML_To_Date(.selectSingleNode("./DT_FIM_VIGE_REGR_TRAP").Text) < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                    dtpDataFimVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("./DT_FIM_VIGE_REGR_TRAP").Text)
                    dtpDataInicioVigencia.Enabled = True
                Else
                    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
                End If
                
                strIndicadorConsulta = .selectSingleNode("./IN_CONS").Text
                
                If strIndicadorConsulta = "S" Then
                    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
                    tlbCadastro.Buttons("Salvar").Enabled = False
                End If

            Else
               dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.Value
               dtpDataFimVigencia.Value = dtpDataInicioVigencia.Value
               dtpDataFimVigencia.Value = Null
            End If
        
        End With
    Next
        
    'Carrega Spread
    Set xmlFormato = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlLeitura.selectSingleNode("//Grupo_RegraTransporte/TX_REGR_TRNF_MESG") Is Nothing Then
        xmlFormato.loadXML xmlLeitura.selectSingleNode("//Grupo_RegraTransporte/TX_REGR_TRNF_MESG").Text
    End If
     
    For Each xmlNodeFormato In xmlFormato.selectNodes("//*/*/*")
    
        Select Case CLng(xmlLeitura.selectSingleNode("//Grupo_RegraTransporte/TP_FORM_MESG_ENTR").Text)
            
            Case enumTipoEntradaMensagem.EntradaCSV
                
                sprRegra.Row = flLinhaSpread(flTipoXMLParteSaidaToLong(xmlNodeFormato.parentNode.nodeName), _
                                             xmlNodeFormato.nodeName)
                sprRegra.Col = 2
                sprRegra.Text = xmlNodeFormato.selectSingleNode("@Indice").Text
                    
            Case enumTipoEntradaMensagem.EntradaXML
                
                Set xmlNodeTipoSaida = xmlNodeFormato.parentNode
                lngLinhaSpread = flLinhaSpread(flTipoXMLParteSaidaToLong(xmlNodeTipoSaida.nodeName), _
                                               xmlNodeFormato.nodeName)
                Do Until lngLinhaSpread <> 0
                    
                    Set xmlNodeTipoSaida = xmlNodeTipoSaida.parentNode
                    
                    If xmlNodeTipoSaida Is Nothing Then
                        Exit Do
                    End If
                    
                    If Not xmlNodeTipoSaida Is Nothing Then
                    
                        lngLinhaSpread = flLinhaSpread(flTipoXMLParteSaidaToLong(xmlNodeTipoSaida.nodeName), _
                                                       xmlNodeFormato.nodeName)
                    End If
                Loop
                
                If lngLinhaSpread <> 0 Then
                    Set xmlNodeTipoSaida = Nothing
                    sprRegra.Row = lngLinhaSpread
                    sprRegra.Col = 2
                    
                    sprRegra.Text = xmlNodeFormato.selectSingleNode("@TargetTag").Text
                
                    If Not xmlNodeFormato.selectSingleNode("@RepetTag") Is Nothing Then
                        sprRegra.Col = 3
                        sprRegra.Value = xmlNodeFormato.selectSingleNode("@RepetTag").Text
                    End If
                
                    If Not xmlNodeFormato.selectSingleNode("@Default") Is Nothing Then
                        sprRegra.Col = 13
                        sprRegra.Text = xmlNodeFormato.selectSingleNode("@Default").Text
                    End If
                
                    If Not xmlNodeFormato.selectSingleNode("@DefaultObrigatorio") Is Nothing Then
                        sprRegra.Col = 14
                        sprRegra.Value = xmlNodeFormato.selectSingleNode("@DefaultObrigatorio").Text
                    End If
                End If
                      
            Case enumTipoEntradaMensagem.EntradaString
                
                Set xmlNodeTipoSaida = xmlNodeFormato.parentNode
                lngLinhaSpread = flLinhaSpread(flTipoXMLParteSaidaToLong(xmlNodeTipoSaida.nodeName), _
                                               xmlNodeFormato.nodeName)
                Do Until lngLinhaSpread <> 0
                    
                    If xmlNodeTipoSaida Is Nothing Then
                        Exit Do
                    End If
                    
                    Set xmlNodeTipoSaida = xmlNodeTipoSaida.parentNode
                                        
                    If Not xmlNodeTipoSaida Is Nothing Then
                    
                        lngLinhaSpread = flLinhaSpread(flTipoXMLParteSaidaToLong(xmlNodeTipoSaida.nodeName), _
                                                       xmlNodeFormato.nodeName)
                    End If
                Loop
                
                Set xmlNodeTipoSaida = Nothing
                
                If lngLinhaSpread <> 0 Then
                    sprRegra.Row = lngLinhaSpread
                    
                    sprRegra.Col = 2
                    sprRegra.Text = xmlNodeFormato.selectSingleNode("@Inicio").Text
                    
                    sprRegra.Col = 3
                    sprRegra.Text = xmlNodeFormato.selectSingleNode("@TamanhoOriginal").Text
                
                    If Not xmlNodeFormato.selectSingleNode("@Default") Is Nothing Then
                        sprRegra.Col = 13
                        sprRegra.Text = xmlNodeFormato.selectSingleNode("@Default").Text
                    End If
                
                    If Not xmlNodeFormato.selectSingleNode("@DefaultObrigatorio") Is Nothing Then
                        sprRegra.Col = 14
                        sprRegra.Value = xmlNodeFormato.selectSingleNode("@DefaultObrigatorio").Text
                    End If
                End If
        End Select
        
    Next
    
    If Not xmlFormato.selectSingleNode("*/@FilaDestino") Is Nothing Then
        txtFilaDestino.Text = xmlFormato.selectSingleNode("//@FilaDestino").Text
    Else
        txtFilaDestino.Text = vbNullString
    End If
    
    'Carrega Informações dos Sistemas Destino
    For Each objTRVNode In treSistemaDestino.Nodes
        If objTRVNode.Parent Is Nothing Then
            objTRVNode.Tag = "0"
        End If
        objTRVNode.Checked = False
    Next
    
    For Each xmlNode In xmlLeitura.selectNodes("//Repeat_SistemaDestino/*")
        With xmlNode
            strKey = "E" & Format$(.selectSingleNode(".//CO_EMPR_DEST").Text, "00000") & _
                     "S" & .selectSingleNode(".//SG_SIST_DEST").Text
            
            treSistemaDestino.Nodes(strKey).Checked = True
            treSistemaDestino.Nodes(strKey).Parent.Tag = CStr(CLng(treSistemaDestino.Nodes(strKey).Parent.Tag) + 1)
        End With
    Next
    
    'Marca o Nó pai quando necessário
    For Each objTRVNode In treSistemaDestino.Nodes
        If objTRVNode.Parent Is Nothing Then
            If objTRVNode.children = CLng(objTRVNode.Tag) Then
                objTRVNode.Checked = True
            End If
        End If
    Next
    
    Set xmlFormato = Nothing
    Set xmlLeitura = Nothing
    
    Exit Sub

ErrorHandler:
    Set xmlFormato = Nothing
    Set xmlLeitura = Nothing
    
    Call fgRaiseError(App.EXEName, Me.Name, "flMoveObjetToInterface", lngCodigoErroNegocio)
    
End Sub

'Converter as literais de Parte da saída da regra de transporte para o domínio numérico.
Private Function flTipoXMLParteSaidaToLong(pstrXMLParteSaida As String) As Long
    
    Select Case pstrXMLParteSaida
        Case "IDOutPut"
            flTipoXMLParteSaidaToLong = enumTipoParteSaida.ParteId
        Case "XMLOutPut"
            flTipoXMLParteSaidaToLong = enumTipoParteSaida.ParteXML
        Case "STROutPut"
            flTipoXMLParteSaidaToLong = enumTipoParteSaida.ParteSTR
        Case "CSVOutPut"
            flTipoXMLParteSaidaToLong = enumTipoParteSaida.ParteCSV
    End Select

End Function

'Retornar a linha do spread de acordo com a parte de saída e nome de atributo recebido.
Private Function flLinhaSpread(plngParteSaida As Long, _
                               pstrAtributo As String) As Long

Dim lngRow                                   As Long
Dim strAtributoAtual                         As String
Dim lngParteSaidaAtual                       As Long
    
    With sprRegra
        For lngRow = 1 To .MaxRows
            .Row = lngRow
            
            .Col = 5
            strAtributoAtual = Trim$(.Text)
            
            .Col = 10
            lngParteSaidaAtual = CLng(.Text)
            
            If strAtributoAtual = pstrAtributo And _
               lngParteSaidaAtual = plngParteSaida Then
                flLinhaSpread = lngRow
                Exit For
            End If
        Next
    End With

End Function
                               
'Montar XML com formato. Este XML é utilizado na regra de transporte.
Private Function flMontarXMLFormato() As String

Dim xmlFormato                              As MSXML2.DOMDocument40
Dim lngRow                                  As Long
Dim strContexto                             As String

Dim strNodePai(3, 10)                       As String
Dim lngNivel                                As Long
Dim lngNivelAtual                           As Long
Dim lngTipoParteSaida                       As Long

    On Error GoTo ErrorHandler

    Set xmlFormato = CreateObject("MSXML2.DOMDocument.4.0")
    
    fgAppendNode xmlFormato, "", "Formato", ""
    If Trim$(txtFilaDestino.Text) <> vbNullString Then
        fgAppendAttribute xmlFormato, "Formato", "FilaDestino", Trim$(txtFilaDestino.Text)
    End If
    
    lngNivelAtual = 1
    
    With sprRegra
        For lngRow = 1 To .MaxRows
            
            .Row = lngRow
            
            'Define o tipo da parte de saida
            .Col = 10
            lngTipoParteSaida = CLng(.Text)
            
            'Obtem o nivel para formatação
            .Col = 11
            lngNivel = CLng(.Text)
            
            If lngNivel = lngNivelAtual Then
                'Incluir no mesmo nivel
                
                .Col = 10
                If lngTipoParteSaida = enumTipoParteSaida.ParteId Then
                    'ID
                    If xmlFormato.selectNodes("//IDOutPut").length = 0 Then
                        fgAppendNode xmlFormato, "Formato", "IDOutPut", vbNullString
                    End If
                    strContexto = "IDOutPut"
                    strNodePai(enumTipoParteSaida.ParteId, 1) = "IDOutPut"
                ElseIf lngTipoParteSaida = enumTipoParteSaida.ParteSTR Then
                    'String
                    If xmlFormato.selectNodes("//STROutPut").length = 0 Then
                        fgAppendNode xmlFormato, "Formato", "STROutPut", vbNullString
                    End If
                    strContexto = "STROutPut"
                    strNodePai(enumTipoParteSaida.ParteSTR, 1) = "STROutPut"
                ElseIf lngTipoParteSaida = enumTipoParteSaida.ParteCSV Then
                    'Apendar indice
                    If xmlFormato.selectNodes("//CSVOutPut").length = 0 Then
                        
                        fgAppendNode xmlFormato, "Formato", "CSVOutPut", vbNullString
                        fgAppendAttribute xmlFormato, "CSVOutPut", "Delimitador", strDelimitadorCSV
                                     
                    End If
                    strContexto = "CSVOutPut"
                    strNodePai(enumTipoParteSaida.ParteCSV, 1) = "CSVOutPut"
                Else
                    'XML
                    'OutPutName = Nome do XML de saida
                    If xmlFormato.selectNodes("//XMLOutPut").length = 0 Then
                        fgAppendNode xmlFormato, "Formato", "XMLOutPut", vbNullString
                        fgAppendAttribute xmlFormato, "XMLOutPut", "OutPutName", "SaidaXML"
                    End If
                    strContexto = "XMLOutPut"
                    strNodePai(enumTipoParteSaida.ParteXML, 1) = "XMLOutPut"
                End If
                
                'Append com nome fisico do atributo
                .Col = 5
                fgAppendNode xmlFormato, strNodePai(lngTipoParteSaida, lngNivelAtual), .Text, vbNullString
                strContexto = strNodePai(lngTipoParteSaida, lngNivelAtual) & "/" & .Text
                
                'Append do atributo para tipo de dado
                .Col = 6
                fgAppendAttribute xmlFormato, strContexto, "Tipo", flTipoDadoToFormato(CLng(.Text))
                
                'Append do atributo para tamanho do atributo = utilizado para formatação
                .Col = 7
                fgAppendAttribute xmlFormato, strContexto, "Tamanho", .Text
        
                'Append do atributo para quantidade de decimais
                .Col = 8
                fgAppendAttribute xmlFormato, strContexto, "Decimais", .Text
                                  
                'Append do atributo para indicador de obrigatoriedade
                .Col = 9
                fgAppendAttribute xmlFormato, strContexto, "Obrigatorio", .Text
            
                Select Case cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex)
                    Case enumTipoEntradaMensagem.EntradaCSV
                        .Col = 2
                        fgAppendAttribute xmlFormato, strContexto, "Indice", .Text
                    Case enumTipoEntradaMensagem.EntradaString
                        .Col = 2
                        fgAppendAttribute xmlFormato, strContexto, "Inicio", .Text
                        .Col = 3
                        fgAppendAttribute xmlFormato, strContexto, "TamanhoOriginal", .Text
                    Case enumTipoEntradaMensagem.EntradaXML
                        .Col = 2
                        fgAppendAttribute xmlFormato, strContexto, "TargetTag", Trim$(.Text)
                        .Col = 3
                        fgAppendAttribute xmlFormato, strContexto, "RepetTag", .Value
                        fgAppendAttribute xmlFormato, strContexto, "UltimaPosicao", 0
                End Select
            
                'Append do atributo para quantidade de repetições
                .Col = 12
                fgAppendAttribute xmlFormato, strContexto, "Repeticoes", .Text
        
                .Col = 13
                fgAppendAttribute xmlFormato, strContexto, "Default", .Text
        
                .Col = 14
                fgAppendAttribute xmlFormato, strContexto, "DefaultObrigatorio", .Value
        
            ElseIf lngNivel > lngNivelAtual Then
                'Adiciona novo pai para inclusao das tags filhos (ultima tag incluida)
                .Row = lngRow - 1
                .Col = 5
                strNodePai(lngTipoParteSaida, lngNivel) = .Text
                lngRow = lngRow - 1
                lngNivelAtual = lngNivel
                
            ElseIf lngNivel < lngNivelAtual Then
                'Volta para o nivel anterioe
                'Pai já incluso no array
                lngNivelAtual = lngNivel
                lngRow = lngRow - 1
            
            End If
        Next
    End With
    
    flMontarXMLFormato = xmlFormato.xml
    
    Set xmlFormato = Nothing
    
    'INDICES
    ' 5 - NO_ATRB_MESG
    ' 6 - TP_DADO_ATRB_MESG
    ' 7 - QT_CTER_ATRB
    ' 8 - QT_CASA_DECI_ATRB
    ' 9 - IN_OBRI_ATRB
    '10 - TP_FORM_MESG
    '11 - NU_NIVE_MESG_ATRB
    '12 - QT_REPE
    '13 - VL_DEFA
    '14 - IN_OBRI_DEFA
    
    Exit Function
ErrorHandler:
    
    Set xmlFormato = Nothing
    
    Call fgRaiseError(App.EXEName, Me.Name, "flMontarXMLFormato", lngCodigoErroNegocio)
    
End Function

Private Sub objInclusaoRegra_EventoEscolhido(ByVal pstrTipoMensagem As String, _
                                             ByVal pstrDescricaoTipoMensagem As String, _
                                             ByVal plngTipoFormatoMensagemSaida As Long, _
                                             ByVal plngNaturezaMensagem As Long)

    flLimparCampos
    cboTipoEntrada.Enabled = True
    fraDetalhe.Enabled = True
    treSistemaDestino.Enabled = True
    chkIndicadorTraducao.Enabled = True
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = False
    
    strCodigoMensagem = pstrTipoMensagem
    lngTipoFormatoMensagemSaida = plngTipoFormatoMensagemSaida
    
    txtTipoMensagem.Text = pstrTipoMensagem & " - " & pstrDescricaoTipoMensagem
    txtTipoFormatoMensagemSaida.Text = flTipoSaidaToSTR(plngTipoFormatoMensagemSaida)
    
    Select Case plngNaturezaMensagem
        
        Case enumNaturezaMensagem.MensagemConsulta
            txtNatureza.Text = "Consulta"
            chkIndicadorTraducao.Enabled = True
            cboTipoEntrada.Enabled = True
        Case enumNaturezaMensagem.MensagemECO
            txtNatureza.Text = "Eco"
            chkIndicadorTraducao.Value = vbUnchecked
            chkIndicadorTraducao.Enabled = False
            cboTipoEntrada.Enabled = False
        Case enumNaturezaMensagem.MensagemEnvio
            txtNatureza.Text = "Envio de dados"
            chkIndicadorTraducao.Enabled = True
            cboTipoEntrada.Enabled = True
    End Select
    
    If Not lstRegra.SelectedItem Is Nothing Then
        lstRegra.SelectedItem.Selected = False
    End If
    strOperacao = "Incluir"

End Sub

Private Sub sprRegra_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

Dim vntNomeTag                            As Variant
    
    If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaXML Then
       
        If ChangeMade And Col <> 3 And Col <> 13 And Col <> 14 Then
            
            sprRegra.GetText Col, Row, vntNomeTag
           
            If Trim(vntNomeTag) <> vbNullString Then
           
               If Not flValidaTag(vntNomeTag) Then
        
                   blnTagInvalida = True
                   sprRegra.Row = Row
                   sprRegra.Col = Col
                   
                   MsgBox "Nome de Tag Inválido."
                   sprRegra.SetFocus
                   sprRegra.Text = vntNomeTag
               Else
                   blnTagInvalida = False
               End If
            End If
        
        Else
            blnTagInvalida = False
            
        End If
    
    Else
        blnTagInvalida = False
    
    End If
    
End Sub

Private Sub sprRegra_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
    If cboTipoEntrada.ListIndex <> -1 Then
    
        If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaXML Then
            If blnTagInvalida Then
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler

    Select Case Button.Key
        Case "Limpar"
        
            If treSistema.SelectedItem Is Nothing Then Exit Sub
            
            If treSistema.SelectedItem.Parent Is Nothing Then Exit Sub
                        
            If treSistema.SelectedItem.ForeColor = vbRed Then
                frmMural.txtMural = "Sistema de origem não está vigente."
                frmMural.Show
                Exit Sub
            End If
                        
            txtTipoMensagem.SetFocus
                        
            Set objInclusaoRegra = New frmInclusaoRegra
            objInclusaoRegra.CodigoBanco = CLng(Mid$(treSistema.SelectedItem.Key, 2, 5))
            objInclusaoRegra.SistemaOrigem = Mid$(treSistema.SelectedItem.Key, 8)
            objInclusaoRegra.Show vbModal
            Set objInclusaoRegra = Nothing
            
            fgCursor False
  
        Case "Excluir"
            flExcluir
        Case "Salvar"
            flSalvar
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select

    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
   
    fgCursor False
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "tlbCadastro_ButtonClick")
    
    Call flCarregaListaTipoMensagem
    
    fgCursor False
        
    If strOperacao = "Excluir" Then
        flLimparCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
    fgCursor False
    
End Sub

'Excluir a regra de trnaporte corrente.
Private Sub flExcluir()

Dim xmlProcessamento                        As MSXML2.DOMDocument40
    
    On Error GoTo ErrorHandler
    
    If MsgBox("Confirma Exclusão da Regra de Transporte selecionada ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    Set xmlProcessamento = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlProcessamento, "", "Repeat_Processamento", "")
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "TP_MESG", Mid$(lstRegra.SelectedItem.Key, 2, 9))
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "TP_FORM_MESG_SAID", CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4)))
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "SG_SIST_ORIG", Mid$(lstRegra.SelectedItem.Key, 15, 3))
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "CO_EMPR_ORIG", CLng(Mid$(lstRegra.SelectedItem.Key, 18, 5)))
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "DH_INIC_VIGE_REGR_TRAP", fgDtHr_To_Xml(dtpDataInicioVigencia.Value & " " & dtpHoraInicioVigencia.Value))
    Call fgAppendNode(xmlProcessamento, "Repeat_Processamento", "DH_ULTI_ATLZ", Mid$(lstRegra.SelectedItem.Key, 37, 14))
    
    Call xmlProcessamento.loadXML(fgMIUExecutarGenerico("Excluir", "A7Server.clsRegraTransporte", xmlProcessamento))
    Call flCarregaListaTipoMensagem
    
    fgCursor
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    Exit Sub

ErrorHandler:
    Call fgRaiseError(App.EXEName, Me.Name, "flExcluir", lngCodigoErroNegocio)

End Sub

Private Sub treSistema_NodeClick(ByVal Node As MSComctlLib.Node)

Dim strSiglaSistema                         As String
Dim xmlRegra                                As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strRegras                               As String

    On Error GoTo ErrorHandler

    fgCursor True
    
    fgLockWindow Me.hwnd
    
    If Not Node.Parent Is Nothing Then
        flCarregartreSistemasDestino
    End If
    
    flCarregaListaTipoMensagem
    
    flLimparCampos
    
    fgCursor False
    
    fgLockWindow 0
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    fgLockWindow 0
    
    Set xmlRegra = Nothing
    Call mdiBUS.uctLogErros.MostrarErros(Err, "treSistema_NodeClick")

End Sub

'Converter o domínio numérico de tipo de entrada para as literais.
Private Function flTipoEntradaToSTR(ByVal plngTipoEntrada As Long) As String

    Select Case plngTipoEntrada
        Case enumTipoEntradaMensagem.EntradaCSV
            flTipoEntradaToSTR = "Entrada CSV"
        Case enumTipoEntradaMensagem.EntradaString
            flTipoEntradaToSTR = "Entrada String"
        Case enumTipoEntradaMensagem.EntradaXML
            flTipoEntradaToSTR = "Entrada XML"
    End Select

End Function

'Salvar as informações correntes da regra de transporte.
Private Sub flSalvar()

Dim xmlPropriedades                         As MSXML2.DOMDocument40
Dim xmlNodeSistemaDestino                   As MSXML2.IXMLDOMNode
Dim xmlNodeSistemaDestinoAux                As MSXML2.IXMLDOMNode
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlCDATA                                As MSXML2.IXMLDOMCDATASection
Dim objTRWItem                              As MSComctlLib.Node

Dim strCodigoMensagemAux                    As String
Dim lngTipoFormatoMensagemSaidaAux          As Long
Dim strSistemaOrigem                        As String
Dim lngCodigoEmpresa                        As Long
Dim strDataInicioVigencia                   As String
Dim strRetorno                              As String

    On Error GoTo ErrorHandler
    
    strRetorno = flValidarCampos()
    
    If strRetorno <> vbNullString Then
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If
    
    If Not IsNull(dtpDataFimVigencia.Value) Then
        If xmlRegraTransporte.documentElement.selectSingleNode("//DT_FIM_VIGE_REGR_TRAP").Text = gstrDataVazia Then
            If MsgBox("Deseja desativar o registro a partir do dia " & dtpDataFimVigencia.Value & " ?", vbYesNo, "Atributos Mensagens") = vbNo Then Exit Sub
        Else
            If fgDtXML_To_Date(xmlRegraTransporte.documentElement.selectSingleNode("//DT_FIM_VIGE_REGR_TRAP").Text) <> dtpDataFimVigencia.Value Then
                If MsgBox("Deseja desativar o registro a partir do dia " & dtpDataFimVigencia.Value & " ?", vbYesNo, "Atributos Mensagens") = vbNo Then Exit Sub
            End If
        End If
    End If
    
    fgCursor True
    
    If strOperacao = "Incluir" Then
        strCodigoMensagemAux = strCodigoMensagem
        lngTipoFormatoMensagemSaidaAux = lngTipoFormatoMensagemSaida
    Else
        strCodigoMensagemAux = Mid$(lstRegra.SelectedItem.Key, 2, 9)
        lngTipoFormatoMensagemSaidaAux = CLng(Mid$(lstRegra.SelectedItem.Key, 11, 4))
    End If
    
    If treSistema.SelectedItem.Parent Is Nothing Then
        strSistemaOrigem = Mid(lstRegra.SelectedItem.Key, (Len(lstRegra.SelectedItem.Key) - 15), 3)
        lngCodigoEmpresa = CLng(Mid$(lstRegra.SelectedItem.Key, (Len(lstRegra.SelectedItem.Key) - 12), 5))
    Else
        lngCodigoEmpresa = CLng(Mid$(treSistema.SelectedItem.Key, 2, 5))
        strSistemaOrigem = Mid$(treSistema.SelectedItem.Key, 8)
    End If
    
    strDataInicioVigencia = fgDtHr_To_Xml(dtpDataInicioVigencia.Value & " " & dtpHoraInicioVigencia.Value)

    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlPropriedades.appendChild xmlRegraTransporte.selectSingleNode("//Grupo_RegraTransporteMensagem").cloneNode(True)

    xmlPropriedades.selectSingleNode("//Grupo_RegraTransporte/@Operacao").Text = strOperacao
    
    With xmlPropriedades.selectSingleNode("//Grupo_RegraTransporte")
        .selectSingleNode("TP_MESG").Text = strCodigoMensagemAux
        .selectSingleNode("TP_FORM_MESG_SAID").Text = lngTipoFormatoMensagemSaidaAux
        .selectSingleNode("SG_SIST_ORIG").Text = strSistemaOrigem
        .selectSingleNode("CO_EMPR_ORIG").Text = lngCodigoEmpresa
        .selectSingleNode("DH_INIC_VIGE_REGR_TRAP").Text = strDataInicioVigencia
        .selectSingleNode("IN_EXIS_REGR_TRNF").Text = IIf(CBool(chkIndicadorTraducao.Value), "S", "N")
        
        If cboTipoEntrada.ListIndex <> -1 Then
            .selectSingleNode("TP_FORM_MESG_ENTR").Text = cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex)
        Else
            .selectSingleNode("TP_FORM_MESG_ENTR").Text = 0
        End If
        
        .selectSingleNode("TP_CTER_DELI").Text = cboDelimitador.Text
        
        If Not lstRegra.SelectedItem Is Nothing Then
            .selectSingleNode("CO_TEXT_XML").Text = lstRegra.SelectedItem.Tag
        End If
        
        If dtpDataFimVigencia.Value <> vbNull Then
            .selectSingleNode("DT_FIM_VIGE_REGR_TRAP").Text = fgDt_To_Xml(dtpDataFimVigencia.Value)
        Else
            .selectSingleNode("DT_FIM_VIGE_REGR_TRAP").Text = vbNullString
        End If
        .selectSingleNode("CO_USUA_ULTI_ATLZ").Text = ""
        .selectSingleNode("CO_ETCA_TRAB_ULTI_ATLZ").Text = ""
        
        .selectSingleNode("TX_REGR_TRNF_MESG").Text = ""
        Set xmlCDATA = xmlPropriedades.createCDATASection(flMontarXMLFormato())
        .selectSingleNode("TX_REGR_TRNF_MESG").appendChild xmlCDATA
        Set xmlCDATA = Nothing
        
        If Not lstRegra.SelectedItem Is Nothing Then
            .selectSingleNode("DH_ULTI_ATLZ").Text = Mid$(lstRegra.SelectedItem.Key, 37, 14)
        Else
            .selectSingleNode("DH_ULTI_ATLZ").Text = vbNullString
        End If
    End With

    Set xmlNodeSistemaDestinoAux = xmlMapaRegra.selectSingleNode("//Grupo_SistemaDestino").cloneNode(True)

    'Remove o node ded sistema destino da propriedade (para caso não tenha sistema destino selecionado
    For Each xmlNode In xmlPropriedades.selectSingleNode("//Repeat_SistemaDestino").childNodes
        xmlPropriedades.selectSingleNode("//Repeat_SistemaDestino").removeChild xmlNode 'xmlPropriedades.selectSingleNode("//Grupo_SistemaDestino")
    Next
    
    With xmlNodeSistemaDestinoAux
        .selectSingleNode("TP_MESG").Text = strCodigoMensagemAux
        .selectSingleNode("TP_FORM_MESG_SAID").Text = lngTipoFormatoMensagemSaidaAux
        .selectSingleNode("SG_SIST_ORIG").Text = strSistemaOrigem
        .selectSingleNode("CO_EMPR_ORIG").Text = lngCodigoEmpresa
        .selectSingleNode("DH_INIC_VIGE_REGR_TRAP").Text = strDataInicioVigencia
    End With

    For Each objTRWItem In treSistemaDestino.Nodes

        If Not objTRWItem.Parent Is Nothing Then
            If objTRWItem.Checked Then
                Set xmlNodeSistemaDestino = xmlNodeSistemaDestinoAux.cloneNode(True)
                
                With xmlNodeSistemaDestino
                    .selectSingleNode("CO_EMPR_DEST").Text = CLng(Mid$(objTRWItem.Key, 2, 5))
                    .selectSingleNode("SG_SIST_DEST").Text = Mid$(objTRWItem.Key, 8)
                    
                    xmlPropriedades.selectSingleNode("//Repeat_SistemaDestino").appendChild xmlNodeSistemaDestino
                End With
                Set xmlNodeSistemaDestino = Nothing
            End If
        End If

    Next

    Call fgMIUExecutarGenerico(strOperacao, "A7Server.clsRegraTransporte", xmlPropriedades)
    Call flCarregaListaTipoMensagem
    
    fgCursor
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    cboTipoEntrada.Enabled = True
    treSistemaDestino.Enabled = True
    chkIndicadorTraducao.Enabled = True
        
    Call flHabilitaSalvar
    
    strCodigoMensagem = 0
    strOperacao = "Alterar"
    
    Set xmlPropriedades = Nothing
    Set xmlNodeSistemaDestino = Nothing
    Set xmlNodeSistemaDestinoAux = Nothing
    
    Exit Sub
    
ErrorHandler:
    Set xmlPropriedades = Nothing
    Set xmlNodeSistemaDestino = Nothing
    Set xmlNodeSistemaDestinoAux = Nothing
    Set xmlCDATA = Nothing
    
    Call fgRaiseError(App.EXEName, Me.Name, "flSalvar", lngCodigoErroNegocio)
    
End Sub

Private Sub treSistemaDestino_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim objNodeAux                              As MSComctlLib.Node
    
    If Node.Parent Is Nothing Then
        For Each objNodeAux In treSistemaDestino.Nodes
            If Not objNodeAux.Parent Is Nothing Then
                If objNodeAux.Parent.Key = Node.Key Then
                    If objNodeAux.Parent.Checked Then
                        objNodeAux.Checked = vbChecked
                    Else
                        objNodeAux.Checked = vbUnchecked
                    End If
                End If
            End If
        Next
    ElseIf Not Node.Checked Then
        Node.Parent.Checked = False
    End If
        
End Sub

'Validar os valores informados para a regra de transporte.
Private Function flValidarCampos() As String

Dim objTRVNode                              As MSComctlLib.Node
Dim blnSistemaDestino                       As Boolean
Dim datServidor                             As Date

    On Error GoTo ErrorHandler
    
    datServidor = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    If Trim(txtNatureza.Text) <> "Eco" Then
        If cboTipoEntrada.ListIndex = -1 Then
            flValidarCampos = "Informe o tipo de entrada da mensagem."
            cboTipoEntrada.SetFocus
            Exit Function
        End If
    End If
    
    If Trim(txtNatureza.Text) <> "Eco" Then
        If cboTipoEntrada.ItemData(cboTipoEntrada.ListIndex) = enumTipoEntradaMensagem.EntradaCSV Then
            If Trim$(cboDelimitador) = vbNullString Then
                flValidarCampos = "Selecione o caracter delimitador para a entrada CSV."
                cboDelimitador.SetFocus
                Exit Function
            End If
        End If
    End If
    
    For Each objTRVNode In treSistema.Nodes
        If objTRVNode.Selected Then
            If Not objTRVNode.Parent Is Nothing Then
                If objTRVNode.Tag = "N" Then
                    flValidarCampos = "Sistema Origem: " & objTRVNode.Text & " não está vigente."
                    Exit Function
                End If
            End If
        End If
    Next
    
    blnSistemaDestino = False
    For Each objTRVNode In treSistemaDestino.Nodes
        If objTRVNode.Checked Then
            If Not objTRVNode.Parent Is Nothing Then
                blnSistemaDestino = True
                Exit For
            End If
        End If
    Next
    
    If Not blnSistemaDestino Then
        flValidarCampos = "Informe pelo menos 1 sistema destino."
        treSistemaDestino.SetFocus
        Exit Function
    End If
    
    For Each objTRVNode In treSistemaDestino.Nodes
        If objTRVNode.Checked Then
            If Not objTRVNode.Parent Is Nothing Then
                If objTRVNode.Tag = "N" Then
                    flValidarCampos = "Sistema Destino: " & objTRVNode.Text & " não está vigente."
                    Exit Function
                End If
            End If
        End If
    Next
    
    
    flValidarCampos = ""
    
    Exit Function

ErrorHandler:
    
    fgRaiseError App.EXEName, "frmRegraTransporte", "flValidarCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Carregar listview com os tipos de mensagens que contém regra de acordo com a empresa e sistema de origem selecionados.
Private Sub flCarregaListaTipoMensagem()

Dim lngCodigoEmpresa    As Long
Dim strSiglaSistema     As String
Dim lngRegraAtiva       As Long
Dim xmlRegra            As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode

    On Error GoTo ErrorHandler

    If treSistema.SelectedItem Is Nothing Then Exit Sub
    
    lngCodigoEmpresa = CLng(Mid$(treSistema.SelectedItem.Key, 2, 5))
    lngRegraAtiva = chkIndicadorRegraAtiva.Value

    If Not treSistema.SelectedItem.Parent Is Nothing Then
       flHabilitaSalvar
       strSiglaSistema = Mid$(treSistema.SelectedItem.Key, 8)
       dtpDataFimVigencia.Enabled = True
    Else
       flDesabilitaSalvar
       strSiglaSistema = ""
       flLimparCampos
       lstRegra.ListItems.Clear
       treSistemaDestino.Nodes.Clear
       dtpDataFimVigencia.Enabled = False
       Exit Sub
    End If

    Set xmlRegra = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlRegra, "", "Repeat_ParametrosLeitura", "")
    Call fgAppendNode(xmlRegra, "Repeat_ParametrosLeitura", "CO_EMPR_ORIG", lngCodigoEmpresa)
    Call fgAppendNode(xmlRegra, "Repeat_ParametrosLeitura", "SG_SIST_ORIG", strSiglaSistema)
    Call fgAppendNode(xmlRegra, "Repeat_ParametrosLeitura", "IN_REGR_ATIV", lngRegraAtiva)
    
    Call xmlRegra.loadXML(fgMIUExecutarGenerico("LerTodos", "A7Server.clsRegraTransporte", xmlRegra))
    
    lstRegra.ListItems.Clear
    
    For Each xmlNode In xmlRegra.selectNodes("//Repeat_Regra/*")
        
        With lstRegra.ListItems.Add(, "R" & _
                                      Left$(xmlNode.selectSingleNode("./TP_MESG").Text & String(9, " "), 9) & _
                                      Format(xmlNode.selectSingleNode("./TP_FORM_MESG_SAID").Text, "0000") & _
                                      Left$(xmlNode.selectSingleNode("./SG_SIST_ORIG").Text & "   ", 3) & _
                                      Format$(xmlNode.selectSingleNode("./CO_EMPR_ORIG").Text, "00000") & _
                                      xmlNode.selectSingleNode("./DH_INIC_VIGE_REGR_TRAP").Text & _
                                      xmlNode.selectSingleNode("./DH_ULTI_ATLZ").Text, _
                                      xmlNode.selectSingleNode("./TP_MESG").Text, , _
                                      "Evento")
            
            .Tag = xmlNode.selectSingleNode("./CO_TEXT_XML").Text
            .SubItems(1) = xmlNode.selectSingleNode("./NO_TIPO_MESG").Text
            .SubItems(2) = flTipoSaidaToSTR(CLng(xmlNode.selectSingleNode("./TP_FORM_MESG_SAID").Text))
            
            Select Case CLng(xmlNode.selectSingleNode("./TP_NATZ_MESG").Text)
                Case enumNaturezaMensagem.MensagemConsulta
                   .SubItems(3) = "Consulta"
                Case enumNaturezaMensagem.MensagemEnvio
                   .SubItems(3) = "Envio de dados"
                Case enumNaturezaMensagem.MensagemECO
                   .SubItems(3) = "ECO"
            End Select
        
            .SubItems(4) = flTipoEntradaToSTR(CLng(xmlNode.selectSingleNode("./TP_FORM_MESG_ENTR").Text))
            .SubItems(5) = fgDtHrXML_To_Interface(xmlNode.selectSingleNode("./DH_INIC_VIGE_REGR_TRAP").Text)
            
            If CStr(xmlNode.selectSingleNode("./DT_FIM_VIGE_REGR_TRAP").Text) <> gstrDataVazia Then
                .SubItems(6) = Format(fgDtXML_To_Date(xmlNode.selectSingleNode("./DT_FIM_VIGE_REGR_TRAP").Text), gstrMascaraDataDtp)
            Else
                .SubItems(6) = ""
            End If
         
        End With
    Next
    
    fgCursor False
    Set xmlRegra = Nothing
    Exit Sub

ErrorHandler:
    Set xmlRegra = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListaTipoMensagem", 0
    
End Sub

'Verifica se as datas informadas se encontram em um período vigente.
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

'Posicionar item no listview tipos de mensagem.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

    If lstRegra.ListItems.Count = 0 Then
        flLimparCampos
        Exit Sub
    End If
    
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstRegra.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstRegra_ItemClick objListItem
           lstRegra.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimparCampos
    End If

End Sub

'Validar tag do XML.
Private Function flValidaTag(ByVal pstrNomeTag As String) As Boolean

    On Error GoTo ErrorHandler
    
    xmlValidaTag.validateOnParse = False
    xmlValidaTag.resolveExternals = False
    
    xmlValidaTag.loadXML ""
    Call fgAppendNode(xmlValidaTag, "", pstrNomeTag, "")
    
    flValidaTag = True
    
    Exit Function
ErrorHandler:
    
    Err.Clear
    
    flValidaTag = False
    
End Function

'Carregar TreeView com Sistemas de destino de acordo com a empresa de origem selecionada.
Private Sub flCarregartreSistemasDestino()

Dim xmlSistema          As MSXML2.DOMDocument40
Dim xmlNode             As MSXML2.IXMLDOMNode
Dim strKey              As String
Dim dtmDataServidor     As Date
Dim objNode             As Node
Dim blnVigente          As Boolean
Dim lngCodigoEmpresa    As Long

    On Error GoTo ErrorHandler
    
    Set xmlSistema = CreateObject("MSXML2.DOMDocument.4.0")
        
    lngCodigoEmpresa = CLng(Mid$(treSistema.SelectedItem.Key, 2, 5))
        
    If treSistemaDestino.Nodes.Count > 0 Then
        If CLng(Mid$(treSistemaDestino.Nodes.Item(1).Key, 2, 5)) = lngCodigoEmpresa Then Exit Sub
    End If
        
    treSistemaDestino.Nodes.Clear
        
    Call fgAppendNode(xmlSistema, "", "Repeat_ParametrosLeitura", "")
    Call fgAppendNode(xmlSistema, "Repeat_ParametrosLeitura", "CO_EMPR", lngCodigoEmpresa)
    Call fgAppendNode(xmlSistema, "Repeat_ParametrosLeitura", "TP_VIGE", "N")
    
    Call xmlSistema.loadXML(fgMIUExecutarGenerico("LerTodos", "A6A7A8.clsSistema", xmlSistema))
    
    dtmDataServidor = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    For Each xmlNode In xmlSistema.selectNodes("Repeat_Sistema/*")
        
        On Error Resume Next
            
        strKey = "E" & Format(xmlNode.selectSingleNode("CO_EMPR").Text, "00000")
                    
        treSistemaDestino.Nodes.Add , , strKey, _
                             xmlNode.selectSingleNode("CO_EMPR").Text & " - " & _
                             xmlNode.selectSingleNode("NO_REDU_EMPR").Text, _
                             "Empresa"
    
        treSistemaDestino.Nodes(strKey).Expanded = True
        
        On Error GoTo 0
        
        blnVigente = flRegistroVigente(dtmDataServidor, _
                                       xmlNode.selectSingleNode("DT_INIC_VIGE_SIST").Text, _
                                       xmlNode.selectSingleNode("DT_FIM_VIGE_SIST").Text)
        
        Set objNode = treSistemaDestino.Nodes.Add(strKey, _
                                                  tvwChild, _
                                                  strKey & "S" & xmlNode.selectSingleNode("SG_SIST").Text, _
                                                  xmlNode.selectSingleNode("SG_SIST").Text & " - " & xmlNode.selectSingleNode("NO_SIST").Text, _
                                                  "SistemaDestino")
    
        If blnVigente Then
            objNode.ForeColor = vbRed
            objNode.Tag = "N"
        Else
            objNode.Tag = "S"
        End If
    
    Next
    
    Exit Sub

ErrorHandler:
    Call fgRaiseError(App.EXEName, Me.Name, "flCarregartreSistemas", lngCodigoErroNegocio)

End Sub

'Habilita a gravação para operações de alteração.
Private Sub flHabilitaSalvar()

    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True

End Sub

'Desabilita a gravação para operações diferentes de alteração. (Para inclusão o botão é liberado após a edição).
Private Sub flDesabilitaSalvar()

    tlbCadastro.Buttons("Excluir").Enabled = False
    tlbCadastro.Buttons("Salvar").Enabled = False

End Sub

Private Sub txtFilaDestino_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

'Converter o domínio numérico de Tipo de Saída para literais.
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

'Montar XML com layout do tipo de mensagem selecionado.
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
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "IN_OBRI_ATRB", _
                              xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "TP_DADO_ATRB_MESG", _
                              xmlNode.selectSingleNode("TP_DADO_ATRB_MESG").Text

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "QT_CTER_ATRB", _
                              xmlNode.selectSingleNode("QT_CTER_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "QT_CASA_DECI_ATRB", _
                              xmlNode.selectSingleNode("QT_CASA_DECI_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "QT_REPE", _
                               CLng(xmlNode.selectSingleNode("QT_REPE").Text)
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "TP_FORM_MESG", _
                              CLng(xmlNode.selectSingleNode("TP_FORM_MESG").Text)

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "NU_NIVE_MESG_ATRB", _
                              CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text)

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

    fgRaiseError App.EXEName, "frmRegraTransporte", "flMontarMensagem", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

