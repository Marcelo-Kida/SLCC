VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFiltro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   11580
   ClientLeft      =   4890
   ClientTop       =   3090
   ClientWidth     =   7305
   Icon            =   "frmFiltro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11580
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFiltro 
      Height          =   10935
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7215
      Begin MSComCtl2.DTPicker dtpHoraInicio 
         Height          =   315
         Left            =   3690
         TabIndex        =   67
         Top             =   10500
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211458
         CurrentDate     =   40822
      End
      Begin VB.ComboBox cboMensagemCAM 
         Height          =   315
         Left            =   2550
         TabIndex        =   64
         Text            =   "cboMensagemCAM"
         Top             =   10140
         Width           =   4515
      End
      Begin VB.ComboBox cboMensagemSPB 
         Height          =   315
         Left            =   2550
         TabIndex        =   62
         Text            =   "cboMensagemSPB"
         Top             =   9720
         Width           =   4515
      End
      Begin VB.TextBox txtIdentificadorPessoa 
         Height          =   315
         Left            =   2550
         MaxLength       =   8
         TabIndex        =   55
         Top             =   9345
         Width           =   4515
      End
      Begin VB.TextBox txtCodReemb 
         Height          =   315
         Left            =   2550
         MaxLength       =   20
         TabIndex        =   53
         Top             =   8580
         Width           =   4515
      End
      Begin VB.ComboBox cboCanalVenda 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   8205
         Width           =   1515
      End
      Begin VB.ComboBox cboTipoBackoffice 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   7800
         Width           =   4515
      End
      Begin VB.TextBox txtComando 
         Height          =   315
         Left            =   2550
         TabIndex        =   48
         Top             =   4200
         Width           =   4515
      End
      Begin VB.TextBox txtContaSelic 
         Height          =   315
         Left            =   2550
         MaxLength       =   9
         TabIndex        =   45
         Top             =   7080
         Width           =   4515
      End
      Begin VB.TextBox txtTipoTituloBMA 
         Height          =   315
         Left            =   2550
         MaxLength       =   3
         TabIndex        =   44
         Top             =   7440
         Width           =   4515
      End
      Begin VB.TextBox txtParticipacaoCETIP 
         Height          =   315
         Left            =   2550
         MaxLength       =   8
         TabIndex        =   42
         Top             =   6720
         Width           =   4515
      End
      Begin VB.TextBox txtVeiculoLegal 
         Height          =   315
         Left            =   2550
         MaxLength       =   15
         TabIndex        =   39
         Top             =   2040
         Width           =   1455
      End
      Begin VB.ComboBox cboSituacaoOperacao 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   5640
         Width           =   4515
      End
      Begin VB.ComboBox cboTipoLiquidacao 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   5280
         Width           =   4515
      End
      Begin VB.ComboBox cboTipoOperacao 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   4920
         Width           =   4515
      End
      Begin VB.TextBox txtContraParte 
         Height          =   315
         Left            =   2550
         TabIndex        =   28
         Top             =   3120
         Width           =   4515
      End
      Begin VB.ComboBox cboOperacaoEvento 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3840
         Width           =   4515
      End
      Begin VB.ComboBox cboCamara 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3480
         Width           =   4515
      End
      Begin VB.ComboBox cboAcoes 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   4560
         Width           =   4515
      End
      Begin VB.ComboBox cboItemCaixa 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2760
         Width           =   4515
      End
      Begin VB.ComboBox cboTipoCaixa 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2400
         Width           =   4515
      End
      Begin VB.ComboBox cboVeiculoLegal 
         Height          =   315
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   3015
      End
      Begin VB.ComboBox cboGrupoVeicLegal 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   4515
      End
      Begin VB.ComboBox cboSistema 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4515
      End
      Begin VB.ComboBox cboBancLiqu 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   4515
      End
      Begin VB.ComboBox cboLocalLiqu 
         Height          =   315
         Left            =   2550
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   4515
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   3690
         TabIndex        =   3
         Top             =   960
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   5430
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbData 
         Height          =   330
         Left            =   2550
         TabIndex        =   2
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1429
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpInicio2 
         Height          =   315
         Left            =   3750
         TabIndex        =   35
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComCtl2.DTPicker dtpFim2 
         Height          =   315
         Left            =   5430
         TabIndex        =   36
         Top             =   6000
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbData2 
         Height          =   330
         Left            =   2550
         TabIndex        =   37
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   6000
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1429
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin NumBox.Number numCodigoCNPJ 
         Height          =   315
         Left            =   2550
         TabIndex        =   47
         Top             =   6360
         Width           =   3630
         _ExtentX        =   6403
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
      Begin MSComCtl2.DTPicker dtpDataEventoCambioIni 
         Height          =   330
         Left            =   3750
         TabIndex        =   57
         Top             =   8985
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbDataEventoCambio 
         Height          =   330
         Left            =   2550
         TabIndex        =   58
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   8970
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1429
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpDataEventoCambioFim 
         Height          =   330
         Left            =   5430
         TabIndex        =   60
         Top             =   8985
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         Format          =   88211457
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbHora 
         Height          =   330
         Left            =   2550
         TabIndex        =   66
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   10500
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1429
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpHoraFim 
         Height          =   315
         Left            =   5430
         TabIndex        =   68
         Top             =   10500
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Format          =   88211458
         CurrentDate     =   40822
      End
      Begin VB.Label lblHora 
         Caption         =   "Horário"
         Height          =   195
         Left            =   180
         TabIndex        =   65
         Top             =   10560
         Width           =   1335
      End
      Begin VB.Label lblMensagemCAM 
         Caption         =   "Mensagem CAM"
         Height          =   195
         Left            =   180
         TabIndex        =   63
         Top             =   10200
         Width           =   1815
      End
      Begin VB.Label lblMensagemSPB 
         Caption         =   "Mensagem SPB"
         Height          =   195
         Left            =   180
         TabIndex        =   61
         Top             =   9780
         Width           =   1815
      End
      Begin VB.Label lblDataEventoCambio 
         AutoSize        =   -1  'True
         Caption         =   "Data Evento Câmbio"
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   9030
         Width           =   1470
      End
      Begin VB.Label lblIdentificadorPessoa 
         Caption         =   "Identificador Pessoa"
         Height          =   255
         Left            =   180
         TabIndex        =   56
         Top             =   9405
         Width           =   2265
      End
      Begin VB.Label lblCodReemb 
         Caption         =   "Cod. Reembolso CCR"
         Height          =   255
         Left            =   180
         TabIndex        =   54
         Top             =   8580
         Width           =   2175
      End
      Begin VB.Label lblCanalVenda 
         Caption         =   "Canal de Venda"
         Height          =   255
         Left            =   180
         TabIndex        =   51
         Top             =   8220
         Width           =   1455
      End
      Begin VB.Label lblTipoBackoffice 
         Caption         =   "Tipo de Backoffice"
         Height          =   255
         Left            =   180
         TabIndex        =   50
         Top             =   7860
         Width           =   1455
      End
      Begin VB.Label lblContaSelic 
         Caption         =   "Conta Própria Custódia SELIC"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   7140
         Width           =   2115
      End
      Begin VB.Label lblTipoTituloBMA 
         Caption         =   "Tipo Titular BMA"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   7500
         Width           =   1455
      End
      Begin VB.Label lblParticipacaoCETIP 
         Caption         =   "Identificador Participante CETIP"
         Height          =   255
         Left            =   180
         TabIndex        =   41
         Top             =   6780
         Width           =   2265
      End
      Begin VB.Label lblCodigoCNPJ 
         Caption         =   "CNPJ Veículo Legal"
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   6450
         Width           =   1455
      End
      Begin VB.Label lblData2 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   180
         TabIndex        =   38
         Top             =   6060
         Width           =   1335
      End
      Begin VB.Label lblSituacaoOperacao 
         Caption         =   "Situação Operação"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   5700
         Width           =   1455
      End
      Begin VB.Label lblTipoLiquidacao 
         Caption         =   "Tipo de Liquidação"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   5340
         Width           =   1455
      End
      Begin VB.Label lblTipoOperacao 
         Caption         =   "Tipo de Operação "
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   4980
         Width           =   1455
      End
      Begin VB.Label lblContraParte 
         AutoSize        =   -1  'True
         Caption         =   "Contraparte"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   3180
         Width           =   855
      End
      Begin VB.Label lblOperacaoEvento 
         AutoSize        =   -1  'True
         Caption         =   "Operação/Evento"
         Height          =   195
         Left            =   180
         TabIndex        =   26
         Top             =   3900
         Width           =   1335
      End
      Begin VB.Label lblCamara 
         AutoSize        =   -1  'True
         Caption         =   "Câmara"
         Height          =   195
         Left            =   180
         TabIndex        =   25
         Top             =   3540
         Width           =   555
      End
      Begin VB.Label lblNumeroComando 
         AutoSize        =   -1  'True
         Caption         =   "Número Comando"
         Height          =   195
         Left            =   180
         TabIndex        =   24
         Top             =   4260
         Width           =   1275
      End
      Begin VB.Label lblAcoes 
         AutoSize        =   -1  'True
         Caption         =   "Ações"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   4620
         Width           =   495
      End
      Begin VB.Label lblItemCaixa 
         AutoSize        =   -1  'True
         Caption         =   "Item de Caixa"
         Height          =   195
         Left            =   180
         TabIndex        =   22
         Top             =   2820
         Width           =   975
      End
      Begin VB.Label lblTipoCaixa 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Caixa"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   2460
         Width           =   975
      End
      Begin VB.Label lblVeiculoLegal 
         AutoSize        =   -1  'True
         Caption         =   "Veículo Legal"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   2100
         Width           =   1035
      End
      Begin VB.Label lblGrupoVeicLegal 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Veículo Legal"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   1740
         Width           =   1515
      End
      Begin VB.Label lblLocalLiqu 
         AutoSize        =   -1  'True
         Caption         =   "Local Liquidação"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblBancLiqu 
         AutoSize        =   -1  'True
         Caption         =   "Empresa"
         Height          =   195
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Width           =   615
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label lblSistema 
         AutoSize        =   -1  'True
         Caption         =   "Sistema"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   660
         Width           =   555
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   3960
      Top             =   10995
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltro.frx":000C
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltro.frx":045E
            Key             =   "Aplicar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   4650
      TabIndex        =   13
      Top             =   11115
      Width           =   2550
      _ExtentX        =   4498
      _ExtentY        =   582
      ButtonWidth     =   2196
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aplicar"
            Key             =   "aplicar"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   2
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Filtro"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:06
'-------------------------------------------------
'' Objeto responsável pelo refinamento (filtro) das telas de consultas de dados
'' existentes no sistema.
Option Explicit

Public Event AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

Public TipoFiltro                           As enumTipoFiltroA8
Public FormOwner                            As Form

Private Const PROP_TOP_CAMPO_01             As Integer = 240
Private Const PROP_TOP_CAMPO_02             As Integer = 600
Private Const PROP_TOP_CAMPO_03             As Integer = 960
Private Const PROP_TOP_CAMPO_04             As Integer = 1320
Private Const PROP_TOP_CAMPO_05             As Integer = 1680
Private Const PROP_TOP_CAMPO_06             As Integer = 2040
Private Const PROP_TOP_CAMPO_07             As Integer = 2400
Private Const PROP_TOP_CAMPO_08             As Integer = 2760
Private Const PROP_TOP_CAMPO_09             As Integer = 3120
Private Const PROP_TOP_CAMPO_10             As Integer = 3480
Private Const PROP_TOP_CAMPO_11             As Integer = 3840
Private Const PROP_TOP_CAMPO_12             As Integer = 4200
Private Const PROP_TOP_CAMPO_13             As Integer = 4560
Private Const PROP_TOP_CAMPO_14             As Integer = 4920
Private Const PROP_TOP_CAMPO_15             As Integer = 5280
Private Const PROP_TOP_CAMPO_16             As Integer = 5640
Private Const PROP_TOP_CAMPO_17             As Integer = 6000
Private Const PROP_TOP_CAMPO_18             As Integer = 6360
Private Const PROP_TOP_CAMPO_19             As Integer = 6720
Private Const PROP_TOP_CAMPO_20             As Integer = 7080
Private Const PROP_TOP_CAMPO_21             As Integer = 7440
Private Const PROP_TOP_CAMPO_22             As Integer = 7800
Private Const PROP_TOP_CAMPO_23             As Integer = 8205
Private Const PROP_TOP_CAMPO_24             As Integer = 8645


Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlPropriedades                     As MSXML2.DOMDocument40
Private xmlDOMRegistro                      As MSXML2.DOMDocument40
Private lcControlesFiltro                   As New Collection

Private strFuncionalidade                   As String
Private arrSgSistCoVeicLega()               As Variant

Private blnFiltraHora                       As Boolean
Private blnFiltraHora2                      As Boolean
Private blnFiltraHora3                      As Boolean
Private blnFiltraHorario                    As Boolean
Private blnPrimeiroActivate                 As Boolean
Private blnOcorreuErro                      As Boolean

'Carregar dinamicamente apenas os combos utilizados no filtro

Private Sub flCarregarCombosUtilizados()

    If cboBancLiqu.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboBancLiqu, gxmlCombosFiltro, "Empresa", "CO_EMPR", "NO_REDU_EMPR", True
    End If
    
    If cboGrupoVeicLegal.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboGrupoVeicLegal, gxmlCombosFiltro, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA", True
    End If
    
    If cboLocalLiqu.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboLocalLiqu, gxmlCombosFiltro, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU", True
    End If
    
    If cboTipoOperacao.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboTipoOperacao, gxmlCombosFiltro, "TipoOperacao", "TP_OPER", "NO_TIPO_OPER", True
    End If
    
    If cboTipoLiquidacao.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboTipoLiquidacao, gxmlCombosFiltro, "TipoLiquidacao", "TP_LIQU_OPER_ATIV", "NO_TIPO_LIQU_OPER_ATIV", True
    End If
    
    If cboSituacaoOperacao.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboSituacaoOperacao, gxmlCombosFiltro, "SituacaoProcesso", "CO_SITU_PROC", "DE_SITU_PROC", True
    End If

    If cboTipoBackoffice.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboTipoBackoffice, gxmlCombosFiltro, "TipoBackOffice", "TP_BKOF", "DE_BKOF", True
    End If
    
    If cboMensagemSPB.Top < PROP_TOP_CAMPO_23 Then
        fgCarregarCombos cboMensagemSPB, gxmlCombosFiltro, "MensagemSPB", "CO_MESG_SPB", "NO_MESG", True
    End If

    If cboAcoes.Top < PROP_TOP_CAMPO_23 Then
        Call flPreenchecboAcao
    End If

    If cboCanalVenda.Top < PROP_TOP_CAMPO_23 Then
        Call flPreenchecboCanalVenda
    End If
    
    If txtCodReemb.Top < PROP_TOP_CAMPO_23 Then
        Call flPreenchecboCanalVenda
    End If
    
    If cboMensagemCAM.Top < PROP_TOP_CAMPO_23 Then
        Call flPreencheMensagemCAM
    End If

    Exit Sub

ErrorHandler:
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name & " - flCarregarCombosUtilizados", Me.Caption)

End Sub

'Carregar a pesquisa anterior

Public Sub fgCarregarPesquisaAnterior()

On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml <> vbNullString Then
        Call flAplicarFiltro(False)
    Else
        Call flAplicarFiltro(True)
    End If

Exit Sub
ErrorHandler:
    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - flCarregarPesquisaAnterior", Me.Caption)
End Sub

'Converter a data e hora para seleção no oracle

Private Function flConverteDataHoraOracle(ByVal strDataHora As String) As String

Dim strAno                                  As String
Dim strMes                                  As String
Dim strDia                                  As String
Dim strHora                                 As String
Dim strMinuto                               As String
Dim strSegundo                              As String
Dim strDataHoraConvertida                   As String

On Error GoTo ErrorHandler

    strAno = Format(Year(strDataHora), "0000")
    strMes = Format(Month(strDataHora), "00")
    strDia = Format(Day(strDataHora), "00")
    strHora = Format(Hour(strDataHora), "00")
    strMinuto = Format(Minute(strDataHora), "00")
    strSegundo = Format(Second(strDataHora), "00")
    
    strDataHoraConvertida = strAno & strMes & strDia & strHora & strMinuto & strSegundo
    
    flConverteDataHoraOracle = "TO_DATE('" & strDataHoraConvertida & "','YYYYMMDDHH24MISS')"

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConverteDataHoraOracle", 0

End Function

'Converter a data para seleção no oracle

Private Function flConverteDataOracle(ByVal strData As String) As String

Dim strAno                                  As String
Dim strMes                                  As String
Dim strDia                                  As String
Dim strDataConvertida                       As String

On Error GoTo ErrorHandler

    strAno = Format(Year(strData), "0000")
    strMes = Format(Month(strData), "00")
    strDia = Format(Day(strData), "00")
    
    strDataConvertida = strAno & strMes & strDia
    
    flConverteDataOracle = "TO_DATE('" & strDataConvertida & "','YYYYMMDD')"

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConverteDataOracle", 0

End Function

'Executar o evento do form chamador que irá fazer solicitar a execução do filtro

Private Sub flAplicarFiltro(ByVal pblnNovaPesquisa As Boolean)

Dim strTituloTableCombo                     As String
Dim strDocFiltros                           As String
Dim xmlDomFiltros                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml = vbNullString Or pblnNovaPesquisa Then
        Set xmlDomFiltros = flInterfaceToXml(strTituloTableCombo)
        strDocFiltros = xmlDomFiltros.xml
        Call flGravarSettingsRegistry(xmlDomFiltros, strTituloTableCombo)
    Else
        strDocFiltros = xmlDOMRegistro.selectSingleNode("//Registry/Repeat_Filtros").xml
        strTituloTableCombo = xmlDOMRegistro.selectSingleNode("//Registry/TituloTableCombo").Text
    End If

    Me.Hide
    DoEvents
    
    RaiseEvent AplicarFiltro(strDocFiltros, strTituloTableCombo)
                                           
    Set xmlDomFiltros = Nothing
    
Exit Sub
ErrorHandler:
    Set xmlDomFiltros = Nothing
   
    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - flAplicarFiltro", Me.Caption)

End Sub

'Carregar a interface com o filtro aplicado anteriormente

Private Function flInterfaceToXml(ByRef strTituloTableCombo As String) As MSXML2.DOMDocument40

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim xmlDomControle                          As MSXML2.DOMDocument40

Dim strDocFiltros                           As String
Dim intIndCombo                             As Integer

On Error GoTo ErrorHandler
    
    strTituloTableCombo = vbNullString
            
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    Set xmlDomControle = CreateObject("MSXML2.DOMDocument.4.0")
    
    Select Case TipoFiltro
        Case enumTipoFiltroA8.frmConsultaOperacao
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If Trim$(txtContraParte.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Contraparte", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Contraparte", _
                                                 "Contraparte", Trim$(fgLimpaCaracterInvalido(txtContraParte.Text)))
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", _
                                                 "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", _
                                                 "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
                        
            If cboSituacaoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SituacaoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_SituacaoOperacao", _
                                                 "SituacaoOperacao", fgObterCodigoCombo(Me.cboSituacaoOperacao.Text))
            End If
                        
            If Trim$(txtComando.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroComando", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroComando", _
                                                 "NumeroComando", Trim$(fgLimpaCaracterInvalido(txtComando.Text)))
            End If

            If cboAcoes.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Acoes", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Acoes", _
                                                 "Acoes", fgObterCodigoCombo(Me.cboAcoes.Text))
            End If
                       
            If cboCanalVenda.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboCanalVenda.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
                       
            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
            
            If Trim(txtCodReemb) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_NumeroControleLTR", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_NumeroControleLTR", "NumeroControleLTR", Trim(txtCodReemb))
            End If
            
            
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
            End If
        
        Case enumTipoFiltroA8.frmConsultaMensagemSisbacen
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If cboMensagemCAM.ListIndex > -1 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CodigoMensagem", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CodigoMensagem", _
                                                 "CodigoMensagem", Left(Me.cboMensagemCAM.Text, 7))
            End If
            
            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.value & " 00:00:00"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataHoraOracle(dtpFim.value & " 23:59:59"))
                       
                End Select
            
            End If
            
            If blnFiltraHora3 Then
                Select Case tlbDataEventoCambio.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataEventoCambio", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataIni", flConverteDataOracle(dtpDataEventoCambioIni.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataEventoCambio", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataFim", flConverteDataOracle(dtpDataEventoCambioIni.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataEventoCambio", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataIni", flConverteDataHoraOracle(dtpDataEventoCambioIni.value & " 00:00:00"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataEventoCambio", _
                                                             "DataFim", flConverteDataHoraOracle(dtpDataEventoCambioFim.value & " 23:59:59"))
                       
                End Select
                
            End If
            
            If txtIdentificadorPessoa.Text <> "" Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_IdentificadorPessoa", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_IdentificadorPessoa", _
                                                 "IdentificadorPessoa", Me.txtIdentificadorPessoa.Text)
            End If
            
            
        Case enumTipoFiltroA8.frmConsultaMensagem
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", _
                                                 "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
            
            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
            
            If cboMensagemSPB.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_MensagemSPB", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_MensagemSPB", "MensagemSPB", fgObterCodigoCombo(Me.cboMensagemSPB.Text))
            End If
            
            'RATS 1176 - Considerar o sistema no filtro de mensagens - 05/07/2012
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If
                       
                       
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            If blnFiltraHorario Then
                            
                                Select Case tlbHora.Buttons(1).Caption
                                
                                    Case "Após"
                                       
                                        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                         "DataIni", flConverteDataHoraOracle(dtpInicio.value & " " & dtpHoraInicio.value))
                                        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                         "DataFim", flConverteDataHoraOracle(dtpFim.value & " 23:59:59"))
                                       
                                    Case "Antes"
                                       
                                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.value & " 00:00:00"))
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataFim", flConverteDataHoraOracle(dtpFim.value & " " & dtpHoraInicio.value))
                                       
                                    Case "Entre"
                                       
                                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.value & " " & dtpHoraInicio.value))
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataFim", flConverteDataHoraOracle(dtpFim.value & " " & dtpHoraFim.value))
                                       
                                End Select
                            
                            Else
                                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                 "DataIni", flConverteDataOracle(dtpInicio.value))
                                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                 "DataFim", flConverteDataOracle("31/12/9999"))
                            End If
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            If blnFiltraHorario Then
                            
                                Select Case tlbHora.Buttons(1).Caption
                                
                                    Case "Após"
                                       
                                        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                         "DataIni", flConverteDataHoraOracle(dtpInicio.value & " " & dtpHoraInicio.value))
                                        Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                         "DataFim", flConverteDataHoraOracle(dtpFim.value & " 23:59:59"))
                                       
                                    Case "Antes"
                                       
                                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.value & " 00:00:00"))
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataFim", flConverteDataHoraOracle(dtpFim.value & " " & dtpHoraInicio.value))
                                       
                                    Case "Entre"
                                       
                                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.value & " " & dtpHoraInicio.value))
                                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                             "DataFim", flConverteDataHoraOracle(dtpFim.value & " " & dtpHoraFim.value))
                                       
                                End Select
                                
                            Else
                                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                                
                                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                 "DataIni", flConverteDataHoraOracle(dtpInicio.value & " 00:00:00"))
                                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                                 "DataFim", flConverteDataHoraOracle(dtpFim.value & " 23:59:59"))
                            End If
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmRemessaRejeitada
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            
            End If
            
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If
        
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmConfirmacaoOperacao
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If Trim$(txtContraParte.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Contraparte", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Contraparte", "Contraparte", txtContraParte.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
            
            If cboTipoLiquidacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", fgObterCodigoCombo(Me.cboTipoLiquidacao.Text))
            End If

            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If

            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmLiberacaoOperacaoMensagem
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
            
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If Trim$(txtContraParte.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Contraparte", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Contraparte", "Contraparte", txtContraParte.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
            
            If cboTipoLiquidacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoLiquidacao", "TipoLiquidacao", fgObterCodigoCombo(Me.cboTipoLiquidacao.Text))
            End If

            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If

            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
            If blnFiltraHora2 Then
                
                Select Case tlbData2.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataLiquidacao", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataIni", flConverteDataOracle(dtpInicio2.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataLiquidacao", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataFim", flConverteDataOracle(dtpInicio2.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_DataLiquidacao", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataIni", flConverteDataOracle(dtpInicio2.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_DataLiquidacao", _
                                                             "DataFim", flConverteDataOracle(dtpFim2.value))
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmSuspenderDisponibilizarLancamentoCC
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
            
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", "VeiculoLegal", txtVeiculoLegal.Text)
            
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(0))
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
            
            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
            
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmIntegrarCCOnLine
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
            
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
            
            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
            
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
        Case enumTipoFiltroA8.frmIntegrarCCOnLineEstorno
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
            
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
            
            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
            
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            
                            'Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataHoraOracle(dtpInicio.MinDate))
                                                             
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
                
            End If
            
            
        Case enumTipoFiltroA8.frmReenvioCancelamentoEstornoMsg
                       
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(dtpInicio.value)))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", fgDtXML_To_Oracle(fgDt_To_Xml(dtpFim.MaxDate)))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(dtpInicio.MinDate)))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", fgDtXML_To_Oracle(fgDt_To_Xml(dtpInicio.value)))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", fgDtXML_To_Oracle(fgDt_To_Xml(dtpInicio.value)))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", fgDtXML_To_Oracle(fgDt_To_Xml(dtpFim.value)))
                       
                End Select
                
            Else
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                
                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                 "DataIni", flConverteDataOracle(fgDataHoraServidor(enumFormatoDataHora.Data)))
                Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                 "DataFim", flConverteDataOracle(fgDataHoraServidor(enumFormatoDataHora.Data)))
            End If
            
        Case enumTipoFiltroA8.frmConsultaContaCorrente
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If
            
            
            If cboGrupoVeicLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", _
                                                 "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If cboTipoOperacao.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoOperacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoOperacao", _
                                                 "TipoOperacao", fgObterCodigoCombo(Me.cboTipoOperacao.Text))
            End If
                       
            If cboCanalVenda.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CanalVenda", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CanalVenda", _
                                                 "CanalVenda", fgObterCodigoCombo(Me.cboCanalVenda.Text))
            End If
                       
            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
            If blnFiltraHora Then
                
                Select Case tlbData.Buttons(1).Caption
                
                       Case "Após"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       
                       Case "Antes"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       
                       Case "Entre"
                       
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                       
                End Select
            End If
            
        Case enumTipoFiltroA8.frmConsultaVeiculoLegal
            
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If

            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            
            End If
                
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            If numCodigoCNPJ.Valor <> 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_CodigoCNPJ", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_CodigoCNPJ", _
                                                 "CodigoCNPJ", fgVlr_To_Xml(Me.numCodigoCNPJ.Valor))
            End If
            
            If Trim$(txtParticipacaoCETIP.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_IdParticipacaoCETIP", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_IdParticipacaoCETIP", _
                                                 "IdParticipacaoCETIP", txtParticipacaoCETIP.Text)
            End If
            
            If Trim$(txtContaSelic.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_ContaPadraoSELIC", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_ContaPadraoSELIC", _
                                                 "ContaPadraoSELIC", txtContaSelic.Text)
            End If
            
            If Trim$(txtTipoTituloBMA.Text) <> vbNullString Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_TipoTituloBMA", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_TipoTituloBMA", _
                                                 "TipoTituloBMA", txtTipoTituloBMA.Text)
            End If

            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
       'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação - Adrian - 20/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacao
            
            'Empresa
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            'Data
            If blnFiltraHora Then
                Select Case tlbData.Buttons(1).Caption
                       Case "Após"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       Case "Antes"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       Case "Entre"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                End Select
            End If

            'Sistema
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If

            'Grupo Veículo Legal
            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            'Veículo Legal
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            'Local de Liquidação
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", _
                                                 "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If
            
            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
        'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Operações Rejeitadas - Adrian - 23/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacaoRejeitada
            
            'Empresa
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            'Data
            If blnFiltraHora Then
                Select Case tlbData.Buttons(1).Caption
                       Case "Após"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       Case "Antes"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       Case "Entre"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                End Select
            End If

            'Sistema
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If

        'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Conta Corrente e Contabilidade - Adrian - 24/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacaoCC_HA
            
            'Empresa
            If cboBancLiqu.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboBancLiqu.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BancoLiquidante", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_BancoLiquidante", _
                                                 "BancoLiquidante", fgObterCodigoCombo(Me.cboBancLiqu.Text))
            End If
            
            'Data
            If blnFiltraHora Then
                Select Case tlbData.Buttons(1).Caption
                       Case "Após"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle("31/12/9999"))
                       Case "Antes"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle("01/01/1900"))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpInicio.value))
                       Case "Entre"
                            Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Data", "")
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataIni", flConverteDataOracle(dtpInicio.value))
                            Call fgAppendNode(xmlDomFiltros, "Grupo_Data", _
                                                             "DataFim", flConverteDataOracle(dtpFim.value))
                End Select
            End If

            'Sistema
            If cboSistema.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Sistema", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_Sistema", _
                                                 "Sistema", fgObterCodigoCombo(Me.cboSistema.Text))
            End If

            'Grupo Veículo Legal
            If cboGrupoVeicLegal.ListIndex > 0 Then
                strTituloTableCombo = fgObterDescricaoCombo(Me.cboGrupoVeicLegal.Text)
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_GrupoVeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_GrupoVeiculoLegal", _
                                                 "GrupoVeiculoLegal", fgObterCodigoCombo(Me.cboGrupoVeicLegal.Text))
            End If
                
            'Veículo Legal
            If cboVeiculoLegal.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_VeiculoLegal", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_VeiculoLegal", _
                                                 "VeiculoLegal", txtVeiculoLegal.Text)
            End If
            
            'Local de Liquidação
            If cboLocalLiqu.ListIndex > 0 Then
                Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_LocalLiquidacao", "")
                Call fgAppendNode(xmlDomFiltros, "Grupo_LocalLiquidacao", _
                                                 "LocalLiquidacao", fgObterCodigoCombo(Me.cboLocalLiqu.Text))
            End If

            If gblnExibirTipoBackOffice Then
                If cboTipoBackoffice.ListIndex >= 0 Then
                    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_BackOfficePerfilGeral", "")
                    Call fgAppendNode(xmlDomFiltros, "Grupo_BackOfficePerfilGeral", _
                                                     "BackOfficePerfilGeral", IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(Me.cboTipoBackoffice.Text)))
                End If
            End If
                       
    End Select
                                           
    Set flInterfaceToXml = xmlDomFiltros
                                           
    Set xmlDomFiltros = Nothing
    Set xmlDomControle = Nothing
    
Exit Function
ErrorHandler:
    Set xmlDomFiltros = Nothing
    Set xmlDomControle = Nothing
    
    fgRaiseError App.EXEName, "frmFiltro", "flInterfaceToXml", 0

End Function

'Aplicar o filtro armazeando no Settings do Registry
Private Sub flAplicarSettingsRegistry()
    
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strRegistry                             As String
Dim intIndControles                         As Integer
    
On Error GoTo ErrorHandler
    
    If xmlDOMRegistro.xml = vbNullString Then Exit Sub
    
    intIndControles = 0
    For Each objDomNode In xmlDOMRegistro.documentElement.selectNodes("//Registry/Grupo_ControleFiltro/*")
        
        intIndControles = intIndControles + 1
        If intIndControles > lcControlesFiltro.Count Then Exit For
        
        If TypeName(lcControlesFiltro(intIndControles)) = "ComboBox" Then
            If lcControlesFiltro(intIndControles).ListCount > Val("0" & objDomNode.Text) Then
                lcControlesFiltro(intIndControles).ListIndex = IIf(objDomNode.Text = vbNullString Or Val(objDomNode.Text) < 0, -1, objDomNode.Text)
            Else
                lcControlesFiltro(intIndControles).ListIndex = -1
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "TextBox" Then
            lcControlesFiltro(intIndControles).Text = objDomNode.Text
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "DTPicker" Then
            If InStr(1, CStr(lcControlesFiltro(intIndControles).Name), "Hora") > 0 Then
                If objDomNode.Text = "-1" Then
                    objDomNode.Text = ""
                    If lcControlesFiltro(intIndControles).Name = "dtpHoraFim" Then
                        dtpHoraFim.Visible = False
                    End If
                End If
                lcControlesFiltro(intIndControles).value = fgValidarMaxDateDTPicker(lcControlesFiltro(intIndControles), fgValidarMinDateDTPicker(lcControlesFiltro(intIndControles), fgHrStr_To_Time(objDomNode.Text)))
            Else
                lcControlesFiltro(intIndControles).value = fgValidarMaxDateDTPicker(lcControlesFiltro(intIndControles), fgValidarMinDateDTPicker(lcControlesFiltro(intIndControles), fgDtXML_To_Date(objDomNode.Text)))
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "Toolbar" Then
            If InStr(1, objDomNode.Text, "|") = 0 Then
                lcControlesFiltro(intIndControles).Buttons(1).Image = 0
            Else
                lcControlesFiltro(intIndControles).Buttons(1).Image = Val(Split(objDomNode.Text, "|")(1))
            End If
            
            If lcControlesFiltro(intIndControles).Name = "tlbData" Then
                blnFiltraHora = lcControlesFiltro(intIndControles).Buttons(1).Image <> 0
                flConfiguratlbData Split(objDomNode.Text, "|")(0), tlbData, dtpFim
            ElseIf lcControlesFiltro(intIndControles).Name = "tlbData2" Then
                blnFiltraHora2 = lcControlesFiltro(intIndControles).Buttons(1).Image <> 0
                flConfiguratlbData Split(objDomNode.Text, "|")(0), tlbData2, dtpFim2
            ElseIf lcControlesFiltro(intIndControles).Name = "tlbHora" Then
                If tlbData.Buttons(1).Caption = "Após" Then
                    If dtpInicio.value = fgDataHoraServidor(Data) Then
                        tlbHora.Enabled = True
                        dtpHoraInicio.Enabled = True
                        dtpHoraFim.Enabled = True
                    Else
                        tlbHora.Enabled = False
                        dtpHoraInicio.Enabled = False
                        dtpHoraFim.Enabled = False
                    End If
                ElseIf tlbData.Buttons(1).Caption = "Entre" Then
                    If dtpInicio.value = dtpFim.value Then
                        tlbHora.Enabled = True
                        dtpHoraInicio.Enabled = True
                        dtpHoraFim.Enabled = True
                    Else
                        tlbHora.Enabled = False
                        dtpHoraInicio.Enabled = False
                        dtpHoraFim.Enabled = False
                    End If
                Else
                    tlbHora.Enabled = False
                    dtpHoraInicio.Enabled = False
                    dtpHoraFim.Enabled = False
                End If
                blnFiltraHorario = lcControlesFiltro(intIndControles).Buttons(1).Image <> 0
                flConfiguratlbData Split(objDomNode.Text, "|")(0), tlbHora, dtpHoraFim
            Else
                blnFiltraHora3 = lcControlesFiltro(intIndControles).Buttons(1).Image <> 0
                flConfiguratlbData Split(objDomNode.Text, "|")(0), tlbDataEventoCambio, dtpDataEventoCambioFim
            End If
        ElseIf TypeName(lcControlesFiltro(intIndControles)) = "Number" Then
            lcControlesFiltro(intIndControles).Valor = fgVlrXml_To_Decimal(objDomNode.Text)
        End If
    
    Next
    
    Set objDomNode = Nothing

Exit Sub
ErrorHandler:

    Set objDomNode = Nothing
    
    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - flAplicarSettingsRegistry", Me.Caption)

End Sub

'Gravar o filtro aplicado no Settings do Registry
Private Sub flGravarSettingsRegistry(ByVal objDOMFiltro As MSXML2.DOMDocument40, _
                                     ByVal strTituloTableCombo As String)

Dim objDomGrupoControle                     As MSXML2.DOMDocument40
Dim objControleFiltro                       As Object
    
On Error GoTo ErrorHandler
    
    Set xmlDOMRegistro = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDOMRegistro, "", "Registry", "")

    Set objDomGrupoControle = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(objDomGrupoControle, "", "Grupo_ControleFiltro", "")

    For Each objControleFiltro In lcControlesFiltro
        If TypeName(objControleFiltro) = "ComboBox" Then
            'If objControleFiltro.Name <> cboTipoBackoffice.Name Then
                Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.ListIndex)
            'End If
        ElseIf TypeName(objControleFiltro) = "TextBox" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Text)
        ElseIf TypeName(objControleFiltro) = "DTPicker" Then
            If InStr(1, CStr(objControleFiltro.Name), "Hora") > 0 Then
                Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", fgHr_To_Xml(objControleFiltro.value))
            Else
                Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", fgDt_To_Xml(objControleFiltro.value))
            End If
        ElseIf TypeName(objControleFiltro) = "Toolbar" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", objControleFiltro.Buttons(1).Caption & "|" & objControleFiltro.Buttons(1).Image)
        ElseIf TypeName(objControleFiltro) = "Number" Then
            Call fgAppendNode(objDomGrupoControle, "Grupo_ControleFiltro", "ConteudoControle", fgVlr_To_Xml(objControleFiltro.Valor))
        End If
    Next
    
    Call fgAppendXML(xmlDOMRegistro, "Registry", objDomGrupoControle.xml)
    Call fgAppendXML(xmlDOMRegistro, "Registry", objDOMFiltro.xml)
    Call fgAppendNode(xmlDOMRegistro, "Registry", "TituloTableCombo", strTituloTableCombo)
    
    Set objDomGrupoControle = Nothing
    
    Call SaveSetting("A8LQS", "Form Filtro\" & FormOwner.Name, "Settings", xmlDOMRegistro.xml)
    
Exit Sub
ErrorHandler:
    Set objDomGrupoControle = Nothing

    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - flGravarSettingsRegistry", Me.Caption)

End Sub

'Configurar o filtro conforme o form chamador

Private Sub flConfiguraLayoutForm()
 
Dim colControles                            As New Collection
Dim objControle                             As Object

On Error GoTo ErrorHandler

    With Me
        
        Set colControles = New Collection
        colControles.Add .lblBancLiqu
        colControles.Add .lblSistema
        colControles.Add .lblData
        colControles.Add .lblLocalLiqu
        colControles.Add .lblGrupoVeicLegal
        colControles.Add .lblVeiculoLegal
        colControles.Add .lblTipoCaixa
        colControles.Add .lblItemCaixa
        colControles.Add .lblContraParte
        colControles.Add .lblCamara
        colControles.Add .lblOperacaoEvento
        colControles.Add .lblNumeroComando
        colControles.Add .lblAcoes
        colControles.Add .lblTipoOperacao
        colControles.Add .lblTipoLiquidacao
        colControles.Add .lblSituacaoOperacao
        colControles.Add .lblData2
        colControles.Add .lblCodigoCNPJ
        colControles.Add .lblParticipacaoCETIP
        colControles.Add .lblContaSelic
        colControles.Add .lblTipoTituloBMA
        colControles.Add .lblTipoBackoffice
        colControles.Add .lblCanalVenda
        colControles.Add .lblCodReemb
        colControles.Add .lblSistema
        colControles.Add .lblMensagemSPB
        colControles.Add .lblMensagemCAM
        colControles.Add .lblHora
                
        For Each objControle In colControles
            objControle.Top = PROP_TOP_CAMPO_24 * 3
        Next
        
        Select Case TipoFiltro
            Case enumTipoFiltroA8.frmConsultaOperacao
            
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblContraParte.Top = PROP_TOP_CAMPO_06
                .lblNumeroComando.Top = PROP_TOP_CAMPO_07
                .lblAcoes.Top = PROP_TOP_CAMPO_08
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_09
                .lblSituacaoOperacao.Top = PROP_TOP_CAMPO_10
                .lblCanalVenda.Top = PROP_TOP_CAMPO_11
                
                .lblCodReemb.Top = PROP_TOP_CAMPO_12
                
                If gblnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_13
                    .fraFiltro.Height = PROP_TOP_CAMPO_14
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_13
                End If
                
            'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Operações - Adrian - 20/06/05
            'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Operações Rejeitadas - Adrian - 23/06/05
            'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Conta Corrente e Contabilidade - Adrian - 24/06/05
            Case enumTipoFiltroA8.frmConsultaMovimentacao, _
                 enumTipoFiltroA8.frmConsultaMovimentacaoCC_HA
            
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblSistema.Top = PROP_TOP_CAMPO_02
                .lblData.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                
                If gblnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_06
                    .fraFiltro.Height = PROP_TOP_CAMPO_07
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_06
                End If
                
            Case enumTipoFiltroA8.frmConsultaMensagem
            
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblSistema.Top = PROP_TOP_CAMPO_02
                .lblData.Top = PROP_TOP_CAMPO_03
                .lblHora.Top = PROP_TOP_CAMPO_04
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_05
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_06
                .lblCanalVenda.Top = PROP_TOP_CAMPO_07
                
                If gblnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_08
                    .lblMensagemSPB.Top = PROP_TOP_CAMPO_09
                    .fraFiltro.Height = PROP_TOP_CAMPO_10
                Else
                    .lblMensagemSPB.Top = PROP_TOP_CAMPO_08
                    .fraFiltro.Height = PROP_TOP_CAMPO_09
                End If
                
            Case enumTipoFiltroA8.frmConsultaMensagemSisbacen
            
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblMensagemCAM.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblDataEventoCambio.Top = PROP_TOP_CAMPO_06
                .lblIdentificadorPessoa.Top = PROP_TOP_CAMPO_07
                
                If gblnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_08
                    .fraFiltro.Height = PROP_TOP_CAMPO_09
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_08
                End If
                
            Case enumTipoFiltroA8.frmRemessaRejeitada, _
                 enumTipoFiltroA8.frmConsultaMovimentacaoRejeitada
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblSistema.Top = PROP_TOP_CAMPO_02
                .lblData.Top = PROP_TOP_CAMPO_03
                
                .fraFiltro.Height = PROP_TOP_CAMPO_04
                
            Case enumTipoFiltroA8.frmConfirmacaoOperacao
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblContraParte.Top = PROP_TOP_CAMPO_06
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_07
                .lblTipoLiquidacao.Top = PROP_TOP_CAMPO_08
                .lblCanalVenda.Top = PROP_TOP_CAMPO_09
                
                .fraFiltro.Height = PROP_TOP_CAMPO_10
                
                dtpInicio.MinDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 3, enumPaginacao.Anterior)
    
            Case enumTipoFiltroA8.frmLiberacaoOperacaoMensagem
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblContraParte.Top = PROP_TOP_CAMPO_06
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_07
                .lblTipoLiquidacao.Top = PROP_TOP_CAMPO_08
                .lblData2.Top = PROP_TOP_CAMPO_09
                .lblCanalVenda.Top = PROP_TOP_CAMPO_10
                
                .fraFiltro.Height = PROP_TOP_CAMPO_11
                
                .lblData2.Caption = "Data Liquidação"
                
                dtpInicio.MinDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 3, enumPaginacao.Anterior)
                dtpInicio2.MinDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 3, enumPaginacao.Anterior)
    
            Case enumTipoFiltroA8.frmSuspenderDisponibilizarLancamentoCC, _
                 enumTipoFiltroA8.frmIntegrarCCOnLineEstorno
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_06
                .lblCanalVenda.Top = PROP_TOP_CAMPO_07
                
                .fraFiltro.Height = PROP_TOP_CAMPO_08
                
            Case enumTipoFiltroA8.frmIntegrarCCOnLine
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_06
                .lblCanalVenda.Top = PROP_TOP_CAMPO_07
                
                .fraFiltro.Height = PROP_TOP_CAMPO_08
                
                dtpInicio.MinDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 5, enumPaginacao.Anterior)
                dtpInicio.MaxDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 0, enumPaginacao.proximo)
                dtpFim.MinDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 0, enumPaginacao.Anterior)
                dtpFim.MaxDate = fgAdicionarDiasUteis(fgDataHoraServidor(enumFormatoDataHora.Data), 0, enumPaginacao.proximo)
                
            Case enumTipoFiltroA8.frmReenvioCancelamentoEstornoMsg
            
                .lblData.Top = PROP_TOP_CAMPO_01
                
                .fraFiltro.Height = PROP_TOP_CAMPO_02
                
            Case enumTipoFiltroA8.frmConsultaContaCorrente
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblData.Top = PROP_TOP_CAMPO_02
                .lblLocalLiqu.Top = PROP_TOP_CAMPO_03
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_04
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_05
                .lblTipoOperacao.Top = PROP_TOP_CAMPO_06
                .lblCanalVenda.Top = PROP_TOP_CAMPO_07
                
                If gblnExibirTipoBackOffice Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_08
                    .fraFiltro.Height = PROP_TOP_CAMPO_09
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_08
                End If
                
            Case enumTipoFiltroA8.frmConsultaVeiculoLegal
                
                .lblBancLiqu.Top = PROP_TOP_CAMPO_01
                .lblSistema.Top = PROP_TOP_CAMPO_02
                .lblGrupoVeicLegal.Top = PROP_TOP_CAMPO_03
                .lblVeiculoLegal.Top = PROP_TOP_CAMPO_04
                
                If gblnExibirTipoBackOffice And Not FormOwner.fraCadastro.Visible Then
                    .lblTipoBackoffice.Top = PROP_TOP_CAMPO_05
                    .fraFiltro.Height = PROP_TOP_CAMPO_06
                Else
                    .fraFiltro.Height = PROP_TOP_CAMPO_05
                End If
                
        End Select
            
        .cboBancLiqu.Top = .lblBancLiqu.Top - 60
        .cboSistema.Top = .lblSistema.Top - 60
        .dtpInicio.Top = .lblData.Top - 60
        .tlbData.Top = .lblData.Top - 60
        .cboLocalLiqu.Top = .lblLocalLiqu.Top - 60
        .cboGrupoVeicLegal.Top = .lblGrupoVeicLegal.Top - 60
        .cboVeiculoLegal.Top = .lblVeiculoLegal.Top - 60
        .txtVeiculoLegal.Top = .lblVeiculoLegal.Top - 60
        .cboTipoCaixa.Top = .lblTipoCaixa.Top - 60
        .cboItemCaixa.Top = .lblItemCaixa.Top - 60
        .txtContraParte.Top = .lblContraParte.Top - 60
        .cboCamara.Top = .lblCamara.Top - 60
        .cboOperacaoEvento.Top = .lblOperacaoEvento.Top - 60
        .txtComando.Top = .lblNumeroComando.Top - 60
        .cboAcoes.Top = .lblAcoes.Top - 60
        .cboTipoOperacao.Top = .lblTipoOperacao.Top - 60
        .cboMensagemSPB.Top = .lblMensagemSPB.Top - 60
        .cboTipoLiquidacao.Top = .lblTipoLiquidacao.Top - 60
        .cboSituacaoOperacao.Top = .lblSituacaoOperacao.Top - 60
        .dtpInicio2.Top = .lblData2.Top - 60
        .tlbData2.Top = .lblData2.Top - 60
        .numCodigoCNPJ.Top = .lblCodigoCNPJ.Top - 60
        .txtParticipacaoCETIP.Top = .lblParticipacaoCETIP.Top - 60
        .txtContaSelic.Top = .lblContaSelic.Top - 60
        .txtTipoTituloBMA.Top = .lblTipoTituloBMA.Top - 60
        .cboTipoBackoffice.Top = .lblTipoBackoffice.Top - 60
        .cboCanalVenda.Top = .lblCanalVenda.Top - 60
        .txtCodReemb.Top = .lblCodReemb.Top - 60
        .cboMensagemCAM.Top = .lblMensagemCAM.Top - 60
        .txtIdentificadorPessoa.Top = .lblIdentificadorPessoa.Top - 60
        .tlbDataEventoCambio.Top = .lblDataEventoCambio.Top - 60
        .dtpFim.Top = .lblData.Top - 60
        .dtpFim2.Top = .lblData2.Top - 60
        .dtpDataEventoCambioIni.Top = .lblDataEventoCambio.Top - 60
        .dtpDataEventoCambioFim.Top = .lblDataEventoCambio.Top - 60
        .tlbHora.Top = .lblHora.Top - 60
        .dtpHoraInicio.Top = .lblHora.Top - 60
        .dtpHoraFim.Top = .lblHora.Top - 60
        
        .tlbComandos.Top = .fraFiltro.Top + .fraFiltro.Height + 60
        .Height = (.Height - .ScaleHeight) + .tlbComandos.Top + .tlbComandos.Height
        
    End With

    Call flCarregarCombosUtilizados
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "frmFiltro", "flConfiguraLayoutForm", 0

End Sub

'Preencher o combo de ação
Private Sub flPreenchecboAcao()

    cboAcoes.Clear
    cboAcoes.AddItem "<-- Todos -->"
    cboAcoes.AddItem enumTipoAcao.CancelamentoSolicitado & " - " & fgDescricaoTipoAcao(enumTipoAcao.CancelamentoSolicitado)
    cboAcoes.AddItem enumTipoAcao.CancelamentoEnviado & " - " & fgDescricaoTipoAcao(enumTipoAcao.CancelamentoEnviado)
    cboAcoes.AddItem enumTipoAcao.EstornoSolicitado & " - " & fgDescricaoTipoAcao(enumTipoAcao.EstornoSolicitado)
    cboAcoes.AddItem enumTipoAcao.EstornoEnviado & " - " & fgDescricaoTipoAcao(enumTipoAcao.EstornoEnviado)
    
    cboAcoes.ListIndex = 0

End Sub

'Preencher o combo de Canal de Venda
Private Sub flPreenchecboCanalVenda()

    cboCanalVenda.Clear
    cboCanalVenda.AddItem "<-- Todos -->"
    cboCanalVenda.AddItem enumCanalDeVenda.SGC & " - " & fgDescricaoCanalVenda(enumCanalDeVenda.SGC)
    cboCanalVenda.AddItem enumCanalDeVenda.SGM & " - " & fgDescricaoCanalVenda(enumCanalDeVenda.SGM)
    cboCanalVenda.ListIndex = 0

End Sub

'Valida se a data foi selecionada, pois é o único campo obrigatório do filtro

Private Function flValidarCampos() As String

Dim strRetorno                              As String

    If tlbData.Top < PROP_TOP_CAMPO_23 And tlbData.Buttons("Comparacao").Image <> 2 Then
        strRetorno = "Obrigatória a seleção do filtro DATA."
    End If

    flValidarCampos = strRetorno

End Function

Private Sub cboBancLiqu_Click()

On Error GoTo ErrorHandler

    fgCursor True

    If cboSistema.Top < PROP_TOP_CAMPO_23 Then
        flCarregarSistema
    End If

    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - cboBancLiqu_Click", Me.Caption)

End Sub

' Carrega os sistemas e preencher o combo de sistemas com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsMIU.Executar

Private Sub flCarregarSistema()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim xmlDomSistema           As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
 
On Error GoTo ErrorHandler

    If cboBancLiqu.ListIndex > 0 Then
        
        Set xmlMapaNavegacao = Nothing
    
        Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
        Call fgAppendNode(xmlMapaNavegacao, vbNullString, "Repeat_Leitura", vbNullString)
        Call fgAppendNode(xmlMapaNavegacao, "Repeat_Leitura", "Grupo_Leitura", vbNullString)
        Call fgAppendAttribute(xmlMapaNavegacao, "Grupo_Leitura", "Operacao", "LerTodos")
        Call fgAppendAttribute(xmlMapaNavegacao, "Grupo_Leitura", "Objeto", "A6A7A8.clsSistema")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "TP_VIGE", "S")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "TP_SEGR", "S")
        Call fgAppendNode(xmlMapaNavegacao, "Grupo_Leitura", "CO_EMPR", fgObterCodigoCombo(cboBancLiqu.Text))

        Set xmlDomSistema = CreateObject("MSXML2.DOMDocument.4.0")
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
        If Not xmlDomSistema.loadXML(objMIU.Executar(xmlMapaNavegacao.xml, _
                                                     vntCodErro, _
                                                     vntMensagemErro)) Then
            
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
        
            cboSistema.Clear
            cboSistema.Enabled = False
            Exit Sub
        End If
        Set objMIU = Nothing
        
        Call fgCarregarCombos(cboSistema, xmlDomSistema, "Sistema", "SG_SIST", "NO_SIST", True)
        
        cboSistema.Enabled = True
        
    Else
        cboSistema.Clear
        cboSistema.Enabled = False
    End If

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarSistema", 0

End Sub

Private Sub cboGrupoVeicLegal_Click()

On Error GoTo ErrorHandler

    fgCursor True
    
    If cboGrupoVeicLegal.ListIndex >= 0 Then
        cboVeiculoLegal.Clear
        txtVeiculoLegal.Text = vbNullString
        txtVeiculoLegal.Enabled = True
        cboVeiculoLegal.Enabled = True
    End If

    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, "frmFiltro - cboGrupoVeicLegal_Click", Me.Caption)

End Sub

' Habilitação do combo de Canal de Vendas, apenas na tela de Filtro, não mais depende do tipo de operação.
' Solicitação FABIANA em 16/04/2008.
' Implementação CAS em 17/04/2008.
'
'Private Sub cboTipoOperacao_Click()
'
'Dim strTipoOper                             As String
'Dim blnCboCanalVendaEnable                  As Boolean
'
'On Error GoTo ErrorHandler
'
'    If cboTipoOperacao.ListIndex = 0 Or cboTipoOperacao.ListCount = 0 Then
'        cboCanalVenda.ListIndex = -1
'        cboCanalVenda.Enabled = False
'        Exit Sub
'    End If
'
'    strTipoOper = fgObterCodigoCombo(cboTipoOperacao.Text)
'
'    blnCboCanalVendaEnable = fgIN(CLng(strTipoOper), enumTipoOperacaoLQS.EventosJurosSWAP, _
'                                                     enumTipoOperacaoLQS.RegistroContratoSWAP, _
'                                                     enumTipoOperacaoLQS.RegDadosComplemContratoSWAP, _
'                                                     enumTipoOperacaoLQS.AntecipacaoResgateContratoSWAP, _
'                                                     enumTipoOperacaoLQS.LanctoPUFatorContratoDerivativo, _
'                                                     enumTipoOperacaoLQS.ExercicioOpcaoContratoSWAP)
'
'    If blnCboCanalVendaEnable Then
'        cboCanalVenda.Enabled = blnCboCanalVendaEnable
'    Else
'        cboCanalVenda.ListIndex = 0
'        cboCanalVenda.Enabled = blnCboCanalVendaEnable
'    End If
'
'Exit Sub
'ErrorHandler:
'
'   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboTipoOperacao_Click"
'
'
'End Sub

Private Sub cboVeiculoLegal_Click()
   
On Error GoTo ErrorHandler

    If cboVeiculoLegal.ListIndex > 0 Then
        If UCase$(txtVeiculoLegal.Text) <> UCase$(Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1)) Then
            txtVeiculoLegal.Text = Split(arrSgSistCoVeicLega(cboVeiculoLegal.ListIndex), "k_")(1)
        End If
    Else
        If Not Me.ActiveControl Is txtVeiculoLegal Then
            txtVeiculoLegal.Text = vbNullString
        End If
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_Click"

End Sub

Private Sub cboVeiculoLegal_DropDown()
  On Error GoTo ErrorHandler

    If cboGrupoVeicLegal.ListIndex >= -1 And cboVeiculoLegal.ListCount = 0 Then
        fgCursor True
        Call fgLerCarregarVeiculoLegal(cboGrupoVeicLegal, cboVeiculoLegal, xmlPropriedades, arrSgSistCoVeicLega)
        fgCursor
    End If
    
Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_DropDown"
End Sub

Private Sub cboVeiculoLegal_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - cboVeiculoLegal_KeyPress"
End Sub

Private Sub dtpFim2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dtpInicio2_Change()
    dtpFim2.MinDate = dtpInicio2.value
    If dtpFim2.value < dtpInicio2.value Then
        dtpFim2.MinDate = dtpInicio2.value
        dtpFim2.value = dtpInicio2.value
    End If
End Sub

Private Sub dtpInicio2_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Activate()

On Error GoTo ErrorHandler

    If blnOcorreuErro Then Exit Sub

    If blnPrimeiroActivate Then
        blnPrimeiroActivate = False
    End If
    
    Call flAplicarSettingsRegistry
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    blnPrimeiroActivate = False
    blnOcorreuErro = True
    mdiLQS.uctlogErros.MostrarErros Err, "frmFiltro - Form_Activate", Me.Caption
    
End Sub

' Carrega os registros e preencher a interface com os mesmos

Private Sub flCarregarRegistro()

Dim strRegistry                             As String

On Error GoTo ErrorHandler

    Set xmlDOMRegistro = CreateObject("MSXML2.DOMDocument.4.0")
    
    strRegistry = GetSetting("A8LQS", "Form Filtro\" & FormOwner.Name, "Settings")
    If strRegistry <> vbNullString Then
        If Not xmlDOMRegistro.loadXML(strRegistry) Then
            Call fgErroLoadXML(xmlDOMRegistro, App.EXEName, "frmFiltro", "flCarregarRegistro")
        Else
            If Not xmlDOMRegistro.selectSingleNode("//Grupo_BackOfficePerfilGeral") Is Nothing Then
                Call fgRemoveNode(xmlDOMRegistro, "Grupo_BackOfficePerfilGeral")
            End If
        End If
    End If

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, "frmFiltro", "flCarregarRegistro", 0

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    
    blnPrimeiroActivate = True
    Set Me.Icon = mdiLQS.Icon
    
    'Apresenta o controles de data com Checked, pois este será obrigatório
    tlbData.Buttons("Comparacao").Image = 2
    
    flCarregarRegistro
    
    Select Case TipoFiltro
        Case enumTipoFiltroA8.frmConsultaOperacao
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.txtContraParte
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.txtComando
            lcControlesFiltro.Add Me.cboAcoes
            lcControlesFiltro.Add Me.cboTipoOperacao
            lcControlesFiltro.Add Me.cboSituacaoOperacao
            lcControlesFiltro.Add Me.cboCanalVenda
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_ConsultaOperacao"

        'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Operações - Adrian - 20/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacao
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboSistema
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_ConsultaMovimentacao"
            
        'Inclusão de Tratamento da Nova Tela de Consulta de Movimentação de Operações Rejeitadas - Adrian - 23/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacaoRejeitada
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboSistema
            
            strFuncionalidade = "frmFiltro_ConsultaMovimentacaoRejeitada"
            
        'Inclusão de Tratamento da Nova Tela de Consulta de Integrações de Conta Corrente e Contabilidade - Adrian - 24/06/05
        Case enumTipoFiltroA8.frmConsultaMovimentacaoCC_HA
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboSistema
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboCanalVenda
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_ConsultaMovimentacaoCC_HA"
            
        Case enumTipoFiltroA8.frmConsultaMensagem
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.tlbHora
            lcControlesFiltro.Add Me.dtpHoraInicio
            lcControlesFiltro.Add Me.dtpHoraFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoBackoffice
            lcControlesFiltro.Add Me.cboMensagemSPB
            'RATS 1176 - Considerar o sistema no filtro
            lcControlesFiltro.Add Me.cboSistema
            
            strFuncionalidade = "frmFiltro_ConsultaMensagem"
            
        Case enumTipoFiltroA8.frmConsultaMensagemSisbacen
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboMensagemCAM
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.tlbDataEventoCambio
            lcControlesFiltro.Add Me.dtpDataEventoCambioIni
            lcControlesFiltro.Add Me.dtpDataEventoCambioFim
            lcControlesFiltro.Add Me.txtIdentificadorPessoa
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_ConsultaMensagem"
            
        Case enumTipoFiltroA8.frmRemessaRejeitada
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.cboSistema
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            
            strFuncionalidade = "frmFiltro_RemessaRejeitada"
            
        Case enumTipoFiltroA8.frmConfirmacaoOperacao
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.txtContraParte
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            lcControlesFiltro.Add Me.cboTipoLiquidacao
            
            strFuncionalidade = "frmFiltro_ConfirmacaoOperacao"
        
        Case enumTipoFiltroA8.frmLiberacaoOperacaoMensagem
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.txtContraParte
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            lcControlesFiltro.Add Me.cboTipoLiquidacao
            
            strFuncionalidade = "frmFiltro_frmLiberacaoOperacaoMensagem"
            
        Case enumTipoFiltroA8.frmSuspenderDisponibilizarLancamentoCC
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            
            strFuncionalidade = "frmFiltro_frmSuspenderDisponibilizarLancamentoCC"
            
        Case enumTipoFiltroA8.frmIntegrarCCOnLine
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            lcControlesFiltro.Add Me.cboCanalVenda
            
            strFuncionalidade = "frmFiltro_frmIntegrarCCOnLine"
            
        Case enumTipoFiltroA8.frmIntegrarCCOnLineEstorno
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            
            strFuncionalidade = "frmFiltro_frmIntegrarCCOnLine"
            
        Case enumTipoFiltroA8.frmReenvioCancelamentoEstornoMsg
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            
            strFuncionalidade = "frmFiltro_ReenvioCancelamentoEstornoMsg"
        
        Case enumTipoFiltroA8.frmConsultaContaCorrente
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.tlbData
            lcControlesFiltro.Add Me.dtpInicio
            lcControlesFiltro.Add Me.dtpFim
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboLocalLiqu
            lcControlesFiltro.Add Me.cboTipoOperacao
            lcControlesFiltro.Add Me.cboCanalVenda
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_frmConsultaContaCorrente"
            
        Case enumTipoFiltroA8.frmConsultaVeiculoLegal
            lcControlesFiltro.Add Me.cboBancLiqu
            lcControlesFiltro.Add Me.cboGrupoVeicLegal
            lcControlesFiltro.Add Me.txtVeiculoLegal
            lcControlesFiltro.Add Me.cboVeiculoLegal
            lcControlesFiltro.Add Me.cboSistema
            lcControlesFiltro.Add Me.numCodigoCNPJ
            lcControlesFiltro.Add Me.txtParticipacaoCETIP
            lcControlesFiltro.Add Me.txtTipoTituloBMA
            lcControlesFiltro.Add Me.txtContaSelic
            lcControlesFiltro.Add Me.cboTipoBackoffice
            
            strFuncionalidade = "frmFiltro_ConsultaVeiculoLegal"
            
    End Select
    
    Set xmlPropriedades = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlPropriedades, "", "Grupo_Propriedades", "")
    Call fgAppendAttribute(xmlPropriedades, "Grupo_Propriedades", "Objeto", "")
    Call fgAppendAttribute(xmlPropriedades, "Grupo_Propriedades", "Operacao", "")
    
    Call fgCursor(True)

    Me.dtpInicio.value = fgValidarMinDateDTPicker(Me.dtpInicio, fgDataHoraServidor(Data))
    Me.dtpFim.value = fgValidarMinDateDTPicker(Me.dtpFim, Me.dtpInicio.value)
    Me.dtpInicio2.value = fgValidarMinDateDTPicker(Me.dtpInicio2, Me.dtpInicio.value)
    Me.dtpFim2.value = fgValidarMinDateDTPicker(Me.dtpFim2, Me.dtpInicio.value)
    Me.dtpDataEventoCambioIni = fgValidarMinDateDTPicker(Me.dtpDataEventoCambioIni, Me.dtpInicio.value)
    Me.dtpDataEventoCambioFim = fgValidarMinDateDTPicker(Me.dtpDataEventoCambioFim, Me.dtpInicio.value)
    Me.dtpHoraInicio = CDate("00:00:00")
    Me.dtpHoraFim = CDate("00:00:00")
    
    Call flConfiguraLayoutForm
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmFiltro - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlPropriedades = Nothing
    Set xmlDOMRegistro = Nothing
    Set xmlMapaNavegacao = Nothing
End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

Dim strRetorno                              As String

    fgCursor True

    Select Case Button.Key
        Case "cancelar"
            Me.Hide
            
        Case "aplicar"
            strRetorno = flValidarCampos
            If strRetorno = vbNullString Then
                If cboBancLiqu.Top < PROP_TOP_CAMPO_23 Then
                    cboBancLiqu.SetFocus
                End If
                Call flAplicarFiltro(True)
            Else
                frmMural.Caption = Me.Caption
                frmMural.Display = strRetorno
                frmMural.Show vbModal
            End If
        
    End Select
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmFiltro - tlbComandos_ButtonClick", Me.Caption
End Sub

Private Sub dtpInicio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dtpFim_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dtpInicio_Change()

    dtpFim.MinDate = dtpInicio.value
    If dtpFim.value < dtpInicio.value Then
        dtpFim.MinDate = dtpInicio.value
        dtpFim.value = dtpInicio.value
    End If
    
    If tlbData.Buttons(1).Caption = "Após" Then
        If dtpInicio.value = fgDataHoraServidor(Data) Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    ElseIf tlbData.Buttons(1).Caption = "Entre" Then
        If dtpInicio.value = dtpFim.value Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    Else
        blnFiltraHorario = False
        tlbHora.Enabled = False
        dtpHoraInicio.Enabled = False
        dtpHoraFim.Enabled = False
    End If
    
End Sub

Private Sub dtpFim_Change()
    
    If tlbData.Buttons(1).Caption = "Após" Then
        If dtpInicio.value = fgDataHoraServidor(Data) Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    ElseIf tlbData.Buttons(1).Caption = "Entre" Then
        If dtpInicio.value = dtpFim.value Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    Else
        blnFiltraHorario = False
        tlbHora.Enabled = False
        dtpHoraInicio.Enabled = False
        dtpHoraFim.Enabled = False
    End If
    
End Sub

Private Sub tlbData_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    If Button.Image = 2 Then
        Button.Image = 0
        blnFiltraHora = False
    Else
        blnFiltraHora = True
        Button.Image = 2
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData_ButtonClick"
End Sub

Private Sub tlbData_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo ErrorHandler

    flConfiguratlbData ButtonMenu.Text, tlbData, dtpFim
    
    If tlbData.Buttons(1).Caption = "Após" Then
        If dtpInicio.value = fgDataHoraServidor(Data) Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    ElseIf tlbData.Buttons(1).Caption = "Entre" Then
        If dtpInicio.value = dtpFim.value Then
            blnFiltraHorario = True
            tlbHora.Enabled = True
            dtpHoraInicio.Enabled = True
            dtpHoraFim.Enabled = True
        Else
            blnFiltraHorario = False
            tlbHora.Enabled = False
            dtpHoraInicio.Enabled = False
            dtpHoraFim.Enabled = False
        End If
    Else
        blnFiltraHorario = False
        tlbHora.Enabled = False
        dtpHoraInicio.Enabled = False
        dtpHoraFim.Enabled = False
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData_ButtonMenuClick"
End Sub

'Configura o lable das Datas

Private Sub flConfiguratlbData(ByVal pstrCaption As String, _
                               ByRef tlbControl As Object, _
                               ByRef dtpDataFim As Object)
        
On Error GoTo ErrorHandler

    With tlbControl
        
        Select Case pstrCaption
            
            Case "Após"
                dtpDataFim.Visible = False
                .Buttons(1).ButtonMenus(1).Text = "Antes"
                .Buttons(1).ButtonMenus(2).Text = "Entre"
                .Buttons(1).Caption = "Após"
                
            Case "Antes"
                dtpDataFim.Visible = False
                .Buttons(1).ButtonMenus(1).Text = "Após"
                .Buttons(1).ButtonMenus(2).Text = "Entre"
                .Buttons(1).Caption = "Antes"
            
            Case "Entre"
                dtpDataFim.Visible = True
                .Buttons(1).ButtonMenus(1).Text = "Após"
                .Buttons(1).ButtonMenus(2).Text = "Antes"
                .Buttons(1).Caption = "Entre"
                
        End Select
    
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flConfiguratlbData", 0

End Sub

Private Sub tlbData2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    If Button.Image = 2 Then
        blnFiltraHora2 = False
        Button.Image = 0
    Else
        blnFiltraHora2 = True
        Button.Image = 2
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData2_ButtonClick"
End Sub

Private Sub tlbData2_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo ErrorHandler

    flConfiguratlbData ButtonMenu.Text, tlbData2, dtpFim2

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData2_ButtonMenuClick"
End Sub

Private Sub tlbDataEventoCambio_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler

    If Button.Image = 2 Then
        blnFiltraHora3 = False
        Button.Image = 0
    Else
        blnFiltraHora3 = True
        Button.Image = 2
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData2_ButtonClick"

End Sub

Private Sub tlbHora_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler

    If Button.Image = 2 Then
        blnFiltraHorario = False
        Button.Image = 0
    Else
        blnFiltraHorario = True
        Button.Image = 2
    End If

Exit Sub
ErrorHandler:
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbHora_ButtonClick"

End Sub

Private Sub tlbDataEventoCambio_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrorHandler

    flConfiguratlbData ButtonMenu.Text, tlbDataEventoCambio, dtpDataEventoCambioFim

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbData2_ButtonMenuClick"

End Sub

Private Sub tlbHora_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
On Error GoTo ErrorHandler

    flConfiguratlbData ButtonMenu.Text, tlbHora, dtpHoraFim

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - tlbHora_ButtonMenuClick"

End Sub

Private Sub txtContaSelic_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtContaSelic_KeyPress"

End Sub

Private Sub txtContraParte_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Executa a busca de veículo legal
Private Function flBuscaVeiculoLegal(ByVal pstrCodigo As String) As Boolean

Dim intCont                                 As Integer

On Error GoTo ErrorHandler

    If cboVeiculoLegal.ListCount <= 1 Then
        flBuscaVeiculoLegal = False
        Exit Function
    End If

    For intCont = 1 To UBound(arrSgSistCoVeicLega)
        If UCase$(Split(arrSgSistCoVeicLega(intCont), "k_")(1)) = UCase$(pstrCodigo) Then
            If cboVeiculoLegal.ListIndex <> intCont Then
                cboVeiculoLegal.ListIndex = intCont
            End If
            flBuscaVeiculoLegal = True
            Exit Function

        End If
    Next intCont
    
    flBuscaVeiculoLegal = False

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flBuscaVeiculoLegal", 0
End Function

Private Sub txtParticipacaoCETIP_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler
    
    If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And _
       (KeyAscii <> vbKeyBack) Then
       KeyAscii = 0
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtParticipacaoCETIP_KeyPress"
End Sub

Private Sub txtTipoTituloBMA_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtVeiculoLegal_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtVeiculoLegal_Change()

On Error GoTo ErrorHandler

    If cboVeiculoLegal.ListCount > 0 Then
        If Not flBuscaVeiculoLegal(txtVeiculoLegal.Text) Then
            cboVeiculoLegal.ListIndex = 0
        End If
    Else
        Call cboVeiculoLegal_DropDown
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - txtVeiculoLegal_Change"
End Sub

'Preencher o combo de Codigo Mensagem Sisbacen
Private Sub flPreencheMensagemCAM()
    
    With cboMensagemCAM
        .Clear
        .AddItem "CAM0042 - IF consulta contratos em ser", 0
        .AddItem "CAM0043 - IF consulta eventos de um dia", 1
        .AddItem "CAM0044 - IF consulta detalhamento de contrato interbancário", 2
        .AddItem "CAM0045 - IF consulta eventos de um contrato do mercado primário", 3
        .AddItem "CAM0046 - Corretora consulta eventos de um contrato intermediário no mercado primário", 4
        .AddItem "CAM0047 - IF consulta histórico de incorporações", 5
        .AddItem "CAM0048 - IF consulta contratos da incorporação", 6
        .AddItem "CAM0049 - IF consulta cadeia de incorporações de um contrato", 7
        .AddItem "CAM0050 - IF consulta posição de câmbio por moeda", 8
        .AddItem "CAM0052 - IF consulta instruções de pagamento", 9
        .ListIndex = -1
    End With

End Sub
