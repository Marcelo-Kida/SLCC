VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEntradaManualSLCCBeta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ferramentas - Entrada Manual - Operação"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   13380
   Begin VB.CheckBox chkParametro 
      Caption         =   "Carregar operação com valores pré-cadastrados"
      Height          =   240
      Left            =   135
      TabIndex        =   22
      Top             =   7770
      Value           =   1  'Checked
      Width           =   3750
   End
   Begin VB.Frame fraFiltro 
      Height          =   7635
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4815
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   450
         Width           =   4635
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1035
         Width           =   4635
      End
      Begin VB.ComboBox cboServico 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1635
         Width           =   4650
      End
      Begin VB.ComboBox cboEvento 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2250
         Width           =   4650
      End
      Begin MSComctlLib.TreeView treMensagem 
         Height          =   4890
         Left            =   105
         TabIndex        =   11
         Top             =   2655
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   8625
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   406
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgLstMensagem"
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   19
         Top             =   210
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   18
         Top             =   795
         Width           =   1125
      End
      Begin VB.Label Label3 
         Caption         =   "Serviço"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   17
         Top             =   1410
         Width           =   1125
      End
      Begin VB.Label Label4 
         Caption         =   "Evento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   16
         Top             =   2010
         Width           =   1125
      End
   End
   Begin VB.Frame Frame2 
      Height          =   7650
      Left            =   4845
      TabIndex        =   0
      Top             =   0
      Width           =   8520
      Begin VB.TextBox txtDescrticaoMensagem 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   420
         Width           =   8265
      End
      Begin FPSpread.vaSpread sprMensagemSLCC 
         Height          =   6795
         Left            =   120
         TabIndex        =   2
         Top             =   765
         Width           =   8280
         _Version        =   196608
         _ExtentX        =   14605
         _ExtentY        =   11986
         _StockProps     =   64
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
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
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         MaxCols         =   5
         MaxRows         =   1
         MoveActiveOnFocus=   0   'False
         NoBorder        =   -1  'True
         ProcessTab      =   -1  'True
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmEntradaManualSLCCBeta.frx":0000
         ScrollBarTrack  =   3
      End
      Begin VB.Frame fraContingencia 
         Caption         =   "Sistema XYN em contingência."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   570
         Left            =   135
         TabIndex        =   3
         Top             =   7020
         Visible         =   0   'False
         Width           =   8250
         Begin VB.CheckBox chkPrevisaoPJ 
            Caption         =   "Previsão PJ"
            Height          =   210
            Left            =   1080
            TabIndex        =   7
            Top             =   270
            Width           =   1170
         End
         Begin VB.CheckBox chkRealizadoPJ 
            Caption         =   "Realizado PJ"
            Height          =   210
            Left            =   2540
            TabIndex        =   6
            Top             =   270
            Width           =   1350
         End
         Begin VB.CheckBox chkPrevistoA6 
            Caption         =   "Previsão A6"
            Height          =   210
            Left            =   4180
            TabIndex        =   5
            Top             =   270
            Width           =   1170
         End
         Begin VB.CheckBox chkRealizadoA6 
            Caption         =   "Realizado A6"
            Height          =   210
            Left            =   5640
            TabIndex        =   4
            Top             =   270
            Width           =   1350
         End
         Begin VB.Label Label7 
            Caption         =   "Enviar : "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   180
            TabIndex        =   8
            Top             =   255
            Width           =   1365
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Descrição da Mensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   105
         TabIndex        =   9
         Top             =   195
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   7155
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
            Picture         =   "frmEntradaManualSLCCBeta.frx":0389
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":06A3
            Key             =   "Padrao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":0AF5
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":0E0F
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":1129
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":1443
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":1895
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":1CE7
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   9270
      TabIndex        =   20
      Top             =   7695
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar Parametrização Padrão"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Enviar"
            Key             =   "Enviar"
            Object.ToolTipText     =   "Enviar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstMensagem 
      Left            =   675
      Top             =   7155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":2139
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":2453
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEntradaManualSLCCBeta.frx":2D2D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Caption         =   "Atributos em negrito são obrigatórios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4980
      TabIndex        =   21
      Top             =   7740
      Width           =   5745
   End
End
Attribute VB_Name = "frmEntradaManualSLCCBeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'' Objeto responsável pelo envio da solicitação (Entrada manual de operação) à
'' camada controladora de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMiu
''      A8MIU.clsMensagem

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private strOperacao                         As String

Private Const strFuncionalidade             As String = "frmEntradaManual"

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private blnLimpar                           As Boolean

Private Const intColunmNomeFisico           As Integer = 1
Private Const intColunmNomeLogico           As Integer = 2
Private Const intColunmEdicao               As Integer = 3
Private Const intColunmTabela               As Integer = 4
Private Const intColunmButton               As Integer = 5

Private strSepMilhar                        As String
Private strSepDecimal                       As String
Private gstrSiglaSistema                    As String

Private xmlMsgEntradaManual                 As MSXML2.DOMDocument40
Private xmlMensagemBase                     As MSXML2.DOMDocument40

Private xmlDominioVeicLega                  As MSXML2.DOMDocument40
Private xmlDominioTipoAtivMerc              As MSXML2.DOMDocument40

Private lngRowControleBMC                   As Long
Private lngRowValor                         As Long
Private lngRowVeiculoLegal                  As Long
Private strChaveMensagemBMC0112             As String
Private lngRowChacam                        As Long
Private lngRowDataOperacao                  As Long
Private lngRowTipoOperacaoCambio            As Long
Private lngRowValorTaxaCambio               As Long
Private lngRowValorMoedaNac                 As Long
Private lngRowValorMoedaEstr                As Long
Private lngRowDataLiquOper                  As Long
Private lngRowRegistroOperCamb              As Long
Private lngRowRegistroOperCamb2             As Long
Private lngRowCnpjBaseIf                    As Long
Private lngRowMoedIso                       As Long
Private lngRowDataEntrMoedNac               As Long
Private lngRowDataEntrMoedEstr              As Long

Private Const NET_ENTRADA_MANUAL_ATIVA      As Boolean = True

Private Const NET_CO_SERV                   As Integer = 9000
Private Const NET_CO_EVEN_MULTI             As Integer = 9000
Private Const NET_CO_EVEN_BILAT             As Integer = 9001

Private Const NET_DE_SERV                   As String = "Entrada Manual de NET"
Private Const NET_DE_EVEN_MULTI             As String = "Entrada Manual de NET Multilateral"
Private Const NET_DE_EVEN_BILAT             As String = "Entrada Manual de NET Bilateral"

Private Const NET_TAGS_VISIVEIS_MULTI       As String = "|DT_MESG" & _
                                                        "|HO_MESG" & _
                                                        "|CO_USUA_CADR_OPER" & _
                                                        "|CO_LOCA_LIQU" & _
                                                        "|DT_OPER_ATIV" & _
                                                        "|CO_VEIC_LEGA" & _
                                                        "|VA_OPER_ATIV" & _
                                                        "|IN_OPER_DEBT_CRED"

Private Const NET_TAGS_VISIVEIS_MULTI_BMC   As String = "|DT_MESG" & _
                                                        "|HO_MESG" & _
                                                        "|CO_USUA_CADR_OPER" & _
                                                        "|CO_LOCA_LIQU" & _
                                                        "|DT_LIQU_OPER_ATIV" & _
                                                        "|CO_VEIC_LEGA" & _
                                                        "|VA_OPER_ATIV" & _
                                                        "|IN_OPER_DEBT_CRED" & _
                                                        "|VA_MOED_ESTR_BMC"

Private Const NET_TAGS_VISIVEIS_BILAT       As String = "|DT_MESG" & _
                                                        "|HO_MESG" & _
                                                        "|CO_USUA_CADR_OPER" & _
                                                        "|CO_LOCA_LIQU" & _
                                                        "|DT_OPER_ATIV" & _
                                                        "|CO_VEIC_LEGA" & _
                                                        "|VA_OPER_ATIV" & _
                                                        "|IN_OPER_DEBT_CRED" & _
                                                        "|CO_PARP_CAMR" & _
                                                        "|NO_CNPT" & _
                                                        "|TP_CNPT" & _
                                                        "|CO_CNPT_CAMR" & _
                                                        "|CO_CNPJ_CNPT"

'Retirado CO_ISPB_BANC_LIQU_CNPT por não estar sendo homologado e por causar confusão ao usuário.
'Retornar a partir do momento em que o legado passar a enviar a informação.
'Cassiano - 11/04/2006

'Private Const NET_TAGS_VISIVEIS_BILAT       As String = "|DT_MESG" & _
                                                        "|HO_MESG" & _
                                                        "|CO_USUA_CADR_OPER" & _
                                                        "|CO_LOCA_LIQU" & _
                                                        "|DT_OPER_ATIV" & _
                                                        "|CO_VEIC_LEGA" & _
                                                        "|VA_OPER_ATIV" & _
                                                        "|IN_OPER_DEBT_CRED" & _
                                                        "|CO_PARP_CAMR" & _
                                                        "|CO_ISPB_BANC_LIQU_CNPT" & _
                                                        "|NO_CNPT" & _
                                                        "|TP_CNPT" & _
                                                        "|CO_CNPT_CAMR"

Private Enum enumEventosNetEntradaManual
    CarregarComboServicos = 1
    CarregarComboEventos = 2
    CarregarTreeView = 3
    CarregarSpread = 4
End Enum

Private Sub flTratarNetEntradaManual(ByVal pintTratamento As enumEventosNetEntradaManual)

Dim objNode                                 As Node

Dim strGrupo                                As String
Dim strServico                              As String
Dim strEvento                               As String
Dim strMensagem                             As String

Dim strCodigoGrupo                          As String
Dim lngRow                                  As Long
Dim vntCellText                             As Variant
Dim intModalidadeLiquidacao                 As Integer
Dim strTagsVisiveisAux                      As String

    If Not NET_ENTRADA_MANUAL_ATIVA Then Exit Sub
    
    strCodigoGrupo = fgObterCodigoCombo(Me.cboGrupo.Text)
    
    If strCodigoGrupo <> "BMC" And _
       strCodigoGrupo <> "BMA" And _
       strCodigoGrupo <> "LDL" And _
       strCodigoGrupo <> "CTP" Then
        Exit Sub
    End If
    
    If cboServico.ListIndex >= 0 Then
        If cboServico.ItemData(cboServico.ListIndex) <> NET_CO_SERV Then
            Exit Sub
        End If
    End If
    
    If cboEvento.ListIndex >= 0 Then
        If cboEvento.ItemData(cboEvento.ListIndex) <> NET_CO_EVEN_MULTI And _
           cboEvento.ItemData(cboEvento.ListIndex) <> NET_CO_EVEN_BILAT Then
            Exit Sub
        End If
    End If
    
    If Not treMensagem.SelectedItem Is Nothing Then
        If Val(fgObterCodigoCombo(treMensagem.SelectedItem.Text)) <> NET_CO_EVEN_MULTI And _
           Val(fgObterCodigoCombo(treMensagem.SelectedItem.Text)) <> NET_CO_EVEN_BILAT Then
            Exit Sub
        End If
    End If
    
    Select Case pintTratamento
        Case enumEventosNetEntradaManual.CarregarComboServicos
            cboServico.AddItem NET_DE_SERV
            cboServico.ItemData(cboServico.NewIndex) = NET_CO_SERV
            
        Case enumEventosNetEntradaManual.CarregarComboEventos
            cboEvento.AddItem NET_DE_EVEN_MULTI
            cboEvento.ItemData(cboEvento.NewIndex) = NET_CO_EVEN_MULTI
                        
            If strCodigoGrupo = "CTP" Then
                cboEvento.AddItem NET_DE_EVEN_BILAT
                cboEvento.ItemData(cboEvento.NewIndex) = NET_CO_EVEN_BILAT
            End If
            
        Case enumEventosNetEntradaManual.CarregarTreeView
            On Error Resume Next
            
            strGrupo = "G" & fgCompletaString(cboGrupo.ItemData(cboGrupo.ListIndex), "0", 10, True)
            Set objNode = treMensagem.Nodes.Add(, tvwChild, "G" & strGrupo, strCodigoGrupo, 1)
            objNode.Expanded = True
            objNode.Tag = "G"
    
            On Error GoTo 0
            
            strServico = "S" & fgCompletaString(NET_CO_SERV, "0", 10, True)
            Set objNode = treMensagem.Nodes.Add("G" & strGrupo, tvwChild, "S" & strGrupo & strServico, NET_DE_SERV, 2)
            objNode.Expanded = True
            objNode.Tag = "S"
            
            If cboEvento.ListIndex = -1 Then
                strEvento = "E" & fgCompletaString(NET_CO_EVEN_MULTI, "0", 10, True)
                    
                strMensagem = "M" & fgCompletaString(NET_CO_EVEN_MULTI, "0", 10, True) & fgCompletaString(NET_CO_EVEN_MULTI, " ", 9, True)
                Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem & "TP" & Format(NET_CO_EVEN_MULTI, "0000"), Format(NET_CO_EVEN_MULTI, "0000") & "-" & NET_DE_EVEN_MULTI, 3)
                objNode.Tag = "M"
            Else
                If cboEvento.ItemData(cboEvento.ListIndex) = NET_CO_EVEN_MULTI Then
                    strEvento = "E" & fgCompletaString(NET_CO_EVEN_MULTI, "0", 10, True)
                        
                    strMensagem = "M" & fgCompletaString(NET_CO_EVEN_MULTI, "0", 10, True) & fgCompletaString(NET_CO_EVEN_MULTI, " ", 9, True)
                    Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem & "TP" & Format(NET_CO_EVEN_MULTI, "0000"), Format(NET_CO_EVEN_MULTI, "0000") & "-" & NET_DE_EVEN_MULTI, 3)
                    objNode.Tag = "M"
                End If
            End If
            
            If strCodigoGrupo = "CTP" Then
                If cboEvento.ListIndex = -1 Then
                    strEvento = "E" & fgCompletaString(NET_CO_EVEN_BILAT, "0", 10, True)
                        
                    strMensagem = "M" & fgCompletaString(NET_CO_EVEN_BILAT, "0", 10, True) & fgCompletaString(NET_CO_EVEN_BILAT, " ", 9, True)
                    Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem & "TP" & Format(NET_CO_EVEN_BILAT, "0000"), Format(NET_CO_EVEN_BILAT, "0000") & "-" & NET_DE_EVEN_BILAT, 3)
                    objNode.Tag = "M"
                Else
                    If cboEvento.ItemData(cboEvento.ListIndex) = NET_CO_EVEN_BILAT Then
                        strEvento = "E" & fgCompletaString(NET_CO_EVEN_BILAT, "0", 10, True)
                            
                        strMensagem = "M" & fgCompletaString(NET_CO_EVEN_BILAT, "0", 10, True) & fgCompletaString(NET_CO_EVEN_BILAT, " ", 9, True)
                        Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem & "TP" & Format(NET_CO_EVEN_BILAT, "0000"), Format(NET_CO_EVEN_BILAT, "0000") & "-" & NET_DE_EVEN_BILAT, 3)
                        objNode.Tag = "M"
                    End If
                End If
            End If
                        
        Case enumEventosNetEntradaManual.CarregarSpread
            If strCodigoGrupo = "BMC" Then
                Call flCarregaSpreadMensagem("BMC", 132)
            Else
                Call flCarregaSpreadMensagem("CTP", 86)
            End If
            
            txtDescrticaoMensagem.Text = treMensagem.SelectedItem.Text
            txtDescrticaoMensagem.Tag = Val(Mid(treMensagem.SelectedItem.Key, 36, 10))
            
            If Val(fgObterCodigoCombo(txtDescrticaoMensagem.Text)) = NET_CO_EVEN_MULTI Then
                intModalidadeLiquidacao = enumTipoLiquidacao.Multilateral
                If strCodigoGrupo = "BMC" Then
                    strTagsVisiveisAux = NET_TAGS_VISIVEIS_MULTI_BMC
                Else
                    strTagsVisiveisAux = NET_TAGS_VISIVEIS_MULTI
                End If
            Else
                intModalidadeLiquidacao = enumTipoLiquidacao.Bilateral
                strTagsVisiveisAux = NET_TAGS_VISIVEIS_BILAT
            End If

            With sprMensagemSLCC
                
                .ReDraw = False
                For lngRow = 1 To .MaxRows
                    .GetText intColunmNomeFisico, lngRow, vntCellText
                    If InStr(1, strTagsVisiveisAux, vntCellText) = 0 And .RowHeight(lngRow) <> 0 Then
                        
                        .RowHeight(lngRow) = 0
                    
                        'Preenchimento de campos obrigatórios com valores default
                        .Col = intColunmNomeLogico
                        .Row = lngRow
                        If .FontBold Then
                            .Col = intColunmEdicao
                            If vntCellText = "CO_PROD" Then
                                .SetText intColunmEdicao, lngRow, 1
                            ElseIf vntCellText = "CO_SUB_TIPO_ATIV" Then
                                .TypeComboBoxCurSel = -1
                            ElseIf vntCellText = "NU_ATIV_MERC_CETIP" Then
                                .SetText intColunmEdicao, lngRow, vbNullString
                            ElseIf vntCellText = "NU_ATIV_MERC" Then
                                .SetText intColunmEdicao, lngRow, vbNullString
                            ElseIf vntCellText = "CO_PARP_CAMR" Then
                                .SetText intColunmEdicao, lngRow, vbNullString
                            ElseIf vntCellText = "TP_LIQU_OPER_ATIV" Then
                                .SetText intColunmEdicao, lngRow, intModalidadeLiquidacao
                            ElseIf vntCellText = "TP_CNPT" Then
                                .SetText intColunmEdicao, lngRow, enumTipoContraparte.Externo
                            ElseIf vntCellText = "CO_FORM_LIQU" Then
                                .TypeComboBoxCurSel = 1
                            End If
                        End If
                    
                        Call sprMensagemSLCC_Change(intColunmEdicao, sprMensagemSLCC.Row)
                    
                    End If
                Next
                .ReDraw = True
            
            End With
    
    End Select
            
End Sub

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao        As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    vntCodErro = 0
    
    Set xmlDominioTipoAtivMerc = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlDominioVeicLega = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmEntradaManual", "flInicializar")
    End If
    
    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmAtributo", "flInit", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Function flCarregarComboGrupoMensagem() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboGrupo.Clear
    
    vntCodErro = 0
    
    strMensagem = objMensagem.LerTodosGrupoMensagem(True, _
                                                    vntCodErro, _
                                                    vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboGrupoMensagem"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboGrupo.AddItem xmlNode.selectSingleNode("CO_GRUP").Text & " - " & xmlNode.selectSingleNode("NO_GRUP").Text
        cboGrupo.ItemData(cboGrupo.NewIndex) = xmlNode.selectSingleNode("SQ_GRUP").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboGrupoMensagem", 0
End Function

Private Function flCarregarComboServico() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboServico.Clear
    cboEvento.Clear
    vntCodErro = 0
    
    strMensagem = objMensagem.LerTodosServico(cboGrupo.ItemData(cboGrupo.ListIndex), _
                                              True, _
                                              vntCodErro, _
                                              vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboServico"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboServico.AddItem xmlNode.selectSingleNode("NO_SERV").Text
        cboServico.ItemData(cboServico.NewIndex) = xmlNode.selectSingleNode("SQ_SERV").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboServico", 0
    
End Function

Private Function flCarregarComboEvento() As Boolean

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim strMensagem                             As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    cboEvento.Clear
    vntCodErro = 0
    strMensagem = objMensagem.LerTodosEvento(cboServico.ItemData(cboServico.ListIndex), _
                                             True, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Function
    
    If Not xmlMensagem.loadXML(strMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregarComboEvento"
    End If
    
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        cboEvento.AddItem xmlNode.selectSingleNode("CO_EVEN").Text & " - " & xmlNode.selectSingleNode("NO_EVEN").Text
        cboEvento.ItemData(cboEvento.NewIndex) = xmlNode.selectSingleNode("SQ_EVEN").Text
    Next
    
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    Exit Function
ErrorHandler:
    Set objMensagem = Nothing
    Set xmlMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboEvento", 0
    
End Function

Private Sub flCarregaTreeViewMensagem(Optional plngSequenciaGrupo As Long, _
                                      Optional plngSequenciaServico As Long, _
                                      Optional plngSequenciaEvento As Long)
    
#If EnableSoap = 1 Then
    Dim objMensagem                             As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                             As A8MIU.clsMensagem
#End If

Dim xmlMensagem                                 As MSXML2.DOMDocument40
Dim xmlNode                                     As MSXML2.IXMLDOMNode
Dim strXMLMensagem                              As String

Dim objNode                                     As Node
Dim strGrupo                                    As String
Dim strServico                                  As String
Dim strEvento                                   As String
Dim strMensagem                                 As String
Dim strCodGrup                                  As String
Dim vntCodErro                                  As Variant
Dim vntMensagemErro                             As Variant

On Error GoTo ErrorHandler
    
    treMensagem.Nodes.Clear
        
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    strCodGrup = Trim$(Mid(cboGrupo.Text, 1, InStr(1, cboGrupo.Text, "-") - 1))
    vntCodErro = 0
    strXMLMensagem = objMensagem.LerTodosMensagem(plngSequenciaGrupo, _
                                                  plngSequenciaServico, _
                                                  plngSequenciaEvento, _
                                                  True, _
                                                  strCodGrup, _
                                                  vntCodErro, _
                                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strXMLMensagem = vbNullString Then
        flLimpaCampos
        Exit Sub
    End If
    
    If Not xmlMensagem.loadXML(strXMLMensagem) Then
        fgErroLoadXML xmlMensagem, App.EXEName, "frmEntradaManual", "flCarregaTreeViewMensagem"
    End If
        
    For Each xmlNode In xmlMensagem.documentElement.childNodes
        
        strGrupo = "G" & fgCompletaString(xmlNode.selectSingleNode("SQ_GRUP").Text, "0", 10, True)
        Set objNode = treMensagem.Nodes.Add(, tvwChild, "G" & strGrupo, xmlNode.selectSingleNode("CO_GRUP").Text, 1)
        objNode.Expanded = True
        objNode.Tag = "G"

        strServico = "S" & fgCompletaString(xmlNode.selectSingleNode("SQ_SERV").Text, "0", 10, True)
        Set objNode = treMensagem.Nodes.Add("G" & strGrupo, tvwChild, "S" & strGrupo & strServico, xmlNode.selectSingleNode("NO_SERV").Text, 2)
        objNode.Expanded = True
        objNode.Tag = "S"
        
        strEvento = "E" & fgCompletaString(xmlNode.selectSingleNode("SQ_EVEN").Text, "0", 10, True)
            
        strMensagem = "M" & fgCompletaString(xmlNode.selectSingleNode("TP_MESG_RECB_INTE").Text, "0", 10, True) & fgCompletaString(xmlNode.selectSingleNode("CO_MESG").Text, " ", 9, True)
        Set objNode = treMensagem.Nodes.Add("S" & strGrupo & strServico, tvwChild, "M" & strGrupo & strServico & strEvento & strMensagem & "TP" & Format(xmlNode.selectSingleNode("TP_OPER").Text, "0000"), Format(xmlNode.selectSingleNode("TP_OPER").Text, "0000") & "-" & xmlNode.selectSingleNode("NO_TIPO_OPER").Text, 3)
        objNode.Tag = "M"
    
    Next
    
    Exit Sub
ErrorHandler:
    
    If Err.Number = 35602 Then
         Resume Next
    End If
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub flEnviarMensagem()

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strMensagem                             As String
Dim strRetorno                              As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    strRetorno = flValidarCampos()
        
    If strRetorno <> "" Then
        frmMural.Caption = Me.Caption
        frmMural.Display = strRetorno
        frmMural.Show vbModal
        Exit Sub
    End If

    strMensagem = flMontaMensagem
    
    fgCursor True
    DoEvents
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    vntCodErro = 0
    strRetorno = objMensagem.EnviarMensagem(strMensagem, _
                                            enumMensagemEntradaManual.TratadaSLCC, _
                                            0, _
                                            0, _
                                            0, _
                                            vbNullString, _
                                            vbNullString, _
                                            vbNullString, _
                                            strChaveMensagemBMC0112, _
                                            vntCodErro, _
                                            vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strRetorno <> vbNullString Then
        'Codificar a exibição da mensagem para o usuário
        Call flExibirErros(strRetorno)
    Else
        MsgBox "Operação gerada com sucesso.", vbInformation, "Entrada Manual"
        flLimpaCampos
    End If
    
    Set objMensagem = Nothing
    
    fgCursor False
    
    Exit Sub
ErrorHandler:
        
    fgCursor False
    Set objMensagem = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flEnviarMensagem", 0
    
End Sub

Private Sub flExibirErros(ByVal pstrXmlErros As String)

Dim xmlDomErros                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strTextoErro                            As String

On Error GoTo ErrorHandler

    Set xmlDomErros = CreateObject("MSXML2.DOMDocument.4.0")
    xmlDomErros.loadXML pstrXmlErros
    
    strTextoErro = "**************************************************" & vbCrLf & _
                   "         Ocorreram erros no processamento         " & vbCrLf
    
    For Each objDomNode In xmlDomErros.documentElement.selectNodes("Grupo_ErrorInfo")
        
        If Not objDomNode.selectSingleNode("Numero") Is Nothing Then
        
            strTextoErro = strTextoErro & _
                objDomNode.selectSingleNode("Numero").Text & " - " & objDomNode.selectSingleNode("Descricao").Text & vbCrLf
        Else
            strTextoErro = strTextoErro & _
               objDomNode.selectSingleNode("Number").Text & " - " & objDomNode.selectSingleNode("Description").Text & vbCrLf
        
        End If
            
    Next objDomNode
    
    strTextoErro = strTextoErro & _
                   "**************************************************"
    
    Set xmlDomErros = Nothing
    
    With frmMural
        .Caption = Me.Caption
        .Display = strTextoErro
        .Show vbModal
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flExibirErros", 0

End Sub

Private Function flCarregarComboEmpresa() As Boolean

Dim xmlNode                                 As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    For Each xmlNode In xmlMapaNavegacao.selectSingleNode("frmEntradaManual/Grupo_Dados/Repeat_Empresa").childNodes
        cboEmpresa.AddItem xmlNode.selectSingleNode("CO_EMPR").Text & " - " & xmlNode.selectSingleNode("NO_REDU_EMPR").Text
        cboEmpresa.ItemData(cboEmpresa.NewIndex) = CLng(xmlNode.selectSingleNode("CO_EMPR").Text)
    Next
        
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmEntradaManual", "flCarregarComboEmpresa", 0
    
End Function

Private Sub flCarregarCampoDominioPadrao(ByVal plngPosicaoPai As Long, _
                                         ByVal pstrNomeCampo As String, _
                                         ByVal pstrValor As String)

Dim strValor                                As String
Dim strXPath                                As String
Dim lngPosicaoAlterada                      As Long

On Error GoTo ErrorHandler

    Select Case pstrNomeCampo
        Case "CO_VEIC_LEGA"
            'Preencher o campo CO_CNTA_CUTD_SELIC_VEIC_LEGA
            strXPath = "//*[@Posicao='" & plngPosicaoPai & "']/CO_CNTA_CUTD_SELIC_VEIC_LEGA"
            pstrValor = Trim$(Mid$(pstrValor, 1, InStr(1, pstrValor, "-") - 1))
            If Not xmlMsgEntradaManual.selectSingleNode(strXPath) Is Nothing Then
               
                If Not xmlDominioVeicLega.selectSingleNode("//Grupo_DominioTabela[./CODIGO='" & pstrValor & "' and ./SG_SIST='" & Left$(Trim$(gstrSiglaSistema) & "   ", 3) & "']") Is Nothing Then
                    strValor = xmlDominioVeicLega.selectSingleNode("//Grupo_DominioTabela[./CODIGO='" & pstrValor & "' and ./SG_SIST='" & Left$(Trim$(gstrSiglaSistema) & "   ", 3) & "']/CO_CNTA_CUTD_PADR_SELIC").Text
                Else
                    strValor = vbNullString
                End If
                xmlMsgEntradaManual.selectSingleNode(strXPath).Text = strValor
                
                lngPosicaoAlterada = CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@Posicao").Text) - 1
                
                sprMensagemSLCC.SetText intColunmEdicao, lngPosicaoAlterada, strValor
                            
            End If
            
        Case "CO_CNTA_CUTD_SELIC_VEIC_LEGA"
            'Preencher o campo CO_VEIC_LEGA
            
            strXPath = "//*[@Posicao='" & plngPosicaoPai & "']/CO_VEIC_LEGA"
            If Not xmlMsgEntradaManual.selectSingleNode(strXPath) Is Nothing Then
               
                If Not xmlDominioVeicLega.selectSingleNode("//Grupo_DominioTabela[./CO_CNTA_CUTD_PADR_SELIC='" & pstrValor & "']") Is Nothing Then
                    strValor = xmlDominioVeicLega.selectSingleNode("//Grupo_DominioTabela[./CO_CNTA_CUTD_PADR_SELIC='" & pstrValor & "']/CODIGO").Text & " - " & _
                               xmlDominioVeicLega.selectSingleNode("//Grupo_DominioTabela[./CO_CNTA_CUTD_PADR_SELIC='" & pstrValor & "']/DESCRICAO").Text
                    xmlMsgEntradaManual.selectSingleNode(strXPath).Text = strValor
                    
                    lngPosicaoAlterada = CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@Posicao").Text) - 1
                    
                    sprMensagemSLCC.SetText intColunmEdicao, lngPosicaoAlterada, strValor
                
                End If
                            
            End If
            
        Case "NU_ATIV_MERC"
            ''Preencher o campo DE_ATIV_MERC
            strXPath = "//*[@Posicao='" & plngPosicaoPai & "']/DE_ATIV_MERC"
            If Not xmlMsgEntradaManual.selectSingleNode(strXPath) Is Nothing Then
               
                strValor = vbNullString
                If Not xmlDominioTipoAtivMerc.selectSingleNode("//Grupo_DominioTabela[./NU_ATIV_MERC='" & pstrValor & "']") Is Nothing Then
                    strValor = Left$(xmlDominioTipoAtivMerc.selectSingleNode("//Grupo_DominioTabela[./NU_ATIV_MERC='" & pstrValor & "']/DE_ATIV_MERC").Text, _
                                    CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@QT_CTER_ATRB").Text))
                End If
                                    
                xmlMsgEntradaManual.selectSingleNode(strXPath).Text = strValor
                lngPosicaoAlterada = CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@Posicao").Text) - 1
                sprMensagemSLCC.SetText intColunmEdicao, lngPosicaoAlterada, strValor
                
            End If
        Case "PU_ATIV_MERC", "QT_ATIV_MERC"
            'Calcular o valor financeiro caso haja os campos
            ' - PU_ATIV_MERC
            ' - QT_ATIV_MERC
            ' - VA_OPER_ATIV
        
            If Not xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/PU_ATIV_MERC") Is Nothing And _
               Not xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/QT_ATIV_MERC") Is Nothing And _
               Not xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/VA_OPER_ATIV") Is Nothing Then
                
                strValor = CStr(CDbl(xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/PU_ATIV_MERC").Text) * _
                                CDbl(xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/QT_ATIV_MERC").Text))
            
                'Trunca o valor.
                If InStrRev(strValor, ",") > 0 Then
                    strValor = Mid$(strValor, 1, InStrRev(strValor, ",") + 2)
                ElseIf InStrRev(strValor, ".") > 0 Then
                    'Apenas assegura o funcionamento caso a estação esteja configurada com "." na casa decimal
                    strValor = Mid$(strValor, 1, InStrRev(strValor, ".") + 2)
                End If
            
                xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/VA_OPER_ATIV").Text = strValor
                
                lngPosicaoAlterada = CLng(xmlMsgEntradaManual.selectSingleNode("//*[@Posicao='" & plngPosicaoPai & "']/VA_OPER_ATIV/@Posicao").Text) - 1
                    
                sprMensagemSLCC.SetText intColunmEdicao, lngPosicaoAlterada, strValor
            
            End If
        
    End Select
    
Exit Sub
ErrorHandler:

End Sub

Private Sub flObterDominiosTabelas()

#If EnableSoap = 1 Then
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem         As A8MIU.clsMensagem
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    vntCodErro = 0
    xmlDominioVeicLega.loadXML objMensagem.LerTodosDominioTabela(cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                                                 "A8.TB_VEIC_LEGA", _
                                                                 "", _
                                                                 "", _
                                                                 "", _
                                                                 vntCodErro, _
                                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    xmlDominioTipoAtivMerc.loadXML objMensagem.LerTodosDominioTabela(cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                                                     "A8.TB_TIPO_ATIV_MERC", _
                                                                     "", _
                                                                     "", _
                                                                     "", _
                                                                     vntCodErro, _
                                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing

Exit Sub
ErrorHandler:
    Set objMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

    fgRaiseError App.EXEName, "frmEntradaManual", "flObterDominiosTabelas", 0
End Sub

Private Sub cboEmpresa_Click()

On Error GoTo ErrorHandler

    flLimpaCampos
    flObterDominiosTabelas

Exit Sub
ErrorHandler:
    
    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboEmpresa_Change", Me.Caption

End Sub

Private Sub cboEvento_Click()

On Error GoTo ErrorHandler

    If cboEvento.ListIndex = -1 Then Exit Sub
    
    Call flLimpaCampos
    
    fgCursor True
    
    Call flCarregaTreeViewMensagem(cboGrupo.ItemData(cboGrupo.ListIndex), _
                                   cboServico.ItemData(cboServico.ListIndex), _
                                   cboEvento.ItemData(cboEvento.ListIndex))
    Call flTratarNetEntradaManual(CarregarTreeView)

    fgCursor False
    
Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboEvento_Click", Me.Caption

End Sub

Private Sub cboGrupo_Click()
    
On Error GoTo ErrorHandler

    If cboGrupo.ListIndex = -1 Then Exit Sub
        
    Call flLimpaCampos
        
    fgCursor True
        
    Call flCarregarComboServico
    Call flTratarNetEntradaManual(CarregarComboServicos)
    
    Call flCarregaTreeViewMensagem(cboGrupo.ItemData(cboGrupo.ListIndex))
    Call flTratarNetEntradaManual(CarregarTreeView)
    
    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboGrupo_Click", Me.Caption
    
End Sub

Private Sub cboServico_Click()

    On Error GoTo ErrorHandler

    If cboServico.ListIndex = -1 Then Exit Sub
    
    Call flLimpaCampos
    
    fgCursor True
    
    Call flCarregarComboEvento
    Call flTratarNetEntradaManual(CarregarComboEventos)
    
    Call flCarregaTreeViewMensagem(cboGrupo.ItemData(cboGrupo.ListIndex), _
                                   cboServico.ItemData(cboServico.ListIndex))
    Call flTratarNetEntradaManual(CarregarTreeView)
    
    fgCursor False
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - cboServico_Click", Me.Caption
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiLQS.Icon
    
    fgCenterMe Me
    
    Me.Show
    DoEvents
    
    strSepDecimal = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SDecimal")
    strSepMilhar = GetKeyValue(HKEY_CURRENT_USER, "Control Panel\International", "SThousand")
    
    sprMensagemSLCC.EditEnterAction = EditEnterActionSame
    flLimpaCampos
    
    flInicializar
    
    flCarregarComboEmpresa
    flCarregarComboGrupoMensagem
            
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlMensagemBase = Nothing
    Set xmlMsgEntradaManual = Nothing
    Set xmlDominioTipoAtivMerc = Nothing
    Set xmlDominioVeicLega = Nothing

End Sub

Private Sub sprMensagemSLCC_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Dim objFrmDominioTabela                     As frmDominioTabela
Dim vntNomeTabela                           As Variant
Dim vntRetorno                              As String
Dim arrRetorno()                            As String
Dim strMascara                              As String

On Error GoTo ErrorHandler

    sprMensagemSLCC.GetText intColunmTabela, Row, vntNomeTabela

    If Trim(vntNomeTabela) = "" Then Exit Sub

    Set objFrmDominioTabela = New frmDominioTabela

    'Carrega campos para a Tela de Dominio Tabela
    objFrmDominioTabela.strNomeTabela = CStr(vntNomeTabela)
    objFrmDominioTabela.lngCodigoEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
    
    'Carrega campos para a Tela de Dominio Tabela de acordo com a Operacao
    If Val(Mid(treMensagem.SelectedItem.Text, 1, 4)) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Then
        objFrmDominioTabela.strMensagem = "CAM0006R2"
    ElseIf Val(Mid(treMensagem.SelectedItem.Text, 1, 4)) = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
        objFrmDominioTabela.strMensagem = "CAM0009R2"
    ElseIf Val(Mid(treMensagem.SelectedItem.Text, 1, 4)) = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
        objFrmDominioTabela.strMensagem = "BMC0015"
    ElseIf Val(Mid(treMensagem.SelectedItem.Text, 1, 4)) = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
        objFrmDominioTabela.strMensagem = "CAM0013R2"
    Else
        objFrmDominioTabela.strMensagem = ""
    End If

    objFrmDominioTabela.Show vbModal
    
    If Not objFrmDominioTabela.blnCancel Then
    
        If objFrmDominioTabela.lstDominio.SelectedItem Is Nothing Then
            vntRetorno = ""
        Else
            If vntNomeTabela = "A8.TB_MESG_RECB_ENVI_SPB" Then
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.Text
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowControleBMC, vntRetorno)
                
                'Atualizar o Valor do XML Auxiliar
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowControleBMC
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowControleBMC + 1 & "']").Text = vntRetorno
                
                flCarregarCampoDominioPadrao CLng(xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowControleBMC + 1 & "']").parentNode.selectSingleNode("./@Posicao").Text), _
                                             sprMensagemSLCC.Text, _
                                             vntRetorno
                
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.SubItems(2)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValor, vntRetorno)
                 
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValor
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValor + 1 & "']").Text = vntRetorno
                
                flCarregarCampoDominioPadrao CLng(xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValor + 1 & "']").parentNode.selectSingleNode("./@Posicao").Text), _
                                             sprMensagemSLCC.Text, _
                                             vntRetorno
                 
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.SubItems(1)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowVeiculoLegal, vntRetorno)
                 
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowVeiculoLegal
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowVeiculoLegal + 1 & "']").Text = vntRetorno
                
                flCarregarCampoDominioPadrao CLng(xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowVeiculoLegal + 1 & "']").parentNode.selectSingleNode("./@Posicao").Text), _
                                             sprMensagemSLCC.Text, _
                                             vntRetorno
                 
                strChaveMensagemBMC0112 = objFrmDominioTabela.lstDominio.SelectedItem.Key
                
            ElseIf vntNomeTabela = "ChACAM" Then
                ReDim arrRetorno(7)
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                
                'Atualizar Chacam
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.Text
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowChacam, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowChacam
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowChacam + 1 & "']").Text = vntRetorno
                                             
                'Atualizar Data Operacao
                vntRetorno = arrRetorno(2)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowDataOperacao, Format(fgDtXML_To_Interface(vntRetorno), "DD-MM-YYYY"))
                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowDataOperacao
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowDataOperacao + 1 & "']").Text = vntRetorno
                                             
                'Atualizar Tipo Operacao Cambio
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(3)
                If vntRetorno = "C" Then
                    Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowTipoOperacaoCambio, "C - Compra")
                ElseIf vntRetorno = "V" Then
                    Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowTipoOperacaoCambio, "V - Venda")
                End If
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowTipoOperacaoCambio
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowTipoOperacaoCambio + 1 & "']").Text = vntRetorno
                
                'Atualizar Valor Taxa Cambio
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(4)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorTaxaCambio, Format(vntRetorno, "#,##0.00000000000"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorTaxaCambio
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorTaxaCambio + 1 & "']").Text = vntRetorno
                                             
                'Atualizar Valor Moeda Nacional
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(5)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorMoedaNac, Format(vntRetorno, "##,###,###,###,###,##0.00"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorMoedaNac
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorMoedaNac + 1 & "']").Text = vntRetorno
                                             
                'Atualizar Valor Moeda Estrangeira
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(6)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorMoedaEstr, Format(vntRetorno, "##,###,###,###,###,##0.00"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorMoedaEstr
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorMoedaEstr + 1 & "']").Text = vntRetorno
                
                'Atualizar Data Liquidacao
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(7)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowDataLiquOper, Format(fgDtXML_To_Interface(vntRetorno), "DD-MM-YYYY"))
                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowDataLiquOper
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowDataLiquOper + 1 & "']").Text = vntRetorno
            
            ElseIf vntNomeTabela = "CO_REG_OPER_CAMB" Then
                'Atualizar Registro Operacao Cambial
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.Text
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowRegistroOperCamb, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowRegistroOperCamb
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowRegistroOperCamb + 1 & "']").Text = vntRetorno
                
                'Atualizar CNPJ Base IF
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(3)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowCnpjBaseIf, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowCnpjBaseIf
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowCnpjBaseIf + 1 & "']").Text = vntRetorno
                
                'Atualizar Codigo Moeda ISO
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(4)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowMoedIso, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowMoedIso
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowMoedIso + 1 & "']").Text = vntRetorno
                
                'Atualizar Valor Moeda Estrangeira
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(5)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorMoedaEstr, Format(vntRetorno, "##,###,###,###,###,##0.00"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorMoedaEstr
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorMoedaEstr + 1 & "']").Text = vntRetorno
                
                'Atualizar Valor Taxa Cambio
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(6)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorTaxaCambio, Format(vntRetorno, "#,##0.00000000000"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorTaxaCambio
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorTaxaCambio + 1 & "']").Text = vntRetorno
                
                'Atualizar Valor Moeda Nacional
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(7)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowValorMoedaNac, Format(vntRetorno, "##,###,###,###,###,##0.00"))
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowValorMoedaNac
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowValorMoedaNac + 1 & "']").Text = vntRetorno
                
                'Atualizar Data Entrada Moeda Nacional
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(8)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowDataEntrMoedNac, Format(fgDtXML_To_Interface(vntRetorno), "DD-MM-YYYY"))
                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowDataEntrMoedNac
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowDataEntrMoedNac + 1 & "']").Text = vntRetorno
                
                'Atualizar Data Entrada Moeda Estrangeira
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(9)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowDataEntrMoedEstr, Format(fgDtXML_To_Interface(vntRetorno), "DD-MM-YYYY"))
                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowDataEntrMoedEstr
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowDataEntrMoedEstr + 1 & "']").Text = vntRetorno
                
                'Atualizar Data Liquidacao
                arrRetorno = Split(objFrmDominioTabela.lstDominio.SelectedItem.Key, "|")
                vntRetorno = arrRetorno(10)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowDataLiquOper, Format(fgDtXML_To_Interface(vntRetorno), "DD-MM-YYYY"))
                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowDataLiquOper
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowDataLiquOper + 1 & "']").Text = vntRetorno
            
            ElseIf vntNomeTabela = "CO_REG_OPER_CAMB2" Then
                'Atualizar Registro Operacao Cambial
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.Text
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowRegistroOperCamb, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowRegistroOperCamb
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowRegistroOperCamb + 1 & "']").Text = vntRetorno
                                             
                'Atualizar Registro Operacao Cambial 2
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.SubItems(1)
                Call sprMensagemSLCC.SetText(intColunmEdicao, lngRowRegistroOperCamb2, vntRetorno)
                sprMensagemSLCC.Col = intColunmNomeFisico
                sprMensagemSLCC.Row = lngRowRegistroOperCamb2
                xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & lngRowRegistroOperCamb2 + 1 & "']").Text = vntRetorno
                
            Else
                vntRetorno = objFrmDominioTabela.lstDominio.SelectedItem.Text & " - " & objFrmDominioTabela.lstDominio.SelectedItem.SubItems(1)
                
            End If
            
            If Trim(objFrmDominioTabela.lstDominio.SelectedItem.Tag) <> "" Then
                gstrSiglaSistema = Trim(Mid(objFrmDominioTabela.lstDominio.SelectedItem.Tag, 1, 3))
                
                If Mid(objFrmDominioTabela.lstDominio.SelectedItem.Tag, 4, 1) = enumIndicadorSimNao.Sim Then
                    sprMensagemSLCC.Height = 6225
                    fraContingencia.Caption = "Sistema " & gstrSiglaSistema & " em contingência"
                    fraContingencia.Visible = True
                    chkPrevistoA6.value = vbUnchecked
                    chkPrevisaoPJ.value = vbUnchecked
                    chkRealizadoA6.value = vbUnchecked
                    chkRealizadoPJ.value = vbUnchecked
                Else
                    sprMensagemSLCC.Height = 6790
                    fraContingencia.Caption = ""
                    fraContingencia.Visible = False
                    chkPrevistoA6.value = vbUnchecked
                    chkPrevisaoPJ.value = vbUnchecked
                    chkRealizadoA6.value = vbUnchecked
                    chkRealizadoPJ.value = vbUnchecked
                End If
            End If
        End If
            
        If vntNomeTabela <> "A8.TB_MESG_RECB_ENVI_SPB" _
        And vntNomeTabela <> "ChACAM" _
        And vntNomeTabela <> "CO_REG_OPER_CAMB" _
        And vntNomeTabela <> "CO_REG_OPER_CAMB2" Then
            sprMensagemSLCC.SetText intColunmEdicao, Row, vntRetorno
            
            'Atualizar o Valor do XML Auxiliar
            sprMensagemSLCC.Col = intColunmNomeFisico
            sprMensagemSLCC.Row = Row
            xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & Row + 1 & "']").Text = vntRetorno
            
            flCarregarCampoDominioPadrao CLng(xmlMsgEntradaManual.selectSingleNode("//" & sprMensagemSLCC.Text & "[@Posicao='" & Row + 1 & "']").parentNode.selectSingleNode("./@Posicao").Text), _
                                         sprMensagemSLCC.Text, _
                                         vntRetorno
        
        End If
    End If
    
    Unload objFrmDominioTabela
        
    Set objFrmDominioTabela = Nothing

Exit Sub
ErrorHandler:

    Set objFrmDominioTabela = Nothing

    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - sprMensagemSLCC_ButtonClicked", Me.Caption

End Sub

Private Sub sprMensagemSLCC_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo ErrorHandler

Dim strNomeTag                              As String
Dim lngQtdAtual                             As Long
Dim lngQtdNova                              As Long
Dim strValorAlterado                        As String
    
Dim strXPath                                As String
Dim lngAux                                  As Long
    
Dim xmlNodeAux                              As IXMLDOMNode
    
    'Alterar o valor da tag no documento xmlmsgentradamanual
    'Caso a tag alterada for uma repetição, alterar o valor e carregar novamente o Spread
    
    sprMensagemSLCC.Row = Row
    sprMensagemSLCC.Col = Col
    strValorAlterado = sprMensagemSLCC.Text
    
    sprMensagemSLCC.Col = intColunmNomeFisico
    strNomeTag = sprMensagemSLCC.Text
    
    strXPath = "//" & strNomeTag & "[@Posicao='" & Row + 1 & "']"
    
    If xmlMsgEntradaManual.selectSingleNode(strXPath & "/child::*") Is Nothing And _
       CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@QT_REPE_MAX").Text) = 0 Then
        'Somente altera o valor
        
        xmlMsgEntradaManual.selectSingleNode(strXPath).Text = strValorAlterado
        xmlMsgEntradaManual.selectSingleNode(strXPath).selectSingleNode("./@Alterado").Text = 1
        
        flCarregarCampoDominioPadrao CLng(xmlMsgEntradaManual.selectSingleNode(strXPath).parentNode.selectSingleNode("./@Posicao").Text), _
                                     strNomeTag, _
                                     strValorAlterado

        strXPath = "//" & strNomeTag & "[@Posicao='" & Row + 1 & "']"
        
    Else
        'Alterar o valor e carregar novamente o Spread
        lngQtdAtual = xmlMsgEntradaManual.selectNodes(strXPath & "/child::*").length
        If strValorAlterado <> vbNullString Then
            lngQtdNova = CLng(strValorAlterado)
        Else 'quando a linha é de Repetição, a Col lida tem que ser a =5(intColunmButton) para obter a quantidade de repetições informada pelo usuário (Bruno Oliveira - 09/09/2011 - RATS 1084)
           sprMensagemSLCC.Col = intColunmButton
           strValorAlterado = sprMensagemSLCC.Text
           lngQtdNova = CLng(strValorAlterado)
           sprMensagemSLCC.Col = intColunmNomeFisico
        End If
        
        If CLng(xmlMsgEntradaManual.selectSingleNode(strXPath & "/@IN_OBRI_ATRB").Text) = 1 And _
           lngQtdNova = 0 Then
           'Não permitir alteração. repetição Obrigatória
           
           sprMensagemSLCC.SetText Col, Row, lngQtdAtual
           Exit Sub
        End If
        
        xmlMsgEntradaManual.selectSingleNode(strXPath & "/@QT_REPE_MSG").Text = lngQtdNova
        
        If lngQtdAtual > CLng(strValorAlterado) Then
            'Excluir tags filho
            For lngAux = lngQtdAtual To lngQtdNova + 1 Step -1
                
                xmlMsgEntradaManual.selectSingleNode(strXPath).removeChild _
                        xmlMsgEntradaManual.selectSingleNode(strXPath & "/child::*[position()=last()]")
            
            Next
        Else
            'Incluir nodes filho
            'Obtem o node a ser incluido
            For lngAux = lngQtdAtual + 1 To lngQtdNova
                For Each xmlNodeAux In xmlMensagemBase.selectNodes("//" & strNomeTag & "/*")
                    fgAppendXML xmlMsgEntradaManual, _
                                Replace$(strXPath, "//", vbNullString), _
                                xmlNodeAux.xml, _
                                Replace$(strXPath, "//", vbNullString)
                                
                    'xmlMsgEntradaManual.selectSingleNode(strXPath).appendChild xmlNodeAux
                Next
            Next
        End If
        
        sprMensagemSLCC.ReDraw = False
        sprMensagemSLCC.MaxRows = 0
        
        flRecalcularPosicoes xmlMsgEntradaManual
        
        Call flCarregaSpread(xmlMsgEntradaManual.selectNodes("//XML/*"))
        
        sprMensagemSLCC.Row = Row
        sprMensagemSLCC.Action = ActionGotoCell
                
    End If

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - sprMensagemSLCC_Change", Me.Caption
End Sub

Private Sub flRecalcularPosicoes(ByRef pxmlMensagem As MSXML2.DOMDocument40)

On Error GoTo ErrorHandler

Dim xmlElement                              As IXMLDOMElement
Dim lngX                                    As Long

    lngX = 0
    For Each xmlElement In pxmlMensagem.selectNodes("//*")
        
        lngX = lngX + 1
        If Not xmlElement.selectSingleNode("./@Posicao") Is Nothing Then
            xmlElement.selectSingleNode("./@Posicao").Text = lngX
        Else
            xmlElement.setAttribute "Posicao", lngX
        End If
    Next

Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, "frmEntradaManual", "flRecalcularPosicoes", 0
End Sub

Private Sub sprMensagemSLCC_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler

Dim strNomeTag                              As String
Dim strXPath                                As String
    
    With sprMensagemSLCC
        
        .Col = .ActiveCol
        .Row = .ActiveRow
        
        'If .CellType = CellTypeComboBox Then
                    
            If KeyCode = 46 Then
                
                If .BackColor <> 12648447 Then
                
                    Call .SetText(.ActiveCol, .ActiveRow, "")
                    
                    'Atualiza o valor no XML da Mensagem da Entrada Manual
                    sprMensagemSLCC.Col = intColunmNomeFisico
                    strNomeTag = sprMensagemSLCC.Text
    
                    strXPath = "//" & strNomeTag & "[@Posicao='" & .ActiveRow + 1 & "']"
    
                    xmlMsgEntradaManual.selectSingleNode(strXPath).Text = ""
                    xmlMsgEntradaManual.selectSingleNode(strXPath).selectSingleNode("./@Alterado").Text = 1
                
                End If
            End If
        'End If
    End With

Exit Sub
ErrorHandler:

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Enviar"
            sprMensagemSLCC_Change intColunmEdicao, sprMensagemSLCC.Row
            flEnviarMensagem
        Case gstrSalvar
            sprMensagemSLCC_Change intColunmEdicao, sprMensagemSLCC.Row
            flSalvarPadrao
        Case "Limpar"
            blnLimpar = True
            treMensagem_NodeClick treMensagem.SelectedItem
            blnLimpar = False
        Case gstrSair
            Unload Me
    End Select
        
Exit Sub
ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - tlbCadastro_ButtonClick", Me.Caption

End Sub

Private Sub treMensagem_NodeClick(ByVal Node As MSComctlLib.Node)

Dim llCodigoGrupo                           As Long
Dim llCodigoServico                         As Long
Dim llCodigoEvento                          As Long
Dim llCodigoMensagem                        As Long
Dim strCodigoMensagem                       As String

On Error GoTo ErrorHandler
    
    fgCursor True
    
    flLimpaCampos
    
    If Node Is Nothing Then
        fgCursor
        Exit Sub
    End If
    
    DoEvents
   
    Select Case Left$(Node.Key, 1)
        
        Case "G"
            fgSearchItemCombo cboGrupo, Val(Mid(Node.Key, 3, 10))
        Case "S"
            fgSearchItemCombo cboServico, Val(Mid(Node.Key, 14, 10))
        Case "E"
            fgSearchItemCombo cboEvento, Val(Mid(Node.Key, 25, 10))
        Case "M"
            llCodigoGrupo = Val(Mid(Node.Key, 3, 10))
            llCodigoServico = Val(Mid(Node.Key, 14, 10))
            llCodigoEvento = Val(Mid(Node.Key, 25, 10))
            llCodigoMensagem = Val(Mid(Node.Key, 36, 10))
            txtDescrticaoMensagem.Tag = Val(Mid(Node.Key, 36, 10))
            strCodigoMensagem = Trim(Mid(Node.Key, 46, 9))
            Call flCarregaSpreadMensagem(strCodigoMensagem, llCodigoMensagem)
                        
            tlbCadastro.Buttons("Salvar").Enabled = True
            
            Call flTratarNetEntradaManual(CarregarSpread)
    
    End Select
        
    sprMensagemSLCC.SetFocus
            
    fgCursor False

Exit Sub
ErrorHandler:
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - treMensagem_NodeClick", Me.Caption

End Sub

Private Sub flCarregaSpreadMensagem(ByVal pstrCodigoMensagem As String, _
                           Optional ByVal plngTipoMensagem As Long = 0, _
                           Optional ByVal plngTipoOperacao As Long = 0)

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim strMensagem                             As String
Dim xmlTipoMensagem                         As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim lngTamanhoAtributo                      As Long
Dim lngCasasDecimais                        As Long
Dim strNomeAtributo                         As String
Dim strMasacara                             As String
Dim strDominio                              As String
Dim vntMascaraData                          As Variant
Dim strNomeLogicoAtributo                   As String
Dim plngCodigoEmpresa                       As Long
Dim lngTipoOperacao                         As Long
Dim strParametro                            As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If cboEmpresa.ListIndex = -1 Then
        frmMural.Display = "Selecione a empresa."
        frmMural.Show vbModal
        Exit Sub
    End If
    
    Set xmlTipoMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgEntradaManual = Nothing
    Set xmlMsgEntradaManual = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    plngCodigoEmpresa = cboEmpresa.ItemData(cboEmpresa.ListIndex)
    lngTipoOperacao = Val(Mid(treMensagem.SelectedItem.Text, 1, 4))
    vntCodErro = 0
    strMensagem = objMensagem.LerMensagem(pstrCodigoMensagem, _
                                          plngTipoMensagem, _
                                          plngCodigoEmpresa, _
                                          vntCodErro, _
                                          vntMensagemErro, _
                                          lngTipoOperacao)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strMensagem = "" Then Exit Sub
    
    xmlTipoMensagem.loadXML strMensagem
    
    xmlMsgEntradaManual.loadXML "<XML NO_ATRB_MESG='XML' QT_REPE='0'></XML>"
    
    Call flMontarMensagemXML_X(xmlTipoMensagem.documentElement.selectNodes("//Repeat_Mensagem/Grupo_Mensagem[TP_FORM_MESG='2']"), _
                               xmlMsgEntradaManual)
        
    Set xmlMensagemBase = New DOMDocument40
    xmlMensagemBase.loadXML xmlMsgEntradaManual.xml
            
    If chkParametro.value = vbChecked Then
        strParametro = flObterPadrao()
            
        If Trim$(strParametro) <> vbNullString Then
            xmlMsgEntradaManual.loadXML strParametro
            If Not xmlMsgEntradaManual.selectSingleNode("//XML/@SistemaOrigem") Is Nothing Then
                gstrSiglaSistema = xmlMsgEntradaManual.selectSingleNode("//XML/@SistemaOrigem").Text
            End If
        Else
            flRecalcularPosicoes xmlMsgEntradaManual
        End If
    Else
        flRecalcularPosicoes xmlMsgEntradaManual
    End If
    
    sprMensagemSLCC.ReDraw = False
    sprMensagemSLCC.MaxRows = 0
                
    txtDescrticaoMensagem.Text = xmlMsgEntradaManual.documentElement.selectSingleNode("//@TP_MESG").Text & _
                                 " - " & _
                                 xmlMsgEntradaManual.documentElement.selectSingleNode("//@NO_TIPO_MESG").Text
                                 
    txtDescrticaoMensagem.Tag = xmlMsgEntradaManual.documentElement.selectSingleNode("//@TP_MESG").Text
    
    'Operacao 236
    If CLng("0" & Right$(treMensagem.SelectedItem.Key, 4)) = enumTipoOperacaoLQS.InformaOperacaoArbitragemParceiroPais Then
        
        'Exclui Repeticao Parceiro Exterior
        If xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']").hasChildNodes = True Then
            xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']").removeChild xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']" & "/child::*[position()=last()]")
            flRecalcularPosicoes xmlMsgEntradaManual
            xmlMsgEntradaManual.selectNodes("//XML/REPE_PARC_EXTR/@QT_REPE_MSG").Item(0).Text = 0
        End If
        
    End If
    
    'Operacao 237
    If CLng("0" & Right$(treMensagem.SelectedItem.Key, 4)) = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
        
        'Exclui Repeticao Parceiro Exterior
        If xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']").hasChildNodes = True Then
            xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']").removeChild xmlMsgEntradaManual.selectSingleNode("//REPE_PARC_EXTR[@Posicao='18']" & "/child::*[position()=last()]")
            flRecalcularPosicoes xmlMsgEntradaManual
            xmlMsgEntradaManual.selectNodes("//XML/REPE_PARC_EXTR/@QT_REPE_MSG").Item(0).Text = 0
        End If
        
        'Exclui Repeticao Contrato
        If xmlMsgEntradaManual.selectSingleNode("//REPE_CONTR[@Posicao='19']").hasChildNodes = True Then
            xmlMsgEntradaManual.selectSingleNode("//REPE_CONTR[@Posicao='19']").removeChild xmlMsgEntradaManual.selectSingleNode("//REPE_CONTR[@Posicao='19']" & "/child::*[position()=last()]")
            flRecalcularPosicoes xmlMsgEntradaManual
            xmlMsgEntradaManual.selectNodes("//XML/REPE_CONTR/@QT_REPE_MSG").Item(0).Text = 0
        End If
        
    End If
    
    Call flCarregaSpread(xmlMsgEntradaManual.selectNodes("//XML/*"))
        
    sprMensagemSLCC.ReDraw = True
    
    Set objMensagem = Nothing
    Set xmlTipoMensagem = Nothing
    
Exit Sub
ErrorHandler:
    fgCursor False
    
    Set xmlTipoMensagem = Nothing
    Set xmlMsgEntradaManual = Nothing
    Set objMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    mdiLQS.uctlogErros.MostrarErros Err, "flCarregaSpreadMensagem", Me.Caption

End Sub

Private Sub flCarregaSpread(ByRef pxmlNodeList As IXMLDOMNodeList)

Dim strMensagem                             As String
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim lngTamanhoAtributo                      As Long
Dim lngCasasDecimais                        As Long
Dim strNomeAtributo                         As String
Dim strMasacara                             As String
Dim strDominio                              As String
Dim vntMascaraData                          As Variant
Dim strNomeLogicoAtributo                   As String
Dim strXPath                                As String

Dim lngAux                                  As Long
Dim lngRepeticao                            As Long

Dim lngNivel                                As Long
Dim strValor                                As String

On Error GoTo ErrorHandler
            
    For Each xmlNode In pxmlNodeList
                
        lngNivel = CLng(xmlNode.selectSingleNode("@NU_NIVE_MESG_ATRB").Text)
        
        sprMensagemSLCC.MaxRows = sprMensagemSLCC.MaxRows + 1
        sprMensagemSLCC.Row = sprMensagemSLCC.MaxRows
           
        sprMensagemSLCC.Col = intColunmNomeFisico
        sprMensagemSLCC.Text = Trim(xmlNode.nodeName)
        
        strNomeLogicoAtributo = Trim(xmlNode.selectSingleNode("@NO_TRAP_ATRB").Text)
        
        sprMensagemSLCC.Col = intColunmNomeLogico
        sprMensagemSLCC.Text = String$((lngNivel - 1) * (5 - (CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text) * 1)), " ") & strNomeLogicoAtributo
        
        sprMensagemSLCC.BackColor = flCorNivel(lngNivel)
        sprMensagemSLCC.FontBold = CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text)
        
        sprMensagemSLCC.Col = intColunmEdicao
        sprMensagemSLCC.Text = xmlNode.Text
        
        Select Case xmlNode.nodeName
            Case "VA_OPER_ATIV"
                lngRowValor = sprMensagemSLCC.Row
            Case "CO_IDEF_TRAF"
                lngRowControleBMC = sprMensagemSLCC.Row
            Case "CO_VEIC_LEGA"
                lngRowVeiculoLegal = sprMensagemSLCC.Row
            Case "ChACAM"
                lngRowChacam = sprMensagemSLCC.Row
            Case "DT_OPER_ATIV"
                lngRowDataOperacao = sprMensagemSLCC.Row
            Case "TP_OPER_CAMB"
                lngRowTipoOperacaoCambio = sprMensagemSLCC.Row
            Case "VA_TAXA_CAMB"
                lngRowValorTaxaCambio = sprMensagemSLCC.Row
            Case "VA_MOED_NACIO"
                lngRowValorMoedaNac = sprMensagemSLCC.Row
            Case "VA_MOED_ESTRG"
                lngRowValorMoedaEstr = sprMensagemSLCC.Row
            Case "DT_LIQU_OPER"
                lngRowDataLiquOper = sprMensagemSLCC.Row
            Case "CO_REG_OPER_CAMB"
                lngRowRegistroOperCamb = sprMensagemSLCC.Row
            Case "CO_REG_OPER_CAMB2"
                lngRowRegistroOperCamb2 = sprMensagemSLCC.Row
            Case "CO_CNPJ_BASE_IF"
                lngRowCnpjBaseIf = sprMensagemSLCC.Row
            Case "CO_MOED_ISO"
                lngRowMoedIso = sprMensagemSLCC.Row
            Case "DT_ENTR_MOED_NACIO"
                lngRowDataEntrMoedNac = sprMensagemSLCC.Row
            Case "DT_ENTR_MOED_ESTR"
                lngRowDataEntrMoedEstr = sprMensagemSLCC.Row
        End Select
        
        If xmlNode.selectSingleNode("child::*") Is Nothing And _
           CLng(xmlNode.selectSingleNode("./@QT_REPE_MAX").Text) = 0 Then
            'Tag não tem filho. Apenas formatar e colocar conteudo.
                
            If Trim(xmlNode.nodeName) = "NU_COMD_OPER" Then
                sprMensagemSLCC.FontBold = True
            End If
            
            lngTamanhoAtributo = CLng(xmlNode.selectSingleNode("@QT_CTER_ATRB").Text)
            lngCasasDecimais = CLng(xmlNode.selectSingleNode("@QT_CASA_DECI_ATRB").Text)
            
            lngTamanhoAtributo = lngTamanhoAtributo - lngCasasDecimais
            
            strNomeAtributo = Trim(xmlNode.nodeName)
            
            strDominio = Trim(xmlNode.selectSingleNode("@TX_DOMI").Text)
            
            sprMensagemSLCC.Col = intColunmButton
            sprMensagemSLCC.CellType = CellTypeStaticText
            sprMensagemSLCC.Text = ""
                                   
            If strNomeAtributo = "TP_CNPT" Or _
              (strNomeAtributo = "CO_BANC" And _
               Val(fgObterCodigoCombo(Me.txtDescrticaoMensagem.Text)) <> enumTipoMensagemBUS.DespesasBMC) Then
                strDominio = strNomeAtributo
            End If
            
            If strDominio <> "" Then
                If strNomeAtributo = "TP_SOLI" Then
                    'Para a entrada manual o campo TP_SOLI sempre 2 - Complementar
                    sprMensagemSLCC.Col = intColunmEdicao
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                    sprMensagemSLCC.Text = enumTipoSolicitacao.Complementacao & " - " & "Complementar"
                    
                    xmlNode.Text = sprMensagemSLCC.Text
                    
                    sprMensagemSLCC.BackColor = &HC0FFFF
                    sprMensagemSLCC.CellBorderType = 16
                    sprMensagemSLCC.CellBorderStyle = 0
                    sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                    sprMensagemSLCC.Action = 16
                    sprMensagemSLCC.RowHeight(sprMensagemSLCC.Row) = 0
                                   
                    sprMensagemSLCC.Col = intColunmButton
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                    
                    sprMensagemSLCC.Col = intColunmTabela
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                
                ElseIf strNomeAtributo = "DT_MESG" Or strNomeAtributo = "HO_MESG" Then
                    
                    sprMensagemSLCC.CellType = CellTypeStaticText
                
                ElseIf strNomeAtributo = "CO_USUA_CADR_OPER" Then
                    
                    sprMensagemSLCC.Col = intColunmEdicao
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                    sprMensagemSLCC.Text = strDominio
                    xmlNode.Text = strDominio
                    sprMensagemSLCC.BackColor = &HC0FFFF
                    sprMensagemSLCC.CellBorderType = 16
                    sprMensagemSLCC.CellBorderStyle = 0
                    sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                    sprMensagemSLCC.Action = 16
                                   
                    sprMensagemSLCC.Col = intColunmButton
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                    
                    sprMensagemSLCC.Col = intColunmTabela
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                
                
                ElseIf Left$(strDominio, 6) <> "TABELA" Then
                    
                    sprMensagemSLCC.Col = intColunmEdicao
                    sprMensagemSLCC.CellType = CellTypeComboBox
                    sprMensagemSLCC.TypeComboBoxEditable = False
                    
                    If strNomeAtributo = "TP_CNPT" Then
                        sprMensagemSLCC.TypeComboBoxList = "1 - Mercado" & vbTab & _
                                                           "2 - Book Transfer" & vbTab & _
                                                           "3 - Cliente 1"
                    ElseIf strNomeAtributo = "CO_BANC" Then
                        sprMensagemSLCC.TypeComboBoxList = "033 - Santander" & vbTab & _
                                                           "024 - Bandepe"
                    
                    ElseIf strNomeAtributo = "TP_NEGO" Then
                        
                        Select Case CLng("0" & Right$(treMensagem.SelectedItem.Key, 4))
                            Case enumTipoOperacaoLQS.RegistroOperacaoBMCBalcao
                                strValor = "1"
                            Case enumTipoOperacaoLQS.RegistroOperacaoBMCEletronica
                                strValor = "2"
                            Case enumTipoOperacaoLQS.RegistroOperacoesBMC
                                strValor = "3"
                        End Select
                        
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeStaticText
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        sprMensagemSLCC.Text = flObterDominioXml(strDominio, _
                                                                CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text), _
                                                                strNomeAtributo, _
                                                                strValor)
                        xmlNode.Text = sprMensagemSLCC.Text
                        sprMensagemSLCC.BackColor = &HC0FFFF
                        sprMensagemSLCC.CellBorderType = 16
                        sprMensagemSLCC.CellBorderStyle = 0
                        sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                        sprMensagemSLCC.Action = 16
                        
                    ElseIf strNomeAtributo = "TP_NEGO_INTB" Then
                        
                        Select Case CLng("0" & Right$(treMensagem.SelectedItem.Key, 4))
                            'Operacao 230 e 231
                            Case enumTipoOperacaoLQS.InformaContratacaoCamaraSemTelaCega, enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega
                                strValor = "3"
                            'Operacao 232 e 233
                            Case enumTipoOperacaoLQS.InformaContratacaoInterbancarioSemCamara, enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara
                                strValor = "1"
                            'Operacao 234
                            Case enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega
                                strValor = "2"
                        End Select
                        
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeStaticText
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        sprMensagemSLCC.Text = flObterDominioXml(strDominio, _
                                                                CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text), _
                                                                strNomeAtributo, _
                                                                strValor)
                        xmlNode.Text = sprMensagemSLCC.Text
                        sprMensagemSLCC.BackColor = &HC0FFFF
                        sprMensagemSLCC.CellBorderType = 16
                        sprMensagemSLCC.CellBorderStyle = 0
                        sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                        sprMensagemSLCC.Action = 16
                        
                    ElseIf strNomeAtributo = "TP_NEGO_ARBT" Then
                        
                        Select Case CLng("0" & Right$(treMensagem.SelectedItem.Key, 4))
                            'Operacao 235
                            Case enumTipoOperacaoLQS.InformaContrArbitParceiroExteriorPaisPropriaIF
                                strValor = "2"
                            'Operacao 236 e 237
                            Case enumTipoOperacaoLQS.InformaOperacaoArbitragemParceiroPais, enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais
                                strValor = "1"
                        End Select
                        
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeStaticText
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        sprMensagemSLCC.Text = flObterDominioXml(strDominio, _
                                                                CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text), _
                                                                strNomeAtributo, _
                                                                strValor)
                        xmlNode.Text = sprMensagemSLCC.Text
                        sprMensagemSLCC.BackColor = &HC0FFFF
                        sprMensagemSLCC.CellBorderType = 16
                        sprMensagemSLCC.CellBorderStyle = 0
                        sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                        sprMensagemSLCC.Action = 16
                        
                    Else
                        sprMensagemSLCC.TypeComboBoxList = flObterDominioXml(strDominio, _
                                                                             CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text), _
                                                                             strNomeAtributo)
                    End If
                    
                    sprMensagemSLCC.TypeMaxEditLen = 100
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                
                    sprMensagemSLCC.Col = intColunmButton
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                
                    sprMensagemSLCC.Col = intColunmTabela
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
                    
                Else
                    sprMensagemSLCC.Col = intColunmEdicao
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                    
                    sprMensagemSLCC.Col = intColunmTabela
                    sprMensagemSLCC.CellType = CellTypeStaticText
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                    sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, Mid(strDominio, InStr(1, strDominio, "=") + 1, Len(strDominio))
                    
                    sprMensagemSLCC.Col = intColunmButton
                    sprMensagemSLCC.CellType = CellTypeButton
                    sprMensagemSLCC.TypeButtonText = "..."
                    sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                    sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                End If
            
                If strNomeAtributo = "IN_OPER_DEBT_CRED" Then
                    If Val(fgObterCodigoCombo(Me.txtDescrticaoMensagem.Text)) = enumTipoMensagemBUS.DespesasCETIP Then
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeStaticText
                        sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, "D - Débito"
                        sprMensagemSLCC.RowHeight(sprMensagemSLCC.MaxRows) = 0
                        sprMensagemSLCC_Change intColunmEdicao, sprMensagemSLCC.MaxRows
                    End If
                End If
                
                strDominio = ""
                
            Else
                 
               If xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text = 1 Then
                   If Mid(UCase(strNomeAtributo), 1, 2) = "DH" Then
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypePic
                        sprMensagemSLCC.TypePicMask = "99-99-9999 99:99:99"
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        If Not CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataHoraAux), "dd-mm-yyyy HH:mm:ss")
                            xmlNode.Text = sprMensagemSLCC.Text
                            xmlNode.selectSingleNode("./@Alterado").Text = 1
                        Else
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                        End If
                    ElseIf Mid(UCase(strNomeAtributo), 1, 2) = "HO" Then
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeTime
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignTop
                        sprMensagemSLCC.TypeTime24Hour = TypeTime24Hour24HourClock
                        sprMensagemSLCC.TypeTimeSeconds = False
                        sprMensagemSLCC.TypeSpin = False
                        sprMensagemSLCC.TypeTimeSeparator = Asc(":")
                        sprMensagemSLCC.TypeTimeMin = "000000"
                        sprMensagemSLCC.TypeTimeMax = "235959"
                        If Not CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataHoraAux), "HH:mm")
                            xmlNode.Text = sprMensagemSLCC.Text
                            xmlNode.selectSingleNode("./@Alterado").Text = 1
                        Else
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                        End If
                    ElseIf Mid(UCase(strNomeAtributo), 1, 2) = "DT" Then
                        
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeDate
                        sprMensagemSLCC.TypeDateCentury = True
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        sprMensagemSLCC.TypeDateSeparator = Asc("-")
                        
                        gstrMascaraDataDtp = UCase(gstrMascaraDataDtp)
                        vntMascaraData = Split(gstrMascaraDataDtp, gstrSeparadorData, , vbBinaryCompare)
                                            
                        If Not CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                            If Left(vntMascaraData(0), 1) = "D" And Left(vntMascaraData(1), 1) = "M" Then
                                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataAux), "DD-MM-YYYY")
                            ElseIf Left(vntMascaraData(1), 1) = "D" And Left(vntMascaraData(0), 1) = "M" Then
                                sprMensagemSLCC.TypeDateFormat = TypeDateFormatMMDDYY
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataAux), "MM-DD-YYYY")
                            ElseIf Left(vntMascaraData(0), 1) = "Y" Then
                                sprMensagemSLCC.TypeDateFormat = TypeDateFormatYYMMDD
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataAux), "YYYY-MM-DD")
                            Else
                                sprMensagemSLCC.TypeDateFormat = TypeDateFormatDDMMYY
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, Format(fgDataHoraServidor(DataAux), "DD-MM-YYYY")
                            End If
                            xmlNode.Text = sprMensagemSLCC.Text
                            xmlNode.selectSingleNode("./@Alterado").Text = 1
                        Else
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                        End If
                        
                    Else
                        If lngCasasDecimais <> 0 Then
                            If lngTamanhoAtributo = 17 Then
                                lngTamanhoAtributo = 15
                            ElseIf lngTamanhoAtributo + lngCasasDecimais > 15 Then
                                If lngCasasDecimais >= lngTamanhoAtributo Then
                                    lngCasasDecimais = lngCasasDecimais - ((lngTamanhoAtributo + lngCasasDecimais) - 15)
                                End If
                            End If
                            strMasacara = String(lngTamanhoAtributo, "9") & strSepDecimal & String(lngCasasDecimais, "9")
                            sprMensagemSLCC.Col = intColunmEdicao
                            sprMensagemSLCC.CellType = CellTypeFloat
                            sprMensagemSLCC.TypeFloatMax = strMasacara
                            sprMensagemSLCC.TypeFloatDecimalPlaces = lngCasasDecimais
                            sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                            sprMensagemSLCC.TypeVAlign = TypeVAlignTop
                            sprMensagemSLCC.TypeFloatMoney = False
                            sprMensagemSLCC.TypeFloatSeparator = True
                            sprMensagemSLCC.TypeFloatDecimalChar = Asc(strSepDecimal)
                            sprMensagemSLCC.TypeFloatSepChar = Asc(strSepMilhar)
                        ElseIf lngTamanhoAtributo > 9 Or lngCasasDecimais = 0 Then
                            strMasacara = String(lngTamanhoAtributo, "9")
                            sprMensagemSLCC.Col = intColunmEdicao
                            sprMensagemSLCC.CellType = CellTypeFloat
                            sprMensagemSLCC.TypeFloatDecimalPlaces = 0
                            sprMensagemSLCC.TypeFloatMax = strMasacara
                            sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                            sprMensagemSLCC.TypeVAlign = TypeVAlignTop
                            sprMensagemSLCC.TypeFloatMoney = False
                            sprMensagemSLCC.TypeFloatSeparator = False
                        End If
                        
                        If CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                            sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                        End If
                        
                    End If
                
                ElseIf xmlNode.selectSingleNode("@TP_DADO_ATRB_MESG").Text = 2 Then
                    
                    'Para a entrada manual o campo CO_OPER_ATIV será gerado automaticamente
                    If strNomeAtributo = "CO_OPER_ATIV" Then
                        sprMensagemSLCC.Col = intColunmEdicao
                        sprMensagemSLCC.CellType = CellTypeStaticText
                        sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                        sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                        sprMensagemSLCC.Text = ""
                        sprMensagemSLCC.BackColor = &HC0FFFF
                        sprMensagemSLCC.CellBorderType = 16
                        sprMensagemSLCC.CellBorderStyle = 0
                        sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
                        sprMensagemSLCC.Action = 16
                        sprMensagemSLCC.RowHeight(sprMensagemSLCC.Row) = 0
                    Else
                        If lngTamanhoAtributo >= 1000 Then
                        
                            sprMensagemSLCC.Col = intColunmEdicao
                            sprMensagemSLCC.CellType = CellTypeEdit
                            sprMensagemSLCC.TypeEditCharSet = TypeEditCharSetASCII
                            sprMensagemSLCC.TypeEditCharCase = TypeEditCharCaseSetNone
                            sprMensagemSLCC.TypeHAlign = TypeHAlignLeft
                            sprMensagemSLCC.TypeVAlign = TypeVAlignTop
                            sprMensagemSLCC.TypeEditMultiLine = True
                            sprMensagemSLCC.RowHeight(sprMensagemSLCC.Row) = 200
                            sprMensagemSLCC.TypeEditPassword = False
                            sprMensagemSLCC.TypeMaxEditLen = lngTamanhoAtributo
                            
                            If CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                            End If
                            
                        Else
                            sprMensagemSLCC.Col = intColunmEdicao
                            sprMensagemSLCC.CellType = CellTypeEdit
                            sprMensagemSLCC.TypeEditCharSet = TypeEditCharSetASCII
                            sprMensagemSLCC.TypeEditCharCase = TypeEditCharCaseSetNone
                            sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                            sprMensagemSLCC.TypeVAlign = TypeVAlignTop
                            sprMensagemSLCC.TypeEditMultiLine = False
                            sprMensagemSLCC.TypeEditPassword = False
                            sprMensagemSLCC.TypeMaxEditLen = lngTamanhoAtributo
                            
                            If CBool(CLng(xmlNode.selectSingleNode("./@Alterado").Text)) Then
                                sprMensagemSLCC.SetText intColunmEdicao, sprMensagemSLCC.MaxRows, xmlNode.Text
                            End If
                        End If
                    End If
                End If
            End If
        
        Else
        
            sprMensagemSLCC.Col = intColunmEdicao
            sprMensagemSLCC.CellType = CellTypeStaticText
            sprMensagemSLCC.TypeHAlign = TypeHAlignRight
            sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
            sprMensagemSLCC.Text = strDominio
            sprMensagemSLCC.BackColor = &HC0FFFF
            sprMensagemSLCC.CellBorderType = 16
            sprMensagemSLCC.CellBorderStyle = 0
            sprMensagemSLCC.CellBorderColor = RGB(0, 0, 0)
            sprMensagemSLCC.Action = 16
                           
            lngRepeticao = CLng(xmlNode.selectSingleNode("@QT_REPE_MAX").Text)
            
            If lngRepeticao > 0 Then
            
                'Tag de Repetição / Permitir edição da quantidade de repetições
                sprMensagemSLCC.Col = intColunmButton
                sprMensagemSLCC.CellType = CellTypeInteger
                sprMensagemSLCC.TypeHAlign = TypeHAlignRight
                sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
                
                If CLng(xmlNode.selectSingleNode("@IN_OBRI_ATRB").Text) = 1 Then
                    sprMensagemSLCC.TypeIntegerMin = 1
                Else
                    sprMensagemSLCC.TypeIntegerMin = 0
                End If
                
                sprMensagemSLCC.TypeIntegerMax = CLng(xmlNode.selectSingleNode("@QT_REPE_MAX").Text)
                sprMensagemSLCC.TypeSpin = True
                sprMensagemSLCC.TypeIntegerSpinInc = 1
                sprMensagemSLCC.Text = CLng(xmlNode.selectSingleNode("@QT_REPE_MSG").Text)
                
                sprMensagemSLCC.Col = intColunmTabela
                sprMensagemSLCC.CellType = CellTypeStaticText
                sprMensagemSLCC.SetText intColunmTabela, sprMensagemSLCC.MaxRows, ""
            
            Else
                'Tag de Grupo / Não permitir a configuração da quantidade.
                sprMensagemSLCC.Col = intColunmButton
                sprMensagemSLCC.CellType = CellTypeStaticText

                sprMensagemSLCC.Col = intColunmEdicao
                sprMensagemSLCC.CellType = CellTypeStaticText
                
                lngRepeticao = 1
            
            End If
            
            Call flCarregaSpread(xmlNode.selectNodes("./*"))
            
        End If
    
        If strNomeAtributo = "DT_MESG" Or _
           strNomeAtributo = "HO_MESG" Or _
           strNomeAtributo = "TP_NEGO" Then
            sprMensagemSLCC.CellType = CellTypeStaticText
            sprMensagemSLCC.BackColor = &HC0FFFF
            sprMensagemSLCC.CellBorderStyle = 1
            sprMensagemSLCC.CellBorderColor = RGB(160, 160, 160)
            sprMensagemSLCC.TypeHAlign = TypeHAlignRight
            sprMensagemSLCC.TypeVAlign = TypeVAlignCenter
            sprMensagemSLCC.Action = 16
        End If
    
    Next

    sprMensagemSLCC.ReDraw = True
        
Exit Sub
ErrorHandler:
    fgCursor False
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - flCarregaSpread", Me.Caption

End Sub

Private Function flMontaMensagem() As String

Dim xmlMensagem                             As MSXML2.DOMDocument40
Dim lngTipoMensagem                         As Long

'RATS 948
Dim lngTipoMensagemLQS                      As Long

Dim lngCont                                 As Long
Dim vntConteudo                             As Variant

Dim xmlMsgEntradaManualAux                  As MSXML2.DOMDocument40

Const TAM_FORMATO_DATA                      As Integer = 8
Const TAM_FORMATO_HORA                      As Integer = 6

Dim xmlNode                                 As IXMLDOMNode

On Error GoTo ErrorHandler
    
    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    Set xmlMsgEntradaManualAux = CreateObject("MSXML2.DOMDocument.4.0")
        
    lngTipoMensagem = CLng(txtDescrticaoMensagem.Tag)
    lngTipoMensagemLQS = fgObterCodigoCombo(treMensagem.SelectedItem)
        
    Call fgAppendNode(xmlMensagem, "", "MESG", "")
    
    'RATS 948
    Call fgAppendNode(xmlMensagem, "MESG", "TP_MESG", lngTipoMensagem)
    
    'Adiciona tag indicando que se trata de entrada manual de NET
    'e define o Tipo de Operação acordando com o Local de Liquidação
    If lngTipoMensagem = NET_CO_EVEN_BILAT Or _
       lngTipoMensagem = NET_CO_EVEN_MULTI Then
        
        Call fgAppendNode(xmlMensagem, "MESG", "NET_ENTRADA_MANUAL", vbNullString)
        
        With sprMensagemSLCC
            For lngCont = 1 To .MaxRows
                
                .GetText intColunmNomeFisico, lngCont, vntConteudo
                If vntConteudo = "CO_LOCA_LIQU" Then
                    
                    .GetText intColunmEdicao, lngCont, vntConteudo
                    vntConteudo = Val(fgObterCodigoCombo(vntConteudo))
                    
                    Select Case vntConteudo
                        Case enumLocalLiquidacao.BMA
                            lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMA
                        Case enumLocalLiquidacao.BMC
                            lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC
                        Case enumLocalLiquidacao.BMD
                            lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD
                        Case enumLocalLiquidacao.CETIP
                            If lngTipoMensagem = NET_CO_EVEN_BILAT Then
                                lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP
                            Else
                                lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralCETIP
                            End If
                        Case enumLocalLiquidacao.CLBCAcoes
                            lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralCBLC
                    End Select
                    
                    xmlMensagem.selectSingleNode("//TP_MESG").Text = lngTipoMensagem
                    Exit For
                    
                End If
            
            Next
        End With
        
    End If
    
    If Trim(gstrSiglaSistema) = vbNullString Then gstrSiglaSistema = "A8"
    
    If lngTipoMensagem = enumTipoMensagemBUS.DespesasBMC Then
        gstrSiglaSistema = "E2"
    End If
    
    Call fgAppendNode(xmlMensagem, "MESG", "SG_SIST_ORIG", gstrSiglaSistema)
    Call fgAppendNode(xmlMensagem, "MESG", "SG_SIST_DEST", "A8")
    Call fgAppendNode(xmlMensagem, "MESG", "CO_EMPR", cboEmpresa.ItemData(cboEmpresa.ListIndex))
    
    'BMF
    If lngTipoMensagemLQS <> enumTipoOperacaoLQS.RegistroLiquidacaoMultilateralBMF And _
       lngTipoMensagemLQS <> enumTipoOperacaoLQS.ConsultaOperacaoCCR And _
       lngTipoMensagemLQS <> enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR Then
        
        Call fgAppendNode(xmlMensagem, "MESG", "IN_ENTR_MANU", enumIndicadorSimNao.Sim)
        
    End If

    If fraContingencia.Visible Then
        Call fgAppendNode(xmlMensagem, "MESG", "IN_ENVI_PREV_SIST_PJ", IIf(chkPrevisaoPJ.value = vbChecked, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao))
        Call fgAppendNode(xmlMensagem, "MESG", "IN_ENVI_RELZ_SIST_PJ", IIf(chkRealizadoPJ.value = vbChecked, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao))
        Call fgAppendNode(xmlMensagem, "MESG", "IN_ENVI_PREV_SIST_A6", IIf(chkPrevistoA6.value = vbChecked, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao))
        Call fgAppendNode(xmlMensagem, "MESG", "IN_ENVI_RELZ_SIST_A6", IIf(chkRealizadoA6.value = vbChecked, enumIndicadorSimNao.Sim, enumIndicadorSimNao.Nao))
    End If

    xmlMsgEntradaManualAux.loadXML xmlMsgEntradaManual.xml
    
    Call flFormatarValoresMensagem(xmlMsgEntradaManualAux.selectSingleNode("//XML"))
    
    For Each xmlNode In xmlMsgEntradaManualAux.selectNodes("//XML/*")
        'Usar append xml para incluir os grupos e repetições
        fgAppendXML xmlMensagem, _
                    "MESG", _
                    xmlNode.xml

    Next

    Call flRemoverAtributos(xmlMensagem)
        
    'RATS 948
    'lngTipoMensagem alterado para ----> lngTipoMensagemLQS
    
    If lngTipoMensagemLQS = enumTipoOperacaoLQS.EventosJuros Or _
       lngTipoMensagemLQS = enumTipoOperacaoLQS.EventosAmortização Or _
       lngTipoMensagemLQS = enumTipoOperacaoLQS.EventosResgate Then
       
       xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Inclusao
       
       If Not xmlMensagem.selectSingleNode("//TP_OPER_ROTI_ABER") Is Nothing Then
            xmlMensagem.selectSingleNode("//TP_OPER_ROTI_ABER").Text = Format$(xmlMensagem.selectSingleNode("//TP_OPER_ROTI_ABER").Text, "00")
       End If
       
       xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")
         
    ElseIf lngTipoMensagemLQS = enumTipoMensagemLQS.DespesasSelic Then
    
        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Inclusao
        xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")
        
    ElseIf lngTipoMensagemLQS = enumTipoMensagemLQS.DespesasCETIP Then
        
        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Inclusao
        xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")
    
    ElseIf lngTipoMensagemLQS = enumTipoMensagemLQS.EventoJurosCETIP Then
        
        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Complementacao
        xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")

    ElseIf lngTipoMensagemLQS = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC Then
        
        Call fgAppendNode(xmlMensagem, "MESG", "TP_LIQU_OPER_ATIV", enumTipoLiquidacao.Multilateral)
    
    ElseIf lngTipoMensagem = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMD Then
        
        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Complementacao
        xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")
    
    ElseIf lngTipoMensagemLQS = enumTipoOperacaoLQS.ConsultaOperacaoCCR Or _
           lngTipoMensagemLQS = enumTipoOperacaoLQS.ConsultaLimitesImportacaoCCR Or _
           lngTipoMensagemLQS = enumTipoOperacaoLQS.NegociacaoOperacaoCCR Then

        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Complementacao
        xmlMensagem.selectSingleNode("//CO_OPER_ATIV").Text = "A8" & Format(lngTipoMensagem, "000") & Format$(Now, "YYYYMMDDHHMMSS")

    ElseIf lngTipoMensagemLQS = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega Or _
           lngTipoMensagemLQS = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Or _
           lngTipoMensagemLQS = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then

        xmlMensagem.selectSingleNode("//TP_SOLI").Text = enumTipoSolicitacao.Confirmacao

    End If
    
    'Local de Liquidacao = 22 para todos os layouts da CAM
    If xmlMensagem.selectSingleNode("//CO_LOCA_LIQU") Is Nothing Then
        If lngTipoMensagem = enumTipoMensagemLQS.ContratacaoMercadoPrimario Or lngTipoMensagem = enumTipoMensagemLQS.EdicaoContratacaoMercadoPrimario _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConfirmacaoEdicaoContratacaoMercadoPrimario Or lngTipoMensagem = enumTipoMensagemLQS.AlteracaoContrato _
        Or lngTipoMensagem = enumTipoMensagemLQS.EdicaoAlteracaoContrato Or lngTipoMensagem = enumTipoMensagemLQS.ConfirmacaoEdicaoAlteracaoContrato _
        Or lngTipoMensagem = enumTipoMensagemLQS.LiquidacaoMercadoPrimario Or lngTipoMensagem = enumTipoMensagemLQS.BaixaValorLiquidar _
        Or lngTipoMensagem = enumTipoMensagemLQS.RestabelecimentoBaixa Or lngTipoMensagem = enumTipoMensagemLQS.CancelamentoValorLiquidar _
        Or lngTipoMensagem = enumTipoMensagemLQS.EdicaoCancelamentoValorLiquidar Or lngTipoMensagem = enumTipoMensagemLQS.ConfirmacaoEdicaoCancelamentoValorLiquidar _
        Or lngTipoMensagem = enumTipoMensagemLQS.VinculacaoContratos Or lngTipoMensagem = enumTipoMensagemLQS.AnulacaoEvento _
        Or lngTipoMensagem = enumTipoMensagemLQS.CorretoraRequisitaClausulasEspecificas Or lngTipoMensagem = enumTipoMensagemLQS.IFInformaClausulasEspecificas _
        Or lngTipoMensagem = enumTipoMensagemLQS.ManutencaoCadastroAgenciaCentralizadoraCambio Or lngTipoMensagem = enumTipoMensagemLQS.CredenciamentoDescredenciamentoDispostoRMCCI _
        Or lngTipoMensagem = enumTipoMensagemLQS.IncorporacaoContratos Or lngTipoMensagem = enumTipoMensagemLQS.AceiteRejeicaoIncorporacaoContratos _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaContratosEmSer Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaEventosUmDia _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaDetalhamentoContratoInterbancario Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaEventosContratoMercadoPrimario _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaEventosContratoIntermediadoMercadoPrimario Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaHistoricoIncorporacoes _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaContratosIncorporacao Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaCadeiaIncorporacoesContrato _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaPosicaoCambioMoeda Or lngTipoMensagem = enumTipoMensagemLQS.AtualizaçãoInclusãoInstrucoesPagamento _
        Or lngTipoMensagem = enumTipoMensagemLQS.ConsultaInstrucoesPagamento _
        Or lngTipoMensagem <> enumTipoMensagemLQS.IFInformaTIRemContrapartidaaRagadorouRecebedorPaís Or lngTipoMensagem <> enumTipoMensagemLQS.IFInformaTIRemContrapartidaOutraCDE _
        Or lngTipoMensagem <> enumTipoMensagemLQS.IFInformaTIRemContrapartidaOperacaoCambialPropria Or lngTipoMensagem <> enumTipoMensagemLQS.IFRequisitaInclusaoemCadastroCDE _
        Or lngTipoMensagem <> enumTipoMensagemLQS.IFRequisitaAlteracaoCadastroCDE Or lngTipoMensagem <> enumTipoMensagemLQS.IFRequisitaExclusaoCadastroCDE _
        Or lngTipoMensagem <> enumTipoMensagemLQS.IFInformaAnulacaoRegistroTIR Or lngTipoMensagem <> enumTipoMensagemLQS.IFConsultaCDE _
        Or lngTipoMensagem <> enumTipoMensagemLQS.IFConsultaTIRUmDia Or lngTipoMensagem <> enumTipoMensagemLQS.IFConsultaDetalhamentoTIR Then
           
            Call fgAppendNode(xmlMensagem, "MESG", "CO_LOCA_LIQU", enumLocalLiquidacao.CAM)
            
        End If
    End If
    
    flMontaMensagem = xmlMensagem.xml

    Set xmlMensagem = Nothing
    Set xmlMsgEntradaManualAux = Nothing
    
    Exit Function

ErrorHandler:
    Set xmlMensagem = Nothing
    Set xmlMsgEntradaManualAux = Nothing
    fgRaiseError App.EXEName, "frmEntradaManual", "flMontaMensagem", 0

End Function

Private Sub flFormatarValoresMensagem(ByRef pxmlNode As IXMLDOMNode)


Dim xmlNode                                 As IXMLDOMNode
Dim strNodeName                             As String

10    On Error GoTo ErrorHandler


20        For Each xmlNode In pxmlNode.selectNodes(".//*")
              
30            strNodeName = xmlNode.nodeName & "---> " & xmlNode.Text
              
40            If xmlNode.selectSingleNode("child::*") Is Nothing And _
                 CLng(xmlNode.selectSingleNode("./@QT_REPE_MAX").Text) > 0 Then
                  'Excluir tag de repetição sem filho
              
50                xmlNode.parentNode.removeChild xmlNode
                  
60            Else
70                If xmlNode.selectSingleNode("./@TX_DOMI").Text = "" Then
                  'tratamento do XML para excluir separador de MILHAR
80                    If xmlNode.selectSingleNode("./@TP_DADO_ATRB_MESG").Text = 1 Then
                          'Verifica o formato do campo:
90                        If InStr(1, xmlNode.nodeName, "DH") > 0 Then
100                           If Trim(xmlNode.Text) <> "" And Val(xmlNode.Text) <> 0 Then
110                               xmlNode.Text = fgDateHr_To_DtHrXML(xmlNode.Text)
120                           End If

130                       ElseIf InStr(1, UCase(xmlNode.nodeName), "DT") > 0 Then
140                           If Trim(xmlNode.Text) <> "" And Val(xmlNode.Text) <> 0 Then
150                               xmlNode.Text = fgDate_To_DtXML(xmlNode.Text)
160                           End If
170                       ElseIf InStr(1, xmlNode.nodeName, "HO") > 0 Then
180                           xmlNode.Text = Replace(xmlNode.Text, ":", "")
190                       ElseIf xmlNode.nodeName = "TP_CNPT" Or _
                                (xmlNode.nodeName = "CO_BANC" And _
                                 Val(fgObterCodigoCombo(Me.txtDescrticaoMensagem.Text)) <> enumTipoMensagemBUS.DespesasBMC) Then
200                           xmlNode.Text = Val(fgObterCodigoCombo(xmlNode.Text))
                          ElseIf xmlNode.nodeName = "VA_TAXA_CAMB" Then
                              'Numerico
                              If Trim(xmlNode.Text) <> "" Then
201                              xmlNode.Text = Replace(UCase(Replace(xmlNode.Text, strSepMilhar, vbNullString)), ".", ",")
                                 xmlNode.Text = CDec(xmlNode.Text)
                              End If
210                       Else
                              'Numerico
220                           xmlNode.Text = Replace(UCase(Replace(xmlNode.Text, strSepMilhar, vbNullString)), ".", ",")
230                       End If
240                   End If
                  
250               Else
260                   If xmlNode.nodeName = "TP_SOLI" Then
                          
270                       xmlNode.Text = 2
                      
280                   ElseIf xmlNode.nodeName = "CO_USUA_CADR_OPER" Then
                          
                          'Não há tratamento
                          
290                   ElseIf Left$(xmlNode.selectSingleNode("./@TX_DOMI").Text, 6) <> "TABELA" Then
                          
300                       If Trim(xmlNode.Text) <> "" Then
                              'RATS - 931
310                           If InStr(1, xmlNode.Text, "-") > 0 Then
320                             xmlNode.Text = Trim$(Mid$(xmlNode.Text, 1, InStr(1, xmlNode.Text, "-") - 1))
330                           Else
340                             xmlNode.Text = Trim$(xmlNode.Text)
350                           End If
                          
360                           If xmlNode.nodeName = "IN_OPER_DEBT_CRED" Then
370                               xmlNode.Text = IIf(xmlNode.Text = "C", enumTipoDebitoCredito.Credito, enumTipoDebitoCredito.Debito)
380                           End If
390                       End If
                          
400                   Else
                          
410                       If Trim$(xmlNode.Text) <> "" Then
420                           If InStr(1, xmlNode.Text, "-") > 0 Then
                                  If InStr(1, UCase(xmlNode.nodeName), "DT") > 0 Then
                                     If Trim(xmlNode.Text) <> "" And Val(xmlNode.Text) <> 0 Then
                                        xmlNode.Text = fgDate_To_DtXML(xmlNode.Text)
                                     End If
                                  Else
                                     xmlNode.Text = Trim$(Mid$(xmlNode.Text, 1, InStr(1, xmlNode.Text, "-") - 1))
                                  End If
440                           Else
450                               xmlNode.Text = Trim$(xmlNode.Text)
460                           End If
470                       End If
                      
480                   End If
490               End If
500           End If
510       Next
          
520   Exit Sub


ErrorHandler:

530       fgRaiseError App.EXEName, TypeName(Me), "flFormatarValoresMensagem", 0, 0, "NodeName: " & strNodeName & " - Linha: " & Erl

End Sub

Private Sub flLimpaCampos()

On Error GoTo ErrorHandler

    sprMensagemSLCC.MaxRows = 0
    txtDescrticaoMensagem.Text = ""
    gstrSiglaSistema = ""
    sprMensagemSLCC.Height = 6790
    fraContingencia.Caption = ""
    fraContingencia.Visible = False
    chkPrevistoA6.value = vbUnchecked
    chkPrevisaoPJ.value = vbUnchecked
    chkRealizadoA6.value = vbUnchecked
    chkRealizadoPJ.value = vbUnchecked
    tlbCadastro.Buttons("Salvar").Enabled = False
    fraFiltro.Enabled = True
    Set xmlMsgEntradaManual = Nothing

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

Private Function flValidarCampos() As String

Dim lngCont                                     As Long
Dim vntConteudo                                 As Variant
Dim xmlNode                                     As IXMLDOMNode
Dim vntCNPJContraparte                          As Variant
Dim vntTipoContraparte                          As Variant
    
On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex = -1 Then
        flValidarCampos = "Selecione uma Empresa."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    If treMensagem.Nodes.Count = 0 Then
        flValidarCampos = "Selecione uma mensagem."
        treMensagem.SetFocus
        Exit Function
    End If
    
    If treMensagem.SelectedItem Is Nothing Then
        flValidarCampos = "Selecione uma mensagem."
        treMensagem.SetFocus
        Exit Function
    End If
    
    If Val(fgObterCodigoCombo(txtDescrticaoMensagem.Text)) = NET_CO_EVEN_MULTI Or _
       Val(fgObterCodigoCombo(txtDescrticaoMensagem.Text)) = NET_CO_EVEN_BILAT Then
    
        With sprMensagemSLCC
            For lngCont = 1 To .MaxRows
                    
                If .RowHeight(lngCont) <> 0 Then
                    .Col = intColunmNomeLogico
                    .Row = lngCont
                    
                    .GetText intColunmNomeFisico, lngCont, vntConteudo
                    If vntConteudo = "CO_CNPJ_CNPT" Then
                        .GetText intColunmEdicao, lngCont, vntCNPJContraparte
                    ElseIf vntConteudo = "TP_CNPT" Then
                        .GetText intColunmEdicao, lngCont, vntTipoContraparte
                        vntTipoContraparte = fgObterCodigoCombo(vntTipoContraparte)
                    End If
                    
                    If .FontBold Then
                        .Col = intColunmEdicao
                        If .Text = vbNullString Then
                            .GetText intColunmNomeLogico, lngCont, vntConteudo
                            flValidarCampos = "Atributo " & vntConteudo & " preenchimento obrigatório."
                            Exit Function
                        End If
                    End If
                End If
                
            Next
            
            If Val(vntTipoContraparte) = enumTipoContraparte.Interno And Val(vntCNPJContraparte) = 0 Then
                flValidarCampos = "Código CNPJ Contraparte é obrigatório para tipo contraparte Book Transfer."
                Exit Function
            End If
        
        End With
        
    Else
        
        For Each xmlNode In xmlMsgEntradaManual.selectNodes("//XML//*")
            If xmlNode.nodeName <> "CO_OPER_ATIV" Then
                If CLng(xmlNode.selectSingleNode("./@IN_OBRI_ATRB").Text) = 1 Then
                    If Trim$(xmlNode.Text) = vbNullString Then
                        flValidarCampos = "Atributo " & xmlNode.selectSingleNode("./@NO_TRAP_ATRB").Text & " preenchimento obrigatório."
                        Exit Function
                    End If
                End If
            End If
        Next
        
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

    fgRaiseError App.EXEName, "frmEntradaManualSLCCBeta", "flValidarCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Function flObterDominioXml(ByVal pstrxmlDominio As String, _
                                   ByVal plngObrigatorio As Long, _
                                   ByVal pstrNomeAtributo As String, _
                          Optional ByVal pstrValor As String) As String

Dim strDominio                                  As String
Dim xmlDominio                                  As MSXML2.DOMDocument40
Dim xmlNode                                     As MSXML2.IXMLDOMNode

On Error GoTo ErrorHandler
    
    Set xmlDominio = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDominio.loadXML(pstrxmlDominio) Then
        fgErroLoadXML xmlDominio, App.EXEName, TypeName(Me), "flObterDominioXml"
    End If
    
    If plngObrigatorio <> 1 And pstrNomeAtributo <> "TP_CPRO_OPER_ATIV" Then
        strDominio = " " & vbTab
    End If
        
    If pstrNomeAtributo = "IN_OPER_DEBT_CRED" Then
        For Each xmlNode In xmlDominio.documentElement.childNodes
            strDominio = strDominio & IIf(xmlNode.selectSingleNode("CO_DOMI").Text = 1, "D", "C") & " - " & xmlNode.selectSingleNode("DE_DOMI").Text & vbTab
        Next
    
    ElseIf pstrNomeAtributo = "TP_NEGO" _
        Or pstrNomeAtributo = "TP_NEGO_INTB" _
        Or pstrNomeAtributo = "TP_NEGO_ARBT" Then
        
        For Each xmlNode In xmlDominio.documentElement.childNodes
            If Trim(pstrValor) = Trim(xmlNode.selectSingleNode("CO_DOMI").Text) Then
                strDominio = strDominio & xmlNode.selectSingleNode("CO_DOMI").Text & " - " & xmlNode.selectSingleNode("DE_DOMI").Text
            End If
        Next
    
    Else
        For Each xmlNode In xmlDominio.documentElement.childNodes
            strDominio = strDominio & xmlNode.selectSingleNode("CO_DOMI").Text & " - " & xmlNode.selectSingleNode("DE_DOMI").Text & vbTab
        Next
    End If
    
    flObterDominioXml = strDominio
    
    Set xmlDominio = Nothing
        
Exit Function
ErrorHandler:
    
    Set xmlDominio = Nothing
    fgRaiseError App.EXEName, "frmEntradaManual", "flObterDominioXml", 0

End Function

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

Private Sub flMontarMensagemXML_X(ByRef pxmlNodeList As IXMLDOMNodeList, _
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
                              "NO_TRAP_ATRB", _
                              xmlNode.selectSingleNode("NO_TRAP_ATRB").Text
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "IN_OBRI_ATRB", _
                              xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "TP_MESG", _
                              xmlNode.selectSingleNode("TP_MESG").Text
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "NO_TIPO_MESG", _
                              xmlNode.selectSingleNode("NO_TIPO_MESG").Text
                              
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
                              "QT_REPE_MAX", _
                               CLng(xmlNode.selectSingleNode("QT_REPE").Text)
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "QT_REPE_MSG", _
                              IIf(CLng(xmlNode.selectSingleNode("QT_REPE").Text) = 0, 0, 1)
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "TP_FORM_MESG", _
                              CLng(xmlNode.selectSingleNode("TP_FORM_MESG").Text)

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text & "[position()=last()]", _
                              "NU_NIVE_MESG_ATRB", _
                              CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text)
                        
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "TX_DOMI", _
                              xmlNode.selectSingleNode("TX_DOMI").Text
                              
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "Alterado", _
                              0
                            
            'fgAppendAttribute pxmlDOMLayout, _
            '                  xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
            '                  "Posicao", _
            '                  lngX

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

Private Sub flMontaMensagemXML(ByRef pxmlNodeList As IXMLDOMNodeList, _
                               ByRef pxmlDOMLayout As DOMDocument40)
                             
Dim xmlNodeInclusao                         As MSXML2.IXMLDOMNode
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim lngNivel                                As Long
Dim lngX                                    As Long
Dim strNomeNodeInclusao                     As String

On Error GoTo ErrorHandler

    If pxmlNodeList.length = 0 Then Exit Sub
    
    Set xmlNodeInclusao = pxmlDOMLayout.selectSingleNode("//XML")
    
    lngNivel = 1
    
    For lngX = 0 To pxmlNodeList.length - 1
        
        Set xmlNode = pxmlNodeList.Item(lngX)
        
        If lngNivel = CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) Then
            'Continuar no mesmo nível

            fgAppendNode pxmlDOMLayout, _
                         xmlNodeInclusao.selectSingleNode("NO_ATRB_MESG | @NO_ATRB_MESG").Text, _
                         xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                         vbNullString

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "IN_OBRI_ATRB", _
                              xmlNode.selectSingleNode("IN_OBRI_ATRB").Text
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "TP_MESG", _
                              xmlNode.selectSingleNode("TP_MESG").Text
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "NO_TIPO_MESG", _
                              xmlNode.selectSingleNode("NO_TIPO_MESG").Text
            
            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "NO_TRAP_ATRB", _
                              xmlNode.selectSingleNode("NO_TRAP_ATRB").Text
                              
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
                              "NU_NIVE_MESG_ATRB", _
                              xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text

            fgAppendAttribute pxmlDOMLayout, _
                              xmlNode.selectSingleNode("NO_ATRB_MESG").Text, _
                              "TX_DOMI", _
                              xmlNode.selectSingleNode("TX_DOMI").Text


        ElseIf CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) > lngNivel Then
            'Novo Nível, Maior que o anterior --> Incluir como Filho
            'Muda o Pai
            Set xmlNodeInclusao = pxmlNodeList.Item(lngX - 1)
            lngNivel = CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text)
            lngX = lngX - 1

        ElseIf CLng(xmlNode.selectSingleNode("NU_NIVE_MESG_ATRB").Text) < lngNivel Then
            'Voltar ao contexto anterior
            
            strNomeNodeInclusao = pxmlDOMLayout.selectSingleNode("//" & xmlNodeInclusao.selectSingleNode("NO_ATRB_MESG | @NO_ATRB_MESG").Text).parentNode.nodeName

            Set xmlNodeInclusao = pxmlNodeList.Item(0).selectSingleNode("../*[NO_ATRB_MESG='" & strNomeNodeInclusao & "']")
            If xmlNodeInclusao Is Nothing Then
                Set xmlNodeInclusao = pxmlDOMLayout.selectSingleNode("//XML")
            End If
            
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

Private Function flCompletarepeticao(ByRef pxmlMensagem As MSXML2.DOMDocument40, _
                                     ByVal pstrContexto As String, _
                                     ByVal plngQuantidadeRepet As Long) As String

Dim xmlMensagem                          As MSXML2.DOMDocument40
Dim xmlNode                              As MSXML2.IXMLDOMNode
Dim xmlNodeAux                           As MSXML2.IXMLDOMNode
Dim lngX                                 As Long
Dim lngQtRepe                            As Long

On Error GoTo ErrorHandler

    Set xmlMensagem = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMensagem.loadXML pxmlMensagem.xml
        
    lngQtRepe = xmlMensagem.selectSingleNode("//" & pstrContexto).childNodes.length
        
    Set xmlNode = xmlMensagem.selectSingleNode("//" & pstrContexto)
    Set xmlNodeAux = xmlMensagem.selectSingleNode("//" & pstrContexto).selectSingleNode("descendant::*")
    
    For lngX = 1 To plngQuantidadeRepet
        Call fgAppendXML(xmlMensagem, pstrContexto, xmlNodeAux.xml)
    Next
    
    pxmlMensagem.loadXML xmlMensagem.xml
    
    Set xmlMensagem = Nothing

Exit Function
ErrorHandler:
    Set xmlMensagem = Nothing

    fgRaiseError App.EXEName, "frmTipoMensagem", "flCompletarepeticao", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Sub flRemoverAtributos(ByRef pxmlNode As MSXML2.IXMLDOMNode)

Dim xmlNode                                 As IXMLDOMNode
Dim xmlAttribute                            As IXMLDOMAttribute

    For Each xmlNode In pxmlNode.selectNodes("//*")
        For Each xmlAttribute In xmlNode.attributes
            xmlNode.attributes.removeNamedItem xmlAttribute.nodeName
        Next
    Next

End Sub

Private Sub flSalvarPadrao()

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim lngTipoOperacao                         As Long

On Error GoTo ErrorHandler

    fgCursor True
        
    'Salva a sigla do sistema de origem (obtido na escolha do Veiculo legal
    fgAppendAttribute xmlMsgEntradaManual, _
                      "XML", _
                      "SistemaOrigem", _
                      gstrSiglaSistema
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    lngTipoOperacao = CLng(Mid$(Me.treMensagem.SelectedItem.Key, Len(Me.treMensagem.SelectedItem.Key) - 3))
    
    'Quando for Entrada Manual de NET, não importa qual o Grupo de Mensagem, nem o Local de Liquidação.
    'Grava os dados sempre para o Tipo de Operação NET Entrada Manual Bilateral CETIP (123).
    '(Exceto Multilateral BMC)
    Select Case lngTipoOperacao
        Case NET_CO_EVEN_BILAT
            lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP
        Case NET_CO_EVEN_MULTI
            If fgObterCodigoCombo(Me.cboGrupo.Text) = "BMC" Then
                lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC
            Else
                lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP
            End If
    End Select
    vntCodErro = 0
    objMensagem.GravarParametroEntradaManual cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                             lngTipoOperacao, _
                                             xmlMsgEntradaManual.xml, _
                                             vntCodErro, _
                                             vntMensagemErro
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMensagem = Nothing
    
    MsgBox "Parametrização gravada com sucesso.", vbInformation
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - flSalvarPadrao", Me.Caption
End Sub

Private Function flObterPadrao() As String

#If EnableSoap = 1 Then
    Dim objMensagem                         As MSSOAPLib30.SoapClient30
#Else
    Dim objMensagem                         As A8MIU.clsMensagem
#End If

Dim xmlParametro                            As MSXML2.DOMDocument40
Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim strUsuario                              As String
Dim intTipoMensagem                         As Integer
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant
Dim lngTipoOperacao                         As Long

    On Error GoTo ErrorHandler

    Set xmlParametro = New MSXML2.DOMDocument40
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    
    lngTipoOperacao = CLng(Mid$(Me.treMensagem.SelectedItem.Key, Len(Me.treMensagem.SelectedItem.Key) - 3))
    
    'Quando for Entrada Manual de NET, não importa qual o Grupo de Mensagem, nem o Local de Liquidação.
    'Obtém os dados sempre para o Tipo de Operação NET Entrada Manual Bilateral CETIP (123).
    '(Exceto Multilateral BMC)
    Select Case lngTipoOperacao
        Case NET_CO_EVEN_BILAT
            lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP
        Case NET_CO_EVEN_MULTI
            If fgObterCodigoCombo(Me.cboGrupo.Text) = "BMC" Then
                lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualMultilateralBMC
            Else
                lngTipoOperacao = enumTipoOperacaoLQS.NETEntradaManualBilateralCETIP
            End If
    End Select
    vntCodErro = 0
    xmlParametro.loadXML objMensagem.ObterParametroEntadaManual(cboEmpresa.ItemData(cboEmpresa.ListIndex), _
                                                                lngTipoOperacao, _
                                                                vntCodErro, _
                                                                vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    For Each xmlNode In xmlParametro.selectNodes("//XML//*")
        
        Select Case Mid$(xmlNode.nodeName, 1, 2)
            Case "DH", "HO", "DT"
                xmlNode.selectSingleNode("./@Alterado").Text = 0
            Case Else
                If xmlNode.nodeName = "TP_SOLI" Then
                    intTipoMensagem = Val(xmlNode.selectSingleNode("./@TP_MESG").Text)
                    
                ElseIf xmlNode.nodeName = "CO_USUA_CADR_OPER" Then
                    strUsuario = fgObterUsuario
                    xmlNode.Text = strUsuario
                    xmlNode.selectSingleNode("./@TX_DOMI").Text = strUsuario
                    xmlNode.selectSingleNode("./@Alterado").Text = 1
                    
                ElseIf xmlNode.nodeName = "CO_IDEF_TRAF" Then
                    xmlNode.Text = ""
                    xmlNode.selectSingleNode("./@Alterado").Text = 0
                    xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                
                ElseIf xmlNode.nodeName = "CO_VEIC_LEGA" Then
                    If intTipoMensagem = enumTipoMensagemBUS.DespesasBMC Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                
                ElseIf xmlNode.nodeName = "VA_OPER_ATIV" Then
                    If intTipoMensagem = enumTipoMensagemBUS.DespesasBMC Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "ChACAM" Then
                    'Layout 249 - Operacao 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "DT_OPER_ATIV" Then
                    'Layout 249 - Operacao 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "TP_OPER_CAMB" Then
                    'Layout 249 - Operacao 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "VA_TAXA_CAMB" Then
                    'Layout 249 - Operacao 231, 233 ou 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "VA_MOED_NACIO" Then
                    'Layout 249 - Operacao 231, 233 ou 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "VA_MOED_ESTRG" Then
                    'Layout 249 - Operacao 231, 233 ou 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "DT_LIQU_OPER" Then
                    'Layout 249 - Operacao 231, 233 ou 234
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoContrCamaraTelaCega Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "CO_REG_OPER_CAMB" Then
                    'Layout 249 - Operacao 231 ou 233 ou Layout 253 - Operacao 237
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "CO_REG_OPER_CAMB2" Then
                    'Layout 253 - Operacao 237
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperArbitragemParceiroPais Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                ElseIf xmlNode.nodeName = "CO_CNPJ_BASE_IF" Then
                    'Layout 249 - Operacao 231 ou 233
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If

                ElseIf xmlNode.nodeName = "CO_MOED_ISO" Then
                    'Layout 249 - Operacao 231 ou 233
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If

                ElseIf xmlNode.nodeName = "DT_ENTR_MOED_NACIO" Then
                    'Layout 249 - Operacao 231 ou 233
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If

                ElseIf xmlNode.nodeName = "DT_ENTR_MOED_ESTR" Then
                    'Layout 249 - Operacao 231 ou 233
                    If lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoCamaraSemTelaCega _
                    Or lngTipoOperacao = enumTipoOperacaoLQS.InformaConfirmacaoOperacaoInterbancarioSemCamara Then
                        xmlNode.Text = ""
                        xmlNode.selectSingleNode("./@Alterado").Text = 0
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                
                Else
                    'Obter Dominios da ultima mensagem carregada
                    'pois caso algum dominio tenha sido incluido ele estara disponivel
                    If Not xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI") Is Nothing Then
                        xmlNode.selectSingleNode("./@TX_DOMI").Text = xmlMsgEntradaManual.selectSingleNode("//" & xmlNode.nodeName & "/@TX_DOMI").Text
                    End If
                    
                End If
        End Select
    Next
    
    flObterPadrao = xmlParametro.xml
    
    Set xmlParametro = Nothing
    Set objMensagem = Nothing
    
    Exit Function
ErrorHandler:
    Set xmlParametro = Nothing
    Set objMensagem = Nothing
    
    If vntCodErro = 0 <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmEntradaManual - flCarregaPadrap", Me.Caption
End Function

