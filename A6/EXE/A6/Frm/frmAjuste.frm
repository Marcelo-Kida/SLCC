VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAjuste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sub-reserva - Ajuste de Movimento"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9645
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9645
   Begin VB.ComboBox cboSiglaSistema 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   7950
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   660
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":195E
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":1A70
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":1B6A
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":1C64
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCNPJContraparte 
      Height          =   285
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3600
      Width           =   4935
   End
   Begin VB.TextBox txtNomeContraparte 
      Height          =   285
      Left            =   4560
      MaxLength       =   80
      TabIndex        =   9
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox txtDescricaoAtivo 
      Height          =   315
      Left            =   4560
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2940
      Width           =   4935
   End
   Begin VB.ComboBox cboTipoLiquidacao 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Frame fraAgenda 
      Height          =   3375
      Left            =   4560
      TabIndex        =   23
      Top             =   4500
      Width           =   4935
      Begin VB.Frame fraTipoMovimento 
         Caption         =   "Tipo Movimento"
         Height          =   855
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1575
         Begin VB.OptionButton optEntradaSaida 
            Caption         =   "Saída"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   540
            Width           =   795
         End
         Begin VB.OptionButton optEntradaSaida 
            Caption         =   "Entrada"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Histórico"
         Height          =   1575
         Left            =   120
         TabIndex        =   25
         Top             =   1620
         Width           =   4695
         Begin VB.TextBox txtObs 
            Height          =   1215
            Left            =   120
            MaxLength       =   50
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   4425
         End
      End
      Begin VB.Frame fraSituacaoMovimento 
         Caption         =   "Situação do Movimento"
         Height          =   855
         Left            =   1800
         TabIndex        =   24
         Top             =   720
         Width           =   3015
         Begin VB.OptionButton optMovimentoConfirmado 
            Caption         =   "Realizado Confirmado"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   540
            Width           =   1935
         End
         Begin VB.OptionButton optMovimentoPrevisto 
            Caption         =   "Previsto"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
      End
      Begin NumBox.Number numValor 
         Height          =   315
         Left            =   1800
         TabIndex        =   11
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         Decimais        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AceitaNegativo  =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpAgenda 
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   104464385
         CurrentDate     =   36966
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   375
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   1800
         TabIndex        =   27
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.ComboBox cboTipoOperacao 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1620
      Width           =   4935
   End
   Begin VB.ComboBox cboBancLiqu 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   4935
   End
   Begin VB.ComboBox cboLocalLiqu 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   4935
   End
   Begin VB.ComboBox cboGrupoVeicLegal 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   300
      Width           =   4335
   End
   Begin VB.ComboBox cboVeiculoLegal 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4335
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   6600
      TabIndex        =   17
      Top             =   7980
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   582
      ButtonWidth     =   1720
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Limpar"
            Key             =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlFiltros 
      Left            =   60
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":1D5E
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":1E70
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":21C2
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":22D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAjuste.frx":25EE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treItemCaixa 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   1380
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   11456
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imlIcons"
      Appearance      =   1
   End
   Begin VB.Label Label5 
      Caption         =   "Nome Contraparte"
      Height          =   195
      Left            =   4560
      TabIndex        =   32
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "CNPJ Contraparte"
      Height          =   195
      Left            =   4560
      TabIndex        =   31
      Top             =   3360
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição Ativo"
      Height          =   195
      Left            =   4560
      TabIndex        =   30
      Top             =   2700
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Liquidação"
      Height          =   195
      Left            =   4560
      TabIndex        =   29
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Operação"
      Height          =   225
      Left            =   4560
      TabIndex        =   22
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label lblBancLiqu 
      AutoSize        =   -1  'True
      Caption         =   "Empresa"
      Height          =   195
      Left            =   4560
      TabIndex        =   21
      Top             =   60
      Width           =   615
   End
   Begin VB.Label lblLocalLiqu 
      AutoSize        =   -1  'True
      Caption         =   "Local Liquidação"
      Height          =   195
      Left            =   4560
      TabIndex        =   20
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblGrupoVeicLegal 
      AutoSize        =   -1  'True
      Caption         =   "Grupo de Veículos Legais"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   60
      Width           =   1845
   End
   Begin VB.Label lblVeiculoLegal 
      AutoSize        =   -1  'True
      Caption         =   "Veículo Legal"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   990
   End
End
Attribute VB_Name = "frmAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário o ajuste do movimento de itens de caixa.

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlItemCaixaVeiculoLegal            As MSXML2.DOMDocument40
Private datDataCaixaSubReserva              As Date
Private Const strFuncionalidade             As String = "frmAjuste"

Private lngListIndexcboGrupoVeiculoLegal    As Long 'Código do último grupo selecionado

Private Type udtItemCaixa
    LetraK01                                As String * 1
    TipoBackOffice                          As String * 1
    LetraK02                                As String * 1
    TipoCaixa                               As String * 1
    LetraK03                                As String * 1
    CodigoItemCaixa                         As String * 16
    CodigoItemCaixaPai                      As String * 16
End Type

Private Type udtItemCaixaAux
    NodeKey                                 As String * 20
End Type

Private udtItemCaixa                        As udtItemCaixa
Private udtItemCaixaPai                     As udtItemCaixa

Private udtItemCaixaAux                     As udtItemCaixaAux

' Define o tamanho máximo dos campos a serem preenchidos, de acordo com as propriedades da tabela.

Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement
        txtCNPJContraparte.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_MovimentoFuturo/CO_CNPJ_CNPT/@Tamanho").Text
        txtDescricaoAtivo.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_MovimentoFuturo/DE_ATIV/@Tamanho").Text
        txtNomeContraparte.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_MovimentoFuturo/NO_CNPT/@Tamanho").Text
        txtObs.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_MovimentoFuturo/TX_MOTI_GERA_MOVI/@Tamanho").Text
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
End Sub

' Carrega definições iniciais do formulário.

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A6MIU.clsMIU
#End If

Dim strMapaNavegacao       As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmAjuste", "flInicializar")
    End If

Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0
End Sub

' Transfere dados do formulário para a camada intermediária, para a atualização da tabela.

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A6MIU.clsMIU
#End If

Dim strExecucao            As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    If Not flValidarCampos Then Exit Sub

    flInterfaceToXml

    strExecucao = IIf(udtItemCaixa.TipoCaixa = enumTipoCaixa.CaixaFuturo, "//Grupo_MovimentoFuturo", "//Grupo_MovimentoSubReserva")

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    Call objMIU.Executar(xmlMapaNavegacao.selectSingleNode(strExecucao).xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMIU = Nothing
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
        
    fgRaiseError App.EXEName, Me.Name, "flSalvar", 0
    
End Sub

' Converte os dados da tela para XML, para a atualização da tabela.

Private Function flInterfaceToXml() As String

Dim strItemCaixa                            As String

On Error GoTo ErrorHandler

    If udtItemCaixa.TipoCaixa = enumTipoCaixa.CaixaFuturo Then
        With xmlMapaNavegacao.selectSingleNode("//Grupo_MovimentoFuturo")
            xmlMapaNavegacao.selectSingleNode("//Grupo_MovimentoFuturo/@Operacao").Text = "Incluir"
            
            .selectSingleNode("SG_SIST").Text = fgObterCodigoCombo(cboSiglaSistema.Text)
            .selectSingleNode("DH_REME").Text = fgDt_To_Xml(dtpAgenda.Value)
            .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboBancLiqu.Text)
            .selectSingleNode("DT_FECH_PROC").Text = fgDt_To_Xml(dtpAgenda.Value)
            .selectSingleNode("CO_LOCA_LIQU").Text = fgObterCodigoCombo(cboLocalLiqu.Text)
            .selectSingleNode("DE_TIPO_LIQU").Text = fgObterDescricaoCombo(cboTipoLiquidacao.Text)
            .selectSingleNode("DT_LIQU_OPER").Text = fgDt_To_Xml(dtpAgenda.Value)
            .selectSingleNode("IN_MOVI_ENTR_SAID").Text = IIf(optEntradaSaida(0).Value, enumTipoEntradaSaida.ENTRADA, enumTipoEntradaSaida.Saida)
            .selectSingleNode("VA_LIQU_OPER").Text = fgVlr_To_Xml(numValor.Valor)
            .selectSingleNode("DE_ATIV").Text = fgLimpaCaracterEspecial(txtDescricaoAtivo.Text)
            .selectSingleNode("CO_VEIC_LEGA").Text = fgObterCodigoCombo(cboVeiculoLegal.Text)
            .selectSingleNode("CO_ITEM_CAIX").Text = udtItemCaixa.CodigoItemCaixa
            .selectSingleNode("CO_CNPJ_CNPT").Text = txtCNPJContraparte.Text
            .selectSingleNode("NO_CNPT").Text = txtNomeContraparte.Text
            .selectSingleNode("TP_GERA_MOVI").Text = enumTipoRemessa.Ajuste
            .selectSingleNode("TX_MOTI_GERA_MOVI").Text = txtObs.Text
            .selectSingleNode("CO_PROD").Text = vbNullString
        End With
    Else
        With xmlMapaNavegacao.selectSingleNode("//Grupo_MovimentoSubReserva")
            xmlMapaNavegacao.selectSingleNode("//Grupo_MovimentoSubReserva/@Operacao").Text = "Incluir"
            
            .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboBancLiqu.Text)
            .selectSingleNode("TP_OPER").Text = fgObterCodigoCombo(cboTipoOperacao.Text)
            .selectSingleNode("DH_MOVI_CAIX_SUB_RESE").Text = fgDt_To_Xml(dtpAgenda.Value)
            .selectSingleNode("CO_LOCA_LIQU").Text = fgObterCodigoCombo(cboLocalLiqu.Text)
            .selectSingleNode("SG_SIST").Text = fgObterCodigoCombo(cboSiglaSistema.Text)
            .selectSingleNode("CO_VEIC_LEGA").Text = fgObterCodigoCombo(cboVeiculoLegal.Text)
            .selectSingleNode("CO_ITEM_CAIX").Text = udtItemCaixa.CodigoItemCaixa
            .selectSingleNode("DE_TIPO_LIQU").Text = fgObterDescricaoCombo(cboTipoLiquidacao.Text)
            .selectSingleNode("VA_MOVI_CAIX_SUB_RESE").Text = fgVlr_To_Xml(numValor.Valor)
            .selectSingleNode("IN_MOVI_ENTR_SAID").Text = IIf(optEntradaSaida(0).Value, enumTipoEntradaSaida.ENTRADA, enumTipoEntradaSaida.Saida)
            .selectSingleNode("CO_CNPJ_CNPT").Text = txtCNPJContraparte.Text
            .selectSingleNode("NO_CNPT").Text = txtNomeContraparte.Text
            .selectSingleNode("DE_ATIV").Text = txtDescricaoAtivo.Text
            .selectSingleNode("CO_SITU_MOVI_CAIX_SUB_RESE").Text = IIf(optMovimentoPrevisto.Value, enumTipoMovimento.Previsto, enumTipoMovimento.RealizadoConfirmado)
            .selectSingleNode("NU_SEQU_OPER_ATIV").Text = 0
            .selectSingleNode("TP_GERA_MOVI").Text = enumTipoRemessa.Ajuste
            .selectSingleNode("TX_MOTI_GERA_MOVI").Text = txtObs.Text
        End With
    End If
    
Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flInterfaceToXml", 0
End Function

Private Sub cboVeiculoLegal_Click()

On Error GoTo ErrorHandler

    fgCursor True
    cboSiglaSistema.ListIndex = cboVeiculoLegal.ListIndex
    
    flObterDataCaixa

    If treItemCaixa.SelectedItem Is Nothing Then
        flTravarControles True
        Exit Sub
    End If
    
    If treItemCaixa.SelectedItem.children > 0 Or cboVeiculoLegal.ListIndex < 0 Then
        flTravarControles True
    Else
        flTravarControles False, Val(Split(treItemCaixa.SelectedItem.Key, "K")(2)) = enumTipoCaixa.CaixaSubReserva
    End If
    
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmAjuste - cboVeiculoLegal_Click"

End Sub

' Obtem a data atual do caixa.

Private Sub flObterDataCaixa()

#If EnableSoap = 1 Then
    Dim objAjuste          As MSSOAPLib30.SoapClient30
#Else
    Dim objAjuste          As A6MIU.clsAjuste
#End If

Dim strCodigoVeiculoLegal  As String
Dim strSiglaSistema        As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant
Dim strDataCaixaSubReserva As String

On Error GoTo ErrorHandler

    strCodigoVeiculoLegal = fgObterCodigoCombo(cboVeiculoLegal.Text)
    
    If Trim(strCodigoVeiculoLegal) = vbNullString Then
        datDataCaixaSubReserva = datDataVazia
    End If

    strSiglaSistema = xmlItemCaixaVeiculoLegal.documentElement.selectSingleNode("Grupo_Dados/Repeat_VeiculoLegal/Grupo_VeiculoLegal[CO_VEIC_LEGA='" & strCodigoVeiculoLegal & "']/SG_SIST").Text

    Set objAjuste = fgCriarObjetoMIU("A6MIU.clsAjuste")
    strDataCaixaSubReserva = objAjuste.ObterDataCaixaVeiculoLegal(strCodigoVeiculoLegal, strSiglaSistema, vntCodErro, vntMensagemErro)
    If Trim(strDataCaixaSubReserva) <> "" Then
       datDataCaixaSubReserva = CDate(Format(strDataCaixaSubReserva, "DD/MM/YYYY"))
    End If
    Set objAjuste = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flObterDataCaixa", 0

End Sub

Private Sub dtpAgenda_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True

    fgCenterMe Me
    
    Me.Icon = mdiSBR.Icon
    
    Me.Show
    DoEvents
    
    lngListIndexcboGrupoVeiculoLegal = -1
    
    flInicializar
    
    Call fgCarregarCombos(Me.cboGrupoVeicLegal, xmlMapaNavegacao, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA")
    Call fgCarregarCombos(Me.cboBancLiqu, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_EMPR")
    Call fgCarregarCombos(Me.cboLocalLiqu, xmlMapaNavegacao, "LocalLiquidacao", "CO_LOCA_LIQU", "DE_LOCA_LIQU")
    Call fgCarregarCombos(Me.cboTipoOperacao, xmlMapaNavegacao, "TipoOperacao", "TP_OPER", "NO_TIPO_OPER")
    Call fgCarregarCombos(Me.cboTipoLiquidacao, xmlMapaNavegacao, "TipoLiquidacao", "TP_LIQU_OPER_ATIV", "NO_TIPO_LIQU_OPER_ATIV")
    dtpAgenda.Value = fgDataHoraServidor(enumFormatoDataHora.Data)
    
    Call flDefinirTamanhoMaximoCampos
    
    flTravarControles True
    
    fgCursor

    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, "frmAjuste - Form_Load"

End Sub

' Habilita ou desabilita controles da tela.

Private Sub flTravarControles(ByVal blnTrava As Boolean, _
                     Optional ByVal blnSubCaixa As Boolean = False)

On Error GoTo ErrorHandler

    cboBancLiqu.Enabled = Not blnTrava
    cboLocalLiqu.Enabled = Not blnTrava
    cboTipoOperacao.Enabled = Not blnTrava And blnSubCaixa
    cboTipoLiquidacao.Enabled = Not blnTrava
    txtDescricaoAtivo.Enabled = Not blnTrava
    txtCNPJContraparte.Enabled = Not blnTrava
    txtNomeContraparte.Enabled = Not blnTrava
    dtpAgenda.Enabled = Not blnTrava And Not blnSubCaixa
    
    If blnSubCaixa Then
        dtpAgenda.Value = datDataCaixaSubReserva
    End If
    
    numValor.Enabled = Not blnTrava
    fraTipoMovimento.Enabled = Not blnTrava
    fraSituacaoMovimento.Enabled = Not blnTrava And blnSubCaixa
    txtObs.Enabled = Not blnTrava
            
    If gblnPerfilManutencao Then
        tlbCadastro.Buttons("Salvar").Enabled = Not blnTrava
    Else
        tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flTravarControles", 0

End Sub

Private Sub cboGrupoVeicLegal_Click()

On Error GoTo ErrorHandler

    If cboGrupoVeicLegal.ListIndex >= 0 And cboGrupoVeicLegal.ListIndex <> lngListIndexcboGrupoVeiculoLegal Then
        fgCursor True
        
        Call flCarregarVeiculoLegalItemCaixa(CLng("0" & fgObterCodigoCombo(cboGrupoVeicLegal.Text)))
        lngListIndexcboGrupoVeiculoLegal = cboGrupoVeicLegal.ListIndex
        If Not treItemCaixa.SelectedItem Is Nothing Then
            treItemCaixa_NodeClick treItemCaixa.SelectedItem
        End If
    
        fgCursor
    End If
    
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - cboGrupoVeicLegal_Click"

End Sub

' Carrega combo de veículo legal a partir do grupo de veículo legal selecionado.

Private Sub flCarregarVeiculoLegalItemCaixa(ByVal plngGrupoVeiculo As Long)

#If EnableSoap = 1 Then
    Dim objAjuste          As MSSOAPLib30.SoapClient30
#Else
    Dim objAjuste          As A6MIU.clsAjuste
#End If

Dim strMapaNavegacao       As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler

    Set objAjuste = fgCriarObjetoMIU("A6MIU.clsAjuste")
    strMapaNavegacao = objAjuste.ObterVeiculoLegalItemCaixa(plngGrupoVeiculo, vntCodErro, vntMensagemErro)
    Set objAjuste = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set xmlItemCaixaVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlItemCaixaVeiculoLegal.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlItemCaixaVeiculoLegal, App.EXEName, "frmAjuste", "flCarregarVeiculoLegalItemCaixa")
    End If

    Call fgCarregarCombos(Me.cboVeiculoLegal, xmlItemCaixaVeiculoLegal, "VeiculoLegal", "CO_VEIC_LEGA", "NO_VEIC_LEGA")
    Call fgCarregarCombos(Me.cboSiglaSistema, xmlItemCaixaVeiculoLegal, "VeiculoLegal", "SG_SIST", "CO_VEIC_LEGA")
    
    fgCarregarTreItemCaixa treItemCaixa, xmlItemCaixaVeiculoLegal, Me

    Exit Sub

ErrorHandler:
    Set objAjuste = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, Me.Name & " - flCarregarVeiculoLegalItemCaixa", 0

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
    Case "Salvar"
        flSalvar
      
    Case "Limpar"
        flLimparCampos
    
    Case "Sair"
        Unload Me
    
    End Select
    
    fgCursor
    Exit Sub

ErrorHandler:
    mdiSBR.uctLogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick"

End Sub

' Limpa conteúdo dos campos da tela.

Private Sub flLimparCampos()

On Error GoTo ErrorHandler

    cboGrupoVeicLegal.ListIndex = -1
    cboVeiculoLegal.Clear
    treItemCaixa.Nodes.Clear
    cboBancLiqu.ListIndex = -1
    cboLocalLiqu.ListIndex = -1
    cboTipoOperacao.ListIndex = -1
    cboTipoLiquidacao.ListIndex = -1
    txtDescricaoAtivo.Text = vbNullString
    txtCNPJContraparte.Text = vbNullString
    txtNomeContraparte.Text = vbNullString
    dtpAgenda.Value = Now
    numValor.Valor = 0
    optEntradaSaida(0).Value = False
    optEntradaSaida(1).Value = False
    optMovimentoPrevisto.Value = False
    optMovimentoConfirmado.Value = False
    txtObs.Text = vbNullString
    
    lngListIndexcboGrupoVeiculoLegal = -1
    flTravarControles True
    
    cboGrupoVeicLegal.SetFocus

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparCampos", 0
        
End Sub

' Valida o conteúdo dos campos da tela, antes da gravação.

Private Function flValidarCampos() As Boolean

    flValidarCampos = False
    
    With cboGrupoVeicLegal
        If .ListIndex < 0 Then
            MsgBox "Selecione um Grupo de Veiculos Legais.", vbExclamation, Me.Caption
            .SetFocus
            Exit Function
        End If
    End With
    
    With cboVeiculoLegal
        If .ListIndex < 0 Then
            MsgBox "Selecione um Veiculo Legal.", vbExclamation, Me.Caption
            .SetFocus
            Exit Function
        End If
    End With
    
    With cboBancLiqu
        If .ListIndex < 0 Then
            MsgBox "Selecione um Banco Liquidante.", vbExclamation, Me.Caption
            .SetFocus
            Exit Function
        End If
    End With
        
    With cboLocalLiqu
        If .ListIndex < 0 Then
            MsgBox "Selecione um Local de Liquidação.", vbExclamation, Me.Caption
            .SetFocus
            Exit Function
        End If
    End With
    
    With cboTipoOperacao
        If .Enabled Then
            If .ListIndex < 0 Then
                MsgBox "Selecione um Tipo de Operação.", vbExclamation, Me.Caption
                .SetFocus
                Exit Function
            End If
        End If
    End With

    With txtDescricaoAtivo
        If Trim$(.Text) = vbNullString Then
            MsgBox "Preencha a Descrição do Ativo.", vbExclamation, Me.Caption
            .SelStart = 0
            .SelLength = Len(.Text)
            .SetFocus
            Exit Function
        End If
    End With

    If numValor.Valor = 0 Then
        MsgBox "Valor do Ajuste deve ser diferente de zero.", vbExclamation, Me.Caption
        Exit Function
    End If

    If Not optEntradaSaida(0).Value And Not optEntradaSaida(1).Value Then
        MsgBox "Selecione um Tipo de Movimento.", vbExclamation, Me.Caption
        Exit Function
    End If
    
    If optMovimentoPrevisto.Enabled Then
        If Not optMovimentoPrevisto.Value And Not optMovimentoConfirmado.Value Then
            MsgBox "Selecione uma situação do movimento.", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    
    flValidarCampos = True

End Function

Private Sub treItemCaixa_NodeClick(ByVal Node As MSComctlLib.Node)
    
On Error GoTo ErrorHandler

    fgCursor True
    
    udtItemCaixaAux.NodeKey = Node.Key
    LSet udtItemCaixa = udtItemCaixaAux
    udtItemCaixa.CodigoItemCaixa = udtItemCaixa.TipoCaixa & udtItemCaixa.CodigoItemCaixa
        
    If Node.Parent Is Nothing Then
        udtItemCaixaAux.NodeKey = Node.Key
    Else
        udtItemCaixaAux.NodeKey = Node.Parent.Key
    End If
    
    LSet udtItemCaixaPai = udtItemCaixaAux
    udtItemCaixaPai.CodigoItemCaixa = udtItemCaixaPai.TipoCaixa & udtItemCaixaPai.CodigoItemCaixa
    
    udtItemCaixa.CodigoItemCaixaPai = udtItemCaixaPai.CodigoItemCaixa
    
    If Node.children > 0 Or cboVeiculoLegal.ListIndex < 0 Then
        flTravarControles True
    Else
        flTravarControles False, Val(Split(Node.Key, "K")(2)) = enumTipoCaixa.CaixaSubReserva
    End If
    
    fgCursor
    Exit Sub

ErrorHandler:
   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - treItemCaixa_NodeClick"
    
End Sub

Private Sub txtCNPJContraparte_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorHandler

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
            KeyAscii = 0
        End If
    End If

Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - txtCNPJContraparte_KeyPress"
End Sub
