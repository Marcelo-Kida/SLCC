VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrupoVeicLegalVeiculoLegal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Veículo Legal X Grupo Veículo Legal"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5385
      Left            =   60
      TabIndex        =   28
      Top             =   840
      Width           =   6975
      Begin MSComctlLib.ListView lstVeiculoLegal 
         Height          =   5145
         Left            =   75
         TabIndex        =   1
         Top             =   195
         Width           =   6825
         _ExtentX        =   12039
         _ExtentY        =   9075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Grupo Veículo Legal Atual"
      Height          =   675
      Index           =   3
      Left            =   60
      TabIndex        =   27
      Top             =   120
      Width           =   6975
      Begin VB.ComboBox cboGrupoVeiculoLegalAtual 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   6675
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Grupo Veículo Legal Novo"
      Height          =   675
      Index           =   2
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      Begin VB.ComboBox cboGrupoVeiculoLegalNovo 
         Height          =   315
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame fraMoldura 
      Caption         =   "Detalhe Veículo Legal"
      Height          =   5385
      Index           =   0
      Left            =   7080
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin VB.Frame fraDetalhe 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5100
         Left            =   45
         TabIndex        =   6
         Top             =   240
         Width           =   4815
         Begin VB.TextBox txtTipoTitularBMA 
            Height          =   315
            Left            =   60
            TabIndex        =   33
            Top             =   3165
            Width           =   4680
         End
         Begin VB.TextBox txtCtaCutdSELICPadrao 
            Height          =   315
            Left            =   2610
            TabIndex        =   30
            Top             =   2565
            Width           =   2140
         End
         Begin VB.TextBox txtTipoPartCamrCETIP 
            Height          =   315
            Left            =   60
            TabIndex        =   29
            Top             =   2565
            Width           =   2430
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   60
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   26
            Top             =   240
            Width           =   2370
         End
         Begin VB.TextBox txtNomeVeiculo 
            Height          =   315
            Left            =   60
            TabIndex        =   18
            Top             =   780
            Width           =   4680
         End
         Begin VB.Frame fraMoldura 
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
            Height          =   1035
            Index           =   1
            Left            =   60
            TabIndex        =   12
            Top             =   3555
            Width           =   4695
            Begin VB.Frame fraMoldura 
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   675
               Index           =   4
               Left            =   180
               TabIndex        =   13
               Top             =   240
               Width           =   4395
               Begin VB.TextBox txtDataFim 
                  ForeColor       =   &H80000012&
                  Height          =   315
                  Left            =   2700
                  TabIndex        =   15
                  Top             =   330
                  Width           =   1560
               End
               Begin VB.TextBox txtDataInicio 
                  ForeColor       =   &H80000012&
                  Height          =   315
                  Left            =   120
                  TabIndex        =   14
                  Top             =   330
                  Width           =   1560
               End
               Begin VB.Label lblLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Início "
                  Height          =   195
                  Index           =   6
                  Left            =   120
                  TabIndex        =   17
                  Top             =   60
                  Width           =   450
               End
               Begin VB.Label lblLabel 
                  AutoSize        =   -1  'True
                  Caption         =   "Fim"
                  Height          =   195
                  Index           =   7
                  Left            =   2730
                  TabIndex        =   16
                  Top             =   75
                  Width           =   240
               End
            End
         End
         Begin VB.TextBox txtReduzido 
            Height          =   315
            Left            =   60
            TabIndex        =   11
            Top             =   1935
            Width           =   2475
         End
         Begin VB.TextBox txtCNPJ 
            Height          =   315
            Left            =   2625
            TabIndex        =   10
            Top             =   1935
            Width           =   2130
         End
         Begin VB.TextBox txtUltimaAtlz 
            Height          =   315
            Left            =   2640
            TabIndex        =   9
            Top             =   4710
            Width           =   2100
         End
         Begin VB.TextBox txtSigla 
            Height          =   285
            Left            =   2505
            TabIndex        =   8
            Top             =   240
            Width           =   2250
         End
         Begin VB.TextBox txtEmpresa 
            Height          =   315
            Left            =   60
            TabIndex        =   7
            Top             =   1335
            Width           =   4680
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Títular BMA "
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   34
            Top             =   2970
            Width           =   1260
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Conta Custódia SELIC Padrão"
            Height          =   195
            Index           =   10
            Left            =   2610
            TabIndex        =   32
            Top             =   2340
            Width           =   2130
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cód. Participante Câmara CETIP "
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   31
            Top             =   2340
            Width           =   2355
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   25
            Top             =   585
            Width           =   420
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   24
            Top             =   0
            Width           =   495
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Nome Reduzido"
            Height          =   195
            Index           =   5
            Left            =   75
            TabIndex        =   23
            Top             =   1740
            Width           =   1140
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Código CNPJ"
            Height          =   195
            Index           =   4
            Left            =   2610
            TabIndex        =   22
            Top             =   1740
            Width           =   945
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Data/Hora Última Atualização "
            Height          =   195
            Index           =   8
            Left            =   420
            TabIndex        =   21
            Top             =   4770
            Width           =   2160
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Sigla Sistema Origem"
            Height          =   195
            Index           =   2
            Left            =   2550
            TabIndex        =   20
            Top             =   0
            Width           =   1485
         End
         Begin VB.Label lblLabel 
            AutoSize        =   -1  'True
            Caption         =   "Empresa"
            Height          =   195
            Index           =   3
            Left            =   75
            TabIndex        =   19
            Top             =   1140
            Width           =   615
         End
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   10020
      TabIndex        =   3
      Top             =   6285
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      ButtonWidth     =   1693
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   180
      Top             =   6840
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
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":150C
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":195E
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":1A70
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":1B6A
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeicLegalVeiculoLegal.frx":1C64
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGrupoVeicLegalVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsável pelo cadastro de veículo Legal no Grupo de Veículo Legal através da camada controladora de casos de uso
' MIU, método A8MIU.clsMIU.Executar
'
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlMapaNavegacao                    As MSXML2.DOMDocument40

Private xmlDominioSPB                       As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmGrupoVeicLegalVeiculoLegal"

Private strSG_SIST                          As String
Private lngCO_EMPR                          As Long
Private strCodigoVeiculoLegal               As String
Private intCodigoGrupoVeiculoLegal          As Integer
Private intTipoBackOffice                   As Integer

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private Const COL_EMPRESA                   As Integer = 1
Private Const COL_SISTEMA                   As Integer = 2
Private Const COL_CODIGO                    As Integer = 3
Private Const COL_NOME                      As Integer = 4

'Formatar o Listview
Private Sub flFormataListView()

    With lstVeiculoLegal.ColumnHeaders
        .Add COL_EMPRESA, , "Empresa", 870.2363, lvwColumnLeft
        .Add COL_SISTEMA, , "Sistema", 870, lvwColumnLeft
        .Add COL_CODIGO, , "Código", 870.9371, lvwColumnLeft
        .Add COL_NOME, , "Nome", 4030.47, lvwColumnLeft
    End With

End Sub

'' Encaminhar a solicitação (Leitura de todos os veículos legais cadastrados, para
'' preenchimento do listview) à camada controladora de caso de uso (componente /
'' classe / metodo ) : A8MIU.clsVeiculoLegal.LerTodos
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregaVeiculosLegais(ByVal plngrGrupoVeiculoLegal As Long)

#If EnableSoap = 1 Then
    Dim objVeiculoLegal                     As MSSOAPLib30.SoapClient30
#Else
    Dim objVeiculoLegal                     As A8MIU.clsVeiculoLegal
#End If

Dim xmlVeiculoLegal                         As MSXML2.DOMDocument40
Dim objDomNodeVeiculo                       As MSXML2.IXMLDOMNode
Dim strXMLRetorno                           As String
Dim strIcon                                 As String
Dim objListItem                             As ListItem

On Error GoTo ErrorHandler

    Call fgCursor(True)

    lstVeiculoLegal.ListItems.Clear
    
    Set objVeiculoLegal = fgCriarObjetoMIU("A8MIU.clsVeiculoLegal")

    Set xmlVeiculoLegal = CreateObject("MSXML2.DOMDocument.4.0")
    
    strXMLRetorno = objVeiculoLegal.LerTodos(vbNullString, _
                                             plngrGrupoVeiculoLegal, _
                                             "S", _
                                             intTipoBackOffice, _
                                             "N")
    
    If Trim(strXMLRetorno) <> "" Then
       If Not xmlVeiculoLegal.loadXML(strXMLRetorno) Then
            Call fgErroLoadXML(xmlVeiculoLegal, App.EXEName, Me.Name, "flCarregaVeiculosLegais")
       End If
    Else
       Call fgCursor(False)
       Exit Sub
    End If
    
    For Each objDomNodeVeiculo In xmlVeiculoLegal.documentElement.selectNodes("//Repeat_VeiculoLegal/*")
    
        Set objListItem = lstVeiculoLegal.ListItems.Add(, ";" & objDomNodeVeiculo.selectSingleNode("CO_GRUP_VEIC_LEGA").Text & _
                                                          ";" & objDomNodeVeiculo.selectSingleNode("SG_SIST").Text & _
                                                          ";" & objDomNodeVeiculo.selectSingleNode("CO_VEIC_LEGA").Text)
                                                          
        objListItem.Text = objDomNodeVeiculo.selectSingleNode("CO_EMPR").Text
        objListItem.SubItems(COL_SISTEMA - 1) = objDomNodeVeiculo.selectSingleNode("SG_SIST").Text
        objListItem.SubItems(COL_NOME - 1) = objDomNodeVeiculo.selectSingleNode("NO_VEIC_LEGA").Text
        objListItem.SubItems(COL_CODIGO - 1) = objDomNodeVeiculo.selectSingleNode("CO_VEIC_LEGA").Text

    Next
    
    Set xmlVeiculoLegal = Nothing
    Set objVeiculoLegal = Nothing
    
    Call fgCursor(False)
    
    Exit Sub
    
ErrorHandler:

    Call fgCursor(False)

    Set xmlVeiculoLegal = Nothing
    Set objVeiculoLegal = Nothing

    fgRaiseError App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flCarregaVeiculosLegais", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub
'' Encaminhar a solicitação (Leitura de todos os grupos de veículo legal
'' cadastrados, para carregamento de combos) à camada controladora de caso de uso
'' (componente / classe / metodo ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregaCombo()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    cboGrupoVeiculoLegalNovo.Clear
    cboGrupoVeiculoLegalAtual.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_SEGR").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_BKOF").Text = intTipoBackOffice
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal").xml
    strLerTodos = objMIU.Executar(strPropriedades)
    
    Set objMIU = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)
    
    'Carrega Combo Atual
    fgCarregarCombos cboGrupoVeiculoLegalAtual, xmlLerTodos, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA"
    'Carrega Combo Novo
    fgCarregarCombos cboGrupoVeiculoLegalNovo, xmlLerTodos, "GrupoVeiculoLegal", "CO_GRUP_VEIC_LEGA", "NO_GRUP_VEIC_LEGA"
    
    Set xmlLerTodos = Nothing

    Exit Sub
    
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing

    fgRaiseError App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flCarregaCombo", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub



'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
'' A8MIU.clsMiu.ObterMapaNavegacao
'' O método retornará uma String XML para a camada de interface.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
    Dim objA6A7A8Funcoes                    As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
    Dim objA6A7A8Funcoes                    As A8MIU.clsA6A7A8Funcoes
#End If

Dim strMapaNavegacao                         As String

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade)
    Set objMIU = Nothing
    
    Set xmlDominioSPB = CreateObject("MSXML2.DOMDocument.4.0")
    Set objA6A7A8Funcoes = fgCriarObjetoMIU("A8MIU.clsA6A7A8Funcoes")
    xmlDominioSPB.loadXML objA6A7A8Funcoes.ObterDominioSPB("TpTitlar")
    Set objA6A7A8Funcoes = Nothing
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flInicializar")
    Else
       intTipoBackOffice = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_TipoBackOffice/Grupo_TipoBackOffice/TP_BKOF").Text
    End If
    
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing

    fgRaiseError App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flInicializar", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub
'' É acionado através no botão 'Salvar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Atualização dos dados na tabela) à camada
'' controladora de caso de uso (componente / classe / metodo ) : A8MIU.
'' clsVeiculoLegal.AlterarCodigoGrupo
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsVeiculoLegal
#End If

Dim strRetorno                              As String
Dim strPropriedades                         As String

On Error GoTo ErrorHandler

    If cboGrupoVeiculoLegalNovo.ListIndex = cboGrupoVeiculoLegalAtual.ListIndex Then
        frmMural.Caption = Me.Caption
        frmMural.Display = "Selecione um novo Grupo de Veículo Legal"
        frmMural.IconeExibicao = IconInformation
        frmMural.Show vbModal
        Exit Sub
    ElseIf txtCodigo.Text = vbNullString Then
       Exit Sub
    End If
    
    Call fgCursor(True)
    
    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsVeiculoLegal")
    Call objMIU.AlterarCodigoGrupo(xmlLer.xml)
    Set objMIU = Nothing
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    flLimparCampos
    tlbCadastro.Buttons(gstrSalvar).Enabled = False
    
    If lstVeiculoLegal.ListItems.Count > 0 Then
       lstVeiculoLegal.ListItems.Remove lstVeiculoLegal.SelectedItem.Index
       If lstVeiculoLegal.ListItems.Count > 0 Then
          lstVeiculoLegal.ListItems(1).Selected = True
          cboGrupoVeiculoLegalNovo.ListIndex = -1
          fraMoldura(2).Enabled = False
       End If
    End If
    
    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    Set objMIU = Nothing

    fgRaiseError App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Carregar a interface com as informações do xml
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler
        
    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = gstrOperAlterar
        .selectSingleNode("CO_GRUP_VEIC_LEGA").Text = fgObterCodigoCombo(cboGrupoVeiculoLegalNovo.List(cboGrupoVeiculoLegalNovo.ListIndex))
    End With
    
    Exit Function

ErrorHandler:
    
    fgRaiseError App.EXEName, "frmGrupoVeicLegalVeiculoLegal", "flInterfaceToXml", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

'Limpar todos os campos para uma nova inclusão.

Private Sub flLimparCampos()
    
On Error GoTo ErrorHandler

    txtCNPJ.Text = Space(0)
    txtDataFim.Text = Space(0)
    txtDataInicio.Text = Space(0)
    txtEmpresa.Text = Space(0)
    txtNomeVeiculo.Text = Space(0)
    txtReduzido.Text = Space(0)
    txtSigla.Text = Space(0)
    txtUltimaAtlz.Text = Space(0)
    txtCodigo.Text = vbNullString
    lngCO_EMPR = 0
    strSG_SIST = 0

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparCampos", 0
    
End Sub

' Encaminhar a solicitação (Leitura de detalhes veículo legal)
' à camada controladora de caso de uso (componente / classe / metodo ) : A8MIU.
' clsMiu.Executar
' O método retornará uma String XML para a camada de interface.
Private Sub flLer(ByVal strVeiculoLegal As String, ByVal pstrSG_SIST As String)

#If EnableSoap = 1 Then
    Dim objVeiculoLegal                     As MSSOAPLib30.SoapClient30
#Else
    Dim objVeiculoLegal                     As A8MIU.clsVeiculoLegal
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strXMLLer                               As String
Dim strIcon                                 As String

Dim strXPath                                As String

On Error GoTo ErrorHandler

    Call fgCursor(True)

    Set objVeiculoLegal = fgCriarObjetoMIU("A8MIU.clsVeiculoLegal")

    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    strXMLLer = objVeiculoLegal.Ler(strVeiculoLegal, pstrSG_SIST)
    
    'caso a tabela esteja sem registros não tem como carregar um XML,
    'sendo assim vai para o fim da rotina.
    If Trim(strXMLLer) <> "" Then
       If Not xmlLer.loadXML(strXMLLer) Then
          '100 - Documento XML Inválido.
          lngCodigoErroNegocio = 100
          GoTo ErrorHandler
       End If
    Else
       Call fgCursor(False)
       Exit Sub
    End If
    
    With xmlLer.documentElement
        txtCodigo.Text = .selectSingleNode("CO_VEIC_LEGA").Text
        txtNomeVeiculo.Text = .selectSingleNode("NO_VEIC_LEGA").Text
        txtReduzido.Text = .selectSingleNode("NO_REDU_VEIC_LEGA").Text
        txtCNPJ.Text = .selectSingleNode("CO_CNPJ_VEIC_LEGA").Text
        txtUltimaAtlz.Text = fgDtHrXML_To_Interface(.selectSingleNode("DH_ULTI_ATLZ").Text)
        txtEmpresa.Text = .selectSingleNode("NO_EMPR").Text
        txtSigla.Text = .selectSingleNode("NO_SIST").Text
        txtDataInicio.Text = fgDtXML_To_Interface(.selectSingleNode("DT_INIC_VIGE").Text)
        
        txtTipoPartCamrCETIP.Text = .selectSingleNode("ID_PART_CAMR_CETIP").Text
        txtCtaCutdSELICPadrao.Text = .selectSingleNode("CO_CNTA_CUTD_PADR_SELIC").Text
        strXPath = "//Grupo_DominioAtributo[./NO_TIPO_TAG='TpTitlar' " & _
                   "and normalize-space(./CO_DOMI)='" & .selectSingleNode("TP_TITL_BMA").Text & "']/DE_DOMI"
        
        txtTipoTitularBMA.Text = .selectSingleNode("TP_TITL_BMA").Text
        If Not xmlDominioSPB.selectSingleNode(strXPath) Is Nothing Then
            txtTipoTitularBMA.Text = txtTipoTitularBMA.Text & " - " & xmlDominioSPB.selectSingleNode(strXPath).Text
        End If
        
        If .selectSingleNode("DT_FIM_VIGE").Text <> gstrDataVazia Then
           txtDataFim.Text = fgDtXML_To_Interface(.selectSingleNode("DT_FIM_VIGE").Text)
        Else
           txtDataFim.Text = ""
        End If
        lngCO_EMPR = .selectSingleNode("CO_EMPR").Text
        strSG_SIST = .selectSingleNode("SG_SIST").Text
    End With
     
    Set objVeiculoLegal = Nothing
    Call fgCursor(False)

    Exit Sub

ErrorHandler:

    Set xmlLer = Nothing
    Set objVeiculoLegal = Nothing
    Call fgCursor(False)

    fgRaiseError App.EXEName, TypeName(Me), "flLer", 0

End Sub

Private Sub cboGrupoVeiculoLegalAtual_Click()
    
On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    If cboGrupoVeiculoLegalAtual.ListCount <= 0 Then
       Exit Sub
    End If
    
    If cboGrupoVeiculoLegalNovo.ListCount > 0 Then
        cboGrupoVeiculoLegalNovo.ListIndex = -1
        fraMoldura(2).Enabled = False
    End If
    
    flLimparCampos
    flCarregaVeiculosLegais fgObterCodigoCombo(cboGrupoVeiculoLegalAtual.List(cboGrupoVeiculoLegalAtual.ListIndex))
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeicLegalVeiculoLegal - cboGrupoVeiculoLegalAtual_Click", Me.Caption
        
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
    
    Call fgCursor(True)
    Call fgCenterMe(Me)
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Salvar").Enabled = False
    
    flFormataListView
    Call flInicializar
    Call flCarregaCombo
    fraMoldura(2).Enabled = False
    Call fgCursor(False)
    
    Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeicLegalVeiculoLegal - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrupoVeicLegalVeiculoLegal = Nothing
    Set xmlDominioSPB = Nothing
End Sub

Private Sub lstVeiculoLegal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lstVeiculoLegal.Sorted = True
    lstVeiculoLegal.SortKey = ColumnHeader.Index - 1

    If lstVeiculoLegal.SortOrder = lvwAscending Then
        lstVeiculoLegal.SortOrder = lvwDescending
    Else
        lstVeiculoLegal.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstVeiculoLegal_ColumnClick"

End Sub

Private Sub lstVeiculoLegal_ItemClick(ByVal Item As MSComctlLib.ListItem)

Dim strChaves()                             As String

On Error GoTo ErrorHandler
    
    fgCursor True

    flLimparCampos
    
    If lstVeiculoLegal.ListItems.Count > 0 Then
        strChaves = Split(Item.Key, ";")
        intCodigoGrupoVeiculoLegal = Val(strChaves(1))
        strSG_SIST = strChaves(2)
        strCodigoVeiculoLegal = strChaves(3)
        
        fraMoldura(2).Enabled = True
        
        fgSearchItemCombo cboGrupoVeiculoLegalNovo, , intCodigoGrupoVeiculoLegal
        flLer strCodigoVeiculoLegal, strSG_SIST
        tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    Else
       tlbCadastro.Buttons(gstrSalvar).Enabled = False
    End If
    
    fgCursor
    
Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiLQS.uctlogErros.MostrarErros Err, "lstVeiculoLegal_ItemClick", Me.Caption

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    
    fgCursor True
    
    Select Case Button.Key
        Case gstrSalvar
             Call flSalvar
        Case gstrSair
             fgCursor False
             Unload Me
    End Select
    
    fgCursor False
    
    Exit Sub
    
ErrorHandler:

    fgCursor False
    mdiLQS.uctlogErros.MostrarErros Err, "tlbCadastro_ButtonClick", Me.Caption
    cboGrupoVeiculoLegalAtual_Click
    
End Sub
