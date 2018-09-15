VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGrupoVeiculoLegal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segregação Acesso - Grupo Veículo Legal"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboTipoBackOffice 
      Height          =   315
      Left            =   1440
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   5235
   End
   Begin VB.Frame fraCadastro 
      Height          =   3975
      Left            =   60
      TabIndex        =   7
      Top             =   600
      Width           =   6885
      Begin VB.Frame fraTipoLiquidacao 
         Caption         =   "Grupo Veículo Legal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   11
         Top             =   2040
         Width           =   4275
         Begin VB.TextBox txtDescricaoGrupoVeicLega 
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   1170
            Width           =   4080
         End
         Begin NumBox.Number numCodigoGrupo 
            Height          =   315
            Left            =   90
            TabIndex        =   2
            Top             =   540
            Width           =   615
            _ExtentX        =   1085
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
            SelStart        =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   90
            TabIndex        =   13
            Top             =   330
            Width           =   495
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   105
            TabIndex        =   12
            Top             =   975
            Width           =   720
         End
      End
      Begin VB.Frame fraPeriodo 
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
         Height          =   1695
         Left            =   4620
         TabIndex        =   8
         Top             =   2040
         Width           =   1995
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   315
            Left            =   180
            TabIndex        =   4
            Top             =   540
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   58785793
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   315
            Left            =   180
            TabIndex        =   5
            Top             =   1200
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   58785793
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin VB.Label lblDataFimVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            Top             =   1005
            Width           =   240
         End
         Begin VB.Label lblDataInicioVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Início "
            Height          =   195
            Left            =   240
            TabIndex        =   9
            Top             =   330
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lstGrupoVeiculoLegal 
         Height          =   1605
         Left            =   210
         TabIndex        =   1
         Top             =   330
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2831
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   30
      Top             =   4680
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
            Picture         =   "frmGrupoVeiculoLegal.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoVeiculoLegal.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2940
      TabIndex        =   6
      Top             =   4680
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
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Back Office"
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmGrupoVeiculoLegal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:20
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Administração geral de Grupos de
'' Veículo Legal) à camada controladora de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMiu
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strGrupoPadrao                As String = "Grupo Padrão"
Private Const strFuncionalidade             As String = "frmGrupoVeiculoLegal"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstGrupoVeiculoLegal.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstGrupoVeiculoLegal.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstGrupoVeiculoLegal_ItemClick objListItem
           lstGrupoVeiculoLegal.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimpaCampos
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'Formatar o listview
Private Sub flFormataListView()

On Error GoTo ErrorHandler

    lstGrupoVeiculoLegal.ColumnHeaders.Add 1, , "Código", 1000, lvwColumnLeft
    lstGrupoVeiculoLegal.ColumnHeaders.Add 2, , "Descrição", 2420, lvwColumnLeft
    lstGrupoVeiculoLegal.ColumnHeaders.Add 3, , "Data Início", 1440, lvwColumnLeft
    lstGrupoVeiculoLegal.ColumnHeaders.Add 4, , "Data Fim ", 1440, lvwColumnLeft

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormataListView", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Atualização dos dados na tabela) à camada
'' controladora de caso de uso (componente / classe / metodo ) : A8MIU.clsMiu.
'' Executar
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim strRetorno                              As String
Dim strPropriedades                         As String
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
    
    Call fgCursor(True)
    
    Call flInterfaceToXml

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    Call objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strOperacao = "Incluir" Then
           strKeyItemSelected = "K" & numCodigoGrupo.Valor
        End If
        strOperacao = gstrOperAlterar
        numCodigoGrupo.Enabled = False
    Else
        flLimpaCampos
    End If
    Set objMIU = Nothing
    
    Call flCarregaListView
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
        
    Call fgCursor(False)

    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    Set objMIU = Nothing
    
    If strOperacao <> gstrOperExcluir Then
       With xmlLer.documentElement
            strKeyItemSelected = "K" & .selectSingleNode("CO_GRUP_VEIC_LEGA").Text
       End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Validar os campos obrigatórios para execução da funcionalidade especificada.

Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If numCodigoGrupo = 0 Then
        flValidarCampos = "Digite o código do Grupo Veículo Legal."
        numCodigoGrupo.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescricaoGrupoVeicLega) = "" Then
        flValidarCampos = "Informe a Descrição do Grupo Veículo Legal."
        txtDescricaoGrupoVeicLega.SetFocus
        Exit Function
    End If
            
    If cboTipoBackoffice.ListIndex < 0 Then
       flValidarCampos = "Informe o Tipo de Back Office."
       cboTipoBackoffice.SetFocus
       Exit Function
    End If
    
    If Not IsNull(dtpDataFimVigencia.value) Then
        If dtpDataFimVigencia.value < dtpDataInicioVigencia.value Then
            flValidarCampos = "Data final da vigência anterior à data inicial."
            dtpDataFimVigencia.SetFocus
            Exit Function
        End If
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Limpar todos os campos para uma nova inclusão.

Private Sub flLimpaCampos()

On Error GoTo ErrorHandler
    
    strOperacao = "Incluir"
    
    If fraPeriodo.Enabled = False And fraTipoLiquidacao.Enabled = False Then
       fraPeriodo.Enabled = True
       fraTipoLiquidacao.Enabled = True
    End If
       
    numCodigoGrupo.Valor = 0
    numCodigoGrupo.Enabled = True
    
    txtDescricaoGrupoVeicLega.Text = ""
        
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
    dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
    dtpDataInicioVigencia.Enabled = True
    
    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = Null
    
    If cboTipoBackoffice.ListIndex = 0 Then
       fraPeriodo.Enabled = False
       fraTipoLiquidacao.Enabled = False
       tlbCadastro.Buttons("Limpar").Enabled = False
       tlbCadastro.Buttons("Excluir").Enabled = False
       tlbCadastro.Buttons("Salvar").Enabled = False
    Else
       fraPeriodo.Enabled = True
       fraTipoLiquidacao.Enabled = True
       tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
       tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
       tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    End If
    
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmGrupoVeiculoLegal", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'' Encaminhar a solicitação (Leitura de detalhes do grupo de veículo legal
'' selecionado) à camada controladora de caso de uso (componente / classe / metodo
'' ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strLer                  As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant
        
On Error GoTo ErrorHandler

    numCodigoGrupo.Enabled = False
    dtpDataInicioVigencia.Enabled = False
    
    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = "Ler"
        .selectSingleNode("CO_GRUP_VEIC_LEGA").Text = lstGrupoVeiculoLegal.SelectedItem.Text
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strLer = objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strLer = "" Then
        Exit Sub
    Else
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    End If
    
    Set objMIU = Nothing
    
    With xmlLer.documentElement
   
        numCodigoGrupo.Valor = Val(.selectSingleNode("CO_GRUP_VEIC_LEGA").Text)
        txtDescricaoGrupoVeicLega.Text = .selectSingleNode("NO_GRUP_VEIC_LEGA").Text
                
        fgCarregaDataVigencia dtpDataInicioVigencia, dtpDataFimVigencia, _
                              .selectSingleNode("DT_INIC_VIGE").Text, _
                              .selectSingleNode("DT_FIM_VIGE").Text
        
    End With

Exit Sub
ErrorHandler:
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Sub

'Carregar a interface com as informações do xml
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler
    
    With xmlLer.documentElement
    
         .selectSingleNode("//@Operacao").Text = strOperacao
    
         If strOperacao = "Incluir" Then
            .selectSingleNode("//CO_GRUP_VEIC_LEGA").Text = numCodigoGrupo.Valor
         ElseIf strOperacao = gstrOperExcluir Then
            Exit Function
         End If
         
         .selectSingleNode("//NO_GRUP_VEIC_LEGA").Text = txtDescricaoGrupoVeicLega.Text
         .selectSingleNode("//DT_INIC_VIGE").Text = fgDate_To_DtXML(dtpDataInicioVigencia.value)
         
         If strOperacao <> gstrOperExcluir Then
             If Not IsNull(dtpDataFimVigencia.value) Then
                If .selectSingleNode("//DT_FIM_VIGE").Text <> gstrDataVazia Then
            
                    If fgDtXML_To_Date(.selectSingleNode("//DT_FIM_VIGE").Text) <> dtpDataFimVigencia.value Then
                       If MsgBox("Deseja desativar o registro a partir da data: " & dtpDataFimVigencia.value, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                          .selectSingleNode("//DT_FIM_VIGE").Text = fgDate_To_DtXML(dtpDataFimVigencia.value)
                       End If
                    End If
                Else
                    If MsgBox("Deseja desativar o registro a partir da data: " & dtpDataFimVigencia.value, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                       .selectSingleNode("//DT_FIM_VIGE").Text = fgDate_To_DtXML(dtpDataFimVigencia.value)
                    Else
                       dtpDataFimVigencia.value = Null
                       .selectSingleNode("//DT_FIM_VIGE").Text = ""
                    End If
                End If
             Else
                .selectSingleNode("//DT_FIM_VIGE").Text = ""
             End If
         End If
         
         .selectSingleNode("//TP_BKOF").Text = fgObterCodigoCombo(cboTipoBackoffice.List(cboTipoBackoffice.ListIndex))
        
    End With
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function
'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão, e todos os tipos de Back Office cadastrados, para o preenchimento do
'' combobox) à camada controladora de caso de uso (componente / classe / metodo ) :
'' A8MIU.clsMiu.ObterMapaNavegacaoA8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strMapaNavegacao        As String
Dim xmlLerTodos             As MSXML2.DOMDocument40
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    strMapaNavegacao = objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, _
                                                 strFuncionalidade, _
                                                 vntCodErro, _
                                                 vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")

    If Not xmlMapaNavegacao.loadXML(strMapaNavegacao) Then
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmGrupoVeiculoLegal", "flInicializar")
    End If
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlLerTodos.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoBackOffice").xml
    xmlLerTodos.documentElement.selectSingleNode("@Operacao").Text = gstrOperLerTodos
    xmlLerTodos.documentElement.selectSingleNode("TP_VIGE").Text = "S"
    xmlLerTodos.documentElement.selectSingleNode("TP_SEGR").Text = "N"
    
    xmlLerTodos.loadXML objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    Call fgCarregarCombos(Me.cboTipoBackoffice, xmlLerTodos, "TipoBackOffice", "TP_BKOF", "DE_BKOF", True)
    
    Set xmlLerTodos = Nothing
    
    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_GrupoVeiculoLegal").xml
    End If
    
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoVeiculoLegal", "flInicializar", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub
Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler
                
    With xmlMapaNavegacao.documentElement
        txtDescricaoGrupoVeicLega.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/NO_GRUP_VEIC_LEGA/@Tamanho").Text
    End With
        
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmTipoLiquidacao", "flDefinirTamanhoMaximoCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub
'' Encaminhar a solicitação (Leitura de todos os grupos de veículo legal
'' cadastrados, para preenchimento do listview) à camada controladora de caso de
'' uso (componente / classe / metodo ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim strPropriedades                         As String
Dim strLerTodos                             As String
Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    lstGrupoVeiculoLegal.ListItems.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_SEGR").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/TP_BKOF").Text = IIf(cboTipoBackoffice.ListIndex = 0, 0, fgObterCodigoCombo(cboTipoBackoffice.List(cboTipoBackoffice.ListIndex)))
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoVeiculoLegal").xml
    strLerTodos = objMIU.Executar(strPropriedades, _
                                  vntCodErro, _
                                  vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    If strLerTodos = "" Then Exit Sub
    
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call xmlLerTodos.loadXML(strLerTodos)

    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_GrupoVeiculoLegal/*")
        With xmlDomNode
                
            Set objListItem = lstGrupoVeiculoLegal.ListItems.Add(, "K" & .selectSingleNode("CO_GRUP_VEIC_LEGA").Text, _
                                                                         .selectSingleNode("CO_GRUP_VEIC_LEGA").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_GRUP_VEIC_LEGA").Text
            objListItem.SubItems(2) = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
            
            If CStr(.selectSingleNode("DT_FIM_VIGE").Text) <> gstrDataVazia Then
                objListItem.SubItems(3) = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text)
            End If
        
        End With
    Next

Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoVeiculoLegal", "flCarregaListView", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Sub cboTipoBackOffice_Click()

On Error GoTo ErrorHandler

    strOperacao = "Incluir"

    Call flLimpaCampos

    Call fgCursor(True)
    Call flCarregaListView
    Call fgCursor
    
    If cboTipoBackoffice.ListIndex = 0 Then
       fraPeriodo.Enabled = False
       fraTipoLiquidacao.Enabled = False
       tlbCadastro.Buttons("Limpar").Enabled = False
       tlbCadastro.Buttons("Excluir").Enabled = False
       tlbCadastro.Buttons("Salvar").Enabled = False
    Else
       fraPeriodo.Enabled = True
       fraTipoLiquidacao.Enabled = True
       tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
       tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
       tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    End If
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeiculoLegal - cboTipoBackOffice_Click", Me.Caption
    
End Sub

Private Sub dtpDataFimVigencia_Change()

On Error GoTo ErrorHandler

    fgDataVigenciaFimChange dtpDataInicioVigencia, dtpDataFimVigencia

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - dtpDataFimVigencia_Change"
End Sub

Private Sub dtpDataFimVigencia_KeyPress(KeyAscii As Integer)
    KeyAscii = 0: Beep
End Sub

Private Sub dtpDataInicioVigencia_Change()

On Error GoTo ErrorHandler

    fgDataVigenciaInicioChange dtpDataInicioVigencia, dtpDataFimVigencia

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - dtpDataInicioVigencia_Change"
End Sub

Private Sub dtpDataInicioVigencia_KeyPress(KeyAscii As Integer)
    KeyAscii = 0: Beep
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
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
        
    Call flFormataListView
    Call flInicializar
    Call flLimpaCampos
    Call flDefinirTamanhoMaximoCampos
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeiculoLegal - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGrupoVeiculoLegal = Nothing
End Sub

Private Sub lstGrupoVeiculoLegal_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

    lstGrupoVeiculoLegal.Sorted = True
    lstGrupoVeiculoLegal.SortKey = ColumnHeader.Index - 1

    If lstGrupoVeiculoLegal.SortOrder = lvwAscending Then
        lstGrupoVeiculoLegal.SortOrder = lvwDescending
    Else
        lstGrupoVeiculoLegal.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstGrupoVeiculoLegal_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    If Item.SubItems(1) = strGrupoPadrao Then
        flLimpaCampos
        frmMural.Caption = Me.Caption
        frmMural.Display = "Grupo Padrão não pode ser modificado ou excluído."
        frmMural.IconeExibicao = IconCritical
        frmMural.Show vbModal
        fraPeriodo.Enabled = False
        fraTipoLiquidacao.Enabled = False
        tlbCadastro.Buttons("Salvar").Enabled = False
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
        
        Exit Sub
    End If
    
    fraPeriodo.Enabled = True
    fraTipoLiquidacao.Enabled = True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
            
    Call fgCursor(True)
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface
    
    strKeyItemSelected = Item.Key
    
    numCodigoGrupo.Enabled = False
    
    If numCodigoGrupo > 0 Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeiculoLegal - lstGrupoVeiculoLegal_ItemClick", Me.Caption
    flRecarregar

End Sub

Private Sub flRecarregar()

On Error GoTo ErrorHandler

    flLimpaCampos
    Call flCarregaListView

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flRecarregar"
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
            If fraTipoLiquidacao.Enabled = True And numCodigoGrupo.Enabled = True Then
               numCodigoGrupo.SetFocus
            End If
        Case gstrSalvar
            Call flSalvar
            If strOperacao = gstrOperAlterar Then
               flPosicionaItemListView
            End If
        Case gstrOperExcluir
            If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
               strOperacao = gstrOperExcluir
               Call flSalvar
            End If
        Case gstrSair
            fgCursor False
            Unload Me
            Exit Sub
    End Select
    
    fgCursor False
    
Exit Sub
ErrorHandler:

    fgCursor False

    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoVeiculoLegal - tlbCadastro_ButtonClick", Me.Caption
    
    Call flCarregaListView
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub
