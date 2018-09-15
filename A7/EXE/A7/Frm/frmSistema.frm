VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSistema 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro - Sistemas"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   4140
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   6645
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   6375
      End
      Begin VB.Frame fraVigencia 
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
         Left            =   4500
         TabIndex        =   11
         Top             =   2430
         Width           =   1995
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   540
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            Format          =   22740993
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   5
            Top             =   1170
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   22740993
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin VB.Label lblDataInicioVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Início "
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   330
            Width           =   450
         End
         Begin VB.Label lblDataFimVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   975
            Width           =   240
         End
      End
      Begin VB.Frame fraSistema 
         Caption         =   "Sistema"
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
         Left            =   90
         TabIndex        =   8
         Top             =   2430
         Width           =   4275
         Begin VB.TextBox txtSiglaSistema 
            Height          =   315
            Left            =   90
            MaxLength       =   3
            TabIndex        =   2
            Top             =   540
            Width           =   1275
         End
         Begin VB.TextBox txtNomeSistema 
            Height          =   315
            Left            =   90
            TabIndex        =   3
            Top             =   1170
            Width           =   4080
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   90
            TabIndex        =   10
            Top             =   945
            Width           =   420
         End
         Begin VB.Label lblSigla 
            AutoSize        =   -1  'True
            Caption         =   "Sigla"
            Height          =   195
            Left            =   90
            TabIndex        =   9
            Top             =   330
            Width           =   345
         End
      End
      Begin MSComctlLib.ListView lstSistema 
         Height          =   1605
         Left            =   90
         TabIndex        =   1
         Top             =   765
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código do Sistema"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nome do Sistema"
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
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   4350
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
            Picture         =   "frmSistema.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSistema.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3150
      TabIndex        =   6
      Top             =   4200
      Width           =   3600
      _ExtentX        =   6350
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
            Enabled         =   0   'False
            Caption         =   "Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Excluir"
            Key             =   "Excluir"
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSistema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pelo cadastramento e manutenção de sistemas.
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmSistema"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

'Posicionar item no listview de sistemas.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

    If lstSistema.ListItems.Count = 0 Then Exit Sub
    
    If Len(strKeyItemSelected) = 0 Then
        flLimpaCampos
        Exit Sub
    End If
    
    blnEncontrou = False
    
    For Each objListItem In lstSistema.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstSistema_ItemClick objListItem
           lstSistema.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimpaCampos
    End If

End Sub

'Salvar informações correntes do sistema.
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim strRetorno          As String
Dim strPropriedades     As String
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
        If xmlLer.documentElement.selectSingleNode("DT_FIM_VIGE_SIST").Text <> gstrDataVazia Then
            If fgDtXML_To_Date(xmlLer.documentElement.selectSingleNode("DT_FIM_VIGE_SIST").Text) <> dtpDataFimVigencia.Value Then
                If MsgBox("Deseja desativar o registro a partir do dia " & dtpDataFimVigencia.Value & " ?", vbYesNo, "Atributos Mensagens") = vbNo Then Exit Sub
            End If
        End If
    End If
        
    Call fgCursor(True)

    Call flInterfaceToXml
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    Call objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> "Excluir" Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        If strOperacao = "Incluir" Then
           strKeyItemSelected = ""
        End If
        
        strOperacao = "Alterar"
        txtSiglaSistema.Enabled = False
        strKeyItemSelected = "K" & Trim(txtSiglaSistema)
    Else
        flLimpaCampos
    End If
    
    Set objMiu = Nothing
    
    Call flCarregaListView
    
    If strKey <> "" Then
        DoEvents
        lstSistema.ListItems(strKey).EnsureVisible
        lstSistema.HideSelection = False
        lstSistema.ListItems(strKey).Selected = True
    End If
    
    Call fgCursor(False)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    Set objMiu = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    If strOperacao = "Excluir" Then strOperacao = "Alterar"
    fgRaiseError App.EXEName, "frmSistema", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Validar os valores informados para o sistema.
Private Function flValidarCampos() As String
    
    If cboEmpresa.ListIndex = -1 Then
        flValidarCampos = "Selecione uma empresa."
        cboEmpresa.SetFocus
        Exit Function
    End If
    
    If Trim$(txtSiglaSistema.Text) = vbNullString Then
        flValidarCampos = "Informe o Código do Sistema."
        txtSiglaSistema.SetFocus
        Exit Function
    End If
    
    If Trim(txtNomeSistema.Text) = vbNullString Then
        flValidarCampos = "Informe o Nome do Sistema."
        txtNomeSistema.SetFocus
        Exit Function
    End If
    
    flValidarCampos = ""

End Function

'Limpar campos do formulário.
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler
        
    strOperacao = "Incluir"
    
    txtSiglaSistema.Text = vbNullString
    txtSiglaSistema.Enabled = True
    
    txtNomeSistema.Text = ""
        
    lstSistema.Sorted = False
        
    dtpDataInicioVigencia.Enabled = True
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = Null
    
    tlbCadastro.Buttons.Item("Excluir").Enabled = False
    tlbCadastro.Buttons.Item("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons.Item("Salvar").Enabled = gblnPerfilManutencao
    
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmSistema", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Converter o domínio numérico de Tipo de Dado para literais.
Private Function flObterNomeTipoDado(ByVal plCodigoTipoDados As Long) As String
    
    Select Case plCodigoTipoDados
        Case enumTipoDadoAtributo.Numerico
            flObterNomeTipoDado = "Numérico"
        Case enumTipoDadoAtributo.Alfanumerico
            flObterNomeTipoDado = "Alfanumérico"
    End Select

End Function

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMiu                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu                              As A7Miu.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    txtSiglaSistema.Enabled = False
    dtpDataInicioVigencia.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//SG_SIST").Text = lstSistema.SelectedItem.Text
        .selectSingleNode("//CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa)
    End With
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Call xmlLer.loadXML(objMiu.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
   
    With xmlLer.documentElement
        txtSiglaSistema.Text = .selectSingleNode("SG_SIST").Text
        txtNomeSistema.Text = .selectSingleNode("NO_SIST").Text
        dtpDataInicioVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE_SIST").Text)
        dtpDataInicioVigencia.Value = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE_SIST").Text)
        
        If dtpDataInicioVigencia.Value > fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
            dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
            dtpDataInicioVigencia.Enabled = True
        End If
        
        If Trim(.selectSingleNode("DT_FIM_VIGE_SIST").Text) <> gstrDataVazia Then
            If fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE_SIST").Text) < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                dtpDataFimVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE_SIST").Text)
                dtpDataInicioVigencia.Enabled = True
            Else
                dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
            End If
            dtpDataFimVigencia.Value = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE_SIST").Text)
        Else
            dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.Value
            dtpDataFimVigencia.Value = dtpDataInicioVigencia.Value
            dtpDataFimVigencia.Value = Null
        End If
        
        strUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
    End With
    
    Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0

End Sub

'Proteger chave do sistema em operações de alteração.
Private Sub flTravaCampos(ByVal pblnTravar As Boolean)
    
On Error GoTo ErrorHandler
    
    fraSistema.Enabled = Not pblnTravar
    fraVigencia.Enabled = Not pblnTravar
    lstSistema.Enabled = Not pblnTravar
    
    Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flTravaCampos", 0
    
End Sub
'Mover valores do formulário para XML para envio ao objeto de negócio.
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
        .selectSingleNode("@Operacao").Text = strOperacao
        .selectSingleNode("SG_SIST").Text = txtSiglaSistema.Text
        .selectSingleNode("NO_SIST").Text = txtNomeSistema.Text
        .selectSingleNode("CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa)
        .selectSingleNode("DT_INIC_VIGE_SIST").Text = fgDt_To_Xml(dtpDataInicioVigencia.Value)
    
        If IsNull(dtpDataFimVigencia.Value) Then
            .selectSingleNode("DT_FIM_VIGE_SIST").Text = vbNullString
        Else
            .selectSingleNode("DT_FIM_VIGE_SIST").Text = fgDt_To_Xml(dtpDataFimVigencia.Value)
        End If
    End With
    
    Exit Function
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
End Function

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
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmSistema", "flInicializar")
    End If
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlLer.loadXML(xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_Sistema").xml)
    
    Set objMiu = Nothing
    
    Exit Sub

ErrorHandler:

    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Definir tamanho máximo para valores de informações pertinentes ao sistema.
Private Sub flDefinirTamanhoMaximoCampos()
    
    With xmlMapaNavegacao.documentElement
        txtNomeSistema.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_Sistema/NO_SIST/@Tamanho").Text
    End With

End Sub

'Carregar listview com os sistema cadastrados
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim xmlLerTodos         As MSXML2.DOMDocument40
Dim xmlDomNode          As MSXML2.IXMLDOMNode
Dim objListItem         As MSComctlLib.ListItem
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    lstSistema.ListItems.Clear
    
    lstSistema.HideSelection = False
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/CO_EMPR").Text = fgObterCodigoCombo(cboEmpresa)
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema/@Operacao").Text = "LerTodos"
    
    Call xmlLerTodos.loadXML(objMiu.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Sistema").xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMiu = Nothing
    
    If Not xmlLerTodos.xml = Empty Then
    
        For Each xmlDomNode In xmlLerTodos.selectSingleNode("//Repeat_Sistema").childNodes
            
            With xmlDomNode
                
                Set objListItem = lstSistema.ListItems.Add(, "K" & Trim(.selectSingleNode("SG_SIST").Text), .selectSingleNode("SG_SIST").Text)
                
                objListItem.SubItems(1) = .selectSingleNode("NO_SIST").Text
                objListItem.SubItems(2) = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE_SIST").Text)
                
                If CStr(.selectSingleNode("DT_FIM_VIGE_SIST").Text) <> gstrDataVazia Then
                    objListItem.SubItems(3) = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE_SIST").Text)
                End If
            End With
        Next
    End If

    Set xmlLerTodos = Nothing
    
    Exit Sub
    
ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlLerTodos = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0
    
End Sub

Private Sub cboEmpresa_Click()
    
On Error GoTo ErrorHandler
    
    If cboEmpresa.ListIndex > -1 Then
        fgCursor True
        flTravaCampos False
        flLimpaCampos
        flCarregaListView
        fgCursor False
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, App.EXEName & " - cboEmpresa_Click"
    
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

    dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
    dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
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
        
    Me.Icon = mdiBUS.Icon
    fgCenterMe Me

    DoEvents
    Call flLimpaCampos
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    
    Me.Show
    
    Call fgCursor(True)
    
    fgCarregarCombos cboEmpresa, xmlMapaNavegacao, "Empresa", "CO_EMPR", "NO_EMPR"
    
    flTravaCampos True
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmSistema - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmSistema = Nothing

End Sub

Private Sub lstSistema_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

    lstSistema.Sorted = True
    lstSistema.SortKey = ColumnHeader.Index - 1

    If lstSistema.SortOrder = lvwAscending Then
        lstSistema.SortOrder = lvwDescending
    Else
        lstSistema.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstSistema_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    Call flLimpaCampos
    Call flXmlToInterface
    
    strOperacao = "Alterar"
    strKeyItemSelected = Item.Key
    txtSiglaSistema.Enabled = False
    
    If Not Trim$(txtSiglaSistema.Text) = vbNullString Then
        tlbCadastro.Buttons.Item("Excluir").Enabled = gblnPerfilManutencao 'True
        tlbCadastro.Buttons.Item("Salvar").Enabled = gblnPerfilManutencao
        tlbCadastro.Buttons.Item("Limpar").Enabled = gblnPerfilManutencao
    End If
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmSistema - lstSistema_ItemClick"

    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
            If fraSistema.Enabled = True Then
               txtSiglaSistema.SetFocus
            End If
        Case "Salvar"
            Call flSalvar
        Case "Excluir"
            If MsgBox("Confirma a exclusão do registro?", vbYesNo, "Exclusão de Registro") = vbYes Then
               strOperacao = "Excluir"
               Call flSalvar
               strOperacao = "Alterar"
            End If
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select
    
    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
    Exit Sub

ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmSistema - tlbCadastro_ButtonClick"
    
    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
    
End Sub

Private Sub txtSiglaSistema_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
