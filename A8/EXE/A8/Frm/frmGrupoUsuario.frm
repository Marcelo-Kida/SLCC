VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGrupoUsuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Segregação Acesso - Grupo Usuário"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCadastro 
      Height          =   3870
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   6885
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
         TabIndex        =   10
         Top             =   2070
         Width           =   1995
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   3
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
            Left            =   90
            TabIndex        =   8
            Top             =   1170
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
         Begin VB.Label lblDataInicioVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Início "
            Height          =   195
            Left            =   150
            TabIndex        =   12
            Top             =   330
            Width           =   450
         End
         Begin VB.Label lblDataFimVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   975
            Width           =   240
         End
      End
      Begin VB.Frame fraTipoLiquidacao 
         Caption         =   "Grupo Usuário"
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
         Begin VB.TextBox txtNomeUsuario 
            Height          =   315
            Left            =   90
            TabIndex        =   2
            Top             =   1170
            Width           =   4080
         End
         Begin NumBox.Number numCodigoUsuario 
            Height          =   315
            Left            =   120
            TabIndex        =   1
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
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Nome"
            Height          =   195
            Left            =   105
            TabIndex        =   9
            Top             =   975
            Width           =   420
         End
         Begin VB.Label lblCodigo 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   90
            TabIndex        =   7
            Top             =   330
            Width           =   495
         End
      End
      Begin MSComctlLib.ListView lstGrupoUsuario 
         Height          =   1605
         Left            =   210
         TabIndex        =   0
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
      Left            =   120
      Top             =   4050
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
            Picture         =   "frmGrupoUsuario.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGrupoUsuario.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   3990
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
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmGrupoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:11
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Administração geral do cadastro
'' de Grupos de Usuarios) à camada controladora de caso de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMiu
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private strOperacao                         As String

Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmGrupoUsuario"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer
Private strUltimaAtualizacao                As String

Private Sub flPosicionaItemListView()
Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstGrupoUsuario.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstGrupoUsuario.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstGrupoUsuario_ItemClick objListItem
           lstGrupoUsuario.ListItems(strKeyItemSelected).EnsureVisible
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

Private Sub flFormataListView()

    lstGrupoUsuario.ColumnHeaders.Add 1, , "Código", 1000, lvwColumnLeft
    lstGrupoUsuario.ColumnHeaders.Add 2, , "Nome", 2420, lvwColumnLeft
    lstGrupoUsuario.ColumnHeaders.Add 3, , "Data Início", 1440, lvwColumnLeft
    lstGrupoUsuario.ColumnHeaders.Add 4, , "Data Fim ", 1440, lvwColumnLeft

End Sub
'' É acionado através no botão 'Salvar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Atualização dos dados na tabela) à camada
'' controladora de caso de uso (componente / classe / metodo ) : A8MIU.clsMiu.
'' Executar
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strRetorno              As String
Dim strPropriedades         As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

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
           strKeyItemSelected = "K" & numCodigoUsuario.Valor
        End If
        strOperacao = gstrOperAlterar
        numCodigoUsuario.Enabled = False
    Else
        flLimpaCampos
    End If
    Set objMIU = Nothing
    
    Call flCarregaListView
    Call fgCursor(False)
    
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    Set objMIU = Nothing
    
    If strOperacao <> gstrOperExcluir Then
       With xmlLer.documentElement
            strKeyItemSelected = "K" & .selectSingleNode("CO_GRUP_USUA").Text
       End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, "frmGrupoUsuario", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Validar os campos obrigatórios para execução da funcionalidade especificada.

Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If numCodigoUsuario = 0 Then
        flValidarCampos = "Digite o código do Grupo de Usuário."
        numCodigoUsuario.SetFocus
        Exit Function
    End If
    
    If Trim(txtNomeUsuario) = "" Then
        flValidarCampos = "Informe o Nome do Grupo de Usuário."
        txtNomeUsuario.SetFocus
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
        
    numCodigoUsuario.Valor = 0
    numCodigoUsuario.Enabled = True
    
    txtNomeUsuario.Text = ""
    
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
    dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
    
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(Data)
    dtpDataFimVigencia.value = fgDataHoraServidor(Data)
    dtpDataFimVigencia.value = Null
        
    dtpDataInicioVigencia.Enabled = True
    dtpDataFimVigencia.Enabled = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
    
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, "frmGrupoUsuario", "flLimpaCampos", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'' Encaminhar a solicitação (Leitura de detalhes do grupo de usuário selecionado)
'' à camada controladora de caso de uso (componente / classe / metodo ) : A8MIU.
'' clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    numCodigoUsuario.Enabled = False
    dtpDataInicioVigencia.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//CO_GRUP_USUA").Text = lstGrupoUsuario.SelectedItem.Text
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call xmlLer.loadXML(objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
   
    With xmlLer.documentElement
        numCodigoUsuario.Valor = CLng(.selectSingleNode("CO_GRUP_USUA").Text)
        txtNomeUsuario.Text = .selectSingleNode("NO_GRUP_USUA").Text
        Call fgCarregaDataVigencia(dtpDataInicioVigencia, _
                                   dtpDataFimVigencia, _
                                   .selectSingleNode("DT_INIC_VIGE").Text, _
                                   .selectSingleNode("DT_FIM_VIGE").Text)
                             
        strUltimaAtualizacao = .selectSingleNode("DH_ULTI_ATLZ").Text
    End With
    
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, "frmGrupoUsuario", "flXmlToInterface", lngCodigoErroNegocio, intNumeroSequencialErro
    
End Sub

'Carregar a interface com as informações do xml
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
         .selectSingleNode("@Operacao").Text = strOperacao
         .selectSingleNode("CO_GRUP_USUA").Text = numCodigoUsuario.Valor
         .selectSingleNode("NO_GRUP_USUA").Text = txtNomeUsuario.Text
         .selectSingleNode("DT_INIC_VIGE").Text = fgDate_To_DtXML(dtpDataInicioVigencia.value)
    
         If strOperacao <> gstrOperExcluir Then
             If Not IsNull(dtpDataFimVigencia.value) Then
                If .selectSingleNode("DT_FIM_VIGE").Text <> gstrDataVazia Then
            
                    If fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text) <> dtpDataFimVigencia.value Then
                       If MsgBox("Deseja desativar o registro a partir da data: " & dtpDataFimVigencia.value, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                          .selectSingleNode("DT_FIM_VIGE").Text = fgDate_To_DtXML(dtpDataFimVigencia.value)
                       End If
                    End If
                Else
                    If MsgBox("Deseja desativar o registro a partir da data: " & dtpDataFimVigencia.value, vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                       .selectSingleNode("DT_FIM_VIGE").Text = fgDate_To_DtXML(dtpDataFimVigencia.value)
                    Else
                       dtpDataFimVigencia.value = Null
                       .selectSingleNode("DT_FIM_VIGE").Text = ""
                    End If
                End If
             Else
                .selectSingleNode("DT_FIM_VIGE").Text = ""
             End If
         End If
    
    End With
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function

'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
'' A8MIU.clsMiu.ObterMapaNavegacao
'' O método retornará uma String XML para a camada de interface.
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmGrupoUsuario", "flInicializar")
    End If
    
    Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlLer.loadXML(xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_GrupoUsuario").xml)
    
    Set objMIU = Nothing
    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

Private Sub flDefinirTamanhoMaximoCampos()

On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement
        txtNomeUsuario.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/NO_GRUP_USUA/@Tamanho").Text
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
End Sub

'' Encaminhar a solicitação (Leitura de todos os grupos de usuário cadastrados,
'' para preenchimento do listview) à camada controladora de caso de uso
'' (componente / classe / metodo ) : A8MIU.clsMiu.Executar
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregaListView()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim xmlDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim strNomeTipoLiquidacao                   As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    fgCursor True
    
    lstGrupoUsuario.ListItems.Clear
    lstGrupoUsuario.HideSelection = False
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/TP_SEGR").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario/@Operacao").Text = "LerTodos"
    Call xmlLerTodos.loadXML(objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_GrupoUsuario").xml, _
                                             vntCodErro, _
                                             vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_GrupoUsuario/*")
        With xmlDomNode
                
            Set objListItem = lstGrupoUsuario.ListItems.Add(, "K" & .selectSingleNode("CO_GRUP_USUA").Text, .selectSingleNode("CO_GRUP_USUA").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_GRUP_USUA").Text
            objListItem.SubItems(2) = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
            
            If CStr(.selectSingleNode("DT_FIM_VIGE").Text) <> gstrDataVazia Then
                objListItem.SubItems(3) = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text)
            End If
        
        End With
    Next

    Set xmlLerTodos = Nothing
    fgCursor
    
    Exit Sub
    
ErrorHandler:
    Set objMIU = Nothing
    Set xmlLerTodos = Nothing
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub dtpDataFimVigencia_Change()

On Error GoTo ErrorHandler

    Call fgDataVigenciaFimChange(dtpDataInicioVigencia, dtpDataFimVigencia)

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - dtpDataFimVigencia_Change"
End Sub

Private Sub dtpDataFimVigencia_KeyPress(KeyAscii As Integer)
    KeyAscii = 0: Beep
End Sub

Private Sub dtpDataInicioVigencia_Change()

On Error GoTo ErrorHandler

    Call fgDataVigenciaInicioChange(dtpDataInicioVigencia, dtpDataFimVigencia)

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
    
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    
    Call flLimpaCampos
    
    Call fgCursor(True)
    Call flFormataListView
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    Call flCarregaListView
           
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing
End Sub

Private Sub lstGrupoUsuario_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lstGrupoUsuario.Sorted = True
    lstGrupoUsuario.SortKey = ColumnHeader.Index - 1

    If lstGrupoUsuario.SortOrder = lvwAscending Then
        lstGrupoUsuario.SortOrder = lvwDescending
    Else
        lstGrupoUsuario.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstGrupoUsuario_ColumnClick"

End Sub

Private Sub lstGrupoUsuario_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface
    
    strKeyItemSelected = Item.Key
    
    numCodigoUsuario.Enabled = False
    
    If numCodigoUsuario > 0 Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - lstGrupoUsuario_ItemClick", Me.Caption
    flRecarregar

End Sub

Private Sub flRecarregar()

On Error GoTo ErrorHandler

    fgCursor True

    flLimpaCampos
    Call flCarregaListView

    fgCursor

Exit Sub
ErrorHandler:

    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flRecarregar"
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
            If fraTipoLiquidacao.Enabled = True And numCodigoUsuario.Enabled = True Then
               numCodigoUsuario.SetFocus
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
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmGrupoUsuario - tlbCadastro_ButtonClick", Me.Caption
    
    Call flCarregaListView
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If
    
End Sub

