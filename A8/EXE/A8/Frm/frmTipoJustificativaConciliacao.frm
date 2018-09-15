VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTipoJustificativaConciliacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Tipo Justificativa Conciliação"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6990
   Begin VB.Frame fraCadastro 
      Height          =   3870
      Left            =   30
      TabIndex        =   5
      Top             =   -60
      Width           =   6885
      Begin VB.Frame fraTipoLiquidacao 
         Caption         =   "Justificativa Conciliação"
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
         TabIndex        =   9
         Top             =   2070
         Width           =   4275
         Begin VB.TextBox txtDescricaoTipoJustificativa 
            Height          =   315
            Left            =   90
            TabIndex        =   2
            Top             =   1170
            Width           =   4080
         End
         Begin NumBox.Number numTipoJustificativa 
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
         Begin VB.Label Label01 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   90
            TabIndex        =   11
            Top             =   330
            Width           =   495
         End
         Begin VB.Label lblDescricao 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   105
            TabIndex        =   10
            Top             =   975
            Width           =   720
         End
      End
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
         TabIndex        =   6
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
            Format          =   58851329
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   1170
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   58851329
            CurrentDate     =   37622
            MaxDate         =   73050
            MinDate         =   37622
         End
         Begin VB.Label lblDataFimVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Fim"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   975
            Width           =   240
         End
         Begin VB.Label lblDataInicioVigencia 
            AutoSize        =   -1  'True
            Caption         =   "Início "
            Height          =   195
            Left            =   150
            TabIndex        =   7
            Top             =   330
            Width           =   450
         End
      End
      Begin MSComctlLib.ListView lstTipoJustificativa 
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
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
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   3900
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
            Picture         =   "frmTipoJustificativaConciliacao.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":031A
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":0634
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":094E
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":0C68
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":10BA
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTipoJustificativaConciliacao.frx":150C
            Key             =   "ItemElementar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   2940
      TabIndex        =   12
      Top             =   3840
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
Attribute VB_Name = "frmTipoJustificativaConciliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:35:19
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Administração do cadastro de
'' Tipos de Justificativa de Cancelamento) à camada controladora de caso de uso
'' A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsMIU
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private Const strFuncionalidade             As String = "frmTipoJustificativaConciliacao"
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Reposiciona os items no ListView de acordo com o critério selecionado
Private Sub flPosicionaItemListView()
Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lstTipoJustificativa.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstTipoJustificativa.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstTipoJustificativa_ItemClick objListItem
           lstTipoJustificativa.ListItems(strKeyItemSelected).EnsureVisible
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

'' É acionado através no botão 'Salvar' da barra de ferramentas.
''
'' Tem como função, encaminhar a solicitação (Atualização dos dados na tabela) à
'' camada controladora de caso de uso (componente / classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
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
    Call objMIU.Executar(xmlLer.xml, _
                         vntCodErro, _
                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strOperacao <> gstrOperExcluir Then
        xmlLer.documentElement.selectSingleNode("//@Operacao").Text = "Ler"
        xmlLer.loadXML objMIU.Executar(xmlLer.xml, _
                                       vntCodErro, _
                                       vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If

        If strOperacao = "Incluir" Then
           strKeyItemSelected = "K" & numTipoJustificativa.Valor
        End If
        strOperacao = gstrOperAlterar
        numTipoJustificativa.Enabled = False
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
            strKeyItemSelected = "K" & .selectSingleNode("TP_JUST_CNCL").Text
       End With
    End If
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

'Valida o preenchimento dos campos
Private Function flValidarCampos() As String
    
On Error GoTo ErrorHandler

    If numTipoJustificativa = 0 Then
        flValidarCampos = "Digite o código do Tipo de Justificativa Conciliação."
        numTipoJustificativa.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescricaoTipoJustificativa) = "" Then
        flValidarCampos = "Informe a Descrição do Tipo de Justificativa Conciliação."
        txtDescricaoTipoJustificativa.SetFocus
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

'Limpa o conteúdo dos campos
Private Sub flLimpaCampos()
        
On Error GoTo ErrorHandler

    strOperacao = "Incluir"
    
    numTipoJustificativa.Valor = 0
    numTipoJustificativa.Enabled = True
    
    txtDescricaoTipoJustificativa.Text = ""
        
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
    dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
    
    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = Null
    
    dtpDataInicioVigencia.Enabled = True
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = False

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'' Encaminhar a solicitação (Leitura de detalhes do tipo de justificativa
'' selecionado) à camada controladora de caso de uso (componente / classe / metodo
'' ) :
''
'' A8MIU.clsMiu.Executar
''
'' O método retornará uma String XML para a camada de interface.
''
Private Sub flXmlToInterface()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
        
    numTipoJustificativa.Enabled = False
    dtpDataInicioVigencia.Enabled = False
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//TP_JUST_CNCL").Text = lstTipoJustificativa.SelectedItem.Text
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlLer.loadXML objMIU.Executar(xmlLer.xml, _
                                   vntCodErro, _
                                   vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
   
    With xmlLer.documentElement
   
        numTipoJustificativa.Valor = CLng(.selectSingleNode("TP_JUST_CNCL").Text)
        txtDescricaoTipoJustificativa.Text = .selectSingleNode("NO_TIPO_JUST_CNCL").Text
        dtpDataInicioVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
        dtpDataInicioVigencia.value = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
        If dtpDataInicioVigencia.value > fgDataHoraServidor(Data) Then
            dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
            dtpDataInicioVigencia.Enabled = True
        End If
        
        If Trim(.selectSingleNode("DT_FIM_VIGE").Text) <> gstrDataVazia Then
            If fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text) < fgDataHoraServidor(Data) Then
                dtpDataFimVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text)
                dtpDataInicioVigencia.Enabled = True
            Else
                dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.value, fgDataHoraServidor(Data))
            End If
            dtpDataFimVigencia.value = fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE").Text)
        Else
            dtpDataFimVigencia.value = fgDataHoraServidor(DataAux)
            dtpDataFimVigencia.value = Null
        End If
        
    End With
        
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0

End Sub

'Preenche o conteúdo do XML com o conteúdo dos campos apresentados em tela
Private Function flInterfaceToXml() As String
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
         .selectSingleNode("//@Operacao").Text = strOperacao
    
         If strOperacao = "Incluir" Then
            .selectSingleNode("//TP_JUST_CNCL").Text = numTipoJustificativa.Valor
            
         ElseIf strOperacao = gstrOperExcluir Then
            Exit Function
         End If
         
         .selectSingleNode("//NO_TIPO_JUST_CNCL").Text = txtDescricaoTipoJustificativa.Text
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
        
    End With
    
Exit Function
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Function

'' Encaminhar a solicitação (Obtenção das propriedades da tabela do cadastro em
'' questão) à camada controladora de caso de uso (componente / classe / metodo ) :
''
''
'' A8MIU.clsMiu.ObterMapaNavegacao
''
'' O método retornará uma String XML para a camada de interface.
''
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
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmTipoJustificativaConciliacao", "flInicializar")
    End If
    
    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_TipoJustificativa").xml
    End If
    
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0
    
End Sub

'Configura o número máximo de caracteres permitidos em cada campo
Private Sub flDefinirTamanhoMaximoCampos()
              
On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement
        txtDescricaoTipoJustificativa.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_TipoJustificativa/NO_TIPO_JUST_CNCL/@Tamanho").Text
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
        
End Sub

'' Encaminhar a solicitação (Leitura de todos os tipos de justificativa, para
'' preenchimento do listview) à camada controladora de caso de uso (componente /
'' classe / metodo ) :
''
'' A8MIU.clsMiu.Executar
''
'' O método retornará uma String XML para a camada de interface.
''
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

    lstTipoJustificativa.ListItems.Clear
    
    Set xmlDomNode = Nothing
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoJustificativa/TP_VIGE").Text = "N"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoJustificativa/TP_SEGR").Text = "S"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoJustificativa/@Operacao").Text = "LerTodos"
    strPropriedades = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_TipoJustificativa").xml
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

    For Each xmlDomNode In xmlLerTodos.selectNodes("//Repeat_TipoJustificativa/*")
        With xmlDomNode
                
            Set objListItem = lstTipoJustificativa.ListItems.Add(, "K" & .selectSingleNode("TP_JUST_CNCL").Text, .selectSingleNode("TP_JUST_CNCL").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_TIPO_JUST_CNCL").Text
            objListItem.SubItems(2) = fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE").Text)
            
            If CStr(.selectSingleNode("DT_FIM_VIGE").Text) <> CStr(gstrDataVazia) Then
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
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

Private Sub dtpDataFimVigencia_Change()

On Error GoTo ErrorHandler

    If Not IsNull(dtpDataFimVigencia.value) Then
        If dtpDataFimVigencia.value < fgDataHoraServidor(Data) Then
            dtpDataFimVigencia.value = fgDataHoraServidor(Data)
            dtpDataFimVigencia.MinDate = fgDataHoraServidor(Data)
        End If
    End If
    If dtpDataInicioVigencia.value < fgDataHoraServidor(Data) And dtpDataInicioVigencia.Enabled Then
        dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
        dtpDataInicioVigencia.MinDate = dtpDataInicioVigencia.value
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - dtpDataFimVigencia_Change"
End Sub

Private Sub dtpDataFimVigencia_KeyPress(KeyAscii As Integer)
    KeyAscii = 0: Beep
End Sub

Private Sub dtpDataInicioVigencia_Change()

On Error GoTo ErrorHandler

    If dtpDataInicioVigencia.value < fgDataHoraServidor(Data) Then
        dtpDataInicioVigencia.value = fgDataHoraServidor(Data)
        dtpDataInicioVigencia.MinDate = fgDataHoraServidor(Data)
    End If
    
    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.value
    dtpDataFimVigencia.value = dtpDataFimVigencia.MinDate
    dtpDataFimVigencia.value = Null

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - dtpDataInicioVigencia_Change"
End Sub

Private Sub dtpDataInicioVigencia_KeyPress(KeyAscii As Integer)
    KeyAscii = 0: Beep
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
        
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
    Call flLimpaCampos
    
    Call fgCursor(True)
    Call flInicializar
    Call flDefinirTamanhoMaximoCampos
    Call flCarregaListView
               
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoJustificativaConciliacao - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set xmlMapaNavegacao = Nothing
    Set xmlLer = Nothing
    Set frmTipoJustificativaConciliacao = Nothing
End Sub

Private Sub lstTipoJustificativa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

On Error GoTo ErrorHandler

    lstTipoJustificativa.Sorted = True
    lstTipoJustificativa.SortKey = ColumnHeader.Index - 1

    If lstTipoJustificativa.SortOrder = lvwAscending Then
        lstTipoJustificativa.SortOrder = lvwDescending
    Else
        lstTipoJustificativa.SortOrder = lvwAscending
    End If

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lstTipoJustificativa_ColumnClick"

End Sub

Private Sub lstTipoJustificativa_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler
            
    Call fgCursor(True)
    
    Call flLimpaCampos
    strOperacao = gstrOperAlterar
    Call flXmlToInterface
    
    strKeyItemSelected = Item.Key
    
    numTipoJustificativa.Enabled = False
    
    If numTipoJustificativa > 0 Then
        tlbCadastro.Buttons(gstrOperExcluir).Enabled = gblnPerfilManutencao
    End If
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoJustificativaConciliacao - lstTipoJustificativa_ItemClick", Me.Caption
    flRecarregar

End Sub

'Remonta a tela com os dados mais atuais
Private Sub flRecarregar()

On Error GoTo ErrorHandler

    flLimpaCampos
    flCarregaListView

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
            If fraTipoLiquidacao.Enabled = True And numTipoJustificativa.Enabled = True Then
               numTipoJustificativa.SetFocus
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
        
    mdiLQS.uctlogErros.MostrarErros Err, "frmTipoJustificativaConciliacao - tlbCadastro_ButtonClick", Me.Caption
    
    Call flCarregaListView
    
    If strOperacao = gstrOperExcluir Then
        flLimpaCampos
    ElseIf strOperacao <> gstrOperExcluir Then
        flPosicionaItemListView
    End If

End Sub
