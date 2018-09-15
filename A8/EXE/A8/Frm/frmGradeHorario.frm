VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGradeHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segregação Acesso - Parametrização de Grade de Horário"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10245
   Begin VB.Frame frmAdm 
      Caption         =   "Administração "
      Height          =   1005
      Left            =   120
      TabIndex        =   6
      Top             =   6210
      Width           =   10005
      Begin VB.CheckBox chkGradeAtiva 
         Alignment       =   1  'Right Justify
         Caption         =   "Grade Ativa"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox txtMargemSeguranca 
         Height          =   315
         Left            =   1830
         MaxLength       =   3
         TabIndex        =   7
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblGradeHorario 
         AutoSize        =   -1  'True
         Caption         =   "Margem de Segurança"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   1620
      End
      Begin VB.Label lblGradeHorario 
         AutoSize        =   -1  'True
         Caption         =   "em minutos"
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   9
         Top             =   330
         Width           =   795
      End
   End
   Begin MSComctlLib.ListView lvwMargemSeguranca 
      Height          =   2055
      Left            =   120
      TabIndex        =   3
      Top             =   4020
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imlIcons"
      SmallIcons      =   "imlIcons"
      ColHdrIcons     =   "imlIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Instituição SPB"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo Grade"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Horário Inicial"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Horário Final"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Margem Segurança"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Grade Ativa"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.ComboBox cboGradeHorario 
      Height          =   315
      ItemData        =   "frmGradeHorario.frx":0000
      Left            =   105
      List            =   "frmGradeHorario.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   10005
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   6240
      TabIndex        =   4
      Top             =   7290
      Width           =   4005
      _ExtentX        =   7064
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
            Enabled         =   0   'False
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
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   120
      Top             =   7350
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":006E
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":0388
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":06A2
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":09BC
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":0CD6
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":1128
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":157A
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":19CC
            Key             =   "checked"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGradeHorario.frx":1D50
            Key             =   "unchecked"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMensagens 
      Height          =   2790
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   4921
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label lblGradeHorario 
      AutoSize        =   -1  'True
      Caption         =   "Mensagens SPB relacionadas à Grade de Horário"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3525
   End
   Begin VB.Label lblGradeHorario 
      AutoSize        =   -1  'True
      Caption         =   "Grade de Horário"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmGradeHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:34:09
'-------------------------------------------------
'' Objeto responsável pelo envio da solicitação (Atribuição da margem de segurança,
'' em minutos, para a grade de horário selecionada) à camada controladora de caso
'' de uso A8MIU.
''
'' São consideradas especificamente as classes destino:
''      A8MIU.clsGradeHorario
''      A8MIU.clsMiu
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLer                              As MSXML2.DOMDocument40
Private xmlDetalheGrade                     As MSXML2.DOMDocument40

Private Const strFuncionalidade             As String = "frmGradeHorario"
Private Const COL_INSTITUICAO               As Integer = 0
Private Const COL_TIPO_GRADE                As Integer = 1
Private Const COL_HORA_INI                  As Integer = 2
Private Const COL_HORA_FIM                  As Integer = 3
Private Const COL_MARGEM_SEG                As Integer = 4
Private Const COL_GRADE_ATIVA               As Integer = 5

Private arrChaves()                         As String

Private strKeyItemSelected                  As String

Private strOperacao                         As String
Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

On Error GoTo ErrorHandler

    If lvwMargemSeguranca.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lvwMargemSeguranca.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lvwMargemSeguranca_ItemClick objListItem
           lvwMargemSeguranca.ListItems(strKeyItemSelected).EnsureVisible
           blnEncontrou = True
           Exit For
        End If
    Next
    Set objListItem = Nothing
    
    If Not blnEncontrou Then
       flLimparDetalhe
    End If

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemListView", 0

End Sub

'' Encaminhar a solicitação (Leitura de detalhes da grade de horário selecionada)
'' à camada controladora de caso de uso (componente / classe / metodo ) : A8MIU.
'' clsGradeHorario.ObterDetalhesGradeHorarioO método retornará uma String XML para
'' a camada de interface.
Private Sub flCarregarDetalhesGradeHorario()

#If EnableSoap = 1 Then
    Dim objGradeHorario                     As MSSOAPLib30.SoapClient30
#Else
    Dim objGradeHorario                     As A8MIU.clsGradeHorario
#End If

Dim xmlDomNode                              As IXMLDOMNode

Dim strContrInstituicao                     As String
Dim strContrTipoGrade                       As String
Dim blnAdicionar                            As Boolean
Dim objListItem                             As ListItem
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    If cboGradeHorario.Text = vbNullString Then Exit Sub
    
    Set objGradeHorario = fgCriarObjetoMIU("A8MIU.clsGradeHorario")
    
    Call xmlDetalheGrade.loadXML(objGradeHorario.ObterDetalhesGradeHorario(fgObterCodigoCombo(Me.cboGradeHorario.Text), _
                                                                           vntCodErro, _
                                                                           vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    blnAdicionar = False
    For Each xmlDomNode In xmlDetalheGrade.selectNodes("/Repeat_GradeHorario/*")
        With xmlDomNode
            If strContrInstituicao <> .selectSingleNode("SQ_ISPB").Text Or _
               strContrTipoGrade <> .selectSingleNode("IN_TIPO_GRAD").Text Then
                
                If .selectSingleNode("IN_TIPO_GRAD").Text = enumTipoGradeHorario.Eventual Then
                    If fgDtXML_To_Date(.selectSingleNode("DT_EMIS_GRAD_BACEN").Text) = fgDataHoraServidor(DataAux) Then
                        blnAdicionar = True
                    Else
                        blnAdicionar = False
                    End If
                Else
                    blnAdicionar = True
                End If
            Else
                blnAdicionar = False
            End If
            
            If blnAdicionar Then
                strContrInstituicao = .selectSingleNode("SQ_ISPB").Text
                strContrTipoGrade = .selectSingleNode("IN_TIPO_GRAD").Text
                
                Set objListItem = lvwMargemSeguranca.ListItems.Add(, ";" & strContrInstituicao & _
                                                                     ";" & strContrTipoGrade & _
                                                                     ";" & .selectSingleNode("CO_GRAD_HORA").Text, _
                                                                     .selectSingleNode("NO_ISPB").Text)
                                                                     
                objListItem.SubItems(COL_TIPO_GRADE) = IIf(strContrTipoGrade = enumTipoGradeHorario.Padrao, "Padrão", "Eventual")
                objListItem.SubItems(COL_HORA_INI) = fgDtHrXml_To_Time(.selectSingleNode("HO_ABER").Text)
                objListItem.SubItems(COL_HORA_FIM) = fgDtHrXml_To_Time(.selectSingleNode("HO_ENCE").Text)
                objListItem.SubItems(COL_MARGEM_SEG) = .selectSingleNode("QT_TEMP_MARG_SEGR").Text
                objListItem.SubItems(COL_GRADE_ATIVA) = " "
                
                If .selectSingleNode("QT_TEMP_MARG_SEGR").Text = vbNullString Then
                   objListItem.Tag = gstrOperIncluir
                Else
                   objListItem.Tag = gstrOperNone
                End If
                
                If Val(.selectSingleNode("IN_SITU_GRAD_HORA").Text) <> enumIndicadorSimNao.nao Then
                    objListItem.ListSubItems(COL_GRADE_ATIVA).ReportIcon = "checked"
                Else
                    objListItem.ListSubItems(COL_GRADE_ATIVA).ReportIcon = "unchecked"
                End If
            
            End If
        End With
    Next
    
    Set objGradeHorario = Nothing
    Set xmlDomNode = Nothing
    
Exit Sub
ErrorHandler:
    Set objGradeHorario = Nothing
    Set xmlDomNode = Nothing
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarDetalhesGradeHorario", 0

End Sub

'' Encaminhar a solicitação (Leitura de mensagens SPB cadastradas, para o
'' preenchimento do listview) à camada controladora de caso de uso (componente /
'' classe / metodo ) : A8MIU.clsGradeHorario.ObterMensagensSPB
'' O método retornará uma String XML para a camada de interface.
Private Sub flCarregarMensagensSPB()

#If EnableSoap = 1 Then
    Dim objGradeHorario     As MSSOAPLib30.SoapClient30
#Else
    Dim objGradeHorario     As A8MIU.clsGradeHorario
#End If

Dim xmlMensagemSPB          As MSXML2.DOMDocument40
Dim xmlDomNode              As IXMLDOMNode
Dim objListItem             As ListItem
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler

    If cboGradeHorario.Text = vbNullString Then Exit Sub
    
    Set objGradeHorario = fgCriarObjetoMIU("A8MIU.clsGradeHorario")
    Set xmlMensagemSPB = CreateObject("MSXML2.DOMDocument.4.0")
        
    Call xmlMensagemSPB.loadXML(objGradeHorario.ObterMensagensSPB(fgObterCodigoCombo(Me.cboGradeHorario.Text), _
                                                                  vntCodErro, _
                                                                  vntMensagemErro))
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    For Each xmlDomNode In xmlMensagemSPB.selectNodes("/Repeat_MensagemSPB/*")
        With xmlDomNode
            Set objListItem = lvwMensagens.ListItems.Add(, , .selectSingleNode("CO_MESG").Text)
            objListItem.SubItems(1) = .selectSingleNode("NO_MESG").Text
        End With
    Next
    
    Set objGradeHorario = Nothing
    Set xmlMensagemSPB = Nothing
    Set xmlDomNode = Nothing
    
    Exit Sub

ErrorHandler:

    Set objGradeHorario = Nothing
    Set xmlMensagemSPB = Nothing
    Set xmlDomNode = Nothing
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarMensagensSPB", 0

End Sub

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

    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, Me.Name, "flInicializar")
    End If
    
    If xmlLer Is Nothing Then
       Set xmlLer = CreateObject("MSXML2.DOMDocument.4.0")
       xmlLer.loadXML xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_Propriedades/Grupo_GradeHorario").xml
    End If
    
    Call fgCarregarCombos(Me.cboGradeHorario, xmlMapaNavegacao, "GradeHorario", "CO_DOMI", "DE_DOMI")
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    fgCursor False
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'' Encaminhar a solicitação (Leitura da margem de seguração em minutos, para a
'' grade de horário selecionada) à camada controladora de caso de uso (componente
'' / classe / metodo ) : A8MIU.clsMiu.Executar
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
        
    With xmlLer.documentElement
        .selectSingleNode("//@Operacao").Text = "Ler"
        .selectSingleNode("//CO_GRAD_HORA").Text = arrChaves(3)
        .selectSingleNode("//SQ_ISPB").Text = arrChaves(1)
    End With
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    strLer = objMIU.Executar(xmlLer.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
   
    If strLer <> "" Then
        xmlLer.loadXML strLer
        With xmlLer.documentElement
            txtMargemSeguranca.Text = .selectSingleNode("QT_TEMP_MARG_SEGR").Text
            chkGradeAtiva.value = IIf(Val(.selectSingleNode("IN_SITU_GRAD_HORA").Text) = enumIndicadorSimNao.sim, _
                                                    vbChecked, _
                                                    vbUnchecked)
        End With
    Else
        txtMargemSeguranca.Text = ""
        chkGradeAtiva.value = vbChecked
    End If
    
Exit Sub
ErrorHandler:
    
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0

End Sub

'' É acionado através no botão 'Salvar' da barra de ferramentas.Tem como função,
'' encaminhar a solicitação (Atualização dos dados na tabela) à camada
'' controladora de caso de uso (componente / classe / metodo ) : A8MIU.clsMiu.
'' Executar
Private Sub flInterfaceToXml()
    
On Error GoTo ErrorHandler

    With xmlLer.documentElement
    
        .selectSingleNode("@Operacao").Text = strOperacao
        .selectSingleNode("QT_TEMP_MARG_SEGR").Text = txtMargemSeguranca.Text
        .selectSingleNode("IN_SITU_GRAD_HORA").Text = IIf(chkGradeAtiva.value = vbChecked, _
                                                                enumIndicadorSimNao.sim, _
                                                                enumIndicadorSimNao.nao)
        
        If strOperacao = "Incluir" Then
           .selectSingleNode("CO_GRAD_HORA").Text = arrChaves(3)
           .selectSingleNode("SQ_ISPB").Text = arrChaves(1)
        End If
        
    End With
    
Exit Sub
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0
    
End Sub
'Limpar os detalhes da interface
Private Sub flLimparDetalhe()

On Error GoTo ErrorHandler

    With Me
        .lvwMensagens.ListItems.Clear
        .lvwMargemSeguranca.ListItems.Clear
        .lvwMensagens.ListItems.Clear
        .txtMargemSeguranca.Text = vbNullString
        .txtMargemSeguranca.Enabled = False
        .chkGradeAtiva.value = vbChecked
        .chkGradeAtiva.Enabled = False
        .lblGradeHorario(2).Enabled = False
        .lblGradeHorario(3).Enabled = False
        .tlbCadastro.Buttons(gstrOperExcluir).Enabled = False
        .tlbCadastro.Buttons(gstrSalvar).Enabled = False
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparDetalhe", 0
End Sub

'Limpar interface
Private Sub flLimparTela()

On Error GoTo ErrorHandler

    cboGradeHorario.ListIndex = -1
    Call flLimparDetalhe

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flLimparTela", 0
End Sub

' Salva as alterações efetuadas através da camada controladora de casos de uso
' MIU, método A8MIU.clsMIU.Executar
Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
#End If

Dim strRetorno              As String
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
    
    fgCursor True
    
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
           strKeyItemSelected = lvwMargemSeguranca.SelectedItem.Key
        End If
        strOperacao = gstrOperAlterar
    End If

    flLimparDetalhe
    Call flCarregarMensagensSPB
    Call flCarregarDetalhesGradeHorario
    
    Set objMIU = Nothing
    
    fgCursor False
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
    
Exit Sub
ErrorHandler:
    fgCursor False
    Set objMIU = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro

    fgRaiseError App.EXEName, Me.Name, "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

'Validar os campos obrigatórios para execução da funcionalidade especificada.

Private Function flValidarCampos() As String

On Error GoTo ErrorHandler

    If strOperacao <> gstrOperExcluir Then
        If lvwMensagens.ListItems.Count <= 0 Then
            flValidarCampos = "Favor informar Mensagens SPB relacionadas à Grade de Horário."
            lvwMensagens.SetFocus
            Exit Function
        End If
    
        If Trim$(txtMargemSeguranca.Text) = vbNullString Then
            txtMargemSeguranca.Text = "0"
        End If
    End If
    
    flValidarCampos = ""

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0
End Function

Private Sub cboGradeHorario_Click()

On Error GoTo ErrorHandler

    fgCursor True
    Call flLimparDetalhe
    Call flCarregarMensagensSPB
    Call flCarregarDetalhesGradeHorario
    fgCursor
    
Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)
    
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    fgCursor True
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao
    tlbCadastro.Buttons(gstrSalvar).Enabled = False
    
    txtMargemSeguranca.Enabled = False
    chkGradeAtiva.Enabled = False
    lblGradeHorario(2).Enabled = False
    lblGradeHorario(3).Enabled = False
    
    Set xmlDetalheGrade = CreateObject("MSXML2.DOMDocument.4.0")
    Call flInicializar
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmGradeHorario = Nothing
End Sub

Private Sub lvwMargemSeguranca_ItemClick(ByVal Item As MSComctlLib.ListItem)

On Error GoTo ErrorHandler

    If lvwMargemSeguranca.SelectedItem Is Nothing Then Exit Sub
    
    txtMargemSeguranca.Enabled = True
    chkGradeAtiva.Enabled = True
    lblGradeHorario(2).Enabled = True
    lblGradeHorario(3).Enabled = True
    
    arrChaves = Split(lvwMargemSeguranca.SelectedItem.Key, ";")
    
    If Item.Tag = gstrOperNone Then
        Call flXmlToInterface
    Else
        txtMargemSeguranca.Text = lvwMargemSeguranca.SelectedItem.SubItems(COL_MARGEM_SEG)
        chkGradeAtiva.value = IIf(lvwMargemSeguranca.SelectedItem.ListSubItems(COL_GRADE_ATIVA).ReportIcon = "checked", _
                                                vbChecked, _
                                                vbUnchecked)
    End If
    
    strKeyItemSelected = Item.Key
    
    If txtMargemSeguranca.Text = vbNullString Then
       strOperacao = "Incluir"
    Else
       strOperacao = gstrOperAlterar
    End If
    
    tlbCadastro.Buttons(gstrOperExcluir).Enabled = _
        (IIf(Trim$(txtMargemSeguranca.Text) = vbNullString, False, True) And gblnPerfilManutencao)
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao
    
Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, "lvwMargemSeguranca_ItemClick", Me.Caption)
    
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True

    Select Case Button.Key
    Case "Limpar"
        Call flLimparTela
    Case gstrOperExcluir
        If Not lvwMargemSeguranca.SelectedItem Is Nothing Then
            If MsgBox("Confirma a exclusão do registro ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
                strOperacao = gstrOperExcluir
                Call flSalvar
            End If
        End If
    Case gstrSalvar
        Call flSalvar
    Case gstrSair
        fgCursor
        Unload Me
        Exit Sub
    End Select
    
    flPosicionaItemListView
    
    fgCursor
    
Exit Sub
ErrorHandler:

    fgCursor
    
    Call mdiLQS.uctlogErros.MostrarErros(Err, "tlbCadastro_ButtonClick", Me.Caption)
    flRecarregar
    
End Sub

Public Sub flRecarregar()

On Error GoTo ErrorHandler

    fgCursor True

    flLimparDetalhe
    Call flCarregarMensagensSPB
    Call flCarregarDetalhesGradeHorario
    
    If strOperacao <> gstrOperNone Then
       flPosicionaItemListView
    End If
    
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - flRecarregar"
End Sub

Private Sub txtMargemSeguranca_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> Asc(vbBack) Then KeyAscii = 0
End Sub
