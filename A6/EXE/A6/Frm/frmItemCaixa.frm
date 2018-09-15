VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemCaixa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Itens de Caixa"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8790
   Begin VB.OptionButton optTipoCaixa 
      Caption         =   "Caixa Futuro"
      Height          =   255
      Index           =   2
      Left            =   1500
      TabIndex        =   13
      Tag             =   "CaixaFuturo"
      Top             =   50
      Width           =   1245
   End
   Begin VB.OptionButton optTipoCaixa 
      Caption         =   "Sub Reserva"
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Tag             =   "SubReserva"
      Top             =   50
      Width           =   1335
   End
   Begin VB.Frame fraMoldura 
      Height          =   5715
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   8685
      Begin MSComctlLib.TreeView treCadastro 
         Height          =   5340
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   8460
         _ExtentX        =   14923
         _ExtentY        =   9419
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imlIcons"
         Appearance      =   1
      End
   End
   Begin VB.Frame fraCadastro 
      Height          =   1290
      Left            =   60
      TabIndex        =   6
      Top             =   5760
      Width           =   8685
      Begin VB.OptionButton optTipoItemCaixa 
         Caption         =   "Item de Grupo"
         DownPicture     =   "frmItemCaixa.frx":0000
         Height          =   555
         Index           =   1
         Left            =   975
         Picture         =   "frmItemCaixa.frx":00EA
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   645
         Width           =   1320
      End
      Begin VB.OptionButton optTipoItemCaixa 
         Alignment       =   1  'Right Justify
         Caption         =   "Item Elementar"
         DownPicture     =   "frmItemCaixa.frx":01D4
         Height          =   555
         Index           =   2
         Left            =   2295
         Picture         =   "frmItemCaixa.frx":02BE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   645
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.TextBox txtDescricao 
         Height          =   315
         Left            =   975
         TabIndex        =   0
         Top             =   190
         Width           =   4380
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   7080
         TabIndex        =   3
         Top             =   190
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         Format          =   53870593
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin MSComCtl2.DTPicker dtpFim 
         Height          =   315
         Left            =   7080
         TabIndex        =   4
         Top             =   585
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   53870593
         CurrentDate     =   37622
         MaxDate         =   73050
         MinDate         =   37622
      End
      Begin VB.Label lblDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblDataInicioVigencia 
         AutoSize        =   -1  'True
         Caption         =   "Data Início Vigência"
         Height          =   195
         Left            =   5520
         TabIndex        =   8
         Top             =   315
         Width           =   1455
      End
      Begin VB.Label lblDataFimVigencia 
         AutoSize        =   -1  'True
         Caption         =   "Data Fim Vigência"
         Height          =   195
         Left            =   5685
         TabIndex        =   9
         Top             =   705
         Width           =   1290
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4830
      TabIndex        =   5
      Top             =   7095
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
            ImageIndex      =   4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Excluir"
            Key             =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Salvar"
            Key             =   "Salvar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   6675
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
            Picture         =   "frmItemCaixa.frx":03A8
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":06C2
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":09DC
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":0CF6
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":1010
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":1462
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":18B4
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":1D06
            Key             =   "OpenReservaFuturo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":1E18
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":1F12
            Key             =   "Closed"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemCaixa.frx":200C
            Key             =   "Leaf"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmItemCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' Este componente tem como objetivo, possibilitar ao usuário a administração do cadastro de itens de caixa.

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlItemCaixa                         As MSXML2.DOMDocument40
Private blnItemCaixaExiste                  As Boolean
Private Const strFuncionalidade             As String = "frmItemCaixa"

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

Private Sub dtpFim_Change()

On Error GoTo ErrorHandler

    If Not IsNull(dtpFim.Value) Then
        If dtpFim.Value < fgDataHoraServidor(Data) Then
            dtpFim.Value = fgDataHoraServidor(Data)
            dtpFim.MinDate = fgDataHoraServidor(Data)
        End If
    End If
    If dtpInicio.Value < fgDataHoraServidor(Data) And dtpInicio.Enabled Then
        dtpInicio.Value = fgDataHoraServidor(Data)
        dtpInicio.MinDate = dtpInicio.Value
    End If

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - dtpFim_Change"
End Sub

Private Sub dtpFim_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub dtpInicio_Change()

On Error GoTo ErrorHandler

    If dtpInicio.Value < fgDataHoraServidor(HORA) Then
        dtpInicio.Value = fgDataHoraServidor(HORA)
        dtpInicio.MinDate = fgDataHoraServidor(HORA)
    End If

    dtpFim.MinDate = dtpInicio.Value
    dtpFim.Value = dtpInicio.Value
    dtpFim.Value = Null

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - dtpInicio_Change"
End Sub

Private Sub dtpInicio_KeyPress(KeyAscii As Integer)
    
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

    Me.Icon = mdiSBR.Icon

    fgCursor True
    fgCenterMe Me
    Me.Show
    DoEvents
    
    blnItemCaixaExiste = False
    
    flInicializar
    flDefinirTamanhoMaximoCampos
    flLimpar
    fraCadastro.Enabled = False
    tlbCadastro.Buttons("Limpar").Enabled = False
    tlbCadastro.Buttons("Excluir").Enabled = False
    tlbCadastro.Buttons("Salvar").Enabled = False
    fgCursor

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixa - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set frmItemCaixa = Nothing

End Sub

Private Sub optTipoCaixa_Click(Index As Integer)

#If EnableSoap = 1 Then
    Dim objMIU      As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU      As A6MIU.clsMIU
#End If

Dim xmlLerTodos     As MSXML2.DOMDocument40
Dim xmlDomNode      As MSXML2.IXMLDOMNode
Dim strLerTodos     As String
Dim vntCodErro      As Variant
Dim vntMensagemErro As Variant
    
On Error GoTo ErrorHandler

    fgCursor True
    
    Set xmlDomNode = xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa")
    xmlDomNode.selectSingleNode("@Operacao").Text = "LerTodos"
    xmlDomNode.selectSingleNode("TP_VIGE").Text = "N"
    Select Case Index
           Case enumTipoCaixa.CaixaFuturo
                xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaFuturo
           Case enumTipoCaixa.CaixaSubReserva
                xmlDomNode.selectSingleNode("TP_CAIX").Text = enumTipoCaixa.CaixaSubReserva
    End Select

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    
    strLerTodos = objMIU.Executar(xmlDomNode.xml, vntCodErro, vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If strLerTodos <> "" Then
       Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
       If Not xmlLerTodos.loadXML(strLerTodos) Then
          Call fgErroLoadXML(xmlLerTodos, App.EXEName, "frmItemCaixaTipoOperacao", "optTipoCaixa_Click")
       End If
    Else
       Set objMIU = Nothing
       fgCursor
       Exit Sub
    End If

    Set objMIU = Nothing

    fgCarregarTreItemCaixa treCadastro, xmlLerTodos, Me
    
    Call treCadastro_NodeClick(treCadastro.Nodes(1))

    Me.dtpInicio.Value = Date
    Me.dtpFim.Value = Date
    Me.dtpFim.Value = Null
    
    fraCadastro.Enabled = True
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    Set xmlDomNode = Nothing
    
    fgCursor
    Exit Sub
    
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixaTipoOperacao - optTipoCaixa_Click"

End Sub

Private Sub optTipoItemCaixa_Click(Index As Integer)

On Error GoTo ErrorHandler

    If Not treCadastro.SelectedItem Is Nothing Then
        If treCadastro.SelectedItem.children > 0 And _
            Index = enumTipoItemCaixa.Elementar And _
            blnItemCaixaExiste Then
            frmMural.Display = "Item de Caixa de Grupo não pode ser modificado para elementar se possuir Item de Caixa abaixo "
            frmMural.IconeExibicao = IconInformation
            frmMural.Show vbModal
            optTipoItemCaixa.Item(enumTipoItemCaixa.Grupo).Value = True
        Else
            If fgObterNivelItemCaixa(udtItemCaixa.CodigoItemCaixa) = 5 And _
                Index = enumTipoItemCaixa.Grupo And _
                blnItemCaixaExiste Then
                
                frmMural.Display = "Item de Caixa Elementar não pode ser modificado para Grupo se o mesmo pertencer ao 5º nível "
                frmMural.IconeExibicao = IconInformation
                frmMural.Show vbModal
                optTipoItemCaixa.Item(enumTipoItemCaixa.Grupo).Value = True
            
            End If
        End If
    End If

    Exit Sub
ErrorHandler:

   mdiSBR.uctLogErros.MostrarErros Err, TypeName(Me) & " - optTipoItemCaixa_Click"
        
End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    If Button.Key = "Sair" Then
        fgLockWindow 0
        Unload Me
        Exit Sub
    End If
    
    If treCadastro.SelectedItem Is Nothing Then
        frmMural.Display = "Selecione um Item de Caixa."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        Exit Sub
    End If
         
    fgCursor True
    fgLockWindow Me.hwnd

    Select Case Button.Key
        Case "Limpar"
            If Not xmlItemCaixa Is Nothing And Not treCadastro.SelectedItem.Parent Is Nothing Then
                If Not xmlItemCaixa.documentElement.selectSingleNode("DT_FIM_VIGE").Text = "00:00:00" Then
                    If fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("DT_FIM_VIGE").Text) < fgDataHoraServidor(enumFormatoDataHora.Data) Then
                        frmMural.Display = "Este Item de Grupo não está vigente," & vbCrLf & "não podem ser inseridos Itens de Caixa para este Item de Grupo."
                        frmMural.IconeExibicao = IconInformation
                        frmMural.Show vbModal
                        Exit Sub
                    End If
                End If
            End If
            flLimpar
        Case "Excluir"
            flExcluir
        Case "Salvar"
            flSalvar
    End Select
    
    fgLockWindow 0
    fgCursor

    Exit Sub

ErrorHandler:
    fgLockWindow 0
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixa - tlbCadastro_ButtonClick"
    
    If optTipoCaixa(1).Value = True Then
       optTipoCaixa_Click 1
    Else
       optTipoCaixa_Click 2
    End If
    
    If Button.Key = "Excluir" Then
       flLimpar
    Else
       flPosicionaItemTreeView txtDescricao.Text
    End If
    
End Sub

Private Sub treCadastro_Collapse(ByVal Node As MSComctlLib.Node)
    
On Error GoTo ErrorHandler

    If Node.Parent Is Nothing Then
        'None
    ElseIf Node.Expanded Then
        Node.Image = "Open"
    ElseIf Not Node.Expanded Then
        Node.Image = "Closed"
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixa - treCadastro_NodeClick"

End Sub

Private Sub treCadastro_Expand(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler
    
    If Node.Parent Is Nothing Then
        'None
    ElseIf Node.Expanded Then
        Node.Image = "Open"
    ElseIf Not Node.Expanded Then
        Node.Image = "Closed"
    End If

    Exit Sub

ErrorHandler:
    
    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixa - treCadastro_NodeClick"

End Sub

Private Sub treCadastro_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler
    
    blnItemCaixaExiste = False
    
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
    
    If Node.Parent Is Nothing Then
        xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/@Operacao").Text = "None"
        Exit Sub
    ElseIf Node.Text = gstrItemGenerico Then
        xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/@Operacao").Text = "None"
        frmMural.Display = "Item de Caixa Genérico não pode ser modificado ou excluído."
        frmMural.IconeExibicao = IconInformation
        frmMural.Show vbModal
        
        If udtItemCaixa.TipoCaixa = enumTipoCaixa.CaixaSubReserva Then
            flPosicionaItemTreeView gstrCaixaSubReserva
        Else
            flPosicionaItemTreeView gstrCaixaFuturo
        End If
        
        Exit Sub
        
    End If

    fgCursor True
    flLerItemCaixa udtItemCaixa.CodigoItemCaixa
    fgCursor

    Exit Sub

ErrorHandler:

    mdiSBR.uctLogErros.MostrarErros Err, "frmItemCaixa - treCadastro_NodeClick"

    If Not blnItemCaixaExiste Then
        If optTipoCaixa(1).Value = True Then
           optTipoCaixa_Click 1
        Else
           optTipoCaixa_Click 2
        End If
        flLimpar
    End If
    
End Sub

' Limpa campos da tela para a inclusão de um item de caixa.

Private Sub flLimpar()

On Error GoTo ErrorHandler
    
    If Not treCadastro.SelectedItem Is Nothing Then
        If treCadastro.SelectedItem.Tag <> enumTipoItemCaixa.Grupo Then
            frmMural.Display = "Para inserir um novo Item de Caixa Selecione (Caixa Reserva, Caixa Futuro ou um Item de Grupo)"
            frmMural.IconeExibicao = IconCritical
            frmMural.Show vbModal
            Exit Sub
        End If
    End If
    
    txtDescricao = ""
    optTipoItemCaixa(enumTipoItemCaixa.Grupo).Value = False
    optTipoItemCaixa(enumTipoItemCaixa.Elementar).Value = False
    
    dtpInicio.MinDate = fgDataHoraServidor(Data)
    dtpInicio.Value = dtpInicio.MinDate
    dtpInicio.Enabled = True
    
    dtpFim.MinDate = fgDataHoraServidor(Data)
    dtpFim.Value = dtpFim.MinDate
    dtpFim.Value = Null
    
    'Define o MaxDate e Data Fim Vigência Igual a  ao fim de Vigência do Pai
    If treCadastro.Nodes.Count > 0 Then
        If treCadastro.SelectedItem.Parent Is Nothing Then
            dtpFim.MaxDate = DateSerial(2099, 12, 31)
            dtpFim.Value = Null
        ElseIf xmlItemCaixa Is Nothing Then
            dtpFim.MaxDate = DateSerial(2099, 12, 31)
            dtpFim.Value = Null
        ElseIf Not xmlItemCaixa.documentElement.selectSingleNode("DT_FIM_VIGE").Text = "00:00:00" Then
            dtpFim.MaxDate = fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("DT_FIM_VIGE").Text)
            dtpFim.Value = dtpFim.MaxDate
        Else
            dtpFim.MaxDate = DateSerial(2099, 12, 31)
            dtpFim.Value = Null
        End If
    End If
    
    blnItemCaixaExiste = False
    If Me.Visible Then txtDescricao.SetFocus

Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flLimpar", 0

End Sub

' Define tamanho máximo dos campos da tela, a partir das propriedades da tabela.

Private Sub flDefinirTamanhoMaximoCampos()
On Error GoTo ErrorHandler

    With xmlMapaNavegacao.documentElement
        txtDescricao.MaxLength = .selectSingleNode("Grupo_Propriedades/Grupo_ItemCaixa/DE_ITEM_CAIX/@Tamanho").Text
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flDefinirTamanhoMaximoCampos", 0
End Sub

' Aciona a exclusão de um item de caixa, ou de um grupo de item de caixa.

Private Sub flExcluir()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strMSG              As String
Dim strNomeItemPai      As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    If treCadastro.SelectedItem.Parent Is Nothing Then
        xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/@Operacao").Text = "None"
        frmMural.Display = treCadastro.SelectedItem.Text & " não pode ser modificado ou excluído."
        frmMural.IconeExibicao = IconCritical
        frmMural.Show vbModal
        Exit Sub
    End If
    
    If treCadastro.SelectedItem.Tag = enumTipoItemCaixa.Grupo Then
        strMSG = "Confirma a exclusão deste Item de Caixa?" & vbCrLf & "Todos os Itens de Caixa deste Grupo serão excluídos"
    Else
        strMSG = "Tem certeza que deseja excluir esse Item de Caixa?"
    End If
    
    If MsgBox(strMSG, vbYesNo, Me.Caption) = vbNo Then Exit Sub
    
    strNomeItemPai = treCadastro.SelectedItem.Parent.Text
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    xmlItemCaixa.documentElement.selectSingleNode("//CO_ITEM_CAIX").Text = udtItemCaixa.CodigoItemCaixa
    xmlItemCaixa.documentElement.selectSingleNode("//@Operacao").Text = "Excluir"
    Call objMIU.Executar(xmlItemCaixa.documentElement.xml, vntCodErro, vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    If optTipoCaixa(1).Value = True Then
       optTipoCaixa_Click 1
    Else
       optTipoCaixa_Click 2
    End If

    flPosicionaItemTreeView strNomeItemPai
    flLimpar
        
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

    Exit Sub
ErrorHandler:
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flExcluir", 0
            
End Sub

' Carrega configurações iniciais do formulário.

Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    Set xmlMapaNavegacao = Nothing

    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.SBR, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmItemCaixa", "flInicializar")
    End If
    Set objMIU = Nothing

    Exit Sub
ErrorHandler:

    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Sub

' Aciona a inclusão ou a alteração de um item de caixa.

Private Sub flSalvar()

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler

    If Not flValidarCampos Then Exit Sub
    
    If blnItemCaixaExiste Then
        'Alteração de um Item de Caixa
        
        If treCadastro.SelectedItem.Parent Is Nothing Then
            xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/@Operacao").Text = "None"
            frmMural.Display = treCadastro.SelectedItem.Text & " não pode ser modificado ou excluído."
            frmMural.IconeExibicao = IconCritical
            frmMural.Show vbModal
            Exit Sub
        End If
            
        If treCadastro.SelectedItem.Tag = enumTipoItemCaixa.Grupo And Not IsNull(dtpFim.Value) Then
            If MsgBox("Confirma a finalização da vigência deste Item de Caixa ?" & vbCrLf & _
                      "Todos os Itens de Caixa deste Grupo terão suas vigênicas finalizadas.", _
                      vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
        End If
        
        If Not xmlItemCaixa.loadXML(flInterfaceToXml(xmlItemCaixa, "Alterar")) Then
            fgErroLoadXML xmlItemCaixa, App.EXEName, TypeName(Me), "flSalvar"
        End If
    Else
        If xmlItemCaixa Is Nothing Then
            Set xmlItemCaixa = CreateObject("MSXML2.DOMDocument.4.0")
        End If
        
        If dtpFim.MaxDate <> DateSerial(2099, 12, 31) And IsNull(dtpFim.Value) Then
            frmMural.Display = "Data de Fim de Vigência é obrigatória para este Item de Caixa."
            frmMural.IconeExibicao = IconInformation
            frmMural.Show vbModal
            dtpFim.SetFocus
            Exit Sub
        End If
        
        If Not xmlItemCaixa.loadXML(flInterfaceToXml(xmlMapaNavegacao, "Incluir")) Then
            fgErroLoadXML xmlItemCaixa, App.EXEName, TypeName(Me), "flSalvar"
        End If
    End If

    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    Call objMIU.Executar(xmlItemCaixa.xml, vntCodErro, vntMensagemErro)
    Set objMIU = Nothing
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    fgCursor
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
            
    If optTipoCaixa(1).Value = True Then
       optTipoCaixa_Click 1
    Else
       optTipoCaixa_Click 2
    End If
    
    flPosicionaItemTreeView txtDescricao.Text

    Exit Sub

ErrorHandler:

    Set objMIU = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, Me.Name, "flSalvar", 0
End Sub

' Converte os dados informados na tela, em formato XML, para a atualização da tabela.

Private Function flInterfaceToXml(ByVal xmlEnvio As MSXML2.DOMDocument40, strOperacao As String) As String

On Error GoTo ErrorHandler

    With xmlEnvio.documentElement
    
        .selectSingleNode("//@Operacao").Text = strOperacao
        If strOperacao = "Incluir" Then
                    
            'Código do Item de Caixa será gerado pela Classe
            .selectSingleNode("//CO_ITEM_CAIX").Text = vbNullString
            'Definir Código do Pai
            .selectSingleNode("//CO_ITEM_CAIX_PAI").Text = udtItemCaixa.CodigoItemCaixa
        End If
                
        .selectSingleNode("//TP_BKOF").Text = udtItemCaixa.TipoBackOffice
        .selectSingleNode("//TP_CAIX").Text = udtItemCaixa.TipoCaixa
        .selectSingleNode("//DE_ITEM_CAIX").Text = txtDescricao.Text
    
        If optTipoItemCaixa.Item(enumTipoItemCaixa.Elementar).Value Then
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Elementar
        ElseIf optTipoItemCaixa.Item(enumTipoItemCaixa.Grupo) Then
            .selectSingleNode("//TP_ITEM_CAIX").Text = enumTipoItemCaixa.Grupo
        End If

        .selectSingleNode("//DT_INIC_VIGE").Text = Format(dtpInicio.Value, "YYYYMMDDHHNNSS")

        If IsNull(dtpFim.Value) Then
            .selectSingleNode("//DT_FIM_VIGE").Text = vbNullString
        Else
            .selectSingleNode("//DT_FIM_VIGE").Text = Format(dtpFim.Value, "YYYYMMDDHHNNSS")
        End If
        flInterfaceToXml = .xml
    End With

    Exit Function

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flInterfaceToXml", 0

End Function

' Carrega array com os nodes do treeview.

Private Sub flCarregaExpanded()
Dim lngCont                                 As Long

On Error GoTo ErrorHandler

    ReDim strvExpanded(1 To treCadastro.Nodes.Count, 0 To 1)
    For lngCont = 1 To treCadastro.Nodes.Count
        strvExpanded(lngCont, 0) = treCadastro.Nodes.Item(lngCont).Key
        strvExpanded(lngCont, 1) = treCadastro.Nodes.Item(lngCont).Expanded
    Next

    Exit Sub
ErrorHandler:
    fgRaiseError App.EXEName, Me.Name, "flCarregaExpanded", 0
End Sub

' Expande Nodes do Treeview, caso estes possuam Nodes Child

Private Sub flRetornaExpanded()
Dim objNode                                 As MSComctlLib.Node

    For Each objNode In treCadastro.Nodes
        If objNode.children > 0 Then
            objNode.Expanded = True
        End If
    Next

    Exit Sub
ErrorHandler:

    fgRaiseError App.EXEName, Me.Name, "flRetornaExpanded", 0
    
End Sub

' Valida campos antes da atualização da tabela de itens de caixa.

Private Function flValidarCampos() As Boolean

On Error GoTo ErrorHandler

    If Trim(txtDescricao.Text) = "" Then
        frmMural.Display = "Descrição do Item de Caixa é Obrigatória."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        txtDescricao.SetFocus
        flValidarCampos = False
        Exit Function
    End If

    If optTipoItemCaixa.Item(enumTipoItemCaixa.Grupo) = False And optTipoItemCaixa.Item(enumTipoItemCaixa.Elementar) = False Then
        frmMural.Display = "Tipo do Item de Caixa inválido."
        frmMural.IconeExibicao = IconExclamation
        frmMural.Show vbModal
        flValidarCampos = False
        Exit Function
    End If

    flValidarCampos = True

    Exit Function
ErrorHandler:
    
    flValidarCampos = False
    fgRaiseError App.EXEName, Me.Name, "flValidarCampos", 0

End Function

' Aciona a leitura de um item de caixa, a partir de seu código.

Private Sub flLerItemCaixa(ByVal pstrCodigoItemCaixa As String)

#If EnableSoap = 1 Then
    Dim objMIU          As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU          As A6MIU.clsMIU
#End If

Dim strLer              As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant
    
On Error GoTo ErrorHandler

    blnItemCaixaExiste = False
    
    Set objMIU = fgCriarObjetoMIU("A6MIU.clsMIU")
    Set xmlItemCaixa = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/@Operacao").Text = "Ler"
    xmlMapaNavegacao.selectSingleNode("//Grupo_ItemCaixa/CO_ITEM_CAIX").Text = udtItemCaixa.CodigoItemCaixa
    xmlItemCaixa.loadXML objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("//Grupo_ItemCaixa").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objMIU = Nothing
    
    txtDescricao = xmlItemCaixa.documentElement.selectSingleNode("//DE_ITEM_CAIX").Text
    optTipoItemCaixa.Item(xmlItemCaixa.documentElement.selectSingleNode("//TP_ITEM_CAIX").Text) = True
    dtpInicio.MinDate = fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("//DT_INIC_VIGE").Text)
    dtpInicio.Value = fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("//DT_INIC_VIGE").Text)
    
    If dtpInicio.Value > fgDataHoraServidor(Data) Then
        dtpInicio.MinDate = fgDataHoraServidor(Data)
        dtpInicio.Enabled = True
    Else
        dtpInicio.Enabled = False
    End If

    dtpFim.MaxDate = DateSerial(2099, 12, 31)

    If xmlItemCaixa.documentElement.selectSingleNode("//DT_FIM_VIGE").Text = CStr(datDataVazia) Then
        dtpFim.MinDate = fgMaiorData(dtpInicio.Value, fgDataHoraServidor(Data))
        dtpFim.Value = fgMaiorData(dtpInicio.Value, fgDataHoraServidor(Data))
        dtpFim.Value = Null
    Else
        If xmlItemCaixa.documentElement.selectSingleNode("//DT_FIM_VIGE").Text <> gstrDataVazia Then
            If fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("//DT_FIM_VIGE").Text) < fgDataHoraServidor(Data) Then
                dtpFim.MinDate = fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("//DT_FIM_VIGE").Text)
                dtpInicio.Enabled = True
            Else
                dtpFim.MinDate = fgMaiorData(dtpInicio.Value, fgDataHoraServidor(Data))
            End If
            dtpFim.Value = fgDtXML_To_Date(xmlItemCaixa.documentElement.selectSingleNode("//DT_FIM_VIGE").Text)
        End If
    End If

    xmlItemCaixa.selectSingleNode("//@Operacao").Text = "Alterar"
    
    blnItemCaixaExiste = True

    Exit Sub
ErrorHandler:

    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flLerItemCaixa", 0
 
End Sub

' Posiciona um item de caixa do treeview, a partir de sua descrição.

Private Sub flPosicionaItemTreeView(ByVal pstrDescricao As String)

Dim objNode                                 As MSComctlLib.Node

On Error GoTo ErrorHandler

    For Each objNode In treCadastro.Nodes
        If objNode.Text = pstrDescricao Then
            objNode.Selected = True
            objNode.EnsureVisible
            treCadastro_NodeClick objNode
            Exit Sub
        End If
    Next

    Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flPosicionaItemTreeView", 0
    
End Sub
