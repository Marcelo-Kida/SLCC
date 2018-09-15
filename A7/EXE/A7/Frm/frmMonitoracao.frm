VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMonitoracao 
   Caption         =   "Monitoração de Mensagens"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   9150
      Top             =   6060
   End
   Begin MSComCtl2.UpDown udTimer 
      Height          =   315
      Left            =   8385
      TabIndex        =   1
      Top             =   6120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtTimer"
      BuddyDispid     =   196610
      OrigLeft        =   8700
      OrigTop         =   6120
      OrigRight       =   8940
      OrigBottom      =   6435
      Max             =   60
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtTimer 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7980
      TabIndex        =   0
      Text            =   "10"
      Top             =   6120
      Width           =   405
   End
   Begin VB.Frame Barra 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000C&
      Height          =   360
      Left            =   35
      TabIndex        =   4
      Top             =   0
      Width           =   12015
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Monitoração de Mensagens"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Left            =   60
         TabIndex        =   6
         Top             =   30
         Width           =   3330
      End
      Begin VB.Label lblData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   300
         Left            =   10980
         TabIndex        =   5
         Top             =   30
         Width           =   90
      End
   End
   Begin MSComctlLib.ListView lstMonitoracao 
      Height          =   5685
      Left            =   3560
      TabIndex        =   2
      Top             =   390
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   10028
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgIcons"
      SmallIcons      =   "imgIcons"
      ColHdrIcons     =   "imgIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Alerta"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Seq."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Identificador da Mensagem"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo de Mensagem"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Data do Recebimento"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Sistema Origem"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAtualizar 
      Height          =   330
      Left            =   2805
      TabIndex        =   3
      Top             =   6075
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Atualizar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   6075
      Width           =   2850
      _ExtentX        =   5027
      _ExtentY        =   582
      ButtonWidth     =   2461
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aplicar Filtro"
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Definir Filtro"
            Key             =   "DefinirFiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView treStatus 
      Height          =   5685
      Left            =   30
      TabIndex        =   8
      Top             =   390
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   10028
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "imgIcons"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   10620
      Top             =   5850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":0000
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":0112
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":0464
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":0576
            Key             =   "st_aguardando"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":09C8
            Key             =   "st_postado"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":0E1A
            Key             =   "st_traduzido"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":126C
            Key             =   "st_retirado"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":16BE
            Key             =   "st_recebido"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":1B10
            Key             =   "st_entregue"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":1F62
            Key             =   "st_root_old"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":23B4
            Key             =   "st_root"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":2806
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":2B58
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":2EAA
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":31FC
            Key             =   "vermelho"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":354E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMonitoracao.frx":39A0
            Key             =   "st_cancelado"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo para Refresh automático da tela (em minutos) :"
      Height          =   195
      Left            =   3990
      TabIndex        =   9
      Top             =   6180
      Width           =   3945
   End
   Begin VB.Image imgDummy 
      Height          =   5925
      Left            =   3470
      MousePointer    =   9  'Size W E
      Top             =   435
      Width           =   105
   End
End
Attribute VB_Name = "frmMonitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela exibição de mensagens transitadas no sistema A7.
Option Explicit

Private blnDummy                            As Boolean
Private intContMinutos                      As Integer
Private blnBaseHistorica                    As Boolean

Private blnTimerBypass                      As Boolean

'Carregar listview de mensagens de acordo com o filtro informado.
Private Sub flCarregarListView()

#If EnableSoap = 1 Then
    Dim objMonitoracao  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao  As A7Miu.clsMonitoracao
#End If

Dim xmlMNode            As MSXML2.IXMLDOMNode
Dim xmlMonitoracao      As MSXML2.DOMDocument40
Dim strMonitoracao      As String
Dim strCodOcorrencia    As String
Dim strCorAlerta        As String
Dim objListItem         As ListItem
Dim strDtHrDe           As String
Dim strDtHrAte          As String
Dim strEmpresa          As String
Dim strSistema          As String
Dim strTipoMensagem     As String
Dim strIDMensagem       As String
Dim strQueryXpath       As String
Dim strDescOcorrencia   As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Me.lstMonitoracao.ListItems.Clear
    lstMonitoracao.HideSelection = False
    strCodOcorrencia = ""
    
    If Me.treStatus.Nodes("st" & enumOcorrencia.RecebimentoBemSucedido).Checked Then
        strCodOcorrencia = strCodOcorrencia & ", " & enumOcorrencia.RecebimentoBemSucedido
    End If
    
    If Me.treStatus.Nodes("st" & enumOcorrencia.PostagemBemSucedida).Checked Then
        strCodOcorrencia = strCodOcorrencia & ", " & enumOcorrencia.PostagemBemSucedida
    End If
    
    If Me.treStatus.Nodes("st" & enumOcorrencia.ConfirmacaoEntregaRecebida).Checked Then
        strCodOcorrencia = strCodOcorrencia & ", " & enumOcorrencia.ConfirmacaoEntregaRecebida
    End If
    
    If Me.treStatus.Nodes("st" & enumOcorrencia.ConfirmacaoRetiradaRecebida).Checked Then
        strCodOcorrencia = strCodOcorrencia & ", " & enumOcorrencia.ConfirmacaoRetiradaRecebida
    End If
    
    If Me.treStatus.Nodes("st" & enumOcorrencia.CanceladaErroTradução).Checked Then
        strCodOcorrencia = strCodOcorrencia & ", " & enumOcorrencia.CanceladaErroTradução
    End If
            
    If strCodOcorrencia = "" Then
        'MsgBox "Selecione pelo menos um Status da Mensagem a ser monitorado.", vbInformation, Me.Caption
        Exit Sub
    End If
    
    strCodOcorrencia = Mid$(strCodOcorrencia, 3)
    
    With frmFiltroMonitoracao
        If Me.tlbButtons.Buttons("AplicarFiltro").Value = tbrPressed Then
            If .blnFiltraHora Then
                If .tlbHorario.Buttons(1).Caption = "Entre" Then
                    strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                    strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraAte.Value, "HHmmss")
                Else
                    Select Case .tlbHorario.Buttons(1).Caption
                        Case "Antes"
                            strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "000000"
                            strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                        Case "Após"
                            strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                            strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "235959"
                    End Select
                End If
            Else
                strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "000000"
                strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "235959"
            End If
                
            Me.lblData.Caption = Format(.dtpDataMovimento.Value, gstrMascaraDataDtp)
            
            If .cboEmpresa.ListIndex > 0 Then
                strEmpresa = fgObterCodigoCombo(.cboEmpresa)
                Me.lblEmpresa.Caption = Mid$(.cboEmpresa.Text, 7)
            Else
                Me.lblEmpresa.Caption = "Mensagens Transitadas"
            End If
            
            If .cboSistema.ListIndex > 0 Then
                strSistema = fgObterCodigoCombo(.cboSistema)
            End If
            
            If .cboTipoMensagem.ListIndex > 0 Then
                strTipoMensagem = fgObterCodigoCombo(.cboTipoMensagem)
            End If
            
            strIDMensagem = Trim(.txtIDMensagem.Text)
        Else
            strDtHrDe = fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000"
            strDtHrAte = fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959"
            
            Me.lblData.Caption = fgDataHoraServidor(DataAux)
            strEmpresa = ""
            strSistema = ""
            strIDMensagem = ""
            strTipoMensagem = ""
        
        End If
        
    End With
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    Set xmlMonitoracao = CreateObject("MSXML2.DOMDocument.4.0")
            
    Call fgCursor(True)
    
    vntCodErro = 0
    
    strMonitoracao = objMonitoracao.LerTodos(strDtHrDe, _
                                             strDtHrAte, _
                                             Val(strEmpresa), _
                                             strSistema, _
                                             strTipoMensagem, _
                                             strCodOcorrencia, _
                                             strIDMensagem, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    lstMonitoracao.Sorted = True
    lstMonitoracao.SortKey = 5
    lstMonitoracao.SortOrder = lvwDescending
        
    If strMonitoracao <> "" Then
        If Not xmlMonitoracao.loadXML(strMonitoracao) Then
            Call fgErroLoadXML(xmlMonitoracao, App.EXEName, "frmMonitoracao", "flCarregarListView")
        End If
        
        'todo here
        If Not xmlMonitoracao.selectSingleNode("//OWNER") Is Nothing Then
            blnBaseHistorica = (xmlMonitoracao.selectSingleNode("//OWNER").Text = "A7HIST")
        End If
        
        For Each xmlMNode In xmlMonitoracao.documentElement.selectNodes("//Repeat_Monitoracao/*")
            
            strQueryXpath = xmlMNode.selectSingleNode("CO_OCOR_MESG").Text
            strQueryXpath = "Grupo_OcorrenciaMensagem[CO_OCOR_MESG='" & strQueryXpath & "']/DE_ABRV_OCOR_MESG"
            
            strDescOcorrencia = gxmlOcorrencia.documentElement.selectSingleNode(strQueryXpath).Text
            
            strQueryXpath = xmlMNode.selectSingleNode("CO_OCOR_MESG").Text
            strQueryXpath = "Grupo_OcorrenciaMensagem[CO_OCOR_MESG='" & strQueryXpath & "']/TP_GRAU_SEVE"
            strQueryXpath = gxmlOcorrencia.documentElement.selectSingleNode(strQueryXpath).Text
            
            Select Case strQueryXpath
                Case "1" ' Verde
                    strCorAlerta = "verde"
                Case "2" ' Amarelo
                    strCorAlerta = "amarelo"
                Case "3" ' Laranja
                    strCorAlerta = "laranja"
                Case "4" ' Vermelho
                    strCorAlerta = "vermelho"
            End Select
            
            Set objListItem = Me.lstMonitoracao.ListItems.Add(, , , strCorAlerta, strCorAlerta)
            
            objListItem.Tag = xmlMNode.selectSingleNode("CO_MESG").Text
            
            objListItem.SubItems(1) = xmlMNode.selectSingleNode("CO_MESG").Text
            objListItem.SubItems(2) = xmlMNode.selectSingleNode("CO_CMPO_ATRB_IDEF_MESG").Text
            
            strQueryXpath = "TP_MESG='" & xmlMNode.selectSingleNode("TP_MESG").Text & "' and " & _
                            "TP_FORM_MESG_SAID='" & xmlMNode.selectSingleNode("TP_FORM_MESG_SAID").Text & "'"
            strQueryXpath = "Grupo_TipoMensagem[" & strQueryXpath & "]/NO_TIPO_MESG"
            strQueryXpath = gxmlTipoMensagem.documentElement.selectSingleNode(strQueryXpath).Text
            
            objListItem.SubItems(3) = xmlMNode.selectSingleNode("TP_MESG").Text & " - " & strQueryXpath
            
            objListItem.SubItems(4) = strDescOcorrencia
            
            objListItem.SubItems(5) = Format(fgDtHrStr_To_DateTime(xmlMNode.selectSingleNode("DH_MESG").Text), gstrMascaraDataHoraDtp)
            
            strQueryXpath = fgCompletaString(xmlMNode.selectSingleNode("SG_SIST_ORIG").Text, " ", 3, False)
            strQueryXpath = "Grupo_Sistema[SG_SIST='" & strQueryXpath & "']/NO_SIST"
            strQueryXpath = gxmlSistema.documentElement.selectSingleNode(strQueryXpath).Text
            
            objListItem.SubItems(6) = strQueryXpath
        
        Next
    End If
    
    Set objMonitoracao = Nothing
    Set xmlMonitoracao = Nothing
    Set xmlMNode = Nothing
   
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objMonitoracao = Nothing
    Set xmlMonitoracao = Nothing
    Set xmlMNode = Nothing

    Call fgCursor(False)
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmMonitoracao - flCarregarListView")
    
End Sub

'Carregar TreeView com todos os Status de Eventos
Private Sub flCarregarTreeView()

Dim objNode                         As Node

    With Me.treStatus
        .Nodes.Clear
        
        Set objNode = .Nodes.Add(, , "root", "Mensagens Transitadas", "st_root")
        objNode.Tag = "Mensagens Transitadas"
        
        Set objNode = .Nodes.Add("root", tvwChild, "st" & enumOcorrencia.RecebimentoBemSucedido, "Recebido", "st_recebido")
        objNode.Tag = "Recebido"
        
        objNode.Checked = True
        
        Set objNode = .Nodes.Add("root", tvwChild, "st" & enumOcorrencia.PostagemBemSucedida, "Postado", "st_postado")
        objNode.Tag = "Postado"
        Set objNode = .Nodes.Add("root", tvwChild, "st" & enumOcorrencia.ConfirmacaoEntregaRecebida, "Entregue", "st_entregue")
        objNode.Tag = "Entregue"
        Set objNode = .Nodes.Add("root", tvwChild, "st" & enumOcorrencia.ConfirmacaoRetiradaRecebida, "Retirado", "st_retirado")
        objNode.Tag = "Retirado"
        Set objNode = .Nodes.Add("root", tvwChild, "st" & enumOcorrencia.CanceladaErroTradução, "Cancelado", "st_cancelado")
        objNode.Tag = "Cancelado"
        
        .Nodes.Item("st" & enumOcorrencia.RecebimentoBemSucedido).EnsureVisible
    End With
    
    'Call treStatus_NodeCheck(treStatus.Nodes("root"))
    
End Sub

'Atualizar quantidade de mensagens por status no treeview de status.
Private Sub flAtualizarQtd()

#If EnableSoap = 1 Then
    Dim objMonitoracao  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao  As A7Miu.clsMonitoracao
#End If

Dim xmlMNode            As MSXML2.IXMLDOMNode
Dim xmlMonitoracao      As MSXML2.DOMDocument40
Dim strMonitoracao      As String
'Dim strCodOcorrencia   As String
Dim strCorAlerta        As String
Dim objListItem         As ListItem
Dim strDtHrDe           As String
Dim strDtHrAte          As String
Dim strEmpresa          As String
Dim strSistema          As String
Dim strTipoMensagem     As String
Dim strIDMensagem       As String
Dim objTreeNode         As MSComctlLib.Node
Dim llQtdTotal          As Long
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Me.lstMonitoracao.ListItems.Clear
    lstMonitoracao.HideSelection = False
    
    With frmFiltroMonitoracao
                        
        If Me.tlbButtons.Buttons("AplicarFiltro").Value = tbrPressed Then
            
            If .blnFiltraHora Then
                If .tlbHorario.Buttons(1).Caption = "Entre" Then
                    strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                    strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraAte.Value, "HHmmss")
                Else
                    Select Case .tlbHorario.Buttons(1).Caption
                        Case "Antes"
                            strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "000000"
                            strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                        Case "Após"
                            strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & Format(.dtpHoraDe.Value, "HHmmss")
                            strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "235959"
                    End Select
                End If
            Else
                strDtHrDe = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "000000"
                strDtHrAte = Format(.dtpDataMovimento.Value, gstrMascaraDataXml) & "235959"
            End If
                
            Me.lblData.Caption = Format(.dtpDataMovimento.Value, gstrMascaraDataDtp)
            
            If .cboEmpresa.ListIndex > 0 Then
                strEmpresa = fgObterCodigoCombo(.cboEmpresa)
                Me.lblEmpresa.Caption = Mid$(.cboEmpresa.Text, 7)
            Else
                Me.lblEmpresa.Caption = "Mensagens Transitadas"
            End If
            
            If .cboSistema.ListIndex > 0 Then
                strSistema = fgObterCodigoCombo(.cboSistema)
            End If
            
            If .cboTipoMensagem.ListIndex > 0 Then
                strTipoMensagem = fgObterCodigoCombo(.cboTipoMensagem)
            End If
            
            strIDMensagem = Trim(.txtIDMensagem.Text)
        Else
        
            strDtHrDe = fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "000000"
            strDtHrAte = fgDt_To_Xml(fgDataHoraServidor(DataAux)) & "235959"
            
            Me.lblData.Caption = fgDataHoraServidor(DataAux)
            strEmpresa = ""
            strSistema = ""
            strIDMensagem = ""
            strTipoMensagem = ""
        
        End If
    End With
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    Set xmlMonitoracao = CreateObject("MSXML2.DOMDocument.4.0")
            
    Call fgCursor(True)

    strMonitoracao = objMonitoracao.ObterQtd(strDtHrDe, _
                                             strDtHrAte, _
                                             Val(strEmpresa), _
                                             strSistema, _
                                             strTipoMensagem, _
                                             strIDMensagem, _
                                             vntCodErro, _
                                             vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    'Inicia as quantidades
    For Each objTreeNode In treStatus.Nodes
        objTreeNode.Text = objTreeNode.Tag
    Next
    
    If strMonitoracao <> "" Then
        If Not xmlMonitoracao.loadXML(strMonitoracao) Then
            Call fgErroLoadXML(xmlMonitoracao, App.EXEName, "frmMonitoracao", "flCarregarListView")
        End If
        
        llQtdTotal = 0
        
        For Each xmlMNode In xmlMonitoracao.documentElement.selectNodes("//Repeat_Monitoracao/*")
            
            Set objTreeNode = treStatus.Nodes("st" & xmlMNode.selectSingleNode("CO_OCOR_MESG").Text)

            objTreeNode.Text = objTreeNode.Tag & " (" & xmlMNode.selectSingleNode("QTD").Text & ")"
            
            llQtdTotal = llQtdTotal + CLng(xmlMNode.selectSingleNode("QTD").Text)
        Next
        
        If llQtdTotal > 0 Then
            Me.treStatus.Nodes("root").Text = Me.treStatus.Nodes("root").Tag & " (" & CStr(llQtdTotal) & ")"
        End If
    
    End If
    
    Set objTreeNode = Nothing
    
    Set objMonitoracao = Nothing
    Set xmlMonitoracao = Nothing
    Set xmlMNode = Nothing
    
    Exit Sub
ErrorHandler:

    Set objTreeNode = Nothing
    Set objMonitoracao = Nothing
    Set xmlMonitoracao = Nothing
    Set xmlMNode = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, Me.Name, "flAtualizarQtd", 0
End Sub

'Atualizar a quantidade de mensagens e as mensagens exibidas.
Public Sub fgAtualizar()

On Error GoTo ErrorHandler
    
    blnTimerBypass = True
    fgLockWindow Me.hwnd
    
    flAtualizarQtd
    flCarregarListView
    
    fgLockWindow
    blnTimerBypass = False
    
    Exit Sub
ErrorHandler:
    fgLockWindow
    fgRaiseError App.EXEName, Me.Name, "flAtualizar", 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        
        fgCursor True
        
        fgAtualizar
        
        fgCursor
    End If
    
    Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, ("frmMonitoracao - Form_KeyDown")

End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler

    fgCursor True

    Call fgCenterMe(Me)
    Me.Icon = mdiBUS.Icon
    Me.Show
    DoEvents
    
    Call flCarregarTreeView
    
    frmFiltroMonitoracao.Show vbModal
    
    treStatus.SetFocus
    
    fgCursor

Exit Sub
ErrorHandler:

    fgCursor
   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    Barra.Top = 0
    Barra.Left = 0
    Barra.Width = Me.ScaleWidth
    
    treStatus.Top = Barra.Height
    treStatus.Left = 0
    treStatus.Height = Me.ScaleHeight - Barra.Height - 90 - tlbButtons.Height
    
    tlbButtons.Top = Me.ScaleHeight - tlbButtons.Height - 60
    tlbButtons.Left = 0
    tlbAtualizar.Top = tlbButtons.Top
    
    imgDummy.Top = treStatus.Top
    imgDummy.Left = treStatus.Width
    imgDummy.Height = treStatus.Height
    
    lstMonitoracao.Top = treStatus.Top
    lstMonitoracao.Left = imgDummy.Left + imgDummy.Width
    lstMonitoracao.Width = Me.ScaleWidth - treStatus.Width - 90
    lstMonitoracao.Height = treStatus.Height
    
    txtTimer.Top = lstMonitoracao.Top + lstMonitoracao.Height + 30
    udTimer.Top = txtTimer.Top
    lblTimer.Top = txtTimer.Top + 60
    
    lblData.Left = Barra.Width - lblData.Width - 90

End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub imgDummy_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnDummy = True
    
End Sub

Private Sub imgDummy_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Not blnDummy Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummy.Left = x + imgDummy.Left

    On Error Resume Next
    
    With Me
        'set the width
        If .imgDummy.Left < 1926 Then
            .imgDummy.Left = 1926
        End If
        If .imgDummy.Left > (.Width - 500) And (.Width - 500) > 0 Then
            .imgDummy.Left = .Width - 500
        End If
        
        .treStatus.Width = .imgDummy.Left
        .lstMonitoracao.Left = .imgDummy.Left + .imgDummy.Width
        .lstMonitoracao.Width = .Width - (.imgDummy.Left + 180)
    End With
    
    On Error GoTo 0
End Sub

Private Sub imgDummy_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    blnDummy = False
    
End Sub

'Ordenar as colunas do listview de acordo com a coluna selecionada.
Private Sub lstMonitoracao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
    Call fgClassificarListview(Me.lstMonitoracao, ColumnHeader.Index)
    
End Sub

Private Sub lstMonitoracao_DblClick()

On Error GoTo ErrorHandler

    If lstMonitoracao.SelectedItem Is Nothing Then Exit Sub
    
    Unload frmMonitoracaoDetalhe
    
    With frmMonitoracaoDetalhe
        If blnBaseHistorica Then
            .lngCodigoMensagem = -1 * CLng(Me.lstMonitoracao.SelectedItem.Tag)
        Else
            .lngCodigoMensagem = CLng(Me.lstMonitoracao.SelectedItem.Tag)
        End If
        .Show
        .ZOrder
    End With

Exit Sub
ErrorHandler:

   mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - lstMonitoracao_DblClick"

End Sub

Private Sub lstMonitoracao_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Form_KeyDown KeyCode, 0
    
End Sub

Private Sub tlbAtualizar_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True

    Call fgAtualizar

    fgCursor

Exit Sub
ErrorHandler:
    
    fgCursor
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - tlbAtualizar_ButtonClick"

End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Key = "DefinirFiltro" Then
        frmFiltroMonitoracao.Show vbModal
    End If

End Sub

Private Sub tmrRefresh_Timer()

On Error GoTo ErrorHandler

    If blnTimerBypass Then Exit Sub
    
    If Not IsNumeric(txtTimer.Text) Then Exit Sub
    
    If CLng(txtTimer.Text) = 0 Then Exit Sub
    
    If fgVerificaJanelaVerificacao() Then Exit Sub
    
    fgCursor True

    intContMinutos = intContMinutos + 1
    
    If intContMinutos >= txtTimer.Text Then
        Call fgAtualizar
        intContMinutos = 0
    End If

    fgCursor False

    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - tmrRefresh_Timer"

End Sub

Private Sub treStatus_Collapse(ByVal Node As MSComctlLib.Node)
    
    If Node.Parent Is Nothing Then
        Node.Expanded = True
    End If

End Sub

Private Sub treStatus_KeyDown(KeyCode As Integer, Shift As Integer)
                   
    Form_KeyDown KeyCode, 0

End Sub

Private Sub treStatus_NodeCheck(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler

Dim objNodeAux                              As Node
Dim blnCheckParent                          As Boolean

    If Node.Parent Is Nothing Then
        For Each objNodeAux In treStatus.Nodes
            objNodeAux.Checked = Node.Checked
        Next
    ElseIf Not Node.Checked Then
        treStatus.Nodes(1).Checked = False
    Else
        'Verifica se todos os nós estão selecionados e seleciona o pai
        blnCheckParent = True
        For Each objNodeAux In Me.treStatus.Nodes
            If Not objNodeAux.Parent Is Nothing Then
                If Not objNodeAux.Checked Then
                    blnCheckParent = False
                    Exit For
                End If
            End If
        Next
            
        Me.treStatus.Nodes("root").Checked = blnCheckParent
        
    End If
    
    fgLockWindow Me.hwnd
    
    Call fgAtualizar
    
    fgLockWindow
    
    Exit Sub
ErrorHandler:
    
    fgLockWindow
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - treStatus_NodeCheck"

End Sub

Private Sub treStatus_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler

Dim objNodeAux                               As Node

    Node.Checked = True
    
    If Node.Key <> "root" Then
        For Each objNodeAux In treStatus.Nodes
            If objNodeAux.Key <> Node.Key Then
                objNodeAux.Checked = False
            End If
        Next
    Else
        For Each objNodeAux In treStatus.Nodes
            objNodeAux.Checked = Node.Checked
        Next
    End If

    fgLockWindow Me.hwnd
    
    Call fgAtualizar
    
    fgLockWindow
    
    Exit Sub
ErrorHandler:
    
    fgLockWindow
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - treStatus_NodeClick"
    
End Sub

