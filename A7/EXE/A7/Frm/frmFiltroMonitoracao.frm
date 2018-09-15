VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFiltroMonitoracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmFiltroMonitoracao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   30
      TabIndex        =   6
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtIDMensagem 
         Height          =   315
         Left            =   1740
         MaxLength       =   30
         TabIndex        =   15
         Top             =   1995
         Width           =   4095
      End
      Begin VB.ComboBox cboSistema 
         Height          =   315
         ItemData        =   "frmFiltroMonitoracao.frx":000C
         Left            =   1740
         List            =   "frmFiltroMonitoracao.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1260
         Width           =   4095
      End
      Begin VB.ComboBox cboEmpresa 
         Height          =   315
         ItemData        =   "frmFiltroMonitoracao.frx":0010
         Left            =   1740
         List            =   "frmFiltroMonitoracao.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   4095
      End
      Begin VB.ComboBox cboTipoMensagem 
         Height          =   315
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1620
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpDataMovimento 
         Height          =   315
         Left            =   1740
         TabIndex        =   0
         Top             =   180
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60162049
         CurrentDate     =   36892
         MinDate         =   36892
      End
      Begin MSComCtl2.DTPicker dtpHoraDe 
         Height          =   315
         Left            =   2940
         TabIndex        =   1
         Top             =   540
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60162050
         CurrentDate     =   37322
      End
      Begin MSComCtl2.DTPicker dtpHoraAte 
         Height          =   315
         Left            =   4410
         TabIndex        =   2
         Top             =   540
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60162050
         CurrentDate     =   37322
      End
      Begin MSComctlLib.Toolbar tlbHorario 
         Height          =   330
         Left            =   1740
         TabIndex        =   7
         ToolTipText     =   "Clique para Habilitar/Desabilitar"
         Top             =   525
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         ButtonWidth     =   1349
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imlIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Após"
               Key             =   "Comparacao"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Antes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Entre"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ID Mensagem:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         Caption         =   "Data Recebimento:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblEvento 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Mensagem:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   1185
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   660
      End
      Begin VB.Label lblHorario 
         AutoSize        =   -1  'True
         Caption         =   "Horário:"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lblSistema 
         AutoSize        =   -1  'True
         Caption         =   "Sistema:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   600
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   3705
      TabIndex        =   13
      Top             =   2565
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   582
      ButtonWidth     =   1826
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "OK"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            Key             =   "Cancelar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   15
      Top             =   2415
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0014
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0126
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0440
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0792
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":08A4
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0BBE
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":0ED8
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":11F2
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":150C
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":185E
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":1BB0
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFiltroMonitoracao.frx":1F02
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFiltroMonitoracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pela configuração do filtro da tela de monitoração de mensagens.
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmMonitoracao"

Public blnFiltraHora                        As Boolean

'Aplicar filtro
Private Sub flAplicarFiltro()
    
On Error GoTo ErrorHanlder
    
    Me.Hide
    DoEvents
    
    frmMonitoracao.tlbButtons.Buttons("AplicarFiltro").Value = tbrPressed
    
    Call frmMonitoracao.fgAtualizar
    
    'If frmMonitoracao.lstMonitoracao.ListItems.Count = 0 Then
    '    MsgBox "Não existem mensagens para o filtro definido.", vbInformation, "Monitoração de Mensagens"
    'End If

    Exit Sub
ErrorHanlder:

    mdiBUS.uctLogErros.MostrarErros Err, "frmMonitoracao - flAplicarFiltro"
                                           
End Sub

'Carregar combo de sistemas de origem.
Private Sub flCarregarSistemasOrigem()

#If EnableSoap = 1 Then
    Dim objMonitoracao  As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao  As A7Miu.clsMonitoracao
#End If

Dim objDOMNode          As MSXML2.IXMLDOMNode
Dim xmlSistema          As MSXML2.DOMDocument40
Dim strSistemas         As String
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    Set xmlSistema = CreateObject("MSXML2.DOMDocument.4.0")
        
    DoEvents
    
    Call fgCursor(True)

    strSistemas = objMonitoracao.ObterSistemasOrigem(Val(fgObterCodigoCombo(cboEmpresa)), vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    cboSistema.Clear
    cboSistema.AddItem "<-- Todos -->"
    
    If strSistemas <> "" Then
        If Not xmlSistema.loadXML(strSistemas) Then
            Call fgErroLoadXML(xmlSistema, App.EXEName, "frmFiltroMonitoracao", "flCarregarSistemasOrigem")
        End If
        
        For Each objDOMNode In xmlSistema.documentElement.selectNodes("//Repeat_Sistema/*")
            cboSistema.AddItem objDOMNode.selectSingleNode("SG_SIST").Text & " - " & _
                               objDOMNode.selectSingleNode("NO_SIST").Text
        Next
    End If
    
    cboSistema.ListIndex = 0
    
    If cboSistema.ListCount = 1 Then
        cboSistema.Enabled = False
    Else
        cboSistema.Enabled = True
    End If
    
    Set objMonitoracao = Nothing
    Set xmlSistema = Nothing
    Set objDOMNode = Nothing
    
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:

    Set objMonitoracao = Nothing
    Set xmlSistema = Nothing
    Set objDOMNode = Nothing

    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarSistemasOrigem", 0
    
End Sub

'Obter as propriedades necessárias para o formulário através de interação com a camada controladora de caso de uso MIU.
Private Sub flInit()

#If EnableSoap = 1 Then
    Dim objMiu          As MSSOAPLib30.SoapClient30
#Else
    Dim objMiu          As A7Miu.clsMIU
#End If

Dim objDOMNode          As MSXML2.IXMLDOMNode
Dim vntCodErro          As Variant
Dim vntMensagemErro     As Variant

On Error GoTo ErrorHandler
    
    Set xmlMapaNavegacao = Nothing
    
    Set objMiu = fgCriarObjetoMIU("A7Miu.clsMIU")
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
        
    If Not xmlMapaNavegacao.loadXML(objMiu.ObterMapaNavegacao(enumSistemaSLCC.BUS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmFiltroMonitoracao", "flInit")
    End If
    
    cboEmpresa.AddItem "<-- Todos -->"
    
    For Each objDOMNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_Empresa/*")
        cboEmpresa.AddItem objDOMNode.selectSingleNode("CO_EMPR").Text & " - " & _
                           objDOMNode.selectSingleNode("NO_REDU_EMPR").Text
    Next
    
    cboEmpresa.ListIndex = 0
    
    cboSistema.AddItem "<-- Todos -->"
    cboSistema.ListIndex = 0
    cboSistema.Enabled = False
    
    cboTipoMensagem.AddItem "<-- Todos -->"
    
    For Each objDOMNode In xmlMapaNavegacao.documentElement.selectNodes("//Repeat_TipoMensagem/*")
        cboTipoMensagem.AddItem objDOMNode.selectSingleNode("TP_MESG").Text & " - " & _
                                objDOMNode.selectSingleNode("NO_TIPO_MESG").Text
        
        If IsNumeric(objDOMNode.selectSingleNode("TP_MESG").Text) Then
            cboTipoMensagem.ItemData(cboTipoMensagem.NewIndex) = objDOMNode.selectSingleNode("TP_MESG").Text
        End If
    
    Next
    
    cboTipoMensagem.ListIndex = 0
    
    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    Set objDOMNode = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objMiu = Nothing
    Set xmlMapaNavegacao = Nothing
    Set objDOMNode = Nothing

    Call fgCursor(False)
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarSistemasOrigem", 0
    
End Sub

Private Sub cboEmpresa_Click()
    
On Error GoTo ErrorHandler

    fgCursor True
    
    If cboEmpresa.ListIndex = 0 Then ' Todos
        cboSistema.Clear
        cboSistema.AddItem "<-- Todos -->"
        cboSistema.ListIndex = 0
        cboSistema.Enabled = False
    Else
        Call flCarregarSistemasOrigem
    End If
    
    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - cboEmpresa_Click"
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    Me.Icon = mdiBUS.Icon
    
    Call fgCenterMe(Me)
        
    DoEvents
    
    Call fgCursor(True)
    
    Call flInit

    Me.dtpDataMovimento.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    Me.dtpHoraDe.Value = CDate("00:00:00")
    Me.dtpHoraAte.Value = Time
        
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmFiltroMonitoracao - Form_Load"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case vbKeyEscape
            Me.Hide
        Case vbKeyReturn
            tlbCadastro_ButtonClick tlbCadastro.Buttons("OK")
    End Select

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "OK"
            Call flAplicarFiltro
        Case "Cancelar"
            Me.Hide
    End Select
    
    fgCursor False
    
    Exit Sub
ErrorHandler:

    fgCursor False

    mdiBUS.uctLogErros.MostrarErros Err, "frmFiltroMonitoracao - tlbCadastro_ButtonClick"

End Sub

Private Sub tlbHorario_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    If Button.Image = 5 Then
        Button.Image = 0
        blnFiltraHora = False
    Else
        blnFiltraHora = True
        Button.Image = 5
    End If
    
End Sub

Private Sub tlbHorario_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    Select Case ButtonMenu.Text
        Case "Após"
            dtpHoraAte.Visible = False
            ButtonMenu.Text = tlbHorario.Buttons(1).Caption
            tlbHorario.Buttons(1).Caption = "Após"
        Case "Antes"
            dtpHoraAte.Visible = False
            ButtonMenu.Text = tlbHorario.Buttons(1).Caption
            tlbHorario.Buttons(1).Caption = "Antes"
        Case "Entre"
            dtpHoraAte.Visible = True
            ButtonMenu.Text = tlbHorario.Buttons(1).Caption
            tlbHorario.Buttons(1).Caption = "Entre"
    End Select

End Sub
