VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConsultaOperacao 
   Caption         =   "Consulta - Operações"
   ClientHeight    =   9765
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   14850
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrRefresh 
      Interval        =   60000
      Left            =   5130
      Top             =   8940
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
      Left            =   4230
      TabIndex        =   6
      Text            =   "10"
      Top             =   9000
      Width           =   390
   End
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   9405
      Width           =   14850
      _ExtentX        =   26194
      _ExtentY        =   635
      ButtonWidth     =   2805
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgToolBar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar Filtro"
            Key             =   "AplicarFiltro"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   1
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Definir Filtro"
            Key             =   "showfiltro"
            Object.ToolTipText     =   "Definir Filtro"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh Tela"
            Key             =   "refresh"
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Árvore"
            Key             =   "showtreeview"
            Object.ToolTipText     =   "Mostrar Árvore"
            ImageIndex      =   3
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Mostrar Lista"
            Key             =   "showlist"
            Object.ToolTipText     =   "Mostrar Lista"
            ImageIndex      =   4
            Style           =   1
            Value           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair               "
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstOperacao 
      Height          =   8715
      Left            =   4710
      TabIndex        =   1
      Top             =   0
      Width           =   10155
      _ExtentX        =   17912
      _ExtentY        =   15372
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   31
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data Operação"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Veículo Legal(Parte)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Ação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tipo Operação"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Entrada / Saída"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Data/Hora Mensagem"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Data Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Local de Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Horario Envio Msg"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Mensagem"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Tipo Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   18
         Text            =   "Preço Unitário"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   19
         Text            =   "Quantidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   20
         Text            =   "Taxa Negociação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "ISPB Banco Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Tipo Instituição Deb/Cred"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Agência Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Conta Corrente Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   25
         Text            =   "Valor Retorno"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Tipo BackOffice"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "Cod. Reembolso"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Tipo Comércio CCR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Código Operação Estruturada"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolBar 
      Left            =   120
      Top             =   8880
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
            Picture         =   "frmConsultaOperacao.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaOperacao.frx":0F6C
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8715
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   15372
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Situação"
      TabPicture(0)   =   "frmConsultaOperacao.frx":1286
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "trvOperacaoStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Tipo Operação"
      TabPicture(1)   =   "frmConsultaOperacao.frx":12A2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "trvOperacaoTipo"
      Tab(1).ControlCount=   1
      Begin MSComctlLib.TreeView trvOperacaoStatus 
         Height          =   8235
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   14526
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   365
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView trvOperacaoTipo 
         Height          =   8235
         Left            =   -74940
         TabIndex        =   3
         Top             =   60
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   14526
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   365
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin MSComCtl2.UpDown udTimer 
      Height          =   315
      Left            =   4620
      TabIndex        =   5
      Top             =   9000
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTimer"
      BuddyDispid     =   196610
      OrigLeft        =   4860
      OrigTop         =   4470
      OrigRight       =   5100
      OrigBottom      =   4815
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label lblTimer 
      AutoSize        =   -1  'True
      Caption         =   "Intervalo para Refresh automático da tela (em minutos) :"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   9060
      Width           =   3945
   End
   Begin VB.Image imgDummyV 
      Height          =   8745
      Left            =   4605
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "frmConsultaOperacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
''-------------------------------------------------
' Gerado pelo Enterprise Architect
' Atualização em:      17-set-2004 11:33:52
'-------------------------------------------------
'' Objeto responsável pela consulta das informações sobre uma operação, através de
'' interação com a camada de controle de caso de uso MIU.
''
'' Classes especificamente consideradas de destino:
''   A8MIU.clsMIU
''   A8MIU.clsOperacao
''
Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private Const strFuncionalidade             As String = "frmConsultaOperacao"
Private WithEvents objFiltro                As frmFiltro
Attribute objFiltro.VB_VarHelpID = -1
Private strFiltroXML                        As String
Private blnUtilizaFiltro                    As Boolean
Private blnOrigemBotaoRefresh               As Boolean
Private blnPrimeiraConsulta                 As Boolean
Private intRefresh                          As Integer

'Constantes de Visão da Lista
Private Const VIS_POR_STATUS                As Integer = 1
Private Const VIS_POR_TIPO_OPERACAO         As Integer = 2

'Constantes de Configuração de Colunas
Private Const COL_NUMERO_COMANDO            As Integer = 1
Private Const COL_VEICULO_LEGAL_PARTE       As Integer = 2
Private Const COL_CONTRAPARTE               As Integer = 3
Private Const COL_SITUACAO                  As Integer = 4
Private Const COL_TP_ACAO_OPER_ATIV_EXEC    As Integer = 5
Private Const COL_TIPO_OPERACAO             As Integer = 6
Private Const COL_TIPO_MOVIMENTO            As Integer = 7
Private Const COL_ENTRADA_SAIDA             As Integer = 8
Private Const COL_VALOR                     As Integer = 9
Private Const COL_OPERACAO                  As Integer = 10
Private Const COL_DATA_HORA_MENSAGEM        As Integer = 11
Private Const COL_DATA_LIQUIDACAO           As Integer = 12
Private Const COL_EMPRESA                   As Integer = 13
Private Const COL_LOCAL_LIQUIDACAO          As Integer = 14
Private Const COL_HORARIO_ENVIO_MSG         As Integer = 15
Private Const COL_MENSAGEM                  As Integer = 16
Private Const COL_TIPO_LIQUIDACAO           As Integer = 17
Private Const COL_PRECO_UNITARIO            As Integer = 18
Private Const COL_QUANTIDADE                As Integer = 19
Private Const COL_TAXA_NEGOCIACAO           As Integer = 20
Private Const COL_ISPB_BANCO_CONTRAPARTE    As Integer = 21
Private Const COL_TIPO_INST_DEBT_CRED       As Integer = 22
Private Const COL_AGENCIA_CONTRAPARTE       As Integer = 23
Private Const COL_CONTA_CONTRAPARTE         As Integer = 24
Private Const COL_VALOR_RETORNO             As Integer = 25
Private Const COL_TIPO_BACKOFFICE           As Integer = 26
Private Const COL_CANAL_VENDA               As Integer = 27
Private Const COL_COD_REEMB                 As Integer = 28
Private Const COL_TIPO_COMER_CCR            As Integer = 29

Private Const COL_OPERACAO_ESTRUTURADA      As Integer = 30
Private Const NUM_OPER_PAGINACAO            As Integer = 4000

Private fblnDummyV                          As Boolean

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Controla a abertura/fechamento dos tipos de operação no listview
Private arrAberturaTipoOperacao()
Private xmlAberturaTipoOperacao             As MSXML2.DOMDocument40

'Controla o timer de refresh da tela
Private intContMinutos                      As Integer
Private blnTimerBypass                      As Boolean

Private lngIndexClassifList                 As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        intRefresh = intRefresh + 1
        If intRefresh > 1 Then
            intRefresh = 0
            Exit Sub
        End If
    
        Call fgCursor(True)
        
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - Form_KeyDown", Me.Caption

End Sub

Private Sub Form_Resize()

On Error Resume Next
    
    DoEvents
    With Me
        
        If .tlbButtons.Buttons("showlist").value = tbrPressed Xor .tlbButtons.Buttons("showtreeview").value = tbrPressed Then
            If tlbButtons.Buttons("showlist").value = tbrPressed Then
                lstOperacao.Left = 30
                lstOperacao.Width = Me.ScaleWidth - 80
                lstOperacao.Height = Me.ScaleHeight - 1000
                
                lblTimer.Top = lstOperacao.Top + lstOperacao.Height + 200
                lblTimer.Left = lstOperacao.Left + 100
                txtTimer.Top = lstOperacao.Top + lstOperacao.Height + 150
                udTimer.Top = lstOperacao.Top + lstOperacao.Height + 150
            
            Else
                .SSTab1.Height = Me.ScaleHeight - tlbButtons.Height - 580
                .SSTab1.Width = Me.ScaleWidth - 80
                .trvOperacaoStatus.Height = SSTab1.Height - 430
                .trvOperacaoStatus.Width = SSTab1.Width - 100
                .trvOperacaoTipo.Height = SSTab1.Height - 430
                .trvOperacaoTipo.Width = SSTab1.Width - 100
                
                lblTimer.Top = .SSTab1.Top + .SSTab1.Height + 200
                lblTimer.Left = .SSTab1.Left + 100
                txtTimer.Top = .SSTab1.Top + .SSTab1.Height + 150
                udTimer.Top = .SSTab1.Top + .SSTab1.Height + 150
                
            End If
        Else
            SSTab1.Left = 30

            .SSTab1.Height = Me.ScaleHeight - tlbButtons.Height - 500
            .SSTab1.Width = imgDummyV.Left - 50
            
            .trvOperacaoStatus.Height = SSTab1.Height - 450
            .trvOperacaoStatus.Width = SSTab1.Width - 150
            .trvOperacaoTipo.Height = SSTab1.Height - tlbButtons.Height - 90
            .trvOperacaoTipo.Width = SSTab1.Width - 150
            
            lstOperacao.Left = imgDummyV.Left + 100
            lstOperacao.Height = Me.ScaleHeight - tlbButtons.Height - 550
            lstOperacao.Width = Me.ScaleWidth - imgDummyV.Left - 120
            
            lblTimer.Top = .SSTab1.Top + .SSTab1.Height + 100
            lblTimer.Left = .SSTab1.Left + 100
            txtTimer.Top = lblTimer.Top - 100
            udTimer.Top = lblTimer.Top - 100
            
        End If
        
    End With
    DoEvents
    
End Sub

Public Sub RedimensionarForm()

    Call Form_Resize

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set objFiltro = Nothing
    Set frmConsultaOperacao = Nothing

End Sub

Private Sub imgDummyV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyV = True
End Sub

Private Sub imgDummyV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

On Error GoTo ErrorHandler

    If Not fblnDummyV Or Button = vbRightButton Then
        Exit Sub
    End If
    
    Me.imgDummyV.Left = x + imgDummyV.Left

    On Error Resume Next
    
    With Me
        If .imgDummyV.Left < 3000 Then
            .imgDummyV.Left = 3000
        End If
        If .imgDummyV.Left > (.Width - 500) And (.Width - 500) > 0 Then
            .imgDummyV.Left = .Width - 500
        End If
        
        .SSTab1.Width = .imgDummyV.Left - 50
        
        .trvOperacaoStatus.Width = .imgDummyV.Left - 150
        .trvOperacaoTipo.Width = .imgDummyV.Left - 150
        
        lstOperacao.Left = imgDummyV.Left + 100
        lstOperacao.Width = Me.ScaleWidth - imgDummyV.Left - 120

    End With
    
    On Error GoTo 0

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - imgDummyV_MouseMove"

End Sub

Private Sub imgDummyV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyV = False
End Sub

'Private Sub lstOperacao_BeforeLabelEdit(Cancel As Integer)
'    Cancel = True
'End Sub

Private Sub lstOperacao_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lstOperacao, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - lstOperacao_ColumnClick", Me.Caption

End Sub

Private Sub lstOperacao_DblClick()

On Error GoTo ErrorHandler
    
    If Not lstOperacao.SelectedItem Is Nothing Then
        With frmDetalheOperacao
            .CodigoEmpresa = fgObterCodigoCombo(lstOperacao.SelectedItem.ListSubItems(COL_EMPRESA))
            .SequenciaOperacao = Mid(lstOperacao.SelectedItem.Key, 2)
            .Show vbModal
        End With
    End If
    
Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - lstOperacao_DblClick", Me.Caption
    
End Sub

Private Sub lstOperacao_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        intRefresh = intRefresh + 1
        If intRefresh > 1 Then
            intRefresh = 0
            Exit Sub
        End If
    
        Call fgCursor(True)
        
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - lstOperacao_KeyDown", Me.Caption
    
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

Dim strSelecaoVisual                        As String
Dim strSelecaoFiltro                        As String

On Error GoTo ErrorHandler

    blnTimerBypass = True
    fgCursor True

    strFiltroXML = xmlDocFiltros
    
    If Trim(strFiltroXML) <> vbNullString Then
        If fgMostraFiltro(strFiltroXML, blnPrimeiraConsulta) Then
            blnPrimeiraConsulta = False
            
            If blnOrigemBotaoRefresh Then
                frmMural.Caption = Me.Caption
                frmMural.Display = "Obrigatória a seleção do filtro DATA."
                frmMural.Show vbModal
                
                Exit Sub
            Else
                Call tlbButtons_ButtonClick(tlbButtons.Buttons("showfiltro"))
            End If
        End If
        
        'Pressiona o botão << Aplicar Filtro >> apenas se o filtro for selecionado diretamente
        If Not blnOrigemBotaoRefresh Then
            blnUtilizaFiltro = True
            tlbButtons.Buttons("AplicarFiltro").value = tbrPressed
        End If
        
        If SSTab1.Tab = 0 Then
            strSelecaoVisual = flObterSelecaoTreeview(trvOperacaoStatus, True)
            strSelecaoFiltro = flObterSelecaoTreeview(trvOperacaoStatus)
            
            If Not flCarregarTreeViewSituStatus(trvOperacaoStatus, _
                                                IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                "Todas Situações") Then Exit Sub
            
            If Trim(strSelecaoFiltro) <> "" Then
                fgLockWindow trvOperacaoStatus.hwnd
                Call flRetornarSelecaoAnterior(trvOperacaoStatus, strSelecaoVisual)
                fgLockWindow 0
                
                Call flCarregarLista(strSelecaoFiltro, VIS_POR_STATUS, _
                                        IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
            Else
                Call flLimparLista
            End If
        
        Else
            strSelecaoVisual = flObterSelecaoTreeview(trvOperacaoTipo, True)
            strSelecaoFiltro = flObterSelecaoTreeview(trvOperacaoTipo)
            
            If Not flCarregarTreeViewTipoOperacao(trvOperacaoTipo, _
                                                  IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                  "Todos Tipos de Operações") Then Exit Sub
            
            If Trim(strSelecaoFiltro) <> "" Then
                fgLockWindow trvOperacaoTipo.hwnd
                Call flRetornarSelecaoAnterior(trvOperacaoTipo, strSelecaoVisual)
                fgLockWindow 0
                
                Call flCarregarLista(strSelecaoFiltro, VIS_POR_TIPO_OPERACAO, _
                                        IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
            Else
                Call flLimparLista
            End If
        End If
    End If
    
    fgCursor
    blnTimerBypass = False

Exit Sub
ErrorHandler:
    fgCursor
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - objFiltro_AplicarFiltro", Me.Caption
    
End Sub

Private Sub Form_Load()
    
On Error GoTo ErrorHandler
    
    fgCenterMe Me
    Set Me.Icon = mdiLQS.Icon
    Me.Show
    DoEvents
    
    Call fgCursor(True)
    Call flInicializar
    
    blnPrimeiraConsulta = True
    
    blnUtilizaFiltro = (tlbButtons.Buttons("AplicarFiltro").value = tbrPressed)
    
    Call flCarregarTreeViewSituStatus(trvOperacaoStatus, vbNullString, "Todas Situações", False)
    Call flCarregarTreeViewTipoOperacao(trvOperacaoTipo, vbNullString, "Todos Tipos de Operações", False)
    
    Set objFiltro = New frmFiltro
    Set objFiltro.FormOwner = Me
    objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaOperacao
    Load objFiltro
    
    Call objFiltro.fgCarregarPesquisaAnterior
    
    blnPrimeiraConsulta = False
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - Form_Load", Me.Caption
    
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo ErrorHandler

    Call fgCursor(True)
    Call flLimparLista
    
    Select Case SSTab1.Tab
           Case 0   'Situação
                Call flCarregarTreeViewSituStatus(trvOperacaoStatus, _
                                                  IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                  "Todas Situações")
           
           Case 1   'Tipo Operação
                Call flCarregarTreeViewTipoOperacao(trvOperacaoTipo, _
                                                    IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                    "Todos Tipos de Operações")
    End Select
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "SSTab1_Click", Me.Caption
    
End Sub

Private Sub tlbButtons_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim strJanelas                              As String

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    'Verifica se o filtro deve ser aplicado
    blnUtilizaFiltro = (tlbButtons.Buttons("AplicarFiltro").value = tbrPressed)
    
    If tlbButtons.Buttons("showtreeview").value = tbrPressed Then
        strJanelas = strJanelas & "1"
    End If
    
    If tlbButtons.Buttons("showlist").value = tbrPressed Then
        strJanelas = strJanelas & "2"
    End If
    
    Call flArranjarJanelasExibicao(strJanelas)
    
    Select Case Button.Key
           Case "showfiltro"
                Set objFiltro = Nothing
                Set objFiltro = New frmFiltro
                Set objFiltro.FormOwner = Me
                objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaOperacao
                objFiltro.Show vbModal
                
            Case "refresh"
                blnOrigemBotaoRefresh = True
                objFiltro.fgCarregarPesquisaAnterior
                blnOrigemBotaoRefresh = False
                
            Case gstrSair
                Unload Me
                
    End Select
    
    Call fgCursor(False)
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    blnOrigemBotaoRefresh = False
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - tlbButtons_ButtonClick", Me.Caption
    
End Sub

Private Sub trvOperacaoStatus_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub trvOperacaoStatus_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        intRefresh = intRefresh + 1
        If intRefresh > 1 Then
            intRefresh = 0
            Exit Sub
        End If
    
        Call fgCursor(True)
        
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - trvOperacaoStatus_KeyDown", Me.Caption

End Sub

Private Sub trvOperacaoStatus_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim strSelecao                              As String
    
On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Node.Selected = True
    Call flMarcarNodes(trvOperacaoStatus, (Node.children > 0), Node.Checked)
    
    strSelecao = flObterSelecaoTreeview(trvOperacaoStatus)
    
    If Trim(strSelecao) <> "" Then
        Call flCarregarLista(strSelecao, VIS_POR_STATUS, IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
    Else
        Call flLimparLista
    End If

    Call fgCursor(False)

Exit Sub
ErrorHandler:
    Call fgCursor(False)
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - trvOperacaoTipo_NodeCheck", Me.Caption

End Sub

Private Sub trvOperacaoTipo_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

'' Carrega o treeview com todos os Status de operação, retornando True em caso de
'' sucesso, através de interação com a camada de controle de caso de uso MIU,
'' método: A8MIU.clsOperacao.ObterOperacoesPorStatus
Private Function flCarregarTreeViewSituStatus(ByVal ptreTreeView As TreeView, _
                                              ByVal pstrFiltroXML As String, _
                                     Optional ByVal pstrNomeRoot As String, _
                                     Optional ByVal pblnMostrarQuantidade As Boolean = True) As Boolean

#If EnableSoap = 1 Then
    Dim objOperacao         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao         As A8MIU.clsOperacao
#End If

Dim xmlDocument             As MSXML2.DOMDocument40
Dim objDomNode              As MSXML2.IXMLDOMNode
Dim strCargaStatus          As String
Dim strQtd                  As String
Dim lngTotal                As Long
Dim strFiltros              As String
Dim vntCodErro              As Variant
Dim vntMensagemErro         As Variant

On Error GoTo ErrorHandler
    
    fgLockWindow Me.hwnd
    
    'Verifica se existe filtro...
    If pstrFiltroXML <> vbNullString Then
        If fgMostraFiltro(pstrFiltroXML, False) Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
            fgLockWindow 0
            Exit Function
        End If
        
        strFiltros = pstrFiltroXML
    
    '...se não, envia um filtro vazio
    Else
        If pblnMostrarQuantidade Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
            fgLockWindow 0
            Exit Function
        End If
        
        strFiltros = flMontarFiltro(Not pblnMostrarQuantidade)
        
    End If
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strCargaStatus = objOperacao.ObterOperacoesPorStatus(strFiltros, _
                                                         vntCodErro, _
                                                         vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    Set xmlDocument = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDocument.loadXML(strCargaStatus) Then
        Call fgErroLoadXML(xmlDocument, App.EXEName, TypeName(Me), "flCarregarTreeViewSituStatus")
    End If
    
    With ptreTreeView
    
        .Nodes.Clear
        
        If pstrNomeRoot <> "" Then
           .Nodes.Add , , "root", pstrNomeRoot
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_SituacaoSistema/*")
                If pblnMostrarQuantidade Then
                    If Val(objDomNode.selectSingleNode("NU_QTD").Text) <> 0 Then
                        strQtd = " (" & objDomNode.selectSingleNode("NU_QTD").Text & ")"
                        lngTotal = lngTotal + Val(objDomNode.selectSingleNode("NU_QTD").Text)
                    Else
                        strQtd = vbNullString
                    End If
                End If
           
                .Nodes.Add "root", tvwChild, "k" & objDomNode.selectSingleNode("CO_SITU_PROC").Text, _
                                                   objDomNode.selectSingleNode("DE_SITU_PROC").Text & _
                                                   strQtd
                .Nodes.Item("k" & objDomNode.selectSingleNode("CO_SITU_PROC").Text).EnsureVisible
            Next
            
            If pblnMostrarQuantidade Then
                If lngTotal > 0 Then
                    .Nodes(1).Text = .Nodes(1).Text & " (" & lngTotal & ")"
                End If
            End If
        Else
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_SituacaoSistema/*")
                If pblnMostrarQuantidade Then
                    If Val(objDomNode.selectSingleNode("NU_QTD").Text) <> 0 Then
                        strQtd = " (" & objDomNode.selectSingleNode("NU_QTD").Text & ")"
                    Else
                        strQtd = vbNullString
                    End If
                End If
                
                .Nodes.Add , , "k" & objDomNode.selectSingleNode("CO_SITU_PROC").Text, _
                                    objDomNode.selectSingleNode("DE_SITU_PROC").Text & _
                                    strQtd
                .Nodes.Item("k" & objDomNode.selectSingleNode("CO_SITU_PROC").Text).EnsureVisible
            Next
        End If
    
    End With
    
    If Me.trvOperacaoStatus.Nodes.Count > 0 Then
       Me.trvOperacaoStatus.Nodes(1).EnsureVisible
    End If
    
    Set xmlDocument = Nothing
    Set objDomNode = Nothing
    
    flCarregarTreeViewSituStatus = True
    fgLockWindow 0
    
Exit Function
ErrorHandler:
    fgLockWindow 0
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarTreeViewSituStatus", 0
    
End Function

'' Carrega o treeview com todos os tipos de operação, retornando True em caso de
'' sucesso, através de interação com a camada de controle de caso de uso MIU,
'' método: A8MIU.clsOperacao.ObterOperacoesPorTipoOperacao
Private Function flCarregarTreeViewTipoOperacao(ByVal treTreeView As TreeView, _
                                                ByVal pstrFiltroXML As String, _
                                       Optional ByVal strNomeRoot As String, _
                                       Optional ByVal blnMostrarQuantidade As Boolean = True) As Boolean

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim xmlDocument                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim strCargaTipoOperacao                    As String
Dim strQtd                                  As String
Dim strGrupoAux                             As String
Dim lngTotal                                As Long
Dim lngGrupo                                As Long
Dim strFiltros                              As String
Dim arrGrupo()                              As Long
Dim intCont                                 As Long
Dim intPosArray                             As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    fgLockWindow Me.hwnd
    'Verifica se existe filtro...
    If pstrFiltroXML <> vbNullString Then
        If fgMostraFiltro(pstrFiltroXML, False) Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
            fgLockWindow 0
            Exit Function
        End If
        
        strFiltros = pstrFiltroXML
    
    '...se não, envia um filtro vazio
    Else
        If blnMostrarQuantidade Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
            fgLockWindow 0
            Exit Function
        End If
        
        strFiltros = flMontarFiltro(Not blnMostrarQuantidade)
        
    End If
    
    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
    strCargaTipoOperacao = objOperacao.ObterOperacoesPorTipoOperacao(strFiltros, _
                                                                     vntCodErro, _
                                                                     vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    Set xmlDocument = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDocument.loadXML(strCargaTipoOperacao) Then
        Call fgErroLoadXML(xmlDocument, App.EXEName, TypeName(Me), "flCarregarTreeViewTipoOperacao")
    End If
    
    With treTreeView
    
        .Nodes.Clear
        
        ReDim arrGrupo(0)
        If strNomeRoot <> "" Then
            .Nodes.Add , , "root", strNomeRoot
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_TipoOperacao/*")
                If blnMostrarQuantidade Then
                    If Val(objDomNode.selectSingleNode("NU_QTD").Text) <> 0 Then
                        strQtd = " (" & objDomNode.selectSingleNode("NU_QTD").Text & ")"
                        lngTotal = lngTotal + Val(objDomNode.selectSingleNode("NU_QTD").Text)
                    Else
                        strQtd = vbNullString
                    End If
                End If
                
                'If objDomNode.selectSingleNode("SG_LOCA_LIQU").Text = "SELIC" Then Stop
                
                If strGrupoAux <> objDomNode.selectSingleNode("SG_LOCA_LIQU").Text Then
                    If strGrupoAux <> vbNullString Then
                        arrGrupo(UBound(arrGrupo)) = lngGrupo
                        
                        ReDim Preserve arrGrupo(UBound(arrGrupo) + 1)
                    End If
                    
                    strGrupoAux = objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                    lngGrupo = 0
                    
                    .Nodes.Add "root", tvwChild, "k" & strGrupoAux, _
                                                 objDomNode.selectSingleNode("SG_LOCA_LIQU").Text
                                                 
                    .Nodes.Item("k" & objDomNode.selectSingleNode("SG_LOCA_LIQU").Text).EnsureVisible
                End If
                
                .Nodes.Add "k" & strGrupoAux, tvwChild, "k" & objDomNode.selectSingleNode("TP_OPER").Text, _
                                                    objDomNode.selectSingleNode("NO_TIPO_OPER").Text & _
                                                    strQtd
                If blnMostrarQuantidade Then
                    lngGrupo = lngGrupo + Val(objDomNode.selectSingleNode("NU_QTD").Text)
                End If
            Next
            
            arrGrupo(UBound(arrGrupo)) = lngGrupo
            If blnMostrarQuantidade Then
                If lngTotal > 0 Then
                    .Nodes(1).Text = .Nodes(1).Text & " (" & lngTotal & ")"
                End If
                
                'Atualiza as quantidades dos grupos
                With treTreeView.Nodes
                    For intCont = 1 To .Count
                        If .Item(intCont).Key <> "root" Then
                            If .Item(intCont).Parent.Key = "root" Then
                                If arrGrupo(intPosArray) <> 0 Then
                                    .Item(intCont).Text = .Item(intCont).Text & " (" & arrGrupo(intPosArray) & ")"
                                End If
                                
                                intPosArray = intPosArray + 1
                            End If
                        End If
                    Next
                End With

            End If
        Else
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_TipoOperacao/*")
                If blnMostrarQuantidade Then
                    If Val(objDomNode.selectSingleNode("NU_QTD").Text) <> 0 Then
                        strQtd = " (" & objDomNode.selectSingleNode("NU_QTD").Text & ")"
                    Else
                        strQtd = vbNullString
                    End If
                End If
                
                .Nodes.Add , , "k" & objDomNode.selectSingleNode("TP_OPER").Text, _
                                    objDomNode.selectSingleNode("NO_TIPO_OPER").Text & _
                                    strQtd
                .Nodes.Item("k" & objDomNode.selectSingleNode("TP_OPER").Text).EnsureVisible
            Next
        End If
    
    End With
    
    If Me.trvOperacaoTipo.Nodes.Count > 0 Then
       Me.trvOperacaoTipo.Nodes(1).EnsureVisible
       fgAberturaTreeViewRefresh treTreeView, xmlAberturaTipoOperacao
    End If
    
    Set xmlDocument = Nothing
    Set objDomNode = Nothing
    
    flCarregarTreeViewTipoOperacao = True
    
    fgLockWindow 0

Exit Function
ErrorHandler:
    fgLockWindow 0
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarTreeViewTipoOperacao", 0
    Exit Function
    Resume
End Function

'' Obtém as mensagens pertinenntes ao filtro e as exibe no listview de Operacoes.
'' Utiliza a camada de controle de caso de uso, método A8MIU.clsOperacao.
'' ObterDetalheOperacao
Private Sub flCarregarLista(ByVal pstrSelecaoFiltro As String, _
                            ByVal pintTipoFiltro As Integer, _
                            ByVal pstrFiltroXML As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsOperacao
#End If

Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode

Dim strTagGrupoFiltro                       As String
Dim strTagFiltro                            As String
Dim strDataOperacao                         As String
Dim lngCodigoEmpresa                        As Long

Dim vntSequenciaOperacao                    As Variant

Dim lngCont                                 As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    fgLockWindow Me.hwnd
    
    Call flLimparLista
    
    'Verifica se existe filtro...
    If pstrFiltroXML <> vbNullString Then
        If fgMostraFiltro(pstrFiltroXML, False) Then
            frmMural.Caption = Me.Caption
            frmMural.Display = "Obrigatória a seleção do filtro DATA."
            frmMural.Show vbModal
            fgLockWindow 0
            Exit Sub
        End If
        
    '...se não, retorna mensagem de aviso e sai da rotina
    Else
        frmMural.Caption = Me.Caption
        frmMural.Display = "Obrigatória a seleção do filtro DATA."
        frmMural.Show vbModal
        fgLockWindow 0
        Exit Sub
    End If
    
    'Verifica qual filtro foi selecionado
    Select Case pintTipoFiltro
        Case VIS_POR_STATUS
            strTagGrupoFiltro = "Grupo_Status"
            strTagFiltro = "Status"
            
        Case VIS_POR_TIPO_OPERACAO
            strTagGrupoFiltro = "Grupo_TipoOperacao"
            strTagFiltro = "TipoOperacao"
            
    End Select
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltros.loadXML(pstrFiltroXML) Then
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", strTagGrupoFiltro, "")
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "BaseHistorica", "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(pstrSelecaoFiltro, ";")) To UBound(Split(pstrSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, strTagGrupoFiltro, _
                                         strTagFiltro, Split(pstrSelecaoFiltro, ";")(lngCont))
    Next
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_OrderBy", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_OrderBy", "Order", "NU_SEQU_OPER_ATIV")
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_RowNun", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_RowNun", "RowNun", NUM_OPER_PAGINACAO)

    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_SeqOperPaginacao", "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_SeqOperPaginacao", "SeqOper", "0")

    
    '>>> -------------------------------------------------------------------------------------------
    
    vntSequenciaOperacao = 0
    
    Do
        Set objOperacao = fgCriarObjetoMIU("A8MIU.clsOperacao")
        strRetLeitura = objOperacao.ObterDadosConsultaOperacao(xmlDomFiltros.xml, _
                                                               vntCodErro, _
                                                               vntMensagemErro)
        
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
        
        Set objOperacao = Nothing
        
        If strRetLeitura <> vbNullString Then
            Call flCarregarListaViewOper(strRetLeitura, vntSequenciaOperacao)
        Else
            vntSequenciaOperacao = "0"
        End If
        
        xmlDomFiltros.selectSingleNode("//SeqOper").Text = vntSequenciaOperacao

    Loop Until vntSequenciaOperacao = 0
    
    
    
    
    
    Call fgClassificarListview(Me.lstOperacao, lngIndexClassifList, True)
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing

    fgLockWindow 0

Exit Sub
ErrorHandler:

    fgLockWindow 0
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0

End Sub

Private Sub flCarregarListaViewOper(ByRef pstrRetLeitura As String, _
                                    ByRef plngUltimaSequencial As Variant)


Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim vntSequenciaOperacao                    As Variant
Dim strDataOperacao                         As String
Dim lngCodigoEmpresa                        As Long
Dim lngCont                                 As Long

Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler

    Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    If Not xmlDomLeitura.loadXML(pstrRetLeitura) Then
        Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregarListaViewOper")
    End If
        
    lngCont = xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*").length
    
    If lngCont > 0 Then
        plngUltimaSequencial = xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*").Item(lngCont - 1).selectSingleNode("NU_SEQU_OPER_ATIV").Text
    Else
        plngUltimaSequencial = 0
    End If
        
    For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheOperacao/*")
        strDataOperacao = vbNullString
        
        If objDomNode.selectSingleNode("DT_OPER_ATIV").Text <> gstrDataVazia Then
            strDataOperacao = fgDtXML_To_Interface(objDomNode.selectSingleNode("DT_OPER_ATIV").Text)
        End If
        
        If objDomNode.selectSingleNode("OWNER").Text = "A8HIST" Then
            vntSequenciaOperacao = objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text * -1
        Else
            vntSequenciaOperacao = objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
        End If
        
        With lstOperacao.ListItems.Add(, "k" & vntSequenciaOperacao, strDataOperacao)
            
            .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
            .SubItems(COL_VEICULO_LEGAL_PARTE) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
            .SubItems(COL_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
            .SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
            .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_OPER_DEBT_CRED").Text
            .SubItems(COL_ENTRADA_SAIDA) = objDomNode.selectSingleNode("IN_ENTR_SAID_RECU_FINC").Text
            .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV").Text)
            .SubItems(COL_VALOR_RETORNO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_OPER_ATIV_RETN").Text)
            .SubItems(COL_OPERACAO) = objDomNode.selectSingleNode("CO_OPER_ATIV").Text
            
            If objDomNode.selectSingleNode("DH_MESG_INTE").Text <> gstrDataVazia Then
                .SubItems(COL_DATA_HORA_MENSAGEM) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_MESG_INTE").Text)
            Else
                .SubItems(COL_DATA_HORA_MENSAGEM) = vbNullString
            End If
            
            .SubItems(COL_TP_ACAO_OPER_ATIV_EXEC) = fgDescricaoTipoAcao(CLng("0" & objDomNode.selectSingleNode("TP_ACAO_OPER_ATIV_EXEC").Text))
            .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
            
            .SubItems(COL_TIPO_BACKOFFICE) = objDomNode.selectSingleNode("TP_BKOF").Text
            If Trim$(objDomNode.selectSingleNode("DE_BKOF").Text) <> vbNullString Then
                .SubItems(COL_TIPO_BACKOFFICE) = .SubItems(COL_TIPO_BACKOFFICE) & " - " & objDomNode.selectSingleNode("DE_BKOF").Text
            End If
            
            If objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text <> gstrDataVazia Then
                .SubItems(COL_DATA_LIQUIDACAO) = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_LIQU_OPER_ATIV").Text)
            End If
            
            If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
               Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                'Obtem a descrição da Empresa via QUERY XML
                .SubItems(COL_EMPRESA) = _
                        objDomNode.selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                        objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
            End If
            
            If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
               Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                
                If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                            objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                    
                    'Obtem a descrição do Local de Liquidação via QUERY XML
                    .SubItems(COL_LOCAL_LIQUIDACAO) = _
                        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                            objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
            
                Else
                    
                    vntCodErro = 5
                    vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                      objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                    GoTo ErrorHandler
                    
                End If
            
            End If
            
            If objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text <> gstrDataVazia Then
                .SubItems(COL_HORARIO_ENVIO_MSG) = Format(fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("HO_ENVI_MESG_SPB").Text), "HH:MM")
            End If
            
            .SubItems(COL_MENSAGEM) = objDomNode.selectSingleNode("CO_MESG_SPB_REGT_OPER").Text
            .SubItems(COL_TIPO_LIQUIDACAO) = objDomNode.selectSingleNode("NO_TIPO_LIQU_OPER_ATIV").Text
            .SubItems(COL_PRECO_UNITARIO) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("PU_ATIV_MERC").Text)
            .SubItems(COL_QUANTIDADE) = objDomNode.selectSingleNode("QT_ATIV_MERC").Text
            .SubItems(COL_TAXA_NEGOCIACAO) = fgVlrXml_To_InterfaceDecimais(objDomNode.selectSingleNode("PE_TAXA_NEGO").Text, 8)
            .SubItems(COL_ISPB_BANCO_CONTRAPARTE) = objDomNode.selectSingleNode("CO_ISPB_BANC_LIQU_CNPT").Text
            .SubItems(COL_TIPO_INST_DEBT_CRED) = objDomNode.selectSingleNode("DE_IF_CRED_DEBT").Text
            .SubItems(COL_AGENCIA_CONTRAPARTE) = objDomNode.selectSingleNode("CO_AGEN_COTR").Text
            .SubItems(COL_CONTA_CONTRAPARTE) = objDomNode.selectSingleNode("NU_CC_COTR").Text
            
            'KIDA - SGC
            .SubItems(COL_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
            
            If Not objDomNode.selectSingleNode("CD_OPER_ETRT") Is Nothing Then
                .SubItems(COL_OPERACAO_ESTRUTURADA) = objDomNode.selectSingleNode("CD_OPER_ETRT").Text
            End If
            
            If Val("0" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text = enumLocalLiquidacao.CCR Then
                    .SubItems(COL_COD_REEMB) = objDomNode.selectSingleNode("NU_CTRL_MESG_SPB_ORIG").Text
                    
                    If UCase(Mid(objDomNode.selectSingleNode("NU_ATIV_MERC").Text, 1, 2)) = "EX" Then
                        .SubItems(COL_TIPO_COMER_CCR) = "Exportação"
                    ElseIf UCase(Mid(objDomNode.selectSingleNode("NU_ATIV_MERC").Text, 1, 2)) = "IM" Then
                        .SubItems(COL_TIPO_COMER_CCR) = "Importação"
                    End If
                End If
            End If
            
            
        End With
    Next
    
    Set xmlDomLeitura = Nothing
    Set objDomNode = Nothing
    
    Exit Sub
ErrorHandler:
    Set xmlDomLeitura = Nothing
    Set objDomNode = Nothing


    Err.Raise Err.Number, Err.Source, Err.Description

End Sub


'' Remove todos os itens da lista de operações.
Private Sub flLimparLista()
    Me.lstOperacao.ListItems.Clear
End Sub

'' Retorna uma String com a seleção dos nós do treeview separados por ";" para
'' separação com a função Split
Private Function flObterSelecaoTreeview(ByVal treTreeView As TreeView, _
                               Optional ByVal pblnConsideraGrupo As Boolean = False) As String

    '>>> -----------------------------------------------------------------------
    'Captura todos os nodes selecionados (Checked), exceto o node RAIZ e,
    'retorna uma STRING, com separador ";", a ser decomposta na função SPLIT.
    '>>> -----------------------------------------------------------------------

Dim intCont                                 As Integer
Dim strRetorno                              As String

On Error GoTo ErrorHandler

    With treTreeView.Nodes
        For intCont = 1 To .Count
            If pblnConsideraGrupo Then
                If .Item(intCont).Checked Then
                    strRetorno = strRetorno & Mid(.Item(intCont).Key, 2) & ";"
                End If
            Else
                If .Item(intCont).children = 0 And .Item(intCont).Checked Then
                    strRetorno = strRetorno & Mid(.Item(intCont).Key, 2) & ";"
                End If
            End If
        Next
        
        If Trim(strRetorno) <> "" Then strRetorno = Left$(strRetorno, Len(strRetorno) - 1)
    End With
    
    flObterSelecaoTreeview = strRetorno

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flObterSelecaoTreeview", 0
    
End Function

Private Sub trvOperacaoTipo_Collapse(ByVal Node As MSComctlLib.Node)

    fgAberturaTreeViewSet Node, xmlAberturaTipoOperacao

End Sub

Private Sub trvOperacaoTipo_Expand(ByVal Node As MSComctlLib.Node)

    fgAberturaTreeViewSet Node, xmlAberturaTipoOperacao

End Sub

Private Sub trvOperacaoTipo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        intRefresh = intRefresh + 1
        If intRefresh > 1 Then
            intRefresh = 0
            Exit Sub
        End If
    
        Call fgCursor(True)
        
        Call tlbButtons_ButtonClick(tlbButtons.Buttons("refresh"))
        
        Call fgCursor(False)
    End If
    
Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - trvOperacaoTipo_KeyDown", Me.Caption
    
End Sub

Private Sub trvOperacaoTipo_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim strSelecao                              As String

On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Node.Selected = True
    Call flMarcarNodesGrupo(trvOperacaoTipo, Node.Index, Node.Checked)
    Call flMarcarNodes(trvOperacaoTipo, (Node.children > 0 And Node.Key = "root"), Node.Checked)
    
    strSelecao = flObterSelecaoTreeview(trvOperacaoTipo)
    If Trim(strSelecao) <> "" Then
        If InStr(1, strFiltroXML, "TipoOperacao") <> 0 Then
            strSelecao = vbNullString
        End If
        Call flCarregarLista(strSelecao, VIS_POR_TIPO_OPERACAO, IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
    Else
        Call flLimparLista
    End If

    Call fgCursor(False)

Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - trvOperacaoTipo_NodeCheck", Me.Caption

End Sub

'' Marca ou desmarca (Check) nodes referentes ao TreeView informado.
Private Sub flMarcarNodes(ByVal treTreeView As TreeView, _
                          ByVal blnNodeRoot As Boolean, _
                          ByVal blnMarcar As Boolean)

    '>>> -----------------------------------------------------------------------
    'Marca ou desmarca (Check) nodes referentes ao TreeView informado.
    '
    'Se o Node Root for informado, transfere seu status para todo TreeView,
    'se não, reflete o status do Node Child no Node Root.
    '
    'Obs.:  Utiliza a API LockWindowUpdate, para que o evento do TreeView
    '       << _NodeCheck >> não seja disparado a cada iteração.
    '>>> -----------------------------------------------------------------------

Dim intCont                                 As Integer
Dim blnMarcaNodeRoot                        As Boolean

On Error GoTo ErrorHandler

    fgLockWindow treTreeView.hwnd
    
    With treTreeView.Nodes
        If blnNodeRoot Then
            For intCont = 1 To .Count
                .Item(intCont).Checked = blnMarcar
            Next
        Else
            If blnMarcar Then
                blnMarcaNodeRoot = True
                
                For intCont = 2 To .Count
                    If Not .Item(intCont).Checked Then
                        blnMarcaNodeRoot = False
                        
                        Exit For
                    End If
                Next
                
                If blnMarcaNodeRoot Then .Item(1).Checked = True
            Else
                .Item(1).Checked = False
            End If
        End If
    End With
    
    fgLockWindow 0

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMarcarNodes", 0

End Sub

Private Sub flArranjarJanelasExibicao(ByVal pstrJanelas As String)
    
On Error GoTo ErrorHandler

    Select Case pstrJanelas
           Case ""
                imgDummyV.Visible = False
                SSTab1.Visible = False
                lstOperacao.Visible = False
            
           Case "1"
                imgDummyV.Visible = False
                SSTab1.Visible = True
                lstOperacao.Visible = False
            
           Case "2"
                imgDummyV.Visible = False
                SSTab1.Visible = False
                lstOperacao.Visible = True
                
           Case "12"
                imgDummyV.Visible = True
                SSTab1.Visible = True
                lstOperacao.Visible = True
                
    End Select
    
    Call Form_Resize

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

'' Obtém as propriedades inicias da tela, através de interação com a camada de
'' controle de caso de uso, método   A8MIU.clsMiu.ObterMapaNavegacao
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
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, "frmConsultaOperacao", "flInicializar")
    End If
    
    Set objMIU = Nothing
    
Exit Sub
ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'' Retorna os itens do treeview para a selecao no parametro selecao
Private Sub flRetornarSelecaoAnterior(ByVal treTreeView As TreeView, _
                                      ByVal strSelecao As String)

Dim intCont                                 As Integer
Dim intContAux                              As Integer

On Error GoTo ErrorHandler

    With treTreeView.Nodes
        For intCont = 1 To .Count
            For intContAux = LBound(Split(strSelecao, ";")) To UBound(Split(strSelecao, ";"))
                If Mid(.Item(intCont).Key, 2) = Split(strSelecao, ";")(intContAux) Then
                    .Item(intCont).Checked = True
                    
                    Exit For
                End If
            Next
        Next
    End With

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flRetornarSelecaoAnterior", 0
    
End Sub

'' Formata XML Filtro padrão
Private Function flMontarFiltro(Optional ByVal pblnOcultarQuantidades As Boolean = False) As String

Dim xmlDomFiltros                           As MSXML2.DOMDocument40

On Error GoTo ErrorHandler

    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    
    If pblnOcultarQuantidades Then
        Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_Quantidade", "")
        Call fgAppendNode(xmlDomFiltros, "Grupo_Quantidade", "OcultarQuantidade", 1)
    End If
    '>>> -------------------------------------------------------------------------------------------

    flMontarFiltro = xmlDomFiltros.xml
    
    Set xmlDomFiltros = Nothing

Exit Function
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMontarFiltro", 0

End Function

Private Sub flMarcarNodesGrupo(ByVal treTreeView As TreeView, _
                               ByVal lngNodeIndex As Long, _
                               ByVal blnMarcar As Boolean)

Dim blnMarcaNodeGrupo                       As Boolean
Dim lngContAux                              As Long

On Error GoTo ErrorHandler

    fgLockWindow treTreeView.hwnd

    With treTreeView.Nodes
        If .Item(lngNodeIndex).Key <> "root" Then
            If .Item(lngNodeIndex).children > 0 Then
                lngContAux = lngNodeIndex + 1
                Do
                    .Item(lngContAux).Checked = blnMarcar
                    lngContAux = lngContAux + 1
                    
                    If lngContAux > .Count Then Exit Do
                    If .Item(lngContAux).children > 0 Then Exit Do
                Loop
            Else
                If blnMarcar Then
                    lngContAux = .Item(.Item(lngNodeIndex).Parent.Key).Index + 1
                    
                    blnMarcaNodeGrupo = True
                    Do
                        If Not .Item(lngContAux).Checked Then
                            blnMarcaNodeGrupo = False
                            Exit Do
                        End If
                        
                        lngContAux = lngContAux + 1
                        
                        If lngContAux > .Count Then Exit Do
                        If .Item(lngContAux).children > 0 Then Exit Do
                    Loop
                    
                    If blnMarcaNodeGrupo Then .Item(.Item(lngNodeIndex).Parent.Key).Checked = True
                Else
                    .Item(.Item(lngNodeIndex).Parent.Key).Checked = False
                End If
            End If
        End If
    End With
    
    fgLockWindow 0

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flMarcarNodesGrupo", 0
    
End Sub

Private Sub tmrRefresh_Timer()

Dim strSelecao              As String
Dim strSelecaoVisual        As String
Dim strSelecaoFiltro        As String

On Error GoTo ErrorHandler

    If blnTimerBypass Then Exit Sub
    
    If Not IsNumeric(txtTimer.Text) Then Exit Sub
    
    If CLng(txtTimer.Text) = 0 Then Exit Sub
    
    If fgVerificaJanelaVerificacao() Then Exit Sub
    
    fgCursor True

    intContMinutos = intContMinutos + 1
    
    If intContMinutos >= txtTimer.Text Then

        'Pressiona o botão << Aplicar Filtro >> apenas se o filtro for selecionado diretamente
        If Not blnOrigemBotaoRefresh Then
            blnUtilizaFiltro = True
            tlbButtons.Buttons("AplicarFiltro").value = tbrPressed
        End If
        
        If SSTab1.Tab = 0 Then
            strSelecaoVisual = flObterSelecaoTreeview(trvOperacaoStatus, True)
            strSelecaoFiltro = flObterSelecaoTreeview(trvOperacaoStatus)
            
            If Not flCarregarTreeViewSituStatus(trvOperacaoStatus, _
                                                IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                "Todas Situações") Then Exit Sub
            
            If Trim(strSelecaoFiltro) <> "" Then
                fgLockWindow trvOperacaoStatus.hwnd
                Call flRetornarSelecaoAnterior(trvOperacaoStatus, strSelecaoVisual)
                fgLockWindow 0
                
                Call flCarregarLista(strSelecaoFiltro, VIS_POR_STATUS, _
                                        IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
            Else
                Call flLimparLista
            End If
        
        Else
            strSelecaoVisual = flObterSelecaoTreeview(trvOperacaoTipo, True)
            strSelecaoFiltro = flObterSelecaoTreeview(trvOperacaoTipo)
            
            If Not flCarregarTreeViewTipoOperacao(trvOperacaoTipo, _
                                                  IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                                  "Todos Tipos de Operações") Then Exit Sub
            
            If Trim(strSelecaoFiltro) <> "" Then
                fgLockWindow trvOperacaoTipo.hwnd
                Call flRetornarSelecaoAnterior(trvOperacaoTipo, strSelecaoVisual)
                fgLockWindow 0
                
                Call flCarregarLista(strSelecaoFiltro, VIS_POR_TIPO_OPERACAO, _
                                        IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
            Else
                Call flLimparLista
            End If
        End If
        
        intContMinutos = 0
    End If

    fgCursor False

Exit Sub
ErrorHandler:
    
    fgCursor False
    
    fgRaiseError App.EXEName, TypeName(Me), "tmrRefresh_Timer", 0

End Sub

