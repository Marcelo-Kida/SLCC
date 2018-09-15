VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmConsultaContaCorrente 
   Caption         =   "Consulta Lançamento Conta Corrente"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   13455
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar tlbButtons 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8325
      Width           =   13455
      _ExtentX        =   23733
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
   Begin MSComctlLib.ListView lvwContaCorrente 
      Height          =   8715
      Left            =   4740
      TabIndex        =   1
      Top             =   0
      Width           =   10155
      _ExtentX        =   17886
      _ExtentY        =   15346
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   36
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Data / Hora Operação"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lançamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sistema"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Empresa"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Número Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Veiculo Legal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Situação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tipo Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Local Liquidação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Banco"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Agência"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Número C/C"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Valor Lançamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Tipo Movto."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Tipo Lançamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Sub-tipo Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Finalidade TED"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Conta Contábil Débito"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   18
         Text            =   "Conta Contábil Crédito"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   19
         Text            =   "Código Histórico Contábil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   20
         Text            =   "Descrição Histórico Contábil"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   21
         Text            =   "Tipo BackOffice"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   22
         Text            =   "Código Veículo Legal"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   23
         Text            =   "Código Situação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   24
         Text            =   "Código Tipo Operação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   25
         Text            =   "Código Banco"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   26
         Text            =   "Código Local Liquidação"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   27
         Text            =   "Net Operações"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   28
         Text            =   "NR_SEQU_OPER_ATIV"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   29
         Text            =   "Canal Venda"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   30
         Text            =   "Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   31
         Text            =   "Sub-Produto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   32
         Text            =   "Código Operação Estruturada"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(34) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   33
         Text            =   "Num. Reproc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(35) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   34
         Text            =   "Data/Hora Reproc"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(36) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   35
         Text            =   "Código Lote"
         Object.Width           =   0
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
            Picture         =   "frmConsultaContaCorrente.frx":0000
            Key             =   "aplicarfiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":0112
            Key             =   "showfiltro"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":0224
            Key             =   "showtreeview"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":0576
            Key             =   "showlist"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":08C8
            Key             =   "showdetail"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":0C1A
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaContaCorrente.frx":0F6C
            Key             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvrContaCorrrenteStatus 
      Height          =   8595
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   15161
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   365
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image imgDummyV 
      Height          =   8745
      Left            =   4605
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "frmConsultaContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Objeto responsavel pela consulta de registro do conta corrente,
' através da camada de controle de caso de uso MIU.
'
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

Private xmlRetLeitura                       As MSXML2.DOMDocument40

'Constantes de Visão da Lista
Private Const VIS_POR_STATUS                As Integer = 1

'Constantes de Configuração de Colunas
Private Const COL_DATA_OPERACAO             As Integer = 0
Private Const COL_LANCAMENTO                As Integer = 1
Private Const COL_SISTEMA                   As Integer = 2
Private Const COL_EMPRESA                   As Integer = 3
Private Const COL_NUMERO_COMANDO            As Integer = 4
Private Const COL_VEICULO_LEGAL             As Integer = 5
Private Const COL_SITUACAO                  As Integer = 6
Private Const COL_TIPO_OPERACAO             As Integer = 7
Private Const COL_LOCA_LIQU                 As Integer = 8
Private Const COL_BANCO                     As Integer = 9
Private Const COL_AGENCIA                   As Integer = 10
Private Const COL_CONTA_CORRENTE            As Integer = 11
Private Const COL_VALOR                     As Integer = 12
Private Const COL_TIPO_MOVIMENTO            As Integer = 13
Private Const COL_TIPO_LANCAMENTO           As Integer = 14
Private Const COL_SUB_TIPO_ATIVO            As Integer = 15
Private Const COL_FINALIDADE_TED            As Integer = 16
Private Const COL_CONTA_CONTABIL_DEB        As Integer = 17
Private Const COL_CONTA_CONTABIL_CRED       As Integer = 18
Private Const COL_COD_HIST_CONTABIL         As Integer = 19
Private Const COL_DES_HIST_CONTABIL         As Integer = 20
Private Const COL_TIPO_BACKOFFICE           As Integer = 21
Private Const COL_COD_VEIC_LEGA             As Integer = 22
Private Const COL_COD_SITUACAO              As Integer = 23
Private Const COL_COD_TIPO_OPER             As Integer = 24
Private Const COL_COD_BANCO                 As Integer = 25
Private Const COL_COD_LOCA_LIQU             As Integer = 26
Private Const COL_COND_NET_OPERACOES        As Integer = 27
Private Const COL_NR_SEQU_OPER_ATIV         As Integer = 28
Private Const COL_CANAL_VENDA               As Integer = 29
Private Const COL_CONTRAPARTE               As Integer = 30
Private Const COL_PRODUTO                   As Integer = 31
Private Const COL_OPERACAO_ESTRUTURADA      As Integer = 32
Private Const COL_NUM_REPROC                As Integer = 33
Private Const COL_DATA_HORA_REPROC          As Integer = 34
Private Const COL_COD_LOTE                  As Integer = 35

Private Const KEY_EMPRESA                   As Integer = 1
Private Const KEY_DATA_OPERACAO             As Integer = 2
Private Const KEY_TIPO_OPERACAO             As Integer = 3
Private Const KEY_VEICULO_LEGAL             As Integer = 4
Private Const KEY_LOCA_LIQU                 As Integer = 5
Private Const KEY_BANCO                     As Integer = 6
Private Const KEY_AGENCIA                   As Integer = 7
Private Const KEY_CONTA_CORRENTE            As Integer = 8
Private Const KEY_CO_ULTI_SITU_PROC         As Integer = 9
Private Const KEY_CD_LOTE                   As Integer = 10

Private fblnDummyV                          As Boolean

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private lngIndexClassifList                 As Long
Private xmlFinalidadeTED                    As MSXML2.DOMDocument40

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        'intRefresh = intRefresh + 1
        'If intRefresh > 1 Then
        '    intRefresh = 0
        '    Exit Sub
        'End If
    
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

    With Me

        .tvrContaCorrrenteStatus.Height = .ScaleHeight - .tlbButtons.Height - 50 - .tvrContaCorrrenteStatus.Top
        .tvrContaCorrrenteStatus.Width = IIf(.imgDummyV.Visible, .imgDummyV.Left, .ScaleWidth - 280)
        
        .tlbButtons.Top = .ScaleHeight - .tlbButtons.Height
        
        .imgDummyV.Top = 0
        .imgDummyV.Height = .tvrContaCorrrenteStatus.Height + 100
        
        .lvwContaCorrente.Left = IIf(.imgDummyV.Visible, .imgDummyV.Left + 100, 0)
        .lvwContaCorrente.Top = .tvrContaCorrrenteStatus.Top
        .lvwContaCorrente.Height = .tvrContaCorrrenteStatus.Height
        .lvwContaCorrente.Width = .ScaleWidth - .lvwContaCorrente.Left
    End With

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
        
        .tvrContaCorrrenteStatus.Width = .imgDummyV.Left
        
        .lvwContaCorrente.Left = .imgDummyV.Left + 100
        .lvwContaCorrente.Width = .Width - (.imgDummyV.Left + 330)
    End With
    
    On Error GoTo 0

Exit Sub
ErrorHandler:

   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - imgDummyV_MouseMove"

End Sub

Private Sub imgDummyV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    fblnDummyV = False
End Sub

Private Sub lvwContaCorrente_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub lvwContaCorrente_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    
On Error GoTo ErrorHandler
    
    Call fgClassificarListview(Me.lvwContaCorrente, ColumnHeader.Index)
    lngIndexClassifList = ColumnHeader.Index

Exit Sub
ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - lvwContaCorrente_ColumnClick", Me.Caption

End Sub

Private Sub lvwContaCorrente_DblClick()

Dim vntSequenciaOperacao                    As Variant
Dim lngTipoLancamentoITGR                   As Long
Dim lngCodigoEmpresa                        As Long
Dim intSequenciaLancamento                  As Integer

On Error GoTo ErrorHandler

    If Not lvwContaCorrente.SelectedItem Is Nothing Then
        fgCursor True
        
        vntSequenciaOperacao = Split(lvwContaCorrente.SelectedItem.Key, "|")(1)
        lngTipoLancamentoITGR = Split(lvwContaCorrente.SelectedItem.Key, "|")(2)
        intSequenciaLancamento = Split(lvwContaCorrente.SelectedItem.Key, "|")(5)
        
        'Negativo quando base historica
        If Split(lvwContaCorrente.SelectedItem.Key, "|")(4) = "A8HIST" Then
            vntSequenciaOperacao = vntSequenciaOperacao * -1
        End If
        
        lngCodigoEmpresa = fgObterCodigoCombo(lvwContaCorrente.SelectedItem.SubItems(3))
        
        With frmHistLancamentoCC
            .lngCodigoEmpresa = lngCodigoEmpresa
            .vntSequenciaOperacao = vntSequenciaOperacao
            .lngTipoLancamentoITGR = lngTipoLancamentoITGR
            .intSequenciaLancamento = intSequenciaLancamento
            .strNetOperacoes = lvwContaCorrente.SelectedItem.SubItems(COL_COND_NET_OPERACOES)
            .Show vbModal
        End With
        fgCursor False
    End If

Exit Sub
ErrorHandler:
   fgCursor False
   mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - lvwContaCorrente_DblClick"
End Sub

Private Sub objFiltro_AplicarFiltro(xmlDocFiltros As String, strTituloTableCombo As String)

Dim strSelecaoVisual                        As String
Dim strSelecaoFiltro                        As String

On Error GoTo ErrorHandler

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
        
        strSelecaoVisual = flObterSelecaoTreeview(tvrContaCorrrenteStatus, True)
        strSelecaoFiltro = flObterSelecaoTreeview(tvrContaCorrrenteStatus)
        
        If Not flCarregarTreeViewSituStatus(tvrContaCorrrenteStatus, _
                                            IIf(blnUtilizaFiltro, strFiltroXML, vbNullString), _
                                            "Todas Situações") Then Exit Sub
        
        If Trim(strSelecaoFiltro) <> "" Then
            fgLockWindow tvrContaCorrrenteStatus.hwnd
            Call flRetornarSelecaoAnterior(tvrContaCorrrenteStatus, strSelecaoVisual)
            fgLockWindow 0
            
            Call flCarregarLista(strSelecaoFiltro, VIS_POR_STATUS, _
                                    IIf(blnUtilizaFiltro, strFiltroXML, vbNullString))
        Else
            Call flLimparLista
        End If
        
    End If
    fgCursor
    
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
    
    Call flCarregarTreeViewSituStatus(tvrContaCorrrenteStatus, vbNullString, "Todas Situações", False)
    
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
                objFiltro.TipoFiltro = enumTipoFiltroA8.frmConsultaContaCorrente
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

Private Sub tvrContaCorrrenteStatus_BeforeLabelEdit(Cancel As Integer)
    Cancel = True
End Sub

Private Sub tvrContaCorrrenteStatus_NodeCheck(ByVal Node As MSComctlLib.Node)

Dim strSelecao                              As String
    
On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    
    Node.Selected = True
    Call flMarcarNodes(tvrContaCorrrenteStatus, (Node.children > 0), Node.Checked)
    
    strSelecao = flObterSelecaoTreeview(tvrContaCorrrenteStatus)
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

'Carregar o tree view com os status de conta corrente

Private Function flCarregarTreeViewSituStatus(ByVal ptreTreeView As TreeView, _
                                              ByVal pstrFiltroXML As String, _
                                     Optional ByVal pstrNomeRoot As String, _
                                     Optional ByVal pblnMostrarQuantidade As Boolean = True) As Boolean

#If EnableSoap = 1 Then
    Dim objContaCorrente    As MSSOAPLib30.SoapClient30
#Else
    Dim objContaCorrente    As A8MIU.clsContaCorrente
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
    
    Set objContaCorrente = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
    strCargaStatus = objContaCorrente.ObterLancamentosPorStatus(strFiltros, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objContaCorrente = Nothing
    
    Set xmlDocument = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDocument.loadXML(strCargaStatus) Then
        Call fgErroLoadXML(xmlDocument, App.EXEName, TypeName(Me), "flCarregarTreeViewSituStatus")
    End If
    
    With ptreTreeView
    
        .Nodes.Clear
        
        If pstrNomeRoot <> "" Then
           .Nodes.Add , , "root", pstrNomeRoot
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_SituacaoLancamentos/*")
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
            For Each objDomNode In xmlDocument.documentElement.selectNodes("//Repeat_SituacaoLancamentos/*")
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
    
    If ptreTreeView.Nodes.Count > 0 Then
       ptreTreeView.Nodes(1).EnsureVisible
    End If
    
    Set xmlDocument = Nothing
    Set objDomNode = Nothing
    
    flCarregarTreeViewSituStatus = True
    
    fgLockWindow 0

Exit Function
ErrorHandler:
    
    fgLockWindow 0
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarTreeViewSituStatus", 0
    
End Function

' Carrega os lançamento de conta corrente e preencher a interface com os mesmos,
' através da classe controladora de caso de uso MIU, método A8MIU.clsContaCorrente.ObterDetalheLancamento
Private Sub flCarregarLista(ByVal pstrSelecaoFiltro As String, _
                            ByVal pintTipoFiltro As Integer, _
                            ByVal pstrFiltroXML As String)

#If EnableSoap = 1 Then
    Dim objOperacao                         As MSSOAPLib30.SoapClient30
#Else
    Dim objOperacao                         As A8MIU.clsContaCorrente
#End If

Dim strRetLeitura                           As String
Dim xmlDomLeitura                           As MSXML2.DOMDocument40
Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim xmlDomNodeTED                           As MSXML2.IXMLDOMNode
                         
Dim strTagGrupoFiltro                       As String
Dim strTagFiltro                            As String
Dim strDataOperacao                         As String

Dim lngCont                                 As Long
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

Dim strListItemKey                          As String
Dim strListItemKey2                         As String
Dim blnCorretoras                           As Boolean
Dim dblValorOperacao                        As Double
Dim objListItem                             As MSComctlLib.ListItem
Dim intDebitoCredito                        As Integer
Dim blnExisteGrid                           As Boolean

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
    strTagGrupoFiltro = "Grupo_Status"
    strTagFiltro = "Status"
    
    '>>> Formata XML Filtro padrão -----------------------------------------------------------------
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    If Not xmlDomFiltros.loadXML(pstrFiltroXML) Then
        Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    End If
    
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", strTagGrupoFiltro, "")
    
    'Captura o filtro cumulativo
    For lngCont = LBound(Split(pstrSelecaoFiltro, ";")) To UBound(Split(pstrSelecaoFiltro, ";"))
        Call fgAppendNode(xmlDomFiltros, strTagGrupoFiltro, _
                                         strTagFiltro, Split(pstrSelecaoFiltro, ";")(lngCont))
    Next
    '>>> -------------------------------------------------------------------------------------------

    Set objOperacao = fgCriarObjetoMIU("A8MIU.clsContaCorrente")
    strRetLeitura = objOperacao.ObterConsultaLancamento(xmlDomFiltros.xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If
    
    Set objOperacao = Nothing
    
    If strRetLeitura <> vbNullString Then
        Set xmlDomLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Not xmlDomLeitura.loadXML(strRetLeitura) Then
            Call fgErroLoadXML(xmlDomLeitura, App.EXEName, TypeName(Me), "flCarregaLista")
        End If
        
        Set xmlRetLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        Call xmlRetLeitura.loadXML(strRetLeitura)
        
        For Each objDomNode In xmlDomLeitura.selectNodes("Repeat_DetalheLancamento/*")
            
            blnCorretoras = False
            blnExisteGrid = False
            If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                
                strListItemKey = "|" & objDomNode.selectSingleNode("CO_EMPR").Text & _
                                 "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("TP_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_VEIC_LEGA").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_BANC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_AGEN").Text & _
                                 "|" & objDomNode.selectSingleNode("NU_CC").Text & _
                                 "|" & objDomNode.selectSingleNode("CO_ULTI_SITU_PROC").Text & _
                                 "|" & objDomNode.selectSingleNode("CD_LOTE").Text
                
                dblValorOperacao = flNetOperacoes(strListItemKey)
                
                For Each objListItem In lvwContaCorrente.ListItems
                    strListItemKey2 = "|" & objDomNode.selectSingleNode("CO_EMPR").Text & _
                                      "|" & fgDt_To_Xml(objListItem.Text) & _
                                      "|" & objListItem.SubItems(COL_COD_TIPO_OPER) & _
                                      "|" & objListItem.SubItems(COL_COD_VEIC_LEGA) & _
                                      "|" & objListItem.SubItems(COL_COD_LOCA_LIQU) & _
                                      "|" & objListItem.SubItems(COL_COD_BANCO) & _
                                      "|" & objListItem.SubItems(COL_AGENCIA) & _
                                      "|" & objListItem.SubItems(COL_CONTA_CORRENTE) & _
                                      "|" & objListItem.SubItems(COL_COD_SITUACAO) & _
                                      "|" & objListItem.SubItems(COL_COD_LOTE)
                    
                    If strListItemKey = strListItemKey2 Then
                        blnCorretoras = True
                        objListItem.SubItems(COL_COND_NET_OPERACOES) = objListItem.SubItems(COL_COND_NET_OPERACOES) & _
                                                                       "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                        
                        Call flSubtrairQuantidadeItens(objListItem.SubItems(COL_SITUACAO))
                        blnExisteGrid = True
                        Exit For
                    End If
                Next
                
                If Not blnExisteGrid Then
                    If dblValorOperacao > 0 Then
                        intDebitoCredito = enumTipoDebitoCredito.Credito
                        
'                        If Trim(objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text) = "Crédito" Then
'                            blnCorretoras = False
'                        Else
'                            blnCorretoras = True
'                        End If
                        
                    Else
                        intDebitoCredito = enumTipoDebitoCredito.Debito
                        
'                        If objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text = "Débito" Then
'                            blnCorretoras = False
'                        Else
'                            blnCorretoras = True
'                        End If
                        
                    End If
                End If
                
            End If
            
            If Not blnCorretoras Then
                strListItemKey = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text & _
                                 "|" & objDomNode.selectSingleNode("TP_LANC_ITGR").Text & _
                                 "|" & objDomNode.selectSingleNode("DT_OPER").Text & _
                                 "|" & objDomNode.selectSingleNode("OWNER").Text & _
                                 "|" & objDomNode.selectSingleNode("NR_SEQU_LANC").Text
                
                
                If Not fgExisteItemLvw(Me.lvwContaCorrente, strListItemKey) Then
                    With lvwContaCorrente.ListItems.Add(, strListItemKey)
                        
                        'Empresa
                        If objDomNode.selectSingleNode("CO_EMPR").Text <> vbNullString And _
                           Val(objDomNode.selectSingleNode("CO_EMPR").Text) <> 0 Then
                            'Obtem a descrição da Empresa via QUERY XML
                            .SubItems(COL_EMPRESA) = _
                                    objDomNode.selectSingleNode("CO_EMPR").Text & " - " & xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_Empresa/Grupo_Empresa[CO_EMPR='" & _
                                    objDomNode.selectSingleNode("CO_EMPR").Text & "']/NO_REDU_EMPR").Text
                        End If
                        
                        'Tipo Lançamento
                        .SubItems(COL_LANCAMENTO) = objDomNode.selectSingleNode("DE_LANC_ITGR").Text
                        
                        'Sistema
                        .SubItems(COL_SISTEMA) = objDomNode.selectSingleNode("SG_SIST").Text & " - " & objDomNode.selectSingleNode("NO_SIST").Text
                        
                        'Data Lançamento CC
                        If objDomNode.selectSingleNode("DH_SITU_LANC_CC").Text <> gstrDataVazia Then
                            .Text = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_SITU_LANC_CC").Text)
                        Else
                            .Text = fgDtXML_To_Date(objDomNode.selectSingleNode("DT_OPER").Text)
                        End If
                        
                        'Número do Comando
                        .SubItems(COL_NUMERO_COMANDO) = objDomNode.selectSingleNode("NU_COMD_OPER").Text
                        
                        'Veiculo Legal
                        .SubItems(COL_VEICULO_LEGAL) = objDomNode.selectSingleNode("NO_VEIC_LEGA").Text
                        
                        'Situação
                        .SubItems(COL_SITUACAO) = objDomNode.selectSingleNode("DE_SITU_PROC").Text
                        
                        'Tipo de Operação
                        .SubItems(COL_TIPO_OPERACAO) = objDomNode.selectSingleNode("NO_TIPO_OPER").Text
                        
                        'Local de Liquidação
                        If objDomNode.selectSingleNode("CO_LOCA_LIQU").Text <> vbNullString And _
                           Val(objDomNode.selectSingleNode("CO_LOCA_LIQU").Text) <> 0 Then
                            
                            If Not xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                   objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU") Is Nothing Then
                                        
                                'Obtem a descrição do Local de Liquidação via QUERY XML
                                .SubItems(COL_LOCA_LIQU) = _
                                    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Dados/Repeat_LocalLiquidacao/Grupo_LocalLiquidacao[CO_LOCA_LIQU='" & _
                                        objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "']/SG_LOCA_LIQU").Text
                        
                            Else
                                
                                vntCodErro = 5
                                vntMensagemErro = "Usuário não possui acesso ao Local de Liquidação " & _
                                                  objDomNode.selectSingleNode("CO_LOCA_LIQU").Text & "."
                                GoTo ErrorHandler
                                
                            End If
                        
                        End If
        
                        'Banco
                        .SubItems(COL_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                        
                        'Agência
                        .SubItems(COL_AGENCIA) = objDomNode.selectSingleNode("CO_AGEN").Text
                        
                        'Número C/C
                        .SubItems(COL_CONTA_CORRENTE) = objDomNode.selectSingleNode("NU_CC").Text
                        
                        If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemBUS.OperacoesCorretoras Then
                            'Valor do Lançamento
                            .SubItems(COL_VALOR) = fgVlrXml_To_Interface(fgVlr_To_Xml(Abs(dblValorOperacao)))
                            
                            'Tipo Movto.
                            .SubItems(COL_TIPO_MOVIMENTO) = IIf(intDebitoCredito = enumTipoDebitoCredito.Debito, "Débito", "Crédito")
    
                        Else
                            'Valor do Lançamento
                            .SubItems(COL_VALOR) = fgVlrXml_To_Interface(objDomNode.selectSingleNode("VA_LANC_CC").Text)
                            
                            'Tipo Movto.
                            .SubItems(COL_TIPO_MOVIMENTO) = objDomNode.selectSingleNode("IN_LANC_DEBT_CRED").Text
                        End If
                                            
                        'Tipo de Lançamento
                        .SubItems(COL_TIPO_LANCAMENTO) = IIf(Val(objDomNode.selectSingleNode("TP_LANC_ITGR").Text) = enumTipoLancamentoIntegracao.Estorno, "Estorno", "Normal")
                        
                        If objDomNode.selectSingleNode("TP_MESG_RECB_INTE").Text = enumTipoMensagemLQS.EnvioTEDClientes Then
                        
                            For Each xmlDomNodeTED In xmlFinalidadeTED.selectNodes("//Repeat_DominioAtributo/*")
                                If objDomNode.selectSingleNode("CD_FIND_TED").Text = "0" Or objDomNode.selectSingleNode("CD_FIND_TED").Text = "" Then .SubItems(COL_FINALIDADE_TED) = "0": Exit For
                                If xmlDomNodeTED.selectSingleNode("CO_DOMI").Text = objDomNode.selectSingleNode("CD_FIND_TED").Text Then
                                    .SubItems(COL_FINALIDADE_TED) = IIf(xmlDomNodeTED.selectSingleNode("DE_DOMI").Text = "", "", xmlDomNodeTED.selectSingleNode("CO_DOMI").Text & "-" & xmlDomNodeTED.selectSingleNode("DE_DOMI").Text)
                                    Exit For
                                End If
                            Next
                            .SubItems(COL_SUB_TIPO_ATIVO) = "0"
                        Else
                            'Sub-tipo Ativo
                            .SubItems(COL_SUB_TIPO_ATIVO) = objDomNode.selectSingleNode("CO_SUB_TIPO_ATIV").Text
                        End If
                        'Conta Contábil Débito
                        .SubItems(COL_CONTA_CONTABIL_DEB) = objDomNode.selectSingleNode("CO_CNTA_DEBT").Text
                        
                        'Conta Contábil Crédito
                        .SubItems(COL_CONTA_CONTABIL_CRED) = objDomNode.selectSingleNode("CO_CNTA_CRED").Text
                        
                        'Código Histórico Contábil
                        .SubItems(COL_COD_HIST_CONTABIL) = objDomNode.selectSingleNode("CO_HIST_CNTA_CNTB").Text
                        
                        'Descriçao do Histórico contábil
                        .SubItems(COL_DES_HIST_CONTABIL) = objDomNode.selectSingleNode("DE_HIST_CNTA_CNTB").Text
                        
                        'Tipo de BackOffice
                        .SubItems(COL_TIPO_BACKOFFICE) = objDomNode.selectSingleNode("TP_BKOF").Text
                        If Trim$(objDomNode.selectSingleNode("DE_BKOF").Text) <> vbNullString Then
                            .SubItems(COL_TIPO_BACKOFFICE) = .SubItems(COL_TIPO_BACKOFFICE) & " - " & objDomNode.selectSingleNode("DE_BKOF").Text
                        End If
                        
                        .SubItems(COL_COND_NET_OPERACOES) = "|" & objDomNode.selectSingleNode("NU_SEQU_OPER_ATIV").Text
                        .SubItems(COL_COD_VEIC_LEGA) = objDomNode.selectSingleNode("CO_VEIC_LEGA").Text
                        .SubItems(COL_COD_SITUACAO) = objDomNode.selectSingleNode("CO_SITU_PROC").Text
                        .SubItems(COL_COD_TIPO_OPER) = objDomNode.selectSingleNode("TP_OPER").Text
                        .SubItems(COL_COD_LOCA_LIQU) = objDomNode.selectSingleNode("CO_LOCA_LIQU").Text
                        .SubItems(COL_COD_BANCO) = objDomNode.selectSingleNode("CO_BANC").Text
                        .SubItems(COL_NR_SEQU_OPER_ATIV) = objDomNode.selectSingleNode("NR_SEQU_OPER_ATIV").Text
                        .SubItems(COL_CANAL_VENDA) = fgDescricaoCanalVenda(objDomNode.selectSingleNode("TP_CNAL_VEND").Text)
                        
                        If Not objDomNode.selectSingleNode("NO_CNPT") Is Nothing Then
                            .SubItems(COL_CONTRAPARTE) = objDomNode.selectSingleNode("NO_CNPT").Text
                        Else
                            .SubItems(COL_CONTRAPARTE) = ""
                        End If
                        
                        If Not objDomNode.selectSingleNode("CD_SUB_PROD") Is Nothing Then
                            .SubItems(COL_PRODUTO) = objDomNode.selectSingleNode("CD_SUB_PROD").Text
                        Else
                            .SubItems(COL_PRODUTO) = ""
                        End If
                        
                        If Not objDomNode.selectSingleNode("CD_OPER_ETRT") Is Nothing Then
                            .SubItems(COL_OPERACAO_ESTRUTURADA) = objDomNode.selectSingleNode("CD_OPER_ETRT").Text
                        End If
                        
                        .SubItems(COL_NUM_REPROC) = objDomNode.selectSingleNode("NU_TENT_REPR_CC").Text
                        
                        If Trim(objDomNode.selectSingleNode("DH_ULTI_REPR_CC").Text) <> "00:00:00" Then
                            .SubItems(COL_DATA_HORA_REPROC) = fgDtHrXML_To_Interface(objDomNode.selectSingleNode("DH_ULTI_REPR_CC").Text)
                        End If
                        
                        'Código do Lote
                        If Not objDomNode.selectSingleNode("CD_LOTE") Is Nothing Then
                            .SubItems(COL_COD_LOTE) = objDomNode.selectSingleNode("CD_LOTE").Text
                        Else
                            .SubItems(COL_COD_LOTE) = ""
                        End If
                        
                    End With
                End If
            End If
        Next
    End If
    
    Call fgClassificarListview(Me.lvwContaCorrente, lngIndexClassifList, True)
    
    Set xmlDomFiltros = Nothing
    Set xmlDomLeitura = Nothing

    fgLockWindow 0
    
Exit Sub
ErrorHandler:
    
    fgLockWindow 0
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0
    
End Sub

Private Sub flLimparLista()
    Me.lvwContaCorrente.ListItems.Clear
End Sub

'Captura todos os nodes selecionados (Checked), exceto o node RAIZ e,
'retorna uma STRING, com separador ";", a ser decomposta na função SPLIT.

Private Function flObterSelecaoTreeview(ByVal treTreeView As TreeView, _
                               Optional ByVal pblnConsideraGrupo As Boolean = False) As String

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

'Marca ou desmarca (Check) nodes referentes ao TreeView informado.
'
'Se o Node Root for informado, transfere seu status para todo TreeView,
'se não, reflete o status do Node Child no Node Root.
'
'Obs.:  Utiliza a API LockWindowUpdate, para que o evento do TreeView
'       << _NodeCheck >> não seja disparado a cada iteração.
Private Sub flMarcarNodes(ByVal treTreeView As TreeView, _
                          ByVal blnNodeRoot As Boolean, _
                          ByVal blnMarcar As Boolean)

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
                tvrContaCorrrenteStatus.Visible = False
                lvwContaCorrente.Visible = False
            
           Case "1"
                imgDummyV.Visible = False
                tvrContaCorrrenteStatus.Visible = True
                lvwContaCorrente.Visible = False
            
           Case "2"
                imgDummyV.Visible = False
                tvrContaCorrrenteStatus.Visible = False
                lvwContaCorrente.Visible = True
                
           Case "12"
                imgDummyV.Visible = True
                tvrContaCorrrenteStatus.Visible = True
                lvwContaCorrente.Visible = True
                
    End Select
    
    Call Form_Resize

Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flArranjarJanelasExibicao", 0
    
End Sub

' Carrega as propriedades necessárias a interface frmCompromissadaGenerica, através da
' camada de controle de caso de uso, método A8MIU.clsMIU.ObterMapaNavegacao
Private Sub flInicializar()

#If EnableSoap = 1 Then
    Dim objMIU              As MSSOAPLib30.SoapClient30
    Dim objMensagem         As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU              As A8MIU.clsMIU
    Dim objMensagem         As A8MIU.clsMensagem
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
    
    Set objMensagem = fgCriarObjetoMIU("A8MIU.clsMensagem")
    Set xmlFinalidadeTED = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlFinalidadeTED.loadXML(objMensagem.ObterDominioSPB("FinlddCli", vntCodErro, vntMensagemErro))
    
    Set objMIU = Nothing
    Set objMensagem = Nothing
    Exit Sub

ErrorHandler:
    Set objMIU = Nothing
    Set xmlMapaNavegacao = Nothing
    Set xmlFinalidadeTED = Nothing
    
    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, TypeName(Me), "flInicializar", 0

End Sub

'Retornar a seleção efetuada anteriormente

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

'Montar o xml de filtro para pesquisa pelo perfil do usuário

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

'Calcula o Net da operações
Public Function flNetOperacoes(ByVal strItemKey As String)
    
Dim strExpression                   As String
Dim vntValor                        As Variant
    
    vntValor = 0
    
    strExpression = flMontarCalculoNetOperacoes(strItemKey)
    vntValor = vntValor + Val(fgFuncaoXPath(xmlRetLeitura, strExpression))
    
    flNetOperacoes = vntValor

End Function

'Monta uma expressão XPath para a somatória dos valores de operações
Public Function flMontarCalculoNetOperacoes(ByVal strItemKey As String)
                
Dim strDebito                               As String
Dim strCredito                              As String
    
    On Error GoTo ErrorHandler
    
'    strDebito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Mid(Split(strItemKey, "|")(KEY_DATA_OPERACAO), 1, 8) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../NR_SEQU_OPER_ATIV='" & Split(strItemKey, "|")(KEY_NR_SEQU_OPER_ATIV) & "' "
    
 '   strCredito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Mid(Split(strItemKey, "|")(KEY_DATA_OPERACAO), 1, 8) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../NR_SEQU_OPER_ATIV='" & Split(strItemKey, "|")(KEY_NR_SEQU_OPER_ATIV) & "' "

    strDebito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Credito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Mid(Split(strItemKey, "|")(KEY_DATA_OPERACAO), 1, 8) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_CD_LOTE) & "' "
    
    strCredito = "sum(//VA_LANC_CC_VLRXML[../CO_IN_LANC_DEBT_CRED='" & enumTipoDebitoCredito.Debito & "' " & _
                                     " and ../CO_EMPR='" & Split(strItemKey, "|")(KEY_EMPRESA) & "' " & _
                                     " and ../DT_OPER='" & Mid(Split(strItemKey, "|")(KEY_DATA_OPERACAO), 1, 8) & "' " & _
                                     " and ../TP_OPER='" & Split(strItemKey, "|")(KEY_TIPO_OPERACAO) & "' " & _
                                     " and ../CO_VEIC_LEGA='" & Split(strItemKey, "|")(KEY_VEICULO_LEGAL) & "' " & _
                                     " and ../CO_LOCA_LIQU='" & Split(strItemKey, "|")(KEY_LOCA_LIQU) & "' " & _
                                     " and ../CO_BANC='" & Split(strItemKey, "|")(KEY_BANCO) & "' " & _
                                     " and ../CO_AGEN='" & Split(strItemKey, "|")(KEY_AGENCIA) & "' " & _
                                     " and ../NU_CC='" & Split(strItemKey, "|")(KEY_CONTA_CORRENTE) & "' " & _
                                     " and ../CO_ULTI_SITU_PROC='" & Split(strItemKey, "|")(KEY_CO_ULTI_SITU_PROC) & "' " & _
                                     " and ../CD_LOTE='" & Split(strItemKey, "|")(KEY_CD_LOTE) & "' "

    strDebito = strDebito & "]) - "
    strCredito = strCredito & "]) "
    
    flMontarCalculoNetOperacoes = strDebito & strCredito
    
    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flMontarCalculoNetOperacoes", 0

End Function

Private Sub flSubtrairQuantidadeItens(ByVal strStatus As String)

Dim objNode                                 As MSComctlLib.Node
Dim intQuantidade                           As Integer
Dim intPosicaoQtd                           As Integer

    On Error GoTo ErrorHandler
    
    For Each objNode In Me.tvrContaCorrrenteStatus.Nodes
        With objNode
            intPosicaoQtd = InStr(1, .Text, "(") + 1
            If intPosicaoQtd > 1 Then
                If Left$(.Text, intPosicaoQtd - 3) = strStatus Or InStr(1, .Text, "Todas Situações") <> 0 Then
                    intQuantidade = Val(Left$(Mid$(.Text, intPosicaoQtd), _
                                              Len(Mid$(.Text, intPosicaoQtd)) - 1)) - 1
                                              
                    If InStr(1, .Text, "Todas Situações") <> 0 Then
                        .Text = "Todas Situações" & " (" & intQuantidade & ")"
                    Else
                        .Text = strStatus & " (" & intQuantidade & ")"
                    End If
                End If
            End If
        End With
    Next
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flSubtrairQuantidadeItens", 0

End Sub
