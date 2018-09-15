VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Begin VB.Form frmAtributo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Atributos de Mensagens"
   ClientHeight    =   7305
   ClientLeft      =   4635
   ClientTop       =   1980
   ClientWidth     =   7890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7890
   Begin VB.Frame fraDetalhe 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   4515
      Left            =   15
      TabIndex        =   12
      Top             =   2325
      Width           =   7815
      Begin RichTextLib.RichTextBox txtDescricao 
         Height          =   825
         Left            =   120
         TabIndex        =   3
         Top             =   1710
         Width           =   7530
         _ExtentX        =   13282
         _ExtentY        =   1455
         _Version        =   393217
         ScrollBars      =   2
         MaxLength       =   255
         TextRTF         =   $"frmAtributo.frx":0000
      End
      Begin VB.Frame Frame3 
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
         Height          =   1770
         Left            =   3600
         TabIndex        =   19
         Top             =   2580
         Width           =   4125
         Begin MSComCtl2.DTPicker dtpDataInicioVigencia 
            Height          =   330
            Left            =   240
            TabIndex        =   8
            Top             =   525
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   582
            _Version        =   393216
            Format          =   58327041
            CurrentDate     =   37816
         End
         Begin MSComCtl2.DTPicker dtpDataFimVigencia 
            Height          =   330
            Left            =   240
            TabIndex        =   9
            Top             =   1260
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   58327041
            CurrentDate     =   37816
         End
         Begin VB.Label Label7 
            Caption         =   "Fim"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   21
            Top             =   1005
            Width           =   900
         End
         Begin VB.Label Label6 
            Caption         =   "Início"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   225
            TabIndex        =   20
            Top             =   285
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Dado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1770
         Left            =   90
         TabIndex        =   16
         Top             =   2580
         Width           =   3465
         Begin VB.CheckBox chkIndicadorValorNegativo 
            Caption         =   "Aceita Valor Negativo"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   1440
            Width           =   1905
         End
         Begin VB.OptionButton optTipoDado 
            Caption         =   "&Alfanumérico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   4
            Top             =   300
            Width           =   1485
         End
         Begin VB.OptionButton optTipoDado 
            Caption         =   "&Numérico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   1830
            TabIndex        =   5
            Top             =   300
            Width           =   1485
         End
         Begin NumBox.Number numDecimais 
            Height          =   285
            Left            =   1830
            TabIndex        =   7
            Top             =   990
            Width           =   345
            _ExtentX        =   609
            _ExtentY        =   503
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
         Begin NumBox.Number numTamanho 
            Height          =   285
            Left            =   165
            TabIndex        =   6
            Top             =   1005
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   503
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
         Begin VB.Label Label5 
            Caption         =   "Decimais"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1830
            TabIndex        =   18
            Top             =   750
            Width           =   780
         End
         Begin VB.Label Label2 
            Caption         =   "Tamanho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   165
            TabIndex        =   17
            Top             =   765
            Width           =   900
         End
      End
      Begin VB.TextBox txtNomeLogico 
         Height          =   315
         Left            =   120
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1050
         Width           =   7530
      End
      Begin VB.TextBox txtNomeFisico 
         Height          =   315
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   375
         Width           =   7530
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1470
         Width           =   2070
      End
      Begin VB.Label Label3 
         Caption         =   "Nome Lógico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   14
         Top             =   810
         Width           =   2145
      End
      Begin VB.Label Label4 
         Caption         =   "Nome Físico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   150
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2340
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      Begin MSComctlLib.ListView lstAtributo 
         Height          =   2070
         Left            =   90
         TabIndex        =   11
         Top             =   165
         Width           =   7650
         _ExtentX        =   13494
         _ExtentY        =   3651
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nome Fisico"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nome Lógico"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo Dado"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tamanho"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Decimais"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Data Início"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Data Fim"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   4275
      TabIndex        =   10
      Top             =   6915
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
            Caption         =   "Limpar"
            Key             =   "Limpar"
            ImageKey        =   "Limpar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Excluir"
            Key             =   "Excluir"
            ImageKey        =   "Excluir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            ImageKey        =   "Salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            ImageKey        =   "Sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   60
      Top             =   6885
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0082
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0194
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":04AE
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0800
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0912
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0C2C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":0F46
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAtributo.frx":1260
            Key             =   "Limpar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAtributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável pelo cadastramento e manutenção de atributos de mensagens.

Option Explicit

Private xmlAtributoMensagem                 As MSXML2.DOMDocument40

Private strOperacao                         As String
Private strKeyItemSelected                  As String

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

'Posicionar item no listview de atributos.
Private Sub flPosicionaItemListView()

Dim objListItem                             As ListItem
Dim blnEncontrou                            As Boolean

    On Error GoTo ErrorHandler

    If lstAtributo.ListItems.Count = 0 Then Exit Sub
    If Len(strKeyItemSelected) = 0 Then Exit Sub
    
    blnEncontrou = False
    For Each objListItem In lstAtributo.ListItems
        If objListItem.Key = strKeyItemSelected Then
           objListItem.Selected = True
           lstAtributo_ItemClick objListItem
           lstAtributo.ListItems(strKeyItemSelected).EnsureVisible
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

'Limpar campos do formulário.
Private Sub flLimpaCampos()

On Error GoTo ErrorHandler
        
    strOperacao = "Incluir"
    
    txtNomeFisico.Text = ""
    txtNomeLogico.Text = ""
    txtDescricao.Text = ""
    
    optTipoDado(1).Value = True
    
    numTamanho.Valor = 0
    numDecimais.Valor = 0
        
    dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataInicioVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    
    dtpDataFimVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
    dtpDataFimVigencia.Value = Null
    
    lstAtributo.Sorted = False
    txtNomeFisico.Enabled = True
    dtpDataInicioVigencia.Enabled = True
    
    chkIndicadorValorNegativo.Value = vbUnchecked
    
    tlbCadastro.Buttons("Excluir").Enabled = False
    
    Exit Sub
    
ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

'Salvar as informações correntes do atributo.
Private Sub flSalvar()

Dim strRetorno                              As String
Dim xmlExecucao                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
      
    strRetorno = flValidarCampos()
        
    If strRetorno <> "" Then
        If strRetorno = "Atributo" Then
            If MsgBox("Este atributo já está associado a um tipo de mensagem." & vbCrLf & _
                      "Somente o Nome Lógico e Descrição serão alterados." & vbCrLf & _
                      "Confirma ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then
                Exit Sub
            End If
        Else
            frmMural.Display = strRetorno
            frmMural.Show vbModal
            Exit Sub
       End If
    End If
    
    If Not IsNull(dtpDataFimVigencia.Value) Then
        If MsgBox("Deseja desativar o registro a partir do dia " & dtpDataFimVigencia.Value & " ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
    End If
    
    Set xmlExecucao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlExecucao, "", "Repeat_Execucao", "")
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "NO_ATRB_MESG", Trim$(txtNomeFisico.Text))
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "NO_TRAP_ATRB", Trim$(txtNomeLogico.Text))
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DE_ATRB_MESG", Trim$(txtDescricao.Text))
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "TP_DADO_ATRB_MESG", flObterTipoDado)
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "QT_CTER_ATRB", numTamanho.Valor)
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "QT_CASA_DECI_ATRB", numDecimais.Valor)
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DT_INIC_VIGE_ATRB_MESG", fgDt_To_Xml(dtpDataInicioVigencia.Value))
        
    If IsNull(dtpDataFimVigencia.Value) Then
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DT_FIM_VIGE_ATRB_MESG", "")
    Else
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DT_FIM_VIGE_ATRB_MESG", fgDt_To_Xml(dtpDataFimVigencia.Value))
    End If
    
    If chkIndicadorValorNegativo.Value = vbChecked Then
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "IN_ATRB_PRMT_VALO_NEGT", enumIndicadorSimNao.sim)
    Else
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "IN_ATRB_PRMT_VALO_NEGT", enumIndicadorSimNao.nao)
    End If
    
    If lstAtributo.SelectedItem Is Nothing Then
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DH_ULTI_ATLZ", "")
    Else
        Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DH_ULTI_ATLZ", Split(lstAtributo.SelectedItem.Key, "|")(2))
    End If
    
    Call fgMIUExecutarGenerico(strOperacao, "A7Server.clsAtributoMensagem", xmlExecucao)
    Call flCarregaListView
            
    fgCursor
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption
                
    If strOperacao = "Incluir" Then
        flProtegerChave
    End If
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, "frmAtributo", "flSalvar", lngCodigoErroNegocio, intNumeroSequencialErro

End Sub

Private Sub dtpDataFimVigencia_Change()
    
    If Not IsNull(dtpDataFimVigencia.Value) Then
        If dtpDataFimVigencia.Value < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
            dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
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

    dtpDataFimVigencia.MinDate = dtpDataInicioVigencia.Value
    dtpDataFimVigencia.Value = dtpDataInicioVigencia.Value
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
        
    fgCenterMe Me
    Me.Icon = mdiBUS.Icon
    
    DoEvents
    
    fgCursor True
    
    tlbCadastro.Buttons("Limpar").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
    tlbCadastro.Buttons("Salvar").Enabled = gblnPerfilManutencao 'True
    
    Call flLimpaCampos
    
    Me.Show

    Call flCarregaListView
    Call fgCursor(False)
    
    Exit Sub
    
ErrorHandler:
    
    Call fgCursor(False)
    
    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAtributo = Nothing
End Sub

'Validar os valores informados para o atributo.
Private Function flValidarCampos() As String
    
Dim xmlDomDoc                               As MSXML2.DOMDocument40
Dim xmlLeitura                              As MSXML2.DOMDocument40
Dim strXml                                  As String

    On Error GoTo ErrorHandler
    
    If Trim$(txtNomeFisico) = "" Then
        flValidarCampos = "Digite o nome Físico do atributo."
        txtNomeFisico.SetFocus
        Exit Function
    End If
    
    If InStr(1, txtNomeFisico, "/", vbBinaryCompare) = 0 Then
        Set xmlDomDoc = CreateObject("MSXML2.DOMDocument.4.0")
        strXml = "<ATRB><" & Trim$(txtNomeFisico) & "></" & Trim$(txtNomeFisico) & "></ATRB>"
        xmlDomDoc.loadXML strXml
        
        If xmlDomDoc.parseError.errorCode <> 0 Then
            flValidarCampos = "Nome Físico do atributo inválido."
            txtNomeFisico.SetFocus
            Exit Function
        End If
    End If
    
    Set xmlDomDoc = Nothing
    
    If Trim$(txtNomeLogico) = "" Then
        flValidarCampos = "Infome o nome Lógico do atributo."
        txtNomeLogico.SetFocus
        Exit Function
    End If
            
    If optTipoDado(0) Then
        If numTamanho.Valor = 0 Or _
            numTamanho.Valor = "" Then
            flValidarCampos = "Infome o tamanho do atributo."
            numTamanho.SetFocus
            Exit Function
        End If
        If numDecimais > numTamanho Then
            flValidarCampos = "Número de casas decimais não pode ultrapassar o tamanho do atributo."
            numDecimais.SetFocus
            Exit Function
        End If
    ElseIf optTipoDado(1) Then
        If numTamanho.Valor = 0 Or _
            numTamanho.Valor = "" Then
            flValidarCampos = "Infome o tamanho do atributo."
            numTamanho.SetFocus
            Exit Function
        End If
    End If
    
    If strOperacao = "Alterar" Then
        Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
        
        Call fgAppendNode(xmlLeitura, "", "Repeat_Leitura", "")
        Call fgAppendNode(xmlLeitura, "Repeat_Leitura", "NO_ATRB_MESG", Trim$(txtNomeFisico))
        
        Call xmlLeitura.loadXML(fgMIUExecutarGenerico("Ler", "A7Server.clsTipoMesgAtributo", xmlLeitura))
        
        If xmlLeitura.xml <> vbNullString Then
            flValidarCampos = "Atributo"
            Exit Function
        End If
    End If
    
    flValidarCampos = ""

    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flValidarCampos", 0

End Function

'Carregar os campos do formulário com os valores recebidos da camada de negócio.
Private Sub flXmlToInterface()

Dim xmlLeitura                              As MSXML2.DOMDocument40
    
    On Error GoTo ErrorHandler
        
    txtNomeFisico.Enabled = False
    dtpDataInicioVigencia.Enabled = False
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, "", "Repeat_Leitura", "")
    Call fgAppendNode(xmlLeitura, "Repeat_Leitura", "NO_ATRB_MESG", lstAtributo.SelectedItem.Text)
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("Ler", "A7Server.clsAtributoMensagem", xmlLeitura))
    
    tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao 'True
       
    With xmlLeitura
   
        txtNomeFisico.Text = .selectSingleNode("//NO_ATRB_MESG").Text
        txtNomeLogico.Text = .selectSingleNode("//NO_TRAP_ATRB").Text
        txtDescricao.Text = .selectSingleNode("//DE_ATRB_MESG").Text
        numTamanho.Valor = .selectSingleNode("//QT_CTER_ATRB").Text
        numDecimais.Valor = .selectSingleNode("//QT_CASA_DECI_ATRB").Text
        
        If .selectSingleNode("//IN_ATRB_PRMT_VALO_NEGT").Text = enumIndicadorSimNao.sim Then
            chkIndicadorValorNegativo.Value = vbChecked
        Else
            chkIndicadorValorNegativo.Value = vbUnchecked
        End If
        
        Call flSelecionaTipoDado(Val(.selectSingleNode("//TP_DADO_ATRB_MESG").Text))
        
        dtpDataInicioVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("//DT_INIC_VIGE_ATRB_MESG").Text)
        dtpDataInicioVigencia.Value = fgDtXML_To_Date(.selectSingleNode("//DT_INIC_VIGE_ATRB_MESG").Text)
        
        If dtpDataInicioVigencia.Value > fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
            dtpDataInicioVigencia.MinDate = fgDataHoraServidor(enumFormatoDataHoraAux.DataAux)
            dtpDataInicioVigencia.Enabled = True
        End If
        
        If Trim$(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text) <> gstrDataVazia Then
            If fgDtXML_To_Date(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text) < fgDataHoraServidor(enumFormatoDataHoraAux.DataAux) Then
                dtpDataFimVigencia.MinDate = fgDtXML_To_Date(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text)
                dtpDataInicioVigencia.Enabled = True
            Else
                dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
            End If
            dtpDataFimVigencia.Value = fgDtXML_To_Date(.selectSingleNode("//DT_FIM_VIGE_ATRB_MESG").Text)
        Else
            dtpDataFimVigencia.MinDate = fgMaiorData(dtpDataInicioVigencia.Value, fgDataHoraServidor(enumFormatoDataHoraAux.DataAux))
            dtpDataFimVigencia.Value = dtpDataFimVigencia.MinDate
            dtpDataFimVigencia.Value = Null
        End If
    
    End With
        
    Exit Sub
    
ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flXmlToInterface", 0
    
End Sub

Private Sub lstAtributo_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'Ordenar as colunas do listview

    lstAtributo.Sorted = True
    lstAtributo.SortKey = ColumnHeader.Index - 1

    If lstAtributo.SortOrder = lvwAscending Then
        lstAtributo.SortOrder = lvwDescending
    Else
        lstAtributo.SortOrder = lvwAscending
    End If

End Sub

Private Sub lstAtributo_ItemClick(ByVal Item As MSComctlLib.ListItem)

    On Error GoTo ErrorHandler
    
    Call fgCursor(True)
    Call flLimpaCampos
    strOperacao = "Alterar"
    strKeyItemSelected = Item.Key
    Call flXmlToInterface
    Call fgCursor(False)
    
    Exit Sub

ErrorHandler:
    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - lstAtributo_ItemClick"

    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

Private Sub optTipoDado_Click(Index As Integer)
    
On Error GoTo ErrorHandler
    
    Select Case Index
        Case 0
            numTamanho.Enabled = True
            numDecimais.Enabled = True
            chkIndicadorValorNegativo.Enabled = True
        Case 1
            numTamanho.Enabled = True
            numDecimais.Enabled = False
            numDecimais.Valor = 0
            chkIndicadorValorNegativo.Value = vbUnchecked
            chkIndicadorValorNegativo.Enabled = False
    End Select

    
    Exit Sub
ErrorHandler:

    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - optTipoDado_Click"

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler
    
    fgCursor True
    
    Select Case Button.Key
        Case "Limpar"
            Call flLimpaCampos
        Case "Salvar"
            Call flSalvar
        Case "Excluir"
            strOperacao = "Excluir"
            flExcluir
            Call flLimpaCampos
        Case "Sair"
            Unload Me
            strOperacao = ""
    End Select
    
    If strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If
            
    fgCursor
            
    Exit Sub

ErrorHandler:

    fgCursor

    mdiBUS.uctLogErros.MostrarErros Err, "frmAtributo - tlbCadastro_ButtonClick"
    
    Call flCarregaListView
    
    If strOperacao = "Excluir" Then
        flLimpaCampos
    ElseIf strOperacao = "Alterar" Then
        flPosicionaItemListView
    End If

End Sub

'Carregar listview com os atributos cadastrados.
Private Sub flCarregaListView()

Dim xmlNode                                 As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem
Dim xmlLeitura                              As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
        
    lstAtributo.ListItems.Clear
    lstAtributo.HideSelection = False
    
    Set xmlLeitura = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlLeitura, "", "Repeat_Leitura", "")
    Call fgAppendNode(xmlLeitura, "Repeat_Leitura", "IN_VIGE", "")
    
    Call xmlLeitura.loadXML(fgMIUExecutarGenerico("LerTodos", "A7Server.clsAtributoMensagem", xmlLeitura))
    
    For Each xmlNode In xmlLeitura.selectSingleNode("//Repeat_AtributoMensagem").childNodes
        With xmlNode
                
            Set objListItem = lstAtributo.ListItems.Add(, "|" & .selectSingleNode("NO_ATRB_MESG").Text & _
                                                          "|" & .selectSingleNode("DH_ULTI_ATLZ").Text, .selectSingleNode("NO_ATRB_MESG").Text)
            
            objListItem.SubItems(1) = .selectSingleNode("NO_TRAP_ATRB").Text
            objListItem.SubItems(2) = flObterNomeTipoDado(Val(.selectSingleNode("TP_DADO_ATRB_MESG").Text))
            objListItem.SubItems(3) = .selectSingleNode("QT_CTER_ATRB").Text
            objListItem.SubItems(4) = .selectSingleNode("QT_CASA_DECI_ATRB").Text
            objListItem.SubItems(5) = Format(fgDtXML_To_Date(.selectSingleNode("DT_INIC_VIGE_ATRB_MESG").Text), gstrMascaraDataDtp)
            
            If CStr(.selectSingleNode("DT_FIM_VIGE_ATRB_MESG").Text) <> gstrDataVazia Then
                objListItem.SubItems(6) = Format(fgDtXML_To_Date(.selectSingleNode("DT_FIM_VIGE_ATRB_MESG").Text), gstrMascaraDataDtp)
            End If
        
        End With
    Next

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flCarregaListView", 0

End Sub

'Converter o domínio numérico de tipo de dado para literais.
Private Function flObterNomeTipoDado(ByVal plCodigoTipoDados As Long) As String
    
    Select Case plCodigoTipoDados
        Case enumTipoDadoAtributo.Numerico
            flObterNomeTipoDado = "Numérico"
        Case enumTipoDadoAtributo.Alfanumerico
            flObterNomeTipoDado = "Alfanumérico"
    End Select

End Function

'Converter as literais de tipo de dado para o domínio numérico.
Private Function flObterTipoDado() As Long

    If optTipoDado(0) Then
        flObterTipoDado = enumTipoDadoAtributo.Numerico
    ElseIf optTipoDado(1) Then
        flObterTipoDado = enumTipoDadoAtributo.Alfanumerico
    End If

End Function

'Marca o tipo de dado do atributo de acordo com o valor recebido da camada de negócio.
Private Sub flSelecionaTipoDado(ByVal plCodigoTipoDado As Long)

   Select Case plCodigoTipoDado
        Case enumTipoDadoAtributo.Numerico
            optTipoDado(0).Value = True
        Case enumTipoDadoAtributo.Alfanumerico
            optTipoDado(1).Value = True
    End Select

End Sub

'Excluir o atributo corrente.
Private Sub flExcluir()

Dim xmlExecucao                             As MSXML2.DOMDocument40

    On Error GoTo ErrorHandler
      
    If lstAtributo.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Confirma a exclusão do Atributo de Mensagem selecionado ?", vbQuestion + vbYesNo, Me.Caption) = vbNo Then Exit Sub
   
    Set xmlExecucao = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlExecucao, "", "Repeat_Execucao", "")
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "NO_ATRB_MESG", Split(lstAtributo.SelectedItem.Key, "|")(1))
    Call fgAppendNode(xmlExecucao, "Repeat_Execucao", "DH_ULTI_ATLZ", Split(lstAtributo.SelectedItem.Key, "|")(2))
    
    Call fgMIUExecutarGenerico("Excluir", "A7Server.clsAtributoMensagem", xmlExecucao)
    Call flCarregaListView
        
    fgCursor
    MsgBox "Operação efetuada com sucesso.", vbInformation, Me.Caption

    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flExcluir", 0

End Sub

'Proteger chave do atributo em operações de alteração.
Private Sub flProtegerChave()
    
   strOperacao = "Alterar"
   txtNomeFisico.Enabled = False
   dtpDataInicioVigencia.Enabled = False
   tlbCadastro.Buttons("Excluir").Enabled = gblnPerfilManutencao
 
End Sub

