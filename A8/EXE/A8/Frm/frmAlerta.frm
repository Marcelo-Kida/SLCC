VERSION 5.00
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlerta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disponibilização de Alertas"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11445
   Icon            =   "frmAlerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboFatoGeradorAlerta 
      Height          =   315
      ItemData        =   "frmAlerta.frx":0442
      Left            =   90
      List            =   "frmAlerta.frx":0449
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   990
      Width           =   5085
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   90
      Picture         =   "frmAlerta.frx":0453
      ScaleHeight     =   555
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   45
      Width           =   585
   End
   Begin MSComctlLib.ListView lvwAlerta 
      Height          =   2565
      Left            =   60
      TabIndex        =   0
      Top             =   1380
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo de Ocorrência"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Veículo Legal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Contraparte"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tipo de Operação"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Nr. Comando"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Data/Hora Ocorrência"
         Object.Width           =   2540
      EndProperty
   End
   Begin NumBox.Number numTimer 
      Height          =   330
      Left            =   1920
      TabIndex        =   2
      Top             =   4035
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      AceitaNegativo  =   0   'False
      SelStart        =   1
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2505
      TabIndex        =   3
      Top             =   4035
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      OrigLeft        =   2026
      OrigTop         =   270
      OrigRight       =   2266
      OrigBottom      =   600
      Max             =   60
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   4050
      Top             =   0
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
            Picture         =   "frmAlerta.frx":0895
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":09A7
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":0CC1
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1013
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1125
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":143F
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1759
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1A73
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1D8D
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":20DF
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":2431
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":2783
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   10260
      TabIndex        =   4
      Top             =   4005
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   582
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "OK"
            Key             =   "OK"
            Object.ToolTipText     =   "Fechar formulário"
            ImageKey        =   "Salvar"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblAlerta 
      AutoSize        =   -1  'True
      Caption         =   "Fato Gerador Alerta"
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
      Index           =   3
      Left            =   90
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Atenção "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   750
      TabIndex        =   6
      Top             =   165
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Tempo Atualização"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   90
      TabIndex        =   5
      Top             =   4050
      Width           =   1755
   End
End
Attribute VB_Name = "frmAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
'' Objeto responsável pela disponibilização dos alertas do sistema.
''
'' São consideradas classes de destino:
'' A8MIU.clsMiu
''

Option Explicit

Private xmlMapaNavegacao                    As MSXML2.DOMDocument40
Private xmlLerTodos                         As MSXML2.DOMDocument40

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private Const strFuncionalidade             As String = "frmAlerta"

Private Const COL_TIPO_OCORRENCIA           As Integer = 0
Private Const COL_VEICULO_LEGAL             As Integer = 1
Private Const COL_CONTRAPARTE               As Integer = 2
Private Const COL_TIPO_OPER                 As Integer = 3
Private Const COL_NUMERO_COMANDO            As Integer = 4
Private Const COL_VALOR                     As Integer = 5
Private Const COL_DATAHORA                  As Integer = 6

Private Sub cboFatoGeradorAlerta_Click()

On Error GoTo ErrorHandler
    
    If cboFatoGeradorAlerta.ListIndex <> -1 Then
        fgCursor True
        Call VerificarAlertas("")
        fgCursor
    End If

    Exit Sub
ErrorHandler:
    
    fgCursor
    
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler

    fgCursor True
    fgCenterMe Me
    
    If glngTempoAlerta < UpDown1.Min Or _
       glngTempoAlerta > UpDown1.Max Then
           
        numTimer.Valor = UpDown1.Min
        UpDown1.value = UpDown1.Min
    
    Else
        numTimer.Valor = glngTempoAlerta
        UpDown1.value = glngTempoAlerta
        
    End If
    
    DoEvents
    
    Call flInicializar
    fgCursor

Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

'' Obtém as propriedades da tabela de alertas, através de interação com a camada
'' de controle de caso de uso MIU, solicitando informações para o seguinte
'' componente/classe/métodoA8MIU.clsMiu.ObterMapaNavegacao
Public Function flInicializar() As Boolean

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    If fgVerificaJanelaVerificacao() Then Exit Function
    
    Set xmlMapaNavegacao = CreateObject("MSXML2.DOMDocument.4.0")
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    If Not xmlMapaNavegacao.loadXML(objMIU.ObterMapaNavegacao(enumSistemaSLCC.LQS, strFuncionalidade, vntCodErro, vntMensagemErro)) Then
    
        If vntCodErro <> 0 Then
            GoTo ErrorHandler
        End If
    
        Call fgErroLoadXML(xmlMapaNavegacao, App.EXEName, TypeName(Me), "flInicializar")
    End If
        
    Call fgCarregarCombos(Me.cboFatoGeradorAlerta, _
                          xmlMapaNavegacao, _
                          "FatoGeradorAlerta", _
                          "CO_FATO_GERA_ALER", _
                          "DE_FATO_GERA_ALER", _
                          True)
    
    Set objMIU = Nothing
    
Exit Function
ErrorHandler:
    Set objMIU = Nothing
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, Me.Name, "flInicializar", 0

End Function

Private Sub flLimparLista()

On Error GoTo ErrorHandler
    
    lvwAlerta.ListItems.Clear
    
    Exit Sub
    
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flLimparLista", 0
    
End Sub

'' Retorna True se existirem alertas ativos.
Public Function VerificarAlertas(ByRef pstrTip As String) As Boolean

Dim lngQtdRegistros                         As Long

On Error GoTo ErrorHandler

    If Not flCarregarAlertas Then
        flLimparLista
        Exit Function
    End If

    Call flCarregarLista(lngQtdRegistros)
    pstrTip = "Existe(m) " & lngQtdRegistros & " alerta(s)"

    VerificarAlertas = True
        
    Exit Function
    
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "VerificarAlertas", 0

End Function

'' Preenche a listagem do objeto com os alertas ativos, se existirem
Private Sub flCarregarLista(ByRef plngQtdRegistros As Long)

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim lstItem                                 As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    flLimparLista

    plngQtdRegistros = 0

    For Each objDomNode In xmlLerTodos.documentElement.childNodes
        
        With objDomNode
        
            plngQtdRegistros = plngQtdRegistros + 1
            'Adilson - Acrescentado "K" & plngQtdRegistros para evitar DupKey no ListView
            Set lstItem = lvwAlerta.ListItems.Add(, "K" & .selectSingleNode("NU_SEQU_ALER_EMIT").Text & "K" & plngQtdRegistros)
            
            lstItem.Text = .selectSingleNode("DE_FATO_GERA_ALER").Text
            lstItem.SubItems(COL_VEICULO_LEGAL) = .selectSingleNode("NO_VEIC_LEGA").Text
            lstItem.SubItems(COL_CONTRAPARTE) = .selectSingleNode("NO_CNPT").Text
            lstItem.SubItems(COL_TIPO_OPER) = .selectSingleNode("NO_TIPO_OPER").Text
            lstItem.SubItems(COL_NUMERO_COMANDO) = .selectSingleNode("NU_COMD_OPER").Text
            lstItem.SubItems(COL_VALOR) = fgVlrXml_To_Interface(.selectSingleNode("VA_OPER_ATIV").Text)
            lstItem.SubItems(COL_DATAHORA) = fgDtHrXML_To_Interface(.selectSingleNode("DH_EMIS_ALER").Text)
        End With
    
    Next objDomNode
    
    Exit Sub
    
ErrorHandler:

    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLista", 0

End Sub

'' Carrega todos os alertas ativos e retorna True se existir ao menos um, através
'' de acesso a camada de controle de caso de uso, para a camada de realização de
'' caso de uso (componente / classe / método):  A8MIU.clsMIU.Executar
Private Function flCarregarAlertas() As Boolean

#If EnableSoap = 1 Then
    Dim objMIU             As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU             As A8MIU.clsMIU
#End If

Dim strRetorno             As String
Dim vntCodErro             As Variant
Dim vntMensagemErro        As Variant

On Error GoTo ErrorHandler
    
    If fgVerificaJanelaVerificacao() Then Exit Function
    
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Alerta/@Operacao").Text = "LerTodos"
    xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Alerta/DH_EMIS_ALER").Text = fgDt_To_Xml(fgDataHoraServidor(enumFormatoDataHora.Data))
    
    If cboFatoGeradorAlerta.ListIndex <> 0 Then
        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Alerta/CO_FATO_GERA_ALER").Text = fgObterCodigoCombo(Me.cboFatoGeradorAlerta.Text)
    Else
        xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Alerta/CO_FATO_GERA_ALER").Text = "0"
    End If

    Set objMIU = fgCriarObjetoMIU("A8Miu.clsMIU")
    strRetorno = objMIU.Executar(xmlMapaNavegacao.documentElement.selectSingleNode("Grupo_Propriedades/Grupo_Alerta").xml, vntCodErro, vntMensagemErro)
    
    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objMIU = Nothing

    If strRetorno = vbNullString Then
        flCarregarAlertas = False
        Exit Function
    End If

    If xmlLerTodos Is Nothing Then
        Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
    End If
    Call xmlLerTodos.loadXML(strRetorno)
    
    flCarregarAlertas = True
    
Exit Function
ErrorHandler:
    
    If Not IsEmpty(vntCodErro) Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If
    
    fgRaiseError App.EXEName, Me.Name, "flCarregarAlertas", 0
    'fgRaiseError App.EXEName, Me.Name, "flCarregarAlertas", lngCodigoErroNegocio, intNumeroSequencialErro

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Cancel = True
    Me.Hide

End Sub

Private Sub lvwAlerta_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo ErrorHandler
    
    If KeyCode = vbKeyF5 Then
        Call fgCursor(True)
        Call VerificarAlertas("")
        Call fgCursor(False)
    End If
    
    Exit Sub
ErrorHandler:
    Call fgCursor(False)
    
    mdiLQS.uctlogErros.MostrarErros Err, "frmConsultaOperacao - lstOperacao_KeyDown", Me.Caption

End Sub

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo ErrorHandler
    
    fgCursor True
        
    Select Case Button.Key
        Case "OK"
            SaveSetting "A8LQS", "DisponibilizacaoAlerta", "Intervalo", numTimer.Valor
            Me.Hide
    End Select
    
    fgCursor
    
    Exit Sub
ErrorHandler:
    fgCursor
    Call mdiLQS.uctlogErros.MostrarErros(Err, Me.Name, Me.Caption)

End Sub

Private Sub UpDown1_Change()
    numTimer.Valor = UpDown1.value
End Sub

