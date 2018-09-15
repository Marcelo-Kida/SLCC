VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl TableCombo 
   BackColor       =   &H80000010&
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   3990
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TableCombo.ctx":0000
            Key             =   "Cancelar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TableCombo.ctx":0452
            Key             =   "Aplicar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandos 
      Height          =   330
      Left            =   1470
      TabIndex        =   3
      Top             =   2970
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   582
      ButtonWidth     =   2011
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Aplicar"
            Key             =   "aplicar"
            Object.ToolTipText     =   "Aplicar Filtro"
            ImageIndex      =   2
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Filtro"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ListView lstDados 
      Height          =   2565
      Left            =   -30
      TabIndex        =   2
      Top             =   390
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Selecione a Informação"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Line lnTitulo 
      DrawMode        =   4  'Mask Not Pen
      Index           =   3
      Visible         =   0   'False
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   360
   End
   Begin VB.Line lnTitulo 
      BorderColor     =   &H00404040&
      DrawMode        =   1  'Blackness
      Index           =   2
      Visible         =   0   'False
      X1              =   3930
      X2              =   3930
      Y1              =   15
      Y2              =   360
   End
   Begin VB.Line lnTitulo 
      DrawMode        =   1  'Blackness
      Index           =   1
      Visible         =   0   'False
      X1              =   0
      X2              =   3930
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Line lnTitulo 
      DrawMode        =   4  'Mask Not Pen
      Index           =   0
      Visible         =   0   'False
      X1              =   0
      X2              =   3930
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lblBarra 
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   2940
      Width           =   3885
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos de Veículos Legais"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   3465
   End
   Begin VB.Label lblDrop 
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3660
      TabIndex        =   1
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "TableCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Empresa            : Regerbanc - Participações , Negócios e Serviços LTDA
'Componente         : ctlTableCombo
'Classe             :
'Data Criação       : 18/07/2003
'Objetivo           : OCX armazena conteúdos de filtros aplicados anteriormente, para serem novamente utilizados
'                     pelo usuário, sem que o mesmo precise acessar a tela de filtros.
'Analista           : Carlos Fortes/Cassiano Nicolosi
'
'Programador        : Cassiano Nicolosi
'Data               : 08/08/2003
'
'Data Teste         :
'Autor              :
'
'Data Alteração     :
'Autor              :
'Objetivo           :

Option Explicit

'Declarações de Eventos
Public Event AplicarFiltro(xmlDocFiltros As String)
Public Event DropDown()
Public Event MouseMove()

Private strTagXML                           As String
Private strTagRepeat                        As String
Private strTagCodigo                        As String
Private strTagDescricao                     As String
Private strNomeObjeto                       As String

Private Sub flAplicarFiltro(objListItem As ComctlLib.ListItem)

Dim xmlDomFiltros                           As MSXML2.DOMDocument40
Dim lsDocFiltros                            As String

    Me.TituloCombo = objListItem.Text
    
    Set xmlDomFiltros = CreateObject("MSXML2.DOMDocument.4.0")
    
    Call fgAppendNode(xmlDomFiltros, "", "Repeat_Filtros", "")
    Call fgAppendNode(xmlDomFiltros, "Repeat_Filtros", "Grupo_" & strTagXML, "")
    Call fgAppendNode(xmlDomFiltros, "Grupo_" & strTagXML, strTagXML, Mid$(objListItem.Key, 5))
            
    lsDocFiltros = xmlDomFiltros.xml
    
    RaiseEvent AplicarFiltro(lsDocFiltros)
                                           
End Sub

Public Sub fgCarregarCombo()

#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As Object
#End If

Dim xmlLerTodos                             As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As ComctlLib.ListItem
Dim strClasseMIU                            As String
Dim vntCodErro                              As Variant
Dim vntMensagemErro                         As Variant

On Error GoTo ErrorHandler
    
    If lstDados.ListItems.Count = 0 Then
        fgCursor True
        Call flInicializarControlesInternos
        
        Set xmlLerTodos = CreateObject("MSXML2.DOMDocument.4.0")
        
        If Trim$(App.Title) = "A6" Then
            strClasseMIU = "A6MIU.clsMIU"
        Else
            strClasseMIU = "A8MIU.clsMIU"
        End If
        
        Set objMIU = fgCriarObjetoMIU(strClasseMIU)
        
        Call fgAppendNode(xmlLerTodos, "", "Repeat_Filtros", "")
        Call fgAppendNode(xmlLerTodos, "Repeat_Filtros", "Grupo_Filtros", "")
        Call fgAppendAttribute(xmlLerTodos, "Grupo_Filtros", "Operacao", "LerTodos")
        Call fgAppendAttribute(xmlLerTodos, "Grupo_Filtros", "Objeto", strNomeObjeto)
        
        If xmlLerTodos.loadXML(objMIU.Executar(xmlLerTodos.xml, vntCodErro, vntMensagemErro)) Then
            
            If vntCodErro <> 0 Then
                GoTo ErrorHandler
            End If
            
            For Each objDomNode In xmlLerTodos.documentElement.selectNodes("/Repeat_" & strTagRepeat & "/*")
                Set objListItem = lstDados.ListItems.Add(, "cod_" & objDomNode.selectSingleNode(strTagCodigo).Text, _
                                                                    objDomNode.selectSingleNode(strTagDescricao).Text)
            Next
        End If
        
        fgCursor
    End If
        
    UserControl.Height = IIf(UserControl.Height > 1000, lnTitulo(1).Y1 + 30, 4000)
    
Exit Sub
ErrorHandler:
    fgCursor
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call fgRaiseError(App.EXEName, UserControl.Parent.Name, "ctlTableCombo - fgCarregarCombo()", 0)

End Sub

Private Sub flInicializarControlesInternos()
    
    Select Case UserControl.Parent.Name
    Case "frmControleRemessa", "frmSubReservaD0", "frmRemessaRejeitada", "frmCaixaFuturo"
        strTagXML = "BancoLiquidante"
        strTagRepeat = "Empresa"
        strTagCodigo = "CO_EMPR"
        strTagDescricao = "NO_REDU_EMPR"
        strNomeObjeto = "A6A7A8.clsEmpresa"
    
    Case "frmSubReservaResumo", "frmSubReservaAbertura", "frmSubReservaFechamento", "frmConsultaAberturaFechamento"
        strTagXML = "GrupoVeiculoLegal"
        strTagRepeat = "GrupoVeiculoLegal"
        strTagCodigo = "CO_GRUP_VEIC_LEGA"
        strTagDescricao = "NO_GRUP_VEIC_LEGA"
        strNomeObjeto = "A6A7A8.clsGrupoVeiculoLegal"
    
    End Select
    
End Sub

Public Sub fgMakeButton3D()
    lnTitulo(0).Visible = True
    lnTitulo(1).Visible = True
    lnTitulo(2).Visible = True
    lnTitulo(3).Visible = True
End Sub

Public Sub fgMakeButtonFlat()
    lnTitulo(0).Visible = False
    lnTitulo(1).Visible = False
    lnTitulo(2).Visible = False
    lnTitulo(3).Visible = False
    
    UserControl.Height = lnTitulo(1).Y1 + 30
End Sub

Public Property Get TituloCombo() As String
    TituloCombo = lblTitulo.Caption
End Property

Public Property Let TituloCombo(psTituloCombo As String)
    lblTitulo.Caption = psTituloCombo
End Property

Private Sub lblDrop_Click()
    RaiseEvent DropDown
End Sub

Private Sub lblDrop_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub lblTitulo_Click()
    RaiseEvent DropDown
End Sub

Private Sub lblTitulo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove
End Sub

Private Sub lstDados_DblClick()
    If lstDados.SelectedItem Is Nothing Then Exit Sub
    Call flAplicarFiltro(lstDados.SelectedItem)
    UserControl.Height = lnTitulo(1).Y1 + 30
End Sub

Private Sub tlbComandos_ButtonClick(ByVal Button As MSComctlLib.Button)
    
Dim objItem                                 As ComctlLib.ListItem

    Select Case Button.Key
    Case "aplicar"
        For Each objItem In lstDados.ListItems
            If objItem.Selected Then
                Call flAplicarFiltro(objItem)
                Exit For
            End If
        Next
        
    End Select
    
    UserControl.Height = lnTitulo(1).Y1 + 30

End Sub

Private Sub UserControl_Resize()

    On Error Resume Next
    
    lstDados.Left = 0
    lstDados.Top = 390
    lstDados.Height = IIf((UserControl.ScaleHeight - lstDados.Top - lblBarra.Height) < 0, 0, UserControl.ScaleHeight - lstDados.Top - lblBarra.Height)
    lstDados.Width = UserControl.ScaleWidth
    
    lblBarra.Left = 0
    lblBarra.Top = lstDados.Top + lstDados.Height
    lblBarra.Width = UserControl.ScaleWidth
    
    tlbComandos.Left = UserControl.ScaleWidth - tlbComandos.Width - 60
    tlbComandos.Top = lblBarra.Top + 30
End Sub
