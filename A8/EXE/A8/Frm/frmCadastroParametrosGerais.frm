VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCadastroParametrosGerais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro - Parâmetros Gerais"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8280
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   8280
   Begin VB.Frame fraAdm 
      Height          =   7455
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   810
      Width           =   8095
      Begin FPSpread.vaSpread vasAdministracaoDados 
         Height          =   7065
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   7780
         _Version        =   196608
         _ExtentX        =   13723
         _ExtentY        =   12462
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ArrowsExitEditMode=   -1  'True
         AutoCalc        =   0   'False
         ColHeaderDisplay=   0
         ColsFrozen      =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         GrayAreaBackColor=   16777215
         MaxCols         =   1
         MaxRows         =   1
         NoBorder        =   -1  'True
         ProcessTab      =   -1  'True
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCadastroParametrosGerais.frx":0000
         UnitType        =   2
         UserResize      =   0
         ScrollBarTrack  =   3
      End
   End
   Begin VB.Frame fraAdm 
      Height          =   675
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   8095
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblAdm 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Grupo"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   435
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   90
      Top             =   8070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":0216
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":0530
            Key             =   "Padrao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":0982
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":0C9C
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":0FB6
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":12D0
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":1722
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":1B74
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":1FC6
            Key             =   "checkfalse"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadastroParametrosGerais.frx":2060
            Key             =   "checktrue"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   6480
      TabIndex        =   0
      Top             =   8295
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ButtonWidth     =   1508
      ButtonHeight    =   582
      Wrappable       =   0   'False
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Limpar"
            Key             =   "Limpar"
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Salvar"
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sai&r"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadastroParametrosGerais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3F3BDE440266"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
'Empresa        : Regerbanc
'Pacote         :
'Classe         : frmCadastroParametrosGerais
'Data Criação   : 31/05/2004
'Objetivo       :
'
'Analista       : Michel P. Barros / M. Kida
'
'Programador    : Cassiano
'Data           : 31/05/2004
'
'Teste          :
'Autor          :
'
'Data Alteração :
'Autor          :
'Objetivo       :

Option Explicit

Dim xmlParametrizacao                       As MSXML2.DOMDocument40

Private Enum enumGrupo
    Todos = 0
    BaseHistorica = 1
    DV = 2
    BG = 3
    HA = 4
End Enum

Private Enum enumTipoDado
    Numerico = 0
    Alfanumerico = 1
    NumericoDecimal = 2
End Enum



Private Sub flAtribuirValoresXML()
    
Dim lngCountRows                            As Long
Dim vntNomeTag                              As Variant
Dim vntConteudoTag                          As Variant

On Error GoTo ErrorHandler

    With Me.vasAdministracaoDados
        
        For lngCountRows = 1 To .MaxRows
            .GetText 5, lngCountRows, vntNomeTag
            
            If vntNomeTag <> vbNullString Then
                
                .BlockMode = False
                .Col = 4
                .Row = lngCountRows
                
                .GetText 3, lngCountRows, vntConteudoTag
                xmlParametrizacao.selectSingleNode(vntNomeTag).Text = vntConteudoTag
                
                 If xmlParametrizacao.selectSingleNode(vntNomeTag & "/@OBRIG") Is Nothing Then
                    Call fgAppendAttribute(xmlParametrizacao, vntNomeTag, "OBRIG", IIf(.value = True, "S", "N"))
                 Else
                    xmlParametrizacao.selectSingleNode(vntNomeTag & "/@OBRIG").Text = IIf(.value = True, "S", "N")
                 End If
            
            End If
        Next
    End With
    
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flAtribuirValoresXML", 0

End Sub

Private Sub flCarregarParametrizacao()
    
#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim arrCondFor()                            As String
Dim intCount                                As Integer
Dim strConteudoCampo                        As String

On Error GoTo ErrorHandler

    If cboGrupo.ListIndex = -1 Then Exit Sub
    
    If xmlParametrizacao.xml = vbNullString Then
        Call fgAppendNode(xmlParametrizacao, "", "PARM_GERL", "")
        Call fgAppendAttribute(xmlParametrizacao, "PARM_GERL", "Operacao", "Ler")
        Call fgAppendAttribute(xmlParametrizacao, "PARM_GERL", "Objeto", "A8LQS.clsParametrosGerais")
        Call fgAppendNode(xmlParametrizacao, "PARM_GERL", "CO_TEXT_XML", "0")
        
        fgCursor True
        
        Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
        Call xmlParametrizacao.loadXML(objMIU.Executar(xmlParametrizacao.xml))
    End If
    
    With Me.vasAdministracaoDados
        Call flAtribuirValoresXML
        
        .ReDraw = False
        .MaxRows = 0
    
        Select Case cboGrupo.ListIndex
            Case enumGrupo.Todos
                ReDim arrCondFor(1 To 6, 1 To 2)
                arrCondFor(1, 1) = "PARM_GERL/PARM_BASE_HIST"
                arrCondFor(2, 1) = "PARM_GERL/PARM_CC_DV"
                arrCondFor(3, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_CRED"
                arrCondFor(4, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_DEBT"
                arrCondFor(5, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_ESTO"
                arrCondFor(6, 1) = "PARM_GERL/PARM_CNTB"
        
                arrCondFor(1, 2) = "Base Histórica"
                arrCondFor(2, 2) = "C/C (DV)"
                arrCondFor(3, 2) = "C/C (BG - Crédito)"
                arrCondFor(4, 2) = "C/C (BG - Débito)"
                arrCondFor(5, 2) = "C/C (BG - Estorno)"
                arrCondFor(6, 2) = "Contábil (HA)"
            
            Case enumGrupo.DV
                ReDim arrCondFor(1 To 1, 1 To 2)
                arrCondFor(1, 1) = "PARM_GERL/PARM_CC_DV"
                arrCondFor(1, 2) = "C/C (DV)"
        
            Case enumGrupo.BG
                ReDim arrCondFor(1 To 3, 1 To 2)
                arrCondFor(1, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_CRED"
                arrCondFor(2, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_DEBT"
                arrCondFor(3, 1) = "PARM_GERL/PARM_CC_BG/PARM_CC_BG_ESTO"
        
                arrCondFor(1, 2) = "C/C (BG - Crédito)"
                arrCondFor(2, 2) = "C/C (BG - Débito)"
                arrCondFor(3, 2) = "C/C (BG - Estorno)"
            
            Case enumGrupo.HA
                ReDim arrCondFor(1 To 1, 1 To 2)
                arrCondFor(1, 1) = "PARM_GERL/PARM_CNTB"
                arrCondFor(1, 2) = "Contábil (HA)"
        
            Case enumGrupo.BaseHistorica
                ReDim arrCondFor(1 To 1, 1 To 2)
                arrCondFor(1, 1) = "PARM_GERL/PARM_BASE_HIST"
                arrCondFor(1, 2) = "Base Histórica"
            
        End Select
        
        For intCount = LBound(arrCondFor) To UBound(arrCondFor)
            
            .MaxRows = .MaxRows + IIf(intCount = 1, 1, 2)
            .SetText 1, .MaxRows, arrCondFor(intCount, 2)
            
            .BlockMode = False
            .Col = 1
            .Row = .MaxRows
            .ForeColor = vbBlue
            .FontBold = True
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .TypeVAlign = TypeVAlignCenter
            
            For Each objDomNode In xmlParametrizacao.selectNodes(arrCondFor(intCount, 1) & "/*")
                
                .MaxRows = .MaxRows + 1
                
                .Row = .MaxRows
                
                .Col = 1
                .CellType = CellTypeStaticText
                .TypeHAlign = TypeHAlignLeft
                .TypeVAlign = TypeVAlignCenter
                
                .Col = 2
                .CellType = CellTypeStaticText
                .TypeHAlign = TypeHAlignLeft
                .TypeVAlign = TypeVAlignCenter
                .Text = objDomNode.baseName
                
                If objDomNode.Text = "" Then
                    If InStr(1, objDomNode.selectSingleNode("@VALOR").Text, "#") > 0 Then
                        .SetText 3, .MaxRows, Replace(objDomNode.selectSingleNode("@VALOR").Text, "#", " ")
                    Else
                        .SetText 3, .MaxRows, objDomNode.selectSingleNode("@VALOR").Text
                    End If
                Else
                    
                    'strConteudoCampo = Replace(objDomNode.Text, vbTab, "")
                    'strConteudoCampo = Replace(strConteudoCampo, vbLf, " ")
                     strConteudoCampo = objDomNode.Text
                    
                    .SetText 3, .MaxRows, strConteudoCampo
                    
                End If
                
                .SetText 5, .MaxRows, arrCondFor(intCount, 1) & "/" & objDomNode.baseName

                If objDomNode.selectSingleNode("@TIPO").Text = "N" Then
                    .Row = .MaxRows
                    .Col = 3
                    .CellType = CellTypeFloat
                    .TypeFloatDecimalPlaces = 0
                    .TypeFloatMax = String(objDomNode.selectSingleNode("@LEN").Text, "9")
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignTop
                    .TypeFloatMoney = False
                    .TypeFloatSeparator = False
                Else
                    .Row = .MaxRows
                    .Col = 3
                    .CellType = CellTypeEdit
                    .TypeEditCharSet = TypeEditCharSetASCII
                    .TypeEditCharCase = TypeEditCharCaseSetNone
                    .TypeHAlign = TypeHAlignRight
                    .TypeVAlign = TypeVAlignTop
                    .TypeEditMultiLine = False
                    .TypeEditPassword = False
                    .TypeMaxEditLen = objDomNode.selectSingleNode("@LEN").Text
                End If
                
                .BlockMode = False
                .Col = 4
                .Row = .MaxRows
                .CellType = CellTypeCheckBox
                .TypeCheckCenter = True
                .TypeCheckPicture(0) = imlIcons.ListImages("checkfalse").Picture
                .TypeCheckPicture(1) = imlIcons.ListImages("checktrue").Picture
                
                If Not objDomNode.selectSingleNode("@OBRIG") Is Nothing Then
                    .value = IIf(objDomNode.selectSingleNode("@OBRIG").Text = "S", 1, 0)
                Else
                    .value = False
                End If
                
            Next
        Next
        
        .BlockMode = True
        .Col = 3
        .Row = 1
        .Col2 = 3
        .Row2 = .MaxRows
        .TypeHAlign = TypeHAlignRight
        .BlockMode = False
        
        .Col = 3
        .Row = 1
        .Action = ActionActiveCell
        
        Call vasAdministracaoDados_LeaveCell(3, 2, 3, 1, False)
        
        .ReDraw = True
        
    End With
    
    fgCursor
    
    Set objMIU = Nothing
    
    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarParametrizacao", 0

End Sub

Private Sub flInicializarSpread()
    
On Error GoTo ErrorHandler

    With vasAdministracaoDados
        
        .ReDraw = False
        
        .MaxCols = 5
        .MaxRows = 1
        
        .ColWidth(1) = 600
        .ColWidth(2) = 2800
        .ColWidth(3) = 2000
        .ColWidth(4) = 2000
        .ColWidth(5) = 0
        
        .SetText 1, 0, "Grupo"
        .SetText 2, 0, "Tag"
        .SetText 3, 0, "Conteúdo"
        .SetText 4, 0, "Envio Obrigatório XML"
        
        .EditEnterAction = EditEnterActionDown
        .CursorStyle = CursorStyleArrow
        
        .ReDraw = True
        
    End With
        
    Exit Sub

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flInicializarSpread", 0

End Sub

Private Sub flLimpaCampos()

    On Error GoTo ErrorHandler

    With vasAdministracaoDados
        
        .BlockMode = True
        
        .Col = 3
        .Row = 1
        .Col2 = 4
        .Row2 = .MaxRows
        .Action = ActionClearText
        
        .BlockMode = False
        
    End With

    Exit Sub

ErrorHandler:
    
    fgRaiseError App.EXEName, TypeName(Me), "flLimpaCampos", 0

End Sub

Private Sub flSalvar()
    
#If EnableSoap = 1 Then
    Dim objMIU                              As MSSOAPLib30.SoapClient30
#Else
    Dim objMIU                              As A8MIU.clsMIU
#End If

On Error GoTo ErrorHandler

    If xmlParametrizacao.selectSingleNode("PARM_GERL/@PADRAO").Text = "1" Then
           
        Call fgAppendAttribute(xmlParametrizacao, "PARM_GERL", "Operacao", "Alterar")
        Call fgAppendAttribute(xmlParametrizacao, "PARM_GERL", "Objeto", "A8LQS.clsParametrosGerais")
        Call fgAppendNode(xmlParametrizacao, "PARM_GERL", "CO_TEXT_XML", "0")
        
        xmlParametrizacao.selectSingleNode("PARM_GERL/@PADRAO").Text = "0"
    End If
    
    Call flAtribuirValoresXML
    
    If Not flValidarPreenchimentoTags Then Exit Sub
    
    xmlParametrizacao.selectSingleNode("//*/@Operacao").Text = "Alterar"
    xmlParametrizacao.selectSingleNode("//*/@Objeto").Text = "A8LQS.clsParametrosGerais"
    xmlParametrizacao.selectSingleNode("//*/CO_TEXT_XML").Text = "0"
    
    fgCursor True
    
    Set objMIU = fgCriarObjetoMIU("A8MIU.clsMIU")
    Call objMIU.Executar(xmlParametrizacao.xml)
    Set objMIU = Nothing
    
    fgCursor
    
    If Me.Visible Then Call MsgBox("Parametrização atualizada com sucesso.", vbOKOnly, Me.Caption)
    
    Exit Sub

ErrorHandler:
    
    Set objMIU = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flSalvar", 0

End Sub

Private Function flValidarPreenchimentoTags() As Boolean
    
Dim lngCount                                As Long
Dim vntNomeTag                              As Variant
Dim vntConteudoTag                          As Variant

    On Error GoTo ErrorHandler

    flValidarPreenchimentoTags = False
    
    With Me.vasAdministracaoDados
        For lngCount = 1 To .MaxRows
            
            .GetText 5, lngCount, vntNomeTag
            
            If vntNomeTag <> vbNullString Then
                
                .BlockMode = False
                .Col = 4
                .Row = lngCount
                
                .GetText 3, lngCount, vntConteudoTag
            
                If vntConteudoTag = vbNullString And .value = True Then
                
                    frmMural.Display = "Parâmetro assinaldo como obrigatório sem preenchimento."
                    frmMural.IconeExibicao = IconExclamation
                    frmMural.Show vbModal
                    
                    .BlockMode = False
                    .Col = 3
                    .Row = lngCount
                    .Action = ActionActiveCell
                    .SetFocus
                    
                    Exit Function
                End If
                    
            End If
        Next
    End With
    
    flValidarPreenchimentoTags = True
    
    Exit Function

ErrorHandler:
    fgRaiseError App.EXEName, TypeName(Me), "flValidarPreenchimentoTags", 0

End Function

Private Sub cboGrupo_Click()

On Error GoTo ErrorHandler
    
    Call flCarregarParametrizacao
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - cboGrupo_Click", Me.Caption

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
On Error GoTo ErrorHandler
    
    KeyAscii = fgBloqueiaCaracterEspecial(KeyAscii)
    
    Exit Sub

ErrorHandler:
    
    mdiLQS.uctlogErros.MostrarErros Err, TypeName(Me) & " - Form_KeyPress"

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorHandler
    
    Me.Icon = mdiLQS.Icon
    fgCenterMe Me
    
    Call flInicializarSpread
    
    Set xmlParametrizacao = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlParametrizacao.preserveWhiteSpace = True
    
    With Me.cboGrupo
        .AddItem "<-- Todos -->"
        .AddItem "Base Histórica"
        .AddItem "DV"
        .AddItem "BG"
        .AddItem "HA"
        .ListIndex = enumGrupo.Todos
    End With
    
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - Form_Load", Me.Caption

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set xmlParametrizacao = Nothing
    Set frmCadastroParametrosGerais = Nothing

End Sub

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Salvar"
            Call flSalvar
        
        Case "Limpar"
            Call flLimpaCampos
        
        Case "Sair"
            Unload Me
            
    End Select
        
    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick", Me.Caption

End Sub

Private Sub vasAdministracaoDados_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

Dim vntConteudo                             As Variant

    On Error GoTo ErrorHandler

    With vasAdministracaoDados
        .BlockMode = False

        .Col = NewCol
        .Row = NewRow

        .GetText 2, NewRow, vntConteudo
        
        If NewCol < 3 Or vntConteudo = vbNullString Then
            .Lock = True
            .Protect = True
        Else
            .Lock = False
            .Protect = False
        End If
    End With

    Exit Sub

ErrorHandler:
    mdiLQS.uctlogErros.MostrarErros Err, Me.Name & " - vasAdministracaoDados_LeaveCell", Me.Caption

End Sub
