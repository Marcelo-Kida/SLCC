VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLogTabela 
   Caption         =   "SLCC Trilha de Auditoria"
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   1515
   ClientWidth     =   11310
   Icon            =   "frmLogTabela.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   11310
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      Height          =   4800
      Left            =   5400
      ScaleHeight     =   2090.126
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   72
   End
   Begin VB.Frame fraFiltro 
      Caption         =   "Filtro por data"
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
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdAtualizar 
         Caption         =   "Atualizar"
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
         Left            =   5010
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDataInicio 
         Height          =   330
         Left            =   1020
         TabIndex        =   2
         Top             =   375
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/mm/yyyy"
         Format          =   104005633
         CurrentDate     =   37294
      End
      Begin MSComCtl2.DTPicker dtpDataFim 
         Height          =   330
         Left            =   3315
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   104005633
         CurrentDate     =   37294
      End
      Begin VB.Label Label1 
         Caption         =   "Início :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Fim :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   4
         Top             =   405
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lstLog 
      Height          =   3000
      Left            =   2055
      TabIndex        =   7
      Top             =   1800
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   5292
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.TreeView treLog 
      Height          =   4800
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   8467
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      SingleSel       =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgSplitter 
      Height          =   4785
      Left            =   1965
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   150
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000A&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   2040
      TabIndex        =   9
      Tag             =   " ListView:"
      Top             =   1395
      Width           =   3210
   End
End
Attribute VB_Name = "frmLogTabela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Formulario de consulta às tabelas de log A6, A7, A8 para trilha de auditoria

Option Explicit

Const sglSplitLimit = 500

Private blnMoving                                             As Boolean
Private mOwner                                                As enumSLCCOwner

Public Enum enumTipoOperLog
    LogUpdate = 1
    LogDelete = 2
    LogInsert = 3
End Enum


Private Type udtTabelaLog
    NomeTabela                As String * 50
    NomeTabelaLog             As String * 50
End Type

Private Type udtTabelaLogAux
    String                    As String * 100
End Type

Public Property Let Owner(ByVal pintOwnwer As enumSLCCOwner)
    mOwner = pintOwnwer
End Property
Private Sub cmdAtualizar_Click()

Dim udtLogTabela                                         As udtTabelaLog
Dim udtLogTabelaAux                                      As udtTabelaLogAux

On Error GoTo ErrorHandler
    
    If CDate(dtpDataInicio.Value) > CDate(dtpDataFim.Value) Then
        MsgBox "Data Início deve ser Menor ou Igual a data Fim", vbInformation
        Exit Sub
    End If
   
    If treLog.SelectedItem Is Nothing Then
        MsgBox "Selecione uma tabela.", vbInformation
        Exit Sub
    End If
    
    fgCursor True
    
    If treLog.SelectedItem.Tag <> "Tabelas" And treLog.SelectedItem.Tag <> "Owner" Then
        
        lstLog.ListItems.Clear
        lstLog.ColumnHeaders.Clear
        
        udtLogTabelaAux.String = treLog.SelectedItem.Tag
        LSet udtLogTabela = udtLogTabelaAux
        
        lblTitle(1).Caption = "  Log da Tabela : [ " & Trim(udtLogTabela.NomeTabela) & " ]"
        
        flCarregaCabecalhoLstLog Trim(udtLogTabela.NomeTabelaLog)
        
        flCarregaLogTabela Trim(udtLogTabela.NomeTabelaLog)
    
    Else
        
        lblTitle(1).Caption = ""
        lstLog.ListItems.Clear
        lstLog.ColumnHeaders.Clear
    
    End If

    fgCursor False
    
    Exit Sub
ErrorHandler:
    
    fgCursor False
    
    mdiTrilhaAuditoria.uctLogErros.MostrarErros Err, "frmLogTabela - cmdAtualizar_Click"
    
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    
    picSplitter.Visible = True
    
    blnMoving = True

End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

Dim sglPos As Single
    
    If blnMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If

End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    blnMoving = False

End Sub
Private Sub treLog_DragDrop(Source As Control, x As Single, Y As Single)
    
    If Source = imgSplitter Then
        SizeControls x
    End If

End Sub

'redifinir o tamanho dos componentes dentro da tela


Sub SizeControls(x As Single)

On Error Resume Next

    'set the width
    If x < 2000 Then x = 2500
    If x > (Me.Width - 2500) Then x = Me.Width - 2500
    treLog.Width = x
    
    imgSplitter.Left = x
    fraFiltro.Top = 0
       
    fraFiltro.Left = x + 100
    lstLog.Left = x + 80
    
    fraFiltro.Width = Me.Width - (treLog.Width) - 250
    lstLog.Width = Me.Width - (treLog.Width) - 250
    lblTitle(1).Top = fraFiltro.Height + 20
    
    
    lblTitle(1).Left = lstLog.Left + 20
    lblTitle(1).Width = lstLog.Width - 30

    lstLog.Top = lblTitle(1).Top + lblTitle(1).Height + 30
    'set the top
    treLog.Top = 10

    treLog.Height = Me.Height - 450
    lstLog.Height = Me.Height - 1650
    
    imgSplitter.Top = treLog.Top
    imgSplitter.Height = treLog.Height


End Sub

Private Sub treLog_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo ErrorHandler
    
    lstLog.ListItems.Clear
    lstLog.ColumnHeaders.Clear
    lblTitle(1).Caption = ""
    
    Exit Sub
ErrorHandler:

    mdiTrilhaAuditoria.uctLogErros.MostrarErros Err, "frmLogTabela - treLog"
    
End Sub

Private Sub Form_Load()

On Error GoTo ErrorHandler
    
    fgCenterMe Me

    SizeControls 2500
    
    Select Case mOwner
        
        Case enumSLCCOwner.OwnerA6
            Me.Caption = Me.Caption & " - A6 SBR"
        Case enumSLCCOwner.OwnerA6Coli
            Me.Caption = Me.Caption & " - A6 Coligadas"
        Case enumSLCCOwner.OwnerA7
            Me.Caption = Me.Caption & " - A7 Bus de Interface"
        Case enumSLCCOwner.OwnerA8
            Me.Caption = Me.Caption & " - A8 LQS"
    
    End Select
    
    Me.Caption = Me.Caption
    
    Me.Show
    
    dtpDataInicio.Value = Date
    dtpDataFim.Value = Date
         
    CarregaLogControle
         
    Exit Sub
ErrorHandler:

    mdiTrilhaAuditoria.uctLogErros.MostrarErros Err, "frmLogTabela - Form_Load"
    
End Sub

'Carregar o treeview com as tabelas que possuem log de trilha de auditoria


Public Sub CarregaLogControle()

Dim xmlLogTabela                                        As MSXML2.DOMDocument40
Dim xmlNode                                             As MSXML2.IXMLDOMNode
Dim lNode                                               As MSComctlLib.Node
Dim lsOwner                                             As String
Dim llCont                                              As Long
Dim udtTabelaLog                                        As udtTabelaLog
Dim udtLogTabelaAux                                     As udtTabelaLogAux
Dim lsNomeTabela                                        As String

#If EnableSoap = 1 Then
    Dim objTabelaLog                                    As MSSOAPLib30.SoapClient30
#Else
    Dim objTabelaLog                                    As A6A7A8Miu.clsLogTabela
#End If

On Error GoTo ErrorHandler
    
    treLog.Nodes.Clear
    
    lblTitle(1).Caption = ""
    
    Set objTabelaLog = fgCriarObjetoMIU("A6A7A8Miu.clsLogTabela")
    
    Set xmlLogTabela = CreateObject("MSXML2.DOMDocument.4.0")
        
    lsOwner = "SLCC"
    
    Select Case mOwner
        Case enumSLCCOwner.OwnerA6
            lsOwner = "A6 - SBR"
        Case enumSLCCOwner.OwnerA6Coli
            lsOwner = "A6 - Coligadas"
        Case enumSLCCOwner.OwnerA7
            lsOwner = "A7 - Bus de Interface"
        Case enumSLCCOwner.OwnerA8
            lsOwner = "A8 - LQS"
    
    End Select
    
    
    xmlLogTabela.loadXML objTabelaLog.LerTodosControleLog(mOwner)
    
    Set lNode = treLog.Nodes.Add(, , "Tabelas", "Tabelas")
    lNode.Tag = "Tabelas"
    lNode.Expanded = True
    Set lNode = treLog.Nodes.Add("Tabelas", tvwChild, lsOwner, lsOwner)
    lNode.Tag = "Owner"
    lNode.Expanded = True
    
    
    For Each xmlNode In xmlLogTabela.selectNodes("//Repeat_ControleTabelaLog/*")
        With xmlNode
            Set lNode = treLog.Nodes.Add(lsOwner, tvwChild, .selectSingleNode("NO_TABE").Text, UCase(Trim((.selectSingleNode("NO_LOGI_TABE").Text))))
            udtTabelaLog.NomeTabela = Trim(.selectSingleNode("NO_TABE").Text)
            udtTabelaLog.NomeTabelaLog = Trim(.selectSingleNode("NO_TABE_LOG").Text)
            
            LSet udtLogTabelaAux = udtTabelaLog
            lNode.Tag = UCase(Trim(udtLogTabelaAux.String))
        End With
    Next
    
    Set objTabelaLog = Nothing
    Set xmlLogTabela = Nothing
    
    Exit Sub
ErrorHandler:
    Set objTabelaLog = Nothing
    Set xmlLogTabela = Nothing

    Err.Raise Err.Number, Err.Source, Err.Description

End Sub

'Carregar o cabeçalho do listview dinamicamente com os nomes das colunas de uma tabela

Private Sub flCarregaCabecalhoLstLog(ByVal pstrNomdeTabela As String)

Dim xmlColunas                                          As MSXML2.DOMDocument40
Dim xmlNode                                             As MSXML2.IXMLDOMNode
Dim objCollumnHeader                                    As MSComctlLib.ColumnHeader

#If EnableSoap = 1 Then
    Dim objLogTabela                                    As MSSOAPLib30.SoapClient30
#Else
    Dim objLogTabela                                    As A6A7A8Miu.clsLogTabela
#End If

On Error GoTo ErrorHandler
    
    Set objLogTabela = fgCriarObjetoMIU("A6A7A8Miu.clsLogTabela")
    
    lstLog.ColumnHeaders.Clear
    
    Set xmlColunas = CreateObject("MSXML2.DOMDocument.4.0")
    
    xmlColunas.loadXML objLogTabela.ObterColunasTabela(pstrNomdeTabela)
        
    For Each xmlNode In xmlColunas.selectNodes("//Repeat_ColunaTabela/*")
   
        With xmlNode
            
            Set objCollumnHeader = lstLog.ColumnHeaders.Add(Val(.selectSingleNode("COLUMN_ID").Text), "K" & Trim(.selectSingleNode("COLUMN_ID").Text), Trim(.selectSingleNode("COLUMN_NAME").Text), 2000)
            objCollumnHeader.Tag = Trim(.selectSingleNode("COLUMN_NAME").Text)
            
        End With
    Next

    Set objLogTabela = Nothing
    Set xmlColunas = Nothing
        
    Exit Sub
ErrorHandler:
    Set objLogTabela = Nothing
    Set xmlColunas = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub form_Resize()
    SizeControls 1500
End Sub

'Carregar no listview os registros das tabelas de log

Private Sub flCarregaLogTabela(ByVal pstrNomdeTabelaLog As String)

Dim xmlColunas                                          As MSXML2.DOMDocument
Dim xmlColunasAux                                       As MSXML2.DOMDocument
Dim xmlNode                                             As MSXML2.IXMLDOMNode
Dim lListItem                                           As MSComctlLib.ListItem
Dim llCont                                              As Long
Dim llCodigoEmpresa                                     As Long
Dim llNumCol                                            As Long
Dim lsListText                                          As String
Dim lsTipoOperLog                                       As String
Dim lcColor                                             As ColorConstants
Dim strDataInicio                                       As String
Dim strDataFim                                          As String

#If EnableSoap = 1 Then
    Dim objLogTabela                                    As MSSOAPLib30.SoapClient30
#Else
    Dim objLogTabela                                    As A6A7A8Miu.clsLogTabela
#End If

On Error GoTo ErrorHandler
   
    lstLog.ListItems.Clear
    
    Set objLogTabela = fgCriarObjetoMIU("A6A7A8Miu.clsLogTabela")
    
    Set xmlColunas = CreateObject("MSXML2.DOMDocument")
    
    strDataInicio = CStr(Format(dtpDataInicio.Value, "DD/MM/YYYY"))
    strDataFim = CStr(Format(dtpDataFim.Value, "DD/MM/YYYY"))
    
    xmlColunas.loadXML objLogTabela.LerTodosLogTabela(pstrNomdeTabelaLog, mOwner, strDataInicio, strDataFim)
    
    If mOwner = OwnerA6Coli Then
        Set xmlColunasAux = CreateObject("MSXML2.DOMDocument")
        
        For Each xmlNode In xmlColunas.selectNodes("//Repeat_TabelaLog/*")
            xmlColunasAux.loadXML xmlNode.selectSingleNode("VL_CNTD_ATLZ").Text
        
            With xmlNode
                For llNumCol = 1 To lstLog.ColumnHeaders.Count
                    Select Case Val(.selectSingleNode("TP_ATLZ").Text)
                        Case enumTipoOperLog.LogInsert
                            lcColor = vbBlack
                            lsTipoOperLog = "Insert"
                        Case enumTipoOperLog.LogUpdate
                            lcColor = vbBlue
                            lsTipoOperLog = "Update"
                        Case enumTipoOperLog.LogDelete
                            lcColor = vbRed
                            lsTipoOperLog = "Delete"
                    End Select
                    
                    If llNumCol = 1 Then
                        lsListText = Trim(xmlColunasAux.selectSingleNode("//" & lstLog.ColumnHeaders.Item(llNumCol).Tag).Text)
    
                        If IsDate(Mid(lsListText, 1, 10)) Then
                            lsListText = Replace(lsListText, "T", " ")
                        End If
    
                        Set lListItem = lstLog.ListItems.Add(, , lsListText)
                        lListItem.ForeColor = lcColor
    
                    Else
                        If Not xmlColunasAux.selectSingleNode("//" & lstLog.ColumnHeaders.Item(llNumCol).Tag) Is Nothing Then
                            
                            lsListText = Trim(xmlColunasAux.selectSingleNode("//" & lstLog.ColumnHeaders.Item(llNumCol).Tag).Text)
                             
                            If Trim(xmlColunasAux.selectSingleNode("//" & lstLog.ColumnHeaders.Item(llNumCol).Tag).Text) = gdtmDataVazia Or _
                            Trim(xmlColunasAux.selectSingleNode("//" & lstLog.ColumnHeaders.Item(llNumCol).Tag).Text) = "" Then
                                lsListText = ""
                            ElseIf Left(lstLog.ColumnHeaders.Item(llNumCol).Tag, 3) = "DT_" Then
                                lsListText = fgDtXML_To_Date(lsListText)
                            ElseIf Left(lstLog.ColumnHeaders.Item(llNumCol).Tag, 3) = "DH_" Then
                                lsListText = fgDtHrStr_To_DateTime(lsListText)
                            End If

                            lListItem.SubItems(llNumCol - 1) = lsListText
                            lListItem.ListSubItems.Item(llNumCol - 1).ForeColor = lcColor

                        Else
                            lListItem.SubItems(llNumCol - 1) = ""
                            lListItem.ListSubItems.Item(llNumCol - 1).ForeColor = lcColor
                        End If
                    End If
                
                Next
             End With
            
        Next
        
        Exit Sub
    Else
        For Each xmlNode In xmlColunas.selectNodes("//Repeat_TabelaLog/*")
            
            With xmlNode
            
                For llNumCol = 1 To lstLog.ColumnHeaders.Count
    
                    If Not .selectSingleNode("IN_TIPO_OPER") Is Nothing Then
    
                        Select Case Val(.selectSingleNode("IN_TIPO_OPER").Text)
    
                            Case enumTipoOperLog.LogInsert
                                lcColor = vbBlack
                                lsTipoOperLog = "Insert"
                            Case enumTipoOperLog.LogUpdate
                                lcColor = vbBlue
                                lsTipoOperLog = "Update"
                            Case enumTipoOperLog.LogDelete
                                lcColor = vbRed
                                lsTipoOperLog = "Delete"
                        End Select
                    End If
    
                    If llNumCol = 1 Then
                        lsListText = Trim(.selectSingleNode(lstLog.ColumnHeaders.Item(llNumCol).Tag).Text)
    
                        If IsDate(Mid(lsListText, 1, 10)) Then
                            lsListText = Replace(lsListText, "T", " ")
                        End If
    
                        Set lListItem = lstLog.ListItems.Add(, , lsListText)
                        lListItem.ForeColor = lcColor
    
                    Else
                        If Not .selectSingleNode(lstLog.ColumnHeaders.Item(llNumCol).Tag) Is Nothing Then
                            If Trim(lstLog.ColumnHeaders.Item(llNumCol).Tag) = "IN_TIPO_OPER" Then
                                lListItem.SubItems(llNumCol - 1) = lsTipoOperLog
                                lListItem.ListSubItems.Item(llNumCol - 1).ForeColor = lcColor
                            Else
                                
                                lsListText = Trim(.selectSingleNode(lstLog.ColumnHeaders.Item(llNumCol).Tag).Text)
                                
                                If Trim(.selectSingleNode(lstLog.ColumnHeaders.Item(llNumCol).Tag).Text) = gdtmDataVazia Then
                                    lsListText = ""
                                ElseIf Left(lstLog.ColumnHeaders.Item(llNumCol).Tag, 3) = "DT_" Then
                                    lsListText = fgDtXML_To_Date(lsListText)
                                ElseIf Left(lstLog.ColumnHeaders.Item(llNumCol).Tag, 3) = "DH_" Then
                                    lsListText = fgDtHrStr_To_DateTime(lsListText)
                                End If
    
                                lListItem.SubItems(llNumCol - 1) = lsListText
                                lListItem.ListSubItems.Item(llNumCol - 1).ForeColor = lcColor
                            End If
                        Else
                            lListItem.SubItems(llNumCol - 1) = ""
                            lListItem.ListSubItems.Item(llNumCol - 1).ForeColor = lcColor
                        End If
                    End If
    
                Next
            End With
        Next
    End If
    
    Set objLogTabela = Nothing
    Set xmlColunas = Nothing
    
    Exit Sub
ErrorHandler:
    Set objLogTabela = Nothing
    Set xmlColunas = Nothing
        
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub


