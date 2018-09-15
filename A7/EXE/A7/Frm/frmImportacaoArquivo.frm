VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportacaoArquivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importação de Arquivos"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRegistrosErro 
      Caption         =   "Registros com Erro"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   6975
      Begin VB.ListBox lstErro 
         Appearance      =   0  'Flat
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.Frame fraInfo 
      Caption         =   "Informações"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.Label lblErro 
         AutoSize        =   -1  'True
         Caption         =   "Registro(s) com erro."
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   960
         Width           =   1530
      End
      Begin VB.Label lblQtdeErro 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "00000"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblSucesso 
         AutoSize        =   -1  'True
         Caption         =   "Registro(s) importados com sucesso."
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   2640
      End
      Begin VB.Label lblQtdeSucesso 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "00000"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   450
      End
   End
   Begin MSComDlg.CommonDialog dlgImportarArquivo 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   4710
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
            Picture         =   "frmImportacaoArquivo.frx":0000
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":031A
            Key             =   "Padrao"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":076C
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":0A86
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":0DA0
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":10BA
            Key             =   "ItemGrupoFechado"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":150C
            Key             =   "ItemGrupoAberto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":195E
            Key             =   "ItemElementar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":1DB0
            Key             =   "checkfalse"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmImportacaoArquivo.frx":1E4A
            Key             =   "checktrue"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCadastro 
      Height          =   330
      Left            =   6300
      TabIndex        =   7
      Top             =   4725
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ButtonWidth     =   1376
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
            Object.Visible         =   0   'False
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
Attribute VB_Name = "frmImportacaoArquivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const intQtdeCaracteres             As Integer = 256
Private vntLinhasErro                       As Variant
Private vntQtCaracteres                     As Variant

Private Sub flMostrarRegistrosErro()

Dim intQtdeErro                             As Integer
Dim intErro                                 As Integer
    
    On Error GoTo ErrorHandler
    
    intQtdeErro = UBound(vntLinhasErro)

    For intErro = 0 To intQtdeErro - 1
        lstErro.AddItem "ERRO / Linha: (" & vntLinhasErro(intErro) & ") - Caracteres esperados: 256. Caracteres da linha: " & vntQtCaracteres(intErro) & "."
    Next intErro
    
    Exit Sub

ErrorHandler:
    fgCursor False
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - flMostrarRegistrosErro"

End Sub

Private Sub Form_Load()

Dim strArquivo                              As String
Dim strDefaultPath                          As String
    
    On Error GoTo ErrorHandler
    
    Me.Icon = mdiBUS.Icon
    fgCenterMe Me
    
    Me.Show
    Me.ZOrder
    DoEvents
    
    strDefaultPath = GetSetting("A7", "Configuracao", "strDefaultPath")
    
    With dlgImportarArquivo
        .DialogTitle = Me.Caption
        .CancelError = False
        .Filter = "Arquivo Importação (*.txt)|*.txt"
        If strDefaultPath <> vbNullString Then
            .InitDir = strDefaultPath
        End If
        .ShowOpen
        
        If Len(.FileName) > 0 Then
            
            If MsgBox("Confirma a importação do Arquivo selecionado ?", vbYesNo + vbQuestion, "Confirmação para Importação de Arquivo") = vbYes Then
                
                fgCursor True
            
                strArquivo = .FileName
                strDefaultPath = flGetPath(strArquivo)
                Call SaveSetting("A7", "Configuracao", "strDefaultPath", strDefaultPath)
                Call flImportarRegistros(strArquivo)
            
            End If
            
        Else
            
            MsgBox "Nenhum Arquivo foi selecionado para Importação.", vbExclamation, Me.Caption
        
        End If
        
    End With

    fgCursor False

    Exit Sub

ErrorHandler:
    fgCursor False
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - Form_Load"

End Sub

Public Sub flImportarRegistros(strArquivo As String)

#If EnableSoap = 1 Then
    Dim objMonitoracao          As MSSOAPLib30.SoapClient30
#Else
    Dim objMonitoracao          As A7Miu.clsMonitoracao
#End If

Dim vntCodErro                  As Variant
Dim vntMensagemErro             As Variant
Dim strLinha                    As String
Dim intLinha                    As Integer
Dim intPosicao                  As Integer
            
    On Error GoTo ErrorHandler
            
    vntLinhasErro = Array(1)
    vntQtCaracteres = Array(1)
                
    Open strArquivo For Input As #1
    
        Set objMonitoracao = fgCriarObjetoMIU("A7Miu.clsMonitoracao")
    
        While Not EOF(1)
        
            Line Input #1, strLinha
            
            intLinha = intLinha + 1
            
            If Len(strLinha) = intQtdeCaracteres Then
            
                Call objMonitoracao.PostarMensagemArquivoImportado(Trim$(strLinha), vntCodErro, vntMensagemErro)
                                            
                If vntCodErro <> 0 Then
                    GoTo ErrorHandler
                End If
                
            Else
            
                intPosicao = UBound(vntLinhasErro)
                vntLinhasErro(intPosicao) = intLinha
                vntQtCaracteres(intPosicao) = Len(strLinha)
                intPosicao = intPosicao + 1
                ReDim Preserve vntLinhasErro(intPosicao)
                ReDim Preserve vntQtCaracteres(intPosicao)
                
            End If
        
        Wend
    
    Close #1
    
    Set objMonitoracao = Nothing
    
    lblQtdeSucesso.Caption = intLinha - UBound(vntLinhasErro)
    lblQtdeErro.Caption = UBound(vntLinhasErro)
    
    If UBound(vntLinhasErro) > 0 Then
        fgCursor False
        
        If MsgBox("Ocorreram erros na importação do Arquivo. Deseja visualizar os registros com erro ?", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            Call flMostrarRegistrosErro
        End If
    Else
        Call FileSystem.FileCopy(strArquivo, strArquivo & ".proces")
        Call FileSystem.Kill(strArquivo)
    End If
    
    Exit Sub
            
ErrorHandler:
    If Err.Number = 55 Then
        Close #1
        Resume
    End If
    
    Set objMonitoracao = Nothing
    
    Err.Number = vntCodErro
    Err.Description = vntMensagemErro
    
    Call mdiBUS.uctLogErros.MostrarErros(Err, "frmImportacaoArquivo - flImportarRegistros")

End Sub

Public Function flGetPath(strFile As String) As String
  
    Dim lngIPos          As Long
    Dim lngIPosPrev      As Long
  
    On Error GoTo ErrorHandler
    
    Do
        lngIPosPrev = lngIPos
        lngIPos = InStr(lngIPos + 1, strFile, "\", vbBinaryCompare)
    Loop While lngIPos > 0
  
    If lngIPosPrev > 0 Then
        flGetPath = Left$(strFile, lngIPosPrev)
    Else
        flGetPath = strFile
    End If
  
    Exit Function

ErrorHandler:
    fgCursor False
    mdiBUS.uctLogErros.MostrarErros Err, TypeName(Me) & " - flGetPath"

End Function

Private Sub tlbCadastro_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error GoTo ErrorHandler
    
    Select Case Button.Key
        Case "Sair"
            Unload Me
    End Select
        
    Exit Sub

ErrorHandler:
    mdiBUS.uctLogErros.MostrarErros Err, Me.Name & " - tlbCadastro_ButtonClick", Me.Caption

End Sub
