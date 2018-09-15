VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{01DF10B2-3D7D-11D5-B2F8-0010B5AB2558}#1.1#0"; "NumBox.ocx"
Begin VB.Form frmAlerta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alerta A7 Bus "
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "frmAlerta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lstMensagens 
      Height          =   2565
      Left            =   90
      TabIndex        =   2
      Top             =   660
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   4524
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Excluir"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ocorrência"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Data Hora"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Source"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Erro"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   180
      Picture         =   "frmAlerta.frx":000C
      ScaleHeight     =   555
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   90
      Width           =   585
   End
   Begin NumBox.Number numTimer 
      Height          =   330
      Left            =   1920
      TabIndex        =   3
      Top             =   3270
      Visible         =   0   'False
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
      AceitaNegativo  =   0   'False
      SelStart        =   1
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2505
      TabIndex        =   4
      Top             =   3270
      Visible         =   0   'False
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
      Left            =   4140
      Top             =   45
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
            Picture         =   "frmAlerta.frx":044E
            Key             =   "DefinirFiltro"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":0560
            Key             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":087A
            Key             =   "Atualizar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":0BCC
            Key             =   "AplicarFiltro"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":0CDE
            Key             =   "Salvar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":0FF8
            Key             =   "Excluir"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1312
            Key             =   "Pendente"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":162C
            Key             =   "Limpar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1946
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1C98
            Key             =   "amarelo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":1FEA
            Key             =   "laranja"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAlerta.frx":233C
            Key             =   "vermelho"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbComandosForm 
      Height          =   330
      Left            =   6900
      TabIndex        =   6
      Top             =   3315
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
      Top             =   3300
      Visible         =   0   'False
      Width           =   1755
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
      Left            =   840
      TabIndex        =   0
      Top             =   210
      Width           =   1095
   End
End
Attribute VB_Name = "frmAlerta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Objeto responsável por alertar o usuario sobre existência de incosnsistências no sistema (Acesso MQ, Banco de dados fora do ar)
Option Explicit

Public strxmlInfoAlerta                     As String

Private Sub Form_Load()

    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 100, 200, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
    numTimer.Valor = GetSetting("A7", "Alerta", "Tempo Alerta", 1)
    
    mdiBUS.tmrAlerta.Enabled = False
            
    fgCenterMe Me
    
    Me.Width = 8300
    Me.Height = 4000
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call SaveSetting("A7", "Alerta", "Tempo Alerta", numTimer.Valor)

    mdiBUS.tmrAlerta.Enabled = True
    mdiBUS.tmrAlerta.Interval = 60000

    Unload Me

End Sub

Private Sub tlbComandosForm_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ErrorHandler

    fgCursor True
    
    Call SaveSetting("A7", "Alerta", "Tempo Alerta", numTimer.Valor)
    
    Call flExcluirAlerta

    mdiBUS.tmrAlerta.Enabled = True
    mdiBUS.tmrAlerta.Interval = 30000

    fgCursor

    Unload Me
    
    Exit Sub
ErrorHandler:
    fgCursor
    mdiBUS.uctLogErros.MostrarErros Err, "frmAlerta - tlbComandosForm_ButtonClick"

End Sub

Private Sub UpDown1_Change()
    numTimer.Valor = UpDown1.Value
End Sub

'Excluir os alertas selecionados na lsta de alertas.

Private Sub flExcluirAlerta()

#If EnableSoap = 1 Then
    Dim objAlerta   As MSSOAPLib30.SoapClient30
#Else
    Dim objAlerta   As A7Miu.clsAlerta
#End If

Dim xmlAlerta       As MSXML2.DOMDocument40
Dim xmlNode         As MSXML2.IXMLDOMNode
Dim objListItem     As ListItem
Dim strInfoAlerta   As String
Dim lngCont         As Long
Dim blnExcluir      As Boolean
Dim vntCodErro      As Variant
Dim vntMensagemErro As Variant

On Error GoTo ErrorHandler

    Set xmlAlerta = CreateObject("MSXML2.DOMDocument.4.0")

    If xmlAlerta.loadXML(strxmlInfoAlerta) Then
        blnExcluir = True
    Else
        blnExcluir = False
    End If

    If Trim(strxmlInfoAlerta) = "" Or Trim(strxmlInfoAlerta) = "0" Then Exit Sub

    While blnExcluir = True

        lngCont = 0
        blnExcluir = False

        For Each objListItem In lstMensagens.ListItems
            lngCont = lngCont + 1

            If objListItem.Checked Then
                Set xmlNode = xmlAlerta.selectSingleNode("Alerta/Repet_Alerta/Grupo_Alerta[" & lngCont & "]")

                If Not xmlNode Is Nothing Then
                    xmlAlerta.selectSingleNode("//Alerta/Repet_Alerta").removeChild xmlNode
                    lstMensagens.ListItems.Remove objListItem.Index
                    Exit For
                End If
            End If
        Next objListItem

        For Each objListItem In lstMensagens.ListItems
            If objListItem.Checked Then blnExcluir = True
        Next

    Wend

    If Not xmlAlerta.selectSingleNode("Alerta/Repet_Alerta") Is Nothing Then
        If xmlAlerta.selectSingleNode("Alerta/Repet_Alerta").childNodes.length = 0 Then
            strInfoAlerta = ""
        Else
            strInfoAlerta = xmlAlerta.xml
        End If
    Else
        strInfoAlerta = ""
    End If

    Set xmlAlerta = Nothing

    Set objAlerta = fgCriarObjetoMIU("A7Miu.clsAlerta")

    Call objAlerta.ConfigurarPropriedadeAlerta(strInfoAlerta, vntCodErro, vntMensagemErro)

    If vntCodErro <> 0 Then
        GoTo ErrorHandler
    End If

    Set objAlerta = Nothing

    Exit Sub

ErrorHandler:
   Set objAlerta = Nothing

    If vntCodErro <> 0 Then
        Err.Number = vntCodErro
        Err.Description = vntMensagemErro
    End If

   fgRaiseError App.EXEName, TypeName(Me), "flExcluirAlerta", 0, , strxmlInfoAlerta

End Sub
