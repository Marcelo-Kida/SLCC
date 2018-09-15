VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSobre 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre"
   ClientHeight    =   3960
   ClientLeft      =   2985
   ClientTop       =   1425
   ClientWidth     =   6585
   ClipControls    =   0   'False
   Icon            =   "frmSobre.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2733.263
   ScaleMode       =   0  'User
   ScaleWidth      =   6183.655
   Begin MSComctlLib.ListView lvwSistema 
      Height          =   1755
      Left            =   60
      TabIndex        =   6
      Top             =   1680
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   3096
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
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSistema 
      Cancel          =   -1  'True
      Caption         =   "&Sistema"
      Height          =   375
      Left            =   4620
      TabIndex        =   1
      Top             =   3540
      Width           =   915
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      Default         =   -1  'True
      Height          =   375
      Left            =   5580
      TabIndex        =   0
      Top             =   3540
      Width           =   915
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   120
      Picture         =   "frmSobre.frx":030A
      ScaleHeight     =   331.61
      ScaleMode       =   0  'User
      ScaleWidth      =   331.61
      TabIndex        =   2
      Top             =   180
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      X1              =   56.343
      X2              =   5267.141
      Y1              =   1118.153
      Y2              =   1118.153
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1020
      TabIndex        =   3
      Top             =   900
      Width           =   5475
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
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
      Height          =   255
      Left            =   1020
      TabIndex        =   4
      Top             =   180
      Width           =   5475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   901.49
      X2              =   6112.287
      Y1              =   1118.153
      Y2              =   1118.153
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
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
      Left            =   1020
      TabIndex        =   5
      Top             =   540
      Width           =   5475
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Objeto responsável pela exibição das informações pertinentes aos sistema A6, A7 e A8.

Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Const COL_COMPONENTE                As Integer = 0
Private Const COL_VERSAO                    As Integer = 1
Private Const COL_DATA                      As Integer = 2

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub cmdSistema_Click()
  Call StartSysInfo
End Sub

Private Sub cmdSair_Click()
  Unload Me
End Sub

Private Sub Form_Load()

Dim strMSG                                  As String
    
    fgCursor True
    fgCenterMe Me
    
    Me.Caption = " Sobre o " & App.ProductName
    
    Set Me.Icon = fgIconeApp
    
    Set Me.picIcon = fgIconeApp
    
    lblVersion.Caption = "Versão :" & App.Major & "." & App.Minor & "." & App.Revision
    
    lblTitle.Caption = " Sobre o " & App.ProductName
    
    lblDescription.Caption = "Grupo Santander Banespa"
    flCarregarLvw
    
    fgCursor False
End Sub

'Formatar o listview.
Public Sub flFormatarLvw()

On Error GoTo ErrorHandler

    With lvwSistema.ColumnHeaders
        .Clear
        .Add , , "Componente"
        .Add , , "Versão"
        .Add , , "Data"
    End With


Exit Sub
ErrorHandler:

   fgRaiseError App.EXEName, TypeName(Me), "flFormatarLvw", 0
End Sub
'Carregar o listview com informações sobre versão dos componentes do sistema.
Public Sub flCarregarLvw()

Dim xmlVersoes                              As MSXML2.DOMDocument40
Dim objDomNode                              As MSXML2.IXMLDOMNode
Dim objListItem                             As MSComctlLib.ListItem

On Error GoTo ErrorHandler

    Set xmlVersoes = CreateObject("MSXML2.DOMDocument.4.0")
    Call xmlVersoes.loadXML(fgObterDetalhesVersoes)
    
    lvwSistema.ListItems.Clear
    flFormatarLvw
    
    For Each objDomNode In xmlVersoes.documentElement.childNodes
        Set objListItem = lvwSistema.ListItems.Add(, , objDomNode.selectSingleNode("Title").Text)
        objListItem.SubItems(COL_VERSAO) = objDomNode.selectSingleNode("Major").Text & "." & _
                                           objDomNode.selectSingleNode("Minor").Text & "." & _
                                           objDomNode.selectSingleNode("Revision").Text
        objListItem.SubItems(COL_DATA) = fgDtHrStr_To_DateTime(objDomNode.selectSingleNode("Date").Text)
    Next objDomNode
    
    Set xmlVersoes = Nothing

Exit Sub
ErrorHandler:
    Set xmlVersoes = Nothing
    fgRaiseError App.EXEName, TypeName(Me), "flCarregarLvw", 0
End Sub

'Apresentar a tela de informações do sistema do sistema aoperacional.
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr

    Dim rc As Long
    Dim SysInfoPath As String

    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"

        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If

    Call Shell(SysInfoPath, vbNormalFocus)

    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

'Obter valores de chaves do registro do sistema operacional.
Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...

    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size

    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value

    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors

    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select

    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit

GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function


