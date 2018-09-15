VERSION 5.00
Begin VB.UserControl ctlMenu 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   1815
   Begin VB.Label lblMenu 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu mnuMarcarDesmarcar 
      Caption         =   "mnuMarcarDesmarcar"
      Begin VB.Menu mnuMarcarDesmarcarMarcarTodas 
         Caption         =   "Marcar Todas"
      End
      Begin VB.Menu mnuMarcarDesmarcarDesmarcarTodas 
         Caption         =   "Desmarcar Todas"
      End
   End
End
Attribute VB_Name = "ctlMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ClickMenu(ByVal Retorno As Long)

Public Sub ShowMenuMarcarDesmarcar(Optional ByVal pbAprovar As Boolean = False)
    PopupMenu mnuMarcarDesmarcar
End Sub

Private Sub mnuMarcarDesmarcarDesmarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.DesmarcarTodas)
End Sub

Private Sub mnuMarcarDesmarcarMarcarTodas_Click()
    RaiseEvent ClickMenu(enumTipoSelecao.MarcarTodas)
End Sub
