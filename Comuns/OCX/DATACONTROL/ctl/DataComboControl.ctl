VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl DataComboControl 
   BackColor       =   &H8000000C&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2460
   KeyPreview      =   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   2460
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   450
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   -90
      TabIndex        =   1
      Top             =   360
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   22740993
      CurrentDate     =   37812
   End
   Begin VB.Label lblData 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2003"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   2265
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000010&
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
      Height          =   270
      Left            =   2040
      TabIndex        =   2
      Top             =   60
      Width           =   195
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   330
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      Visible         =   0   'False
      X1              =   -30
      X2              =   2460
      Y1              =   330
      Y2              =   330
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   2430
      X2              =   2430
      Y1              =   0
      Y2              =   330
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   0
      X2              =   2460
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "DataComboControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event Data(DataSelecionada As Date)
'Default Property Values:
Const m_def_Inicializar = 0
'Property Variables:
Dim m_Inicializar As Boolean
'Event Declarations:
Event Click() 'MappingInfo=lblData,lblData,-1,Click


Private Sub lblData_Click()
    RaiseEvent Click

    UserControl.Height = 2745
    UserControl.SetFocus
    MonthView1.ZOrder 0

End Sub

Private Sub lblData_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    If m_Inicializar = True Then
        Timer1.Enabled = False
    
        Line1.BorderColor = vbBlack
        Line4.BorderColor = vbBlack
        Line2.BorderColor = vbWhite
        Line3.BorderColor = vbWhite
        
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
    End If

End Sub

Private Sub lblData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    
    If m_Inicializar = True Then
        Line1.BorderColor = vbWhite
        Line4.BorderColor = vbWhite
        Line2.BorderColor = vbBlack
        Line3.BorderColor = vbBlack
    
        Line3.Visible = True
        Line4.Visible = True
        Line1.Visible = True
        Line2.Visible = True
    
        Timer1.Enabled = True
    End If

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    
    lblData.Caption = DateClicked
    
    RaiseEvent Data(DateClicked)
    
    UserControl.Height = 375
End Sub

Private Sub Timer1_Timer()
    Dim Rec As RECT, Point As POINTAPI
    ' Get Left, Right, Top and Bottom of Form1
    GetWindowRect hWnd, Rec
    
    ' Get the position of the cursor
    GetCursorPos Point

    ' If the cursor is located above the form then
    If Point.X >= Rec.Left And Point.X <= Rec.Right And Point.Y >= Rec.Top And Point.Y <= Rec.Bottom Then
        'Me.Caption = "MouseCursor is on form."
        Line1.Visible = True
        Line2.Visible = True
        Line3.Visible = True
        Line4.Visible = True
    Else
        ' The cursor is not located above the form
        'Me.Caption = "MouseCursor is not on form."
        Line1.Visible = False
        Line2.Visible = False
        Line3.Visible = False
        Line4.Visible = False
    End If
End Sub

Private Sub UserControl_Initialize()
    lblData.Caption = Date
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    lblData.Caption = Date
'    m_hWnd = m_def_hWnd
    m_Inicializar = m_def_Inicializar
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next
    
    If KeyCode = vbKeyEscape Then
        UserControl.Height = 375
    End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'    m_hWnd = PropBag.ReadProperty("hWnd", m_def_hWnd)
    m_Inicializar = PropBag.ReadProperty("Inicializar", m_def_Inicializar)
End Sub

Private Sub UserControl_Resize()
    
    If m_Inicializar = False Then
        UserControl.Height = 375
        UserControl.Width = 2460
    End If
    
End Sub

Private Sub UserControl_Terminate()
    Timer1.Enabled = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'    Call PropBag.WriteProperty("hWnd", m_hWnd, m_def_hWnd)
    Call PropBag.WriteProperty("Inicializar", m_Inicializar, m_def_Inicializar)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,00
Public Property Get Inicializar() As Boolean
    Inicializar = m_Inicializar
End Property

Public Property Let Inicializar(ByVal New_Inicializar As Boolean)
    m_Inicializar = New_Inicializar
    PropertyChanged "Inicializar"
    
    If New_Inicializar Then
        Timer1.Enabled = True
    End If

End Property


