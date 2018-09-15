VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"colButtons"
Option Explicit

Private mvarIndex                           As Long
Private mvarCaption                         As String
Private mvarDescription                     As String
Private mvarKey                             As String
Private mvarVisible                         As Boolean
Private mvarEnabled                         As Boolean

Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

Private mvarcolButtons                      As colButtons

Public Property Let Enabled(ByVal vData As Boolean)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Enabled = 5
    mvarEnabled = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Enabled Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Enabled() As Boolean
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Enabled
    Enabled = mvarEnabled
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Enabled Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Visible(ByVal vData As Boolean)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    mvarVisible = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Visible Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Visible() As Boolean
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Visible Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Buttons() As colButtons
On Error GoTo ErrorHandler

    If mvarcolButtons Is Nothing Then
        Set mvarcolButtons = New colButtons
    End If

    Set Buttons = mvarcolButtons
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Buttons Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Set Buttons(vData As colButtons)
On Error GoTo ErrorHandler

    Set mvarcolButtons = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Buttons Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    Set mvarcolButtons = Nothing
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Class_Terminate Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub

Public Property Let Description(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Description Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Description() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Description Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Key(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Key Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Key() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Key Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Caption(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Caption Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Caption() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Caption Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Index(ByVal vData As Long)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Index Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Index() As Long
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsTab", "Index Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property
