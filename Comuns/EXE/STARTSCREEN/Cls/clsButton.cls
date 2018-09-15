VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private mvarIndex                           As Long
Private mvarCaption                         As String
Private mvarDescription                     As String
Private mvarKey                             As String
Private mvarTag                             As String
Private mvarToolTip                         As String
Private mvarImage                           As Variant
Private mvarVisible                         As Boolean
Private mvarEnabled                         As Boolean

Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

Public Property Let Visible(ByVal vData As Boolean)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Visible = 5
    mvarVisible = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Visible Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Visible() As Boolean
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Visible
    Visible = mvarVisible
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Visible Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Enabled(ByVal vData As Boolean)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Enabled = 5
    mvarEnabled = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Enabled Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Enabled() As Boolean
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Enabled
    Enabled = mvarEnabled
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Enabled Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Index(ByVal vData As Long)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Index Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Index() As Long
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Index Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Image(ByVal vData As Variant)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Image Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Image() As Variant
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Image Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let ToolTip(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ToolTip = 5
    mvarToolTip = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "ToolTip Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get ToolTip() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ToolTip
    ToolTip = mvarToolTip
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "ToolTip Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Key(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Key Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Key() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Key Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Tag(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Tag = 5
    mvarTag = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Tag Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Tag() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Tag
    Tag = mvarTag
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Tag Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Description(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Descroption Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Description() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Description Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Let Caption(ByVal vData As String)
On Error GoTo ErrorHandler

'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Caption = 5
    mvarCaption = vData
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Caption Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Property Get Caption() As String
On Error GoTo ErrorHandler

'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "clsButton", "Caption Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property
