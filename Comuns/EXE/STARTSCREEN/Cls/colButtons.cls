VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Button"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Private llCodigoErroNegocio                 As Long
Private liNumeroSequencialErro              As Integer

'local variable to hold collection
Private mcol                                As Collection

Public Function Add(Optional Index As Long, _
                    Optional Key As String, _
                    Optional Caption As String, _
                    Optional Description As String, _
                    Optional Tag As String, _
                    Optional ToolTip As String, _
                    Optional Image As Variant, _
                    Optional Visible As Boolean, _
                    Optional Enabled As Boolean) As clsButton
           
On Error GoTo ErrorHandler
           
    Dim objNewMember As clsButton
    Set objNewMember = New clsButton


    'set the properties passed into the method
    objNewMember.Index = Index
    objNewMember.Key = Key
    objNewMember.Caption = Caption
    objNewMember.Description = Description
    objNewMember.Tag = Tag
    objNewMember.ToolTip = ToolTip
    objNewMember.Image = Image
    objNewMember.Visible = Visible
    objNewMember.Enabled = Enabled
    
    If Key = Empty Then
        mcol.Add objNewMember
    Else
        mcol.Add objNewMember, Key
    End If

    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing

    Exit Function
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Add Function", llCodigoErroNegocio, liNumeroSequencialErro)

End Function

Public Property Get Item(vntIndexKey As Variant) As clsButton
Attribute Item.VB_UserMemId = 0
On Error GoTo ErrorHandler

Dim objButton                               As clsButton
    
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  
    If IsNumeric(vntIndexKey) Then
        Set Item = mcol(vntIndexKey)
    Else
        For Each objButton In Me
            If objButton.Key = vntIndexKey Then
                Set Item = objButton
                Exit For
            End If
        Next
    End If

    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Item Property", llCodigoErroNegocio, liNumeroSequencialErro)

End Property

Public Property Get Count() As Long
On Error GoTo ErrorHandler
    
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mcol.Count
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Count Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Public Sub Remove(vntIndexKey As Variant)
On Error GoTo ErrorHandler

    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    mcol.Remove vntIndexKey
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Remove Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo ErrorHandler

    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mcol.[_NewEnum]
    
    Exit Property
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "NewEnum Property", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Property

Private Sub Class_Initialize()
On Error GoTo ErrorHandler

    'creates the collection when this class is created
    Set mcol = New Collection
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Class_Initialize Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub

Private Sub Class_Terminate()
On Error GoTo ErrorHandler

    'destroys collection when this class is terminated
    Set mcol = Nothing
    
    Exit Sub
ErrorHandler:
    If llCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "colButtons", "Class_Terminate Sub", llCodigoErroNegocio, liNumeroSequencialErro)
    
End Sub
