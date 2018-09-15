VERSION 5.00
Begin VB.UserControl ctlBarraVertical 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   PropertyPages   =   "ctlBarraVertical.ctx":0000
   ScaleHeight     =   7140
   ScaleWidth      =   2055
   ToolboxBitmap   =   "ctlBarraVertical.ctx":001B
   Begin VB.PictureBox picLeftPane 
      BackColor       =   &H8000000A&
      CausesValidation=   0   'False
      Height          =   6975
      Left            =   60
      MousePointer    =   1  'Arrow
      ScaleHeight     =   6915
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      Begin VB.PictureBox picInnerFrame 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Height          =   5775
         Left            =   195
         ScaleHeight     =   5775
         ScaleWidth      =   1455
         TabIndex        =   2
         Top             =   960
         Width           =   1455
         Begin VB.CommandButton cmdScrollDown 
            Height          =   255
            Left            =   1065
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2880
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CommandButton cmdScrollUp 
            Height          =   255
            Left            =   1065
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   540
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgImage 
            Height          =   495
            Index           =   0
            Left            =   420
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Label"
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
            Height          =   435
            Index           =   0
            Left            =   30
            TabIndex        =   5
            Top             =   780
            Visible         =   0   'False
            Width           =   1365
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CommandButton cmdTab 
         Caption         =   "Tab"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ctlBarraVertical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const SPLITTER_WIDTH = 50
Private Const BTN_HEIGHT = 315
Private Const SCROLL_UP = -BTN_HEIGHT
Private Const SCROLL_DOWN = BTN_HEIGHT
Private Const DRAW_HIDDEN = 0
Private Const DRAW_RAISED = 1
Private Const DRAW_INSET = 2
Private Const BOX_BORDER = 50
Private Const PIC_OFFSET = 300
Private Const PIC_SPACING = 700
Private Const LABEL_SPACING = 100

Private lngCodigoErroNegocio                As Long
Private intNumeroSequencialErro             As Integer

Private mcolTabs                            As colTabs

Private strImageListName                    As String
Private snglLeftPane                        As Single
Private intCurButton                        As Integer
Private intFirstVis                         As Integer
Private intLastCtl                          As Integer
Private intLastStat                         As Integer
Private intCurrentTab                       As Integer
Private intVisibleTabs                       As Integer

Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picLeftPane,picLeftPane,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picLeftPane,picLeftPane,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single) 'MappingInfo=picLeftPane,picLeftPane,-1,MouseUp
Public Event Click(ByVal TabIndex As Integer, ByVal TabKey As String, ByVal ButtonIndex As Integer, ByVal ButtonKey As String)      'MappingInfo=UserControl,UserControl,-1,Click

Public Sub ArrangeControls()

Dim intIndex                                 As Integer
Dim snglRightPane                            As Single

On Error GoTo ErrorHandler

    picLeftPane.Move 0, 0, snglLeftPane, ScaleHeight
    snglRightPane = (ScaleWidth - SPLITTER_WIDTH) - snglLeftPane
    
    If snglRightPane < 0 Then
        snglRightPane = 0
    End If
    
    DrawControls
    
    If cmdTab.UBound > 0 Then
        cmdTab(cmdTab.LBound + 1).Move picLeftPane.ScaleLeft, _
                                       picLeftPane.ScaleTop, _
                                       picLeftPane.ScaleWidth, BTN_HEIGHT
        
        intVisibleTabs = 0
        
        For intIndex = cmdTab.LBound + 1 To cmdTab.UBound
            If mcolTabs(intIndex).Visible Then
                If cmdTab(intIndex).Tag = "TOP" Then
                    cmdTab(intIndex).Move picLeftPane.ScaleLeft, _
                                         picLeftPane.ScaleTop + (BTN_HEIGHT * (intIndex - 1)), _
                                         picLeftPane.ScaleWidth, BTN_HEIGHT
                    intCurButton = intIndex
                Else
                    cmdTab(intIndex).Move picLeftPane.ScaleLeft, _
                                         picLeftPane.ScaleHeight - (BTN_HEIGHT * (cmdTab.Count - intIndex)), _
                                         picLeftPane.ScaleWidth, BTN_HEIGHT
                End If
                
                intVisibleTabs = intVisibleTabs + 1
            End If
        Next
    End If
    
    DrawInnerFrame
    DrawPics
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ArrangeControls Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub cmdScrollDown_Click()

On Error GoTo ErrorHandler
    
    intFirstVis = intFirstVis + 1
    
    If intFirstVis >= (imgImage.Count - 1) Then intFirstVis = (imgImage.Count - 1)
    
    If Not imgImage(intFirstVis).Visible Then
        intFirstVis = intFirstVis - 1
    Else
        DrawPics
    End If
    
    If intFirstVis > 0 Then cmdScrollUp.Visible = True
    
    Exit Sub

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "cmsScrollDown_Click Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub cmdScrollUp_Click()

On Error GoTo ErrorHandler

    If intFirstVis > 1 Then
        intFirstVis = intFirstVis - 1
        
        If intFirstVis = 1 Then
            cmdScrollUp.Visible = False
        End If
    End If
    
    DrawPics
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "cmdScrollUp_Click Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub imgImage_Click(Index As Integer)

Dim objTab                                  As clsTab
Dim objButton                               As clsButton

On Error GoTo ErrorHandler
    
    If intLastStat <> DRAW_HIDDEN Then
        intLastStat = DRAW_HIDDEN
        picInnerFrame.Refresh
    End If
    
    If mcolTabs.Count > 0 And intCurrentTab > 0 Then
        Set objTab = mcolTabs(intCurrentTab)
        
        If objTab.Buttons.Count > 0 And Index > 0 Then
            Set objButton = objTab.Buttons(Index)
            RaiseEvent Click(objTab.Index, objTab.Key, objButton.Index, objButton.Key)
        End If
    End If
    
    Set objTab = Nothing
    Set objButton = Nothing
    
    Exit Sub
ErrorHandler:
    
    Set objTab = Nothing
    Set objButton = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "imgImage_Click Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub SetActiveTab(Index As Integer)

Dim intIndex                                As Integer
Dim intStart                                As Integer
Dim intEnd                                  As Integer
Dim intStep                                 As Integer
Dim intDir                                  As Integer

On Error GoTo ErrorHandler

    intCurrentTab = Index

    If cmdTab(Index).Tag = "TOP" Then
        intStart = Index + 1
        intEnd = cmdTab.UBound
        intStep = 1
        intDir = SCROLL_DOWN
    Else
        intStart = Index
        intEnd = 1
        intStep = -1
        intDir = SCROLL_UP
    End If
    
    For intIndex = intStart To intEnd Step intStep
        ScrollBtn intIndex, intDir
    Next
    
    If intCurButton <> Index Then
        intCurButton = Index
        intFirstVis = 1
        cmdScrollUp.Visible = False
        
        DrawInnerFrame
        DrawPics
    End If

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "SetActiveTab Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub cmdTab_Click(Index As Integer)

On Error GoTo ErrorHandler

    SetActiveTab Index
    picLeftPane.SetFocus

    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "cmdTab_Click Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub picInnerFrame_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub picInnerFrame_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub DrawControls()

Dim objButton                               As clsButton
Dim objTab                                  As clsTab

On Error GoTo ErrorHandler

    For Each objTab In mcolTabs
        If objTab.Index > cmdTab.UBound Then
            Load cmdTab(objTab.Index)
        End If
            
        With cmdTab(objTab.Index)
            .Caption = objTab.Caption
            .Enabled = objTab.Enabled
            .Visible = objTab.Visible
        End With
    
        If objTab.Index = cmdTab(cmdTab.LBound).Index + 1 Then
            cmdTab(objTab.Index).Tag = "TOP"
            intCurrentTab = 1
        Else
            cmdTab(objTab.Index).Tag = "BOTTON"
        End If
        
        For Each objButton In objTab.Buttons
            If objButton.Index > imgImage.UBound Then
                Load imgImage(objButton.Index)
            End If
        
            If objButton.Index > lblLabel.UBound Then
                Load lblLabel(objButton.Index)
            End If
        Next
    Next

    Set objButton = Nothing
    Set objTab = Nothing

    Exit Sub
ErrorHandler:
    Set objButton = Nothing
    Set objTab = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "DrawControls Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler

    Set mcolTabs = New colTabs

    snglLeftPane = 2000
    intCurButton = 1
    intFirstVis = 1

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "UserControl_Initialize Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Dim objDomDocument                          As MSXML2.DOMDocument40
Dim objNodeList                             As MSXML2.IXMLDOMNodeList
Dim strXML                                  As String
Dim intTab                                  As Integer
Dim intButton                               As Integer
Dim objTab                                  As clsTab
Dim objButton                               As clsButton

On Error GoTo ErrorHandler
    
    strXML = PropBag.ReadProperty("Tabs", vbNullString)
    
    strImageListName = PropBag.ReadProperty("ImageListName", vbNullString)
    
    Set Picture = PropBag.ReadProperty("Picture", Nothing)


    Set objDomDocument = CreateObject("MSXML2.DOMDocument.4.0")
    
    objDomDocument.loadXML strXML
    
    Set mcolTabs = New colTabs
    
    Set objNodeList = objDomDocument.getElementsByTagName("TAB")
    
    For intTab = 0 To objNodeList.length - 1
        
        Set objTab = mcolTabs.Add()
        
        With objNodeList.Item(intTab).Attributes
            objTab.Index = .getNamedItem("Index").Text
            objTab.Caption = .getNamedItem("Caption").Text
            objTab.Description = .getNamedItem("Description").Text
            objTab.Key = .getNamedItem("Key").Text
            objTab.Enabled = .getNamedItem("Enabled").Text
            objTab.Visible = .getNamedItem("Visible").Text
        End With
        
        If objNodeList.Item(intTab).hasChildNodes Then
            For intButton = 0 To objNodeList.Item(intTab).childNodes.length - 1
                Set objButton = objTab.Buttons.Add()
            
                With objNodeList.Item(intTab).childNodes.Item(intButton).Attributes
                    objButton.Index = .getNamedItem("Index").Text
                    objButton.Caption = .getNamedItem("Caption").Text
                    objButton.Description = .getNamedItem("Description").Text
                    objButton.Key = .getNamedItem("Key").Text
                    objButton.Tag = .getNamedItem("Tag").Text
                    objButton.Image = .getNamedItem("Image").Text
                    objButton.ToolTip = .getNamedItem("Tooltip").Text
                    objButton.Enabled = .getNamedItem("Enabled").Text
                    objButton.Visible = .getNamedItem("Visible").Text
                End With
            Next
        End If
    Next
    
    ArrangeControls
    
    Set objDomDocument = Nothing
    Set objNodeList = Nothing
    Set objTab = Nothing
    Set objButton = Nothing
    
    Exit Sub
ErrorHandler:
    Set objDomDocument = Nothing
    Set objNodeList = Nothing
    Set objTab = Nothing
    Set objButton = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "UserControl_ReadProperties Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub UserControl_Resize()

On Error GoTo ErrorHandler
    
    snglLeftPane = UserControl.ScaleWidth - 40
    ArrangeControls
    DoEvents
    
    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "UserControl_Resize Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

Private Sub picInnerFrame_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo ErrorHandler

    If intLastStat <> DRAW_HIDDEN Then
        intLastStat = DRAW_HIDDEN
        picInnerFrame.Refresh
    End If

    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "picInnerFrame_MouseMove Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub imgImage_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error GoTo ErrorHandler

    Make3D imgImage(Index), DRAW_INSET

    Exit Sub
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "imgImage_MouseDown Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub imgImage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo ErrorHandler

    If Button = 0 Then
        intLastCtl = Index
        
        If intLastStat <> DRAW_RAISED Then
            intLastStat = DRAW_RAISED
            Make3D imgImage(Index), DRAW_RAISED
        End If
    Else
        If intLastStat = DRAW_HIDDEN Then
            picInnerFrame.Refresh
            intLastStat = DRAW_HIDDEN
        End If
    End If

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "imgImage_MouseMove Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub imgImage_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

On Error GoTo ErrorHandler

    If (x < imgImage(Index).Left Or x > (imgImage(Index).Left + imgImage(Index).Width)) Or _
       (Y < imgImage(Index).Top Or Y > (imgImage(Index).Top + imgImage(Index).Height)) Then
        
        'mouse foi solto fora do object
        If intLastStat <> DRAW_HIDDEN Then
            intLastStat = DRAW_HIDDEN
            picInnerFrame.Refresh
        End If
        Exit Sub
    End If
    
    If Index = intLastCtl Then
        Make3D imgImage(intLastCtl), DRAW_RAISED
    Else
        intLastStat = DRAW_HIDDEN
        picInnerFrame.Refresh
    End If

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "imgImage_MouseUp Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub ScrollBtn(iBtnIndex As Integer, iDir As Integer)

Dim intCont                                 As Integer
Dim intBtnStep                              As Integer
Dim intEndPos                               As Integer
Dim objTab                                  As clsTab
Dim intTabsOnTop                            As Integer

On Error GoTo ErrorHandler
    
    If iDir = SCROLL_UP Then
        
        For Each objTab In mcolTabs
            If objTab.Index >= iBtnIndex Then
                Exit For
            Else
                If objTab.Visible Then
                    intTabsOnTop = intTabsOnTop + 1
                End If
            End If
        Next
    
        intEndPos = picLeftPane.ScaleTop + (BTN_HEIGHT * intTabsOnTop)
        cmdTab(iBtnIndex).Tag = "TOP"
    Else
        intEndPos = picLeftPane.ScaleHeight - (BTN_HEIGHT * (cmdTab.Count - iBtnIndex))
        cmdTab(iBtnIndex).Tag = "BOTTOM"
    End If
    
    For intCont = cmdTab(iBtnIndex).Top To intEndPos Step iDir
        cmdTab(iBtnIndex).Move picLeftPane.ScaleLeft, intCont, picLeftPane.ScaleWidth, BTN_HEIGHT
    Next
    
    If intCont <> intEndPos Then cmdTab(iBtnIndex).Move picLeftPane.ScaleLeft, intEndPos, picLeftPane.ScaleWidth, BTN_HEIGHT

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ScrollBtn Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub Make3D(ctl As Control, iMode As Integer)

Dim lngTopLeftCol                             As Long
Dim lngBotRightCol                            As Long

On Error GoTo ErrorHandler


    Select Case iMode
        Case DRAW_INSET
            lngTopLeftCol = vbBlack
            lngBotRightCol = vbWhite
        Case DRAW_RAISED
            lngTopLeftCol = vbWhite
            lngBotRightCol = vbBlack
        Case Else
            Exit Sub
    
    End Select
    
    picInnerFrame.CurrentX = ctl.Left - BOX_BORDER
    picInnerFrame.CurrentY = ctl.Top - BOX_BORDER
    'left
    picInnerFrame.Line -(ctl.Left - BOX_BORDER, ctl.Top + ctl.Height + BOX_BORDER), lngTopLeftCol
    'bottom
    picInnerFrame.Line -(ctl.Left + ctl.Width + BOX_BORDER, ctl.Top + ctl.Height + BOX_BORDER), lngBotRightCol
    'right
    picInnerFrame.Line -(ctl.Left + ctl.Width + BOX_BORDER, ctl.Top - BOX_BORDER), lngBotRightCol
    'top
    picInnerFrame.Line -(ctl.Left - BOX_BORDER, ctl.Top - BOX_BORDER), lngTopLeftCol

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "Make3D Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Function ImageListValid(pControlName As String) As Boolean

Dim objControl                              As Control

On Error GoTo ErrorHandler

    For Each objControl In Parent.Controls
        If UCase(objControl.Name) = UCase(pControlName) Then
            If TypeName(objControl) = "ImageList" Then
                ImageListValid = True
                Exit For
            End If
        End If
    Next

    Set objControl = Nothing

    Exit Function
ErrorHandler:
    Set objControl = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ImageListValid Function", lngCodigoErroNegocio, intNumeroSequencialErro)

End Function

Private Sub DrawPics()

Dim objButton                               As clsButton
Dim objTab                                  As clsTab
Dim intIndex                                As Integer
Dim intLastVis                              As Integer

Const liButtonsPerTab                       As Integer = 4

On Error GoTo ErrorHandler
   
    If mcolTabs.Count > 0 Then
        Set objTab = mcolTabs(intCurrentTab)
        
        For intIndex = imgImage.LBound + 1 To imgImage.UBound
            
            If intIndex <= objTab.Buttons.Count Then
                Set objButton = objTab.Buttons(intIndex)
                
                If objTab.Visible And objButton.Visible Then
                    
                    If objButton.Index >= intFirstVis And _
                       objButton.Index < intFirstVis + liButtonsPerTab Then
                    
                        With imgImage(intIndex)
                            If objButton.Image <> vbNullString And strImageListName <> vbNullString And ImageListValid(strImageListName) Then
                                Set .Picture = Parent.Controls(strImageListName).ListImages.Item(objButton.Image).Picture
                            End If
                            
                            .Visible = objButton.Visible
                            .Enabled = objButton.Enabled
                            .Tag = objButton.Tag
                            .ToolTipText = objButton.ToolTip
                        End With
                        
                        With lblLabel(intIndex)
                            .Caption = objButton.Caption
                            .Visible = objButton.Visible
                        End With
                    Else
                        imgImage(intIndex).Visible = False
                        lblLabel(intIndex).Visible = False
                    End If
                                                    
                    intLastVis = intIndex
                Else
                    intFirstVis = intFirstVis + 1
                                        
                    imgImage(intIndex).Visible = False
                    lblLabel(intIndex).Visible = False
                End If
            Else
                imgImage(intIndex).Visible = False
                lblLabel(intIndex).Visible = False
            End If
                
            
            If intIndex = intFirstVis Then
                imgImage(intIndex).Move (picInnerFrame.ScaleWidth - imgImage(intIndex).Width) / 2, _
                                        picInnerFrame.ScaleTop + PIC_OFFSET
            ElseIf intIndex > intFirstVis Then
                imgImage(intIndex).Move (picInnerFrame.ScaleWidth - imgImage(intIndex).Width) / 2, _
                                        imgImage(intIndex - 1).Top + imgImage(intIndex - 1).Height + PIC_SPACING
            End If
                
            lblLabel(intIndex).Move (picInnerFrame.ScaleWidth - lblLabel(intIndex).Width) / 2, _
                                    imgImage(intIndex).Top + imgImage(intIndex).Height + LABEL_SPACING
        
        Next
    End If
    
    If intLastVis > intFirstVis And (lblLabel(intLastVis).Top + lblLabel(intLastVis).Height) > picInnerFrame.ScaleHeight Then
        cmdScrollDown.Visible = True
    Else
        cmdScrollDown.Visible = False
    End If

    Set objButton = Nothing
    Set objTab = Nothing

    Exit Sub
ErrorHandler:
    Set objButton = Nothing
    Set objTab = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "DrawPics Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub DrawInnerFrame()

On Error GoTo ErrorHandler

    If cmdTab.UBound > 0 And intVisibleTabs > 0 Then
        picInnerFrame.Move picLeftPane.ScaleLeft, _
                           cmdTab(intCurButton).Top + cmdTab(intCurButton).Height, _
                           picLeftPane.ScaleWidth, _
                           picLeftPane.ScaleHeight - (intVisibleTabs * cmdTab(intCurButton).Height)
    Else
        picInnerFrame.Move picLeftPane.ScaleLeft, _
                           picLeftPane.ScaleTop, _
                           picLeftPane.ScaleWidth, _
                           picLeftPane.ScaleHeight
    End If
    
    cmdScrollUp.Move picInnerFrame.ScaleWidth - cmdScrollUp.Width - 40, _
                     picInnerFrame.ScaleTop + 40
    
    cmdScrollDown.Move picInnerFrame.ScaleWidth - cmdScrollDown.Width - 40, _
                       picInnerFrame.ScaleHeight - cmdScrollDown.Height - 40

    Exit Sub
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "DrawInnerFrame Sub", lngCodigoErroNegocio, intNumeroSequencialErro)

End Sub

Private Sub picLeftPane_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub picLeftPane_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

Private Sub picLeftPane_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Dim strXML                                  As String
Dim objTab                                  As clsTab
Dim objButton                               As clsButton

On Error GoTo ErrorHandler

    strXML = "<XML>" & vbCr
    strXML = strXML & vbTab
    strXML = strXML & "<TABS>" & vbCr
    
    For Each objTab In mcolTabs
        strXML = strXML & vbTab & vbTab
        strXML = strXML & "<TAB Index= '" & objTab.Index & "'" & _
                            " Caption= '" & objTab.Caption & "'" & _
                            " Description= '" & objTab.Description & "'" & _
                            " Key= '" & objTab.Key & "'" & _
                            " Enabled= '" & objTab.Enabled & "'" & _
                            " Visible= '" & objTab.Visible & "'>" & vbCr
        
        For Each objButton In objTab.Buttons
            strXML = strXML & vbTab & vbTab & vbTab
            strXML = strXML & "<BUTTON Index= '" & objButton.Index & "'" & _
                                   " Caption= '" & objButton.Caption & "'" & _
                                   " Description= '" & objButton.Description & "'" & _
                                   " Key= '" & objButton.Key & "'" & _
                                   " Tag= '" & objButton.Tag & "'" & _
                                   " Tooltip= '" & objButton.ToolTip & "'" & _
                                   " Image= '" & objButton.Image & "'" & _
                                   " Enabled= '" & objButton.Enabled & "'" & _
                                   " Visible= '" & objButton.Visible & "' />" & vbCr
        Next
        
        strXML = strXML & vbTab & vbTab
        strXML = strXML & "</TAB>" & vbCr
    Next
    
    strXML = strXML & vbTab
    strXML = strXML & "</TABS>" & vbCr
    strXML = strXML & "</XML>"
    
    
    PropBag.WriteProperty "Tabs", strXML, vbNullString
    PropBag.WriteProperty "ImageListName", strImageListName, vbNullString
    PropBag.WriteProperty "Picture", Picture, Nothing
    
    Set objTab = Nothing
    Set objButton = Nothing
    
    Exit Sub
ErrorHandler:
    Set objTab = Nothing
    Set objButton = Nothing

    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "UserControl_WriteProperties Sub", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picInnerFrame,picInnerFrame,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

On Error GoTo ErrorHandler

    Set Picture = picInnerFrame.Picture
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "Picture Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Set Picture(ByVal New_Picture As Picture)

On Error GoTo ErrorHandler

    Set picInnerFrame.Picture = New_Picture
    PropertyChanged "Picture"
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "Picture Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=24,3,2,
Public Property Get ObjectTabs() As colTabs
Attribute ObjectTabs.VB_MemberFlags = "440"

On Error GoTo ErrorHandler

    If Ambient.UserMode Then Err.Raise 393
    Set ObjectTabs = mcolTabs
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ObjectTabs Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Set ObjectTabs(ByVal New_ObjectTabs As colTabs)

On Error GoTo ErrorHandler

    If Ambient.UserMode Then Err.Raise 382
    Set mcolTabs = New_ObjectTabs
    PropertyChanged "Tabs"
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ObjectTabs Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=25,3,2,0
Public Property Get GetImagesList() As Collection
Attribute GetImagesList.VB_MemberFlags = "440"

Dim objControl                              As Control
Dim colImagesList                           As Collection

On Error GoTo ErrorHandler


    If Ambient.UserMode Then Err.Raise 393
    Set colImagesList = New Collection

    For Each objControl In Parent.Controls
        If TypeName(objControl) = "ImageList" Then
            colImagesList.Add objControl
        End If
    Next

    Set GetImagesList = colImagesList

    Set objControl = Nothing
    Set colImagesList = Nothing

    Exit Property
ErrorHandler:
    
    Set objControl = Nothing
    Set colImagesList = Nothing
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "GetImagesList Property", lngCodigoErroNegocio, intNumeroSequencialErro)

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,3,2,
Public Property Get ImageListName() As String
Attribute ImageListName.VB_MemberFlags = "440"
On Error GoTo ErrorHandler
    
    If Ambient.UserMode Then Err.Raise 393
    ImageListName = strImageListName
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ImageListName Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Let ImageListName(ByVal New_ImageListName As String)

On Error GoTo ErrorHandler

    If Ambient.UserMode Then Err.Raise 382
    strImageListName = New_ImageListName
    PropertyChanged "ImageListName"
    
    Exit Property

ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ImageListName Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Get ActiveTab() As Integer

On Error GoTo ErrorHandler

    ActiveTab = intCurrentTab
    
    Exit Property
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ActiveTab Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property

Public Property Let ActiveTab(vData As Integer)

On Error GoTo ErrorHandler
    
    SetActiveTab vData

    Exit Property
ErrorHandler:
    
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "ActiveTab Property", lngCodigoErroNegocio, intNumeroSequencialErro)

End Property

Public Property Get Tabs() As colTabs

On Error GoTo ErrorHandler

    Set Tabs = mcolTabs
    
    Exit Property
ErrorHandler:
    If lngCodigoErroNegocio <> 0 Then On Error GoTo 0
    Call fgRaiseError(App.EXEName, "ctlBarraVertical", "Tabs Property", lngCodigoErroNegocio, intNumeroSequencialErro)
    
End Property
