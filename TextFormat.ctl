VERSION 5.00
Begin VB.UserControl TextFormat 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   HasDC           =   0   'False
   ScaleHeight     =   265
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "TextFormat.ctx":0000
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.PictureBox Picture2 
      HasDC           =   0   'False
      Height          =   3435
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   4635
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   2235
         Left            =   420
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   233
         TabIndex        =   0
         Top             =   540
         Width           =   3555
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   2835
         LargeChange     =   30
         Left            =   4200
         Max             =   11
         SmallChange     =   10
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Menu mnuHidden 
      Caption         =   "Hidden"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuCopyFormated 
         Caption         =   "Copy Formated"
      End
   End
   Begin VB.Menu mnuHiddenLink 
      Caption         =   "HiddenLink"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyLink 
         Caption         =   "Copy link"
      End
   End
End
Attribute VB_Name = "TextFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum FormatMethods
    UseBrackets
    UseSlash
End Enum

Private Const BLACKNESS As Long = &H42
Private Const DSTINVERT As Long = &H550009
Private Const MERGECOPY As Long = &HC00CA
Private Const MERGEPAINT As Long = &HBB0226
Private Const NOTSRCCOPY As Long = &H330008
Private Const NOTSRCERASE As Long = &H1100A6
Private Const PATCOPY As Long = &HF00021
Private Const PATINVERT As Long = &H5A0049
Private Const PATPAINT As Long = &HFB0A090
Private Const SRCAND As Long = &H8800C6
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCERASE As Long = &H440328
Private Const SRCINVERT As Long = &H660046
Private Const SRCPAINT As Long = &HEE0086
Private Const WHITENESS As Long = &HFF0062
Private Declare Function StretchBlt Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hSrcDC As Long, _
     ByVal xSrc As Long, _
     ByVal ySrc As Long, _
     ByVal nSrcWidth As Long, _
     ByVal nSrcHeight As Long, _
     ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" ( _
     ByVal hDestDC As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal nWidth As Long, _
     ByVal nHeight As Long, _
     ByVal hSrcDC As Long, _
     ByVal xSrc As Long, _
     ByVal ySrc As Long, _
     ByVal dwRop As Long) As Long

Enum aBorderStyleConstants
    None = 0
    FixedSingle = 1
End Enum
Private Const DT_ACCEPT_DBCS As Long = (&H20)
Private Const DT_AGENT As Long = (&H3)
Private Const DT_BOTTOM As Long = &H8
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_CHARSTREAM As Long = 4
Private Const DT_DISPFILE As Long = 6
Private Const DT_DISTLIST As Long = (&H1)
Private Const DT_EDITABLE As Long = (&H2)
Private Const DT_EDITCONTROL As Long = &H2000
Private Const DT_END_ELLIPSIS As Long = &H8000
Private Const DT_EXPANDTABS As Long = &H40
Private Const DT_EXTERNALLEADING As Long = &H200
Private Const DT_FOLDER As Long = (&H1000000)
Private Const DT_FOLDER_LINK As Long = (&H2000000)
Private Const DT_FOLDER_SPECIAL As Long = (&H4000000)
Private Const DT_FORUM As Long = (&H2)
Private Const DT_GLOBAL As Long = (&H20000)
Private Const DT_HIDEPREFIX As Long = &H100000
Private Const DT_INTERNAL As Long = &H1000
Private Const DT_LEFT As Long = &H0
Private Const DT_LOCAL As Long = (&H30000)
Private Const DT_MAILUSER As Long = (&H0)
Private Const DT_METAFILE As Long = 5
Private Const DT_MODIFIABLE As Long = (&H10000)
Private Const DT_MODIFYSTRING As Long = &H10000
Private Const DT_MULTILINE As Long = (&H1)
Private Const DT_NOCLIP As Long = &H100
Private Const DT_NOFULLWIDTHCHARBREAK As Long = &H80000
Private Const DT_NOPREFIX As Long = &H800
Private Const DT_NOT_SPECIFIC As Long = (&H50000)
Private Const DT_ORGANIZATION As Long = (&H4)
Private Const DT_PASSWORD_EDIT As Long = (&H10)
Private Const DT_PATH_ELLIPSIS As Long = &H4000
Private Const DT_PLOTTER As Long = 0
Private Const DT_PREFIXONLY As Long = &H200000
Private Const DT_PRIVATE_DISTLIST As Long = (&H5)
Private Const DT_RASCAMERA As Long = 3
Private Const DT_RASDISPLAY As Long = 1
Private Const DT_RASPRINTER As Long = 2
Private Const DT_REMOTE_MAILUSER As Long = (&H6)
Private Const DT_REQUIRED As Long = (&H4)
Private Const DT_RIGHT As Long = &H2
Private Const DT_RTLREADING As Long = &H20000
Private Const DT_SET_IMMEDIATE As Long = (&H8)
Private Const DT_SET_SELECTION As Long = (&H40)
Private Const DT_SINGLELINE As Long = &H20
Private Const DT_TABSTOP As Long = &H80
Private Const DT_TOP As Long = &H0
Private Const DT_VCENTER As Long = &H4
Private Const DT_WAN As Long = (&H40000)
Private Const DT_WORD_ELLIPSIS As Long = &H40000
Private Const DT_WORDBREAK As Long = &H10
Private Type DRAWTEXTPARAMS
    cbSize As Long
    iTabLength As Long
    iLeftMargin As Long
    iRightMargin As Long
    uiLengthDrawn As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type LinkAreaType
    R(1 To 7) As RECT
    Link As String
End Type

Private Declare Function SetBkMode Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal nBkMode As Long) As Long
Private Const OPAQUE As Long = 2
Private Const TRANSPARENT As Long = 1
Private Declare Function SetBkColor Lib "gdi32" ( _
     ByVal hdc As Long, _
     ByVal crColor As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" ( _
     ByVal hdc As Long, _
     ByVal lpStr As String, _
     ByVal nCount As Long, _
     lpRect As RECT, _
     ByVal wFormat As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" ( _
     ByVal hdc As Long, _
     ByVal lpsz As String, _
     ByVal n As Long, _
     lpRect As RECT, _
     ByVal un As Long, _
     lpDrawTextParams As DRAWTEXTPARAMS) As Long

Private Declare Function GetSystemMetrics Lib "user32" ( _
     ByVal nIndex As Long) As Long
Private Const SM_CXVSCROLL As Long = 2

Private Const IDC_HAND As Long = (32649)
Private Const IDC_ARROW As Long = 32512&
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" ( _
     ByVal hInstance As Long, _
     ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" ( _
     ByVal hCursor As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" _
   Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private TheText As String

Private stackFormat As New clsStack
Private stackFrom As New clsStack
Private stackLen As New clsStack

Private Links() As LinkAreaType

Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private UseFormat As FormatMethods, mEnableSelectText As Boolean
Private MaxHeight As Long, mRightMargin As Long, mAutoRedraw As Boolean

Private xDown As Long
Private yDown As Long
Private FontWas As StdFont
Private FontWasBold As Boolean
Private FontWasItalic As Boolean
Private FontWasUnderline As Boolean
Private FontWasSize As Single
Private LeftMarginWas As Long
Private NextMarginWas As Long
Private PosWas As Long
Private FontCol As New Collection

Public Property Get BackColor() As OLE_COLOR
    BackColor = Picture1.BackColor
End Property

Public Property Let BackColor(NewBackColor As OLE_COLOR)
    Picture1.BackColor = NewBackColor
    PicTemp.BackColor = NewBackColor
    Picture2.BackColor = NewBackColor
End Property

Public Property Get AutoRedraw() As Boolean
    AutoRedraw = mAutoRedraw
    If mAutoRedraw Then DrawTheText
End Property

Public Property Let AutoRedraw(NewAutoRedraw As Boolean)
    mAutoRedraw = NewAutoRedraw
End Property

Public Property Get EnableSelectText() As Boolean
    EnableSelectText = mEnableSelectText
End Property

Public Property Let EnableSelectText(NewEnableSelectText As Boolean)
    mEnableSelectText = NewEnableSelectText
End Property

Public Property Get Fonts() As Collection
    Set Fonts = FontCol
End Property

Public Property Get RightMargin() As Long
    RightMargin = mRightMargin
End Property

Public Property Let RightMargin(NewRightMargin As Long)
    mRightMargin = NewRightMargin
    UserControl_Resize
End Property

Public Property Get PrintAreaMaxHeight() As Long
    PrintAreaMaxHeight = MaxHeight
End Property

Public Property Let PrintAreaMaxHeight(NewPrintAreaMaxHeight As Long)
    MaxHeight = NewPrintAreaMaxHeight
End Property

Public Property Get PointerForLink() As IPictureDisp
    Set PointerForLink = Picture1.MouseIcon
End Property

Public Property Set PointerForLink(NewPointerForLink As IPictureDisp)
    Set Picture1.MouseIcon = NewPointerForLink
End Property

Public Property Get FormatMethod() As FormatMethods
    FormatMethod = UseFormat
End Property

Public Property Let FormatMethod(NewFormatMethod As FormatMethods)
    UseFormat = NewFormatMethod
End Property

Private Sub SubClassHookForm()
   'MSWHEEL_ROLLMSG = RegisterWindowMessage("MSWHEEL_ROLLMSG")
   ' On Windows NT 4.0, Windows 98, and Windows Me, change the above line to
   MSWHEEL_ROLLMSG = &H20A
   m_PrevWndProc = SetWindowLong(Picture1.hwnd, GWL_WNDPROC, _
                                 AddressOf WindowProc)
End Sub

Private Sub SubClassUnHookForm()
   Call SetWindowLong(Picture1.hwnd, GWL_WNDPROC, m_PrevWndProc)
End Sub

Public Property Get BorderStyle() As aBorderStyleConstants
    BorderStyle = Picture2.BorderStyle
End Property

Public Property Let BorderStyle(aVal As aBorderStyleConstants)
    Picture2.BorderStyle = aVal
    UserControl_Resize
End Property

Public Property Get Text() As String
Text = TheText
End Property

Public Property Let Text(new_text As String)
TheText = new_text
DrawTheText
End Property

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(IsEnabled As Boolean)
    UserControl.Enabled = IsEnabled
    VScroll1.Enabled = IsEnabled
    DrawTheText
End Property

Public Property Get MenuCopyVisible() As Boolean
    MenuCopyVisible = mnuCopy.Visible
End Property

Public Property Let MenuCopyVisible(IsVisible As Boolean)
    mnuCopy.Visible = IsVisible
End Property

Public Property Get MenuCopyFormatedVisible() As Boolean
    MenuCopyFormatedVisible = mnuCopyFormated.Visible
End Property

Public Property Let MenuCopyFormatedVisible(IsVisible As Boolean)
    mnuCopyFormated.Visible = IsVisible
End Property

Public Sub Refresh()
Dim OldAutoRedraw As Boolean
OldAutoRedraw = mAutoRedraw
mAutoRedraw = True
DrawTheText
mAutoRedraw = OldAutoRedraw
End Sub

Private Sub mnuCopyLink_Click()
    Clipboard.Clear
    Clipboard.SetText mnuCopyLink.Tag
End Sub

Private Sub PicTemp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
MsgBox x
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not VScroll1.Enabled Then Exit Sub
If KeyCode = vbKeyUp Then
    If VScroll1.Value - VScroll1.SmallChange < 0 Then
        VScroll1.Value = 0
    Else
        VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
    End If
ElseIf KeyCode = vbKeyDown Then
    If VScroll1.Value + VScroll1.SmallChange > VScroll1.Max Then
        VScroll1.Value = VScroll1.Max
    Else
        VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
    End If
ElseIf KeyCode = vbKeyPageUp Then
    If VScroll1.Value - VScroll1.LargeChange < 0 Then
        VScroll1.Value = 0
    Else
        VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
    End If
ElseIf KeyCode = vbKeyPageDown Then
    If VScroll1.Value + VScroll1.LargeChange > VScroll1.Max Then
        VScroll1.Value = VScroll1.Max
    Else
        VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
    End If
ElseIf KeyCode = vbKeyHome Then
    VScroll1.Value = 0
ElseIf KeyCode = vbKeyEnd Then
    VScroll1.Value = VScroll1.Max
End If
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
If Button = 1 And mEnableSelectText Then
    xDown = x
    yDown = y
    LocateTextAtXY PosWas, xDown, yDown, FontWas, LeftMarginWas, NextMarginWas
    FontWasBold = FontWas.Bold
    FontWasItalic = FontWas.Italic
    FontWasUnderline = FontWas.Underline
    FontWasSize = FontWas.Size
End If
    If MouseAtALinkArea(x, y) = 0 Then
        'Call SetCursor(LoadCursor(0, IDC_ARROW))
    Else
        Call SetCursor(LoadCursor(0, IDC_HAND))
    End If

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseMove(Button, Shift, x, y)

If Button = 0 Then
    If MouseAtALinkArea(x, y) = 0 Then
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    Else
        Call SetCursor(LoadCursor(0, IDC_HAND))
    End If
ElseIf Button = 1 And mEnableSelectText Then
    Dim pos As Long, X1 As Long, Y1 As Long, TheFont As StdFont, LMargin As Long, NMargin As Long
    X1 = x
    Y1 = y
    LocateTextAtXY pos, X1, Y1, TheFont, LMargin, NMargin
        
    DrawTheTextFromToPoint xDown, yDown, PosWas, pos, FontWas, LeftMarginWas, NextMarginWas
End If
End Sub

Private Function MouseAtALinkArea(x As Single, y As Single) As Long
Dim i As Long, j As Long

For i = 1 To UBound(Links)
    For j = 1 To UBound(Links(i).R)
        If x > Links(i).R(j).Left And x < Links(i).R(j).Right And _
            y > Links(i).R(j).Top And y < Links(i).R(j).Bottom Then
                MouseAtALinkArea = i
                Exit Function
        End If
    Next j
Next i

End Function

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
RaiseEvent MouseUp(Button, Shift, x, y)
If Button = 2 Then
    i = MouseAtALinkArea(x, y)
    If i = 0 Then
        UserControl.PopupMenu mnuHidden
    Else
        mnuCopyLink.Tag = Links(i).Link
        UserControl.PopupMenu mnuHiddenLink
    End If
ElseIf Button = 1 Then
    i = MouseAtALinkArea(x, y)
    If i <> 0 Then
        Dim iret As Long
        ' open URL into the default internet browser
        Const SW_SHOWNORMAL = 1
        iret = ShellExecute(UserControl.Parent.hwnd, vbNullString, Links(i).Link, _
            vbNullString, "", SW_SHOWNORMAL)
    End If
End If
End Sub

Private Sub UserControl_GotFocus()
Picture1.SetFocus
End Sub

Private Sub UserControl_Initialize()
VScroll1.Width = GetSystemMetrics(SM_CXVSCROLL)
Picture1.BorderStyle = 0
Picture2.Left = 0
Picture2.Top = 0
VScroll1.Top = 0
Set aControl = VScroll1
ReDim Links(0) As LinkAreaType
'mAutoRedraw = True
MaxHeight = 5000
'SubClassHookForm
End Sub

Private Sub UserControl_Resize()
    Picture2.Width = UserControl.ScaleWidth
    Picture2.Height = UserControl.ScaleHeight
    
    Picture1.Left = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    Picture1.Top = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    If UserControl.ScaleWidth - VScroll1.Width - mRightMargin - IIf(Picture2.BorderStyle = 0, 0, 4) < 0 Then
        Picture1.Width = 0
    Else
        Picture1.Width = UserControl.ScaleWidth - VScroll1.Width - mRightMargin - IIf(Picture2.BorderStyle = 0, 0, 4)
    End If
    PicTemp.Width = Picture1.Width
    VScroll1.Top = 0 'IIf(Picture2.BorderStyle = 0, 0, 2)
    VScroll1.Left = Picture1.Width + mRightMargin
    VScroll1.Height = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 0, 4)
    DrawTheText
End Sub

Private Sub UserControl_Terminate()
SubClassUnHookForm
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mnuCopyFormated.Visible = PropBag.ReadProperty("MenuCopyFormatedVisible", True)
    mnuCopy.Visible = PropBag.ReadProperty("MenuCopyVisible", True)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    FormatMethod = PropBag.ReadProperty("FormatMethod", 0)
    PrintAreaMaxHeight = PropBag.ReadProperty("PrintAreaMaxHeight", 5000)
    Set PointerForLink = PropBag.ReadProperty("PointerForLink", Nothing)
    RightMargin = PropBag.ReadProperty("RightMargin", 0)
    mAutoRedraw = PropBag.ReadProperty("AutoRedraw", True)
    BackColor = PropBag.ReadProperty("BackColor", Picture1.BackColor)
    Text = PropBag.ReadProperty("Text", "")
    EnableSelectText = PropBag.ReadProperty("EnableSelectText", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MenuCopyFormatedVisible", mnuCopyFormated.Visible, True
    PropBag.WriteProperty "MenuCopyVisible", mnuCopy.Visible, True
    PropBag.WriteProperty "Enabled", UserControl.Enabled, True
    PropBag.WriteProperty "BorderStyle", Picture2.BorderStyle, 1
    PropBag.WriteProperty "FormatMethod", FormatMethod, ""
    PropBag.WriteProperty "PrintAreaMaxHeight", PrintAreaMaxHeight, ""
    PropBag.WriteProperty "PointerForLink", PointerForLink, ""
    PropBag.WriteProperty "RightMargin", RightMargin, ""
    PropBag.WriteProperty "AutoRedraw", mAutoRedraw, True
    PropBag.WriteProperty "BackColor", Picture1.BackColor
    PropBag.WriteProperty "Text", TheText, ""
    PropBag.WriteProperty "EnableSelectText", EnableSelectText, False
End Sub

Private Sub VScroll1_Change()
Picture1.Top = -VScroll1.Value
PicTemp.Top = -VScroll1.Value
End Sub

Private Sub mnuCopy_Click()
Dim i As Long, printThis As String, aStr As String
Dim nextText As String
    
    If UseFormat = 0 Then
        aStr = Replace(TheText, vbNewLine, "[l]")
    Else
        aStr = Replace(TheText, vbNewLine, "/l ")
    End If
    
    SplitFormat aStr
    For i = 1 To stackLen.stackLevel
        nextText = Mid(aStr, stackFrom.pop, stackLen.pop)
        If UseFormat = 0 Then
            nextText = Replace(nextText, "[[", "[")
            nextText = Replace(nextText, "]]", "]")
        Else
            nextText = Replace(nextText, "//", "/")
        End If
        If stackFormat.pop = "l" Then printThis = printThis & vbNewLine
        printThis = printThis & nextText
    Next i
        
    Clipboard.Clear
    Clipboard.SetText printThis
    
End Sub

Private Sub mnuCopyFormated_Click()
    Clipboard.Clear
    Clipboard.SetText TheText
End Sub

' Use of [ ]
Private Sub SplitFormat(ByRef aStr As String)
Dim pl As Long, pl2 As Long, lastStart As Long
Dim stackFormatNo As New clsStack
Dim stackFromNo As New clsStack
Dim stackLenNo As New clsStack

If UseFormat = UseSlash Then
    SplitFormat2 aStr
    Exit Sub
End If

stackFormat.Clear
stackFrom.Clear
stackLen.Clear
pl = InStr(aStr, "[")
If pl > 1 Then
    While Mid(aStr, pl + 1, 1) = "[" And pl <> 0
        pl = InStr(pl + 2, aStr, "[")
    Wend
    stackFormatNo.push ""
    stackFromNo.push 1
    If pl = 0 Then
        stackLenNo.push Len(aStr)
        pl2 = Len(aStr)
    Else
        stackLenNo.push pl - 1
    End If
    lastStart = 0
End If
Do While pl <> 0
    If Mid(aStr, pl + 1, 1) <> "[" Then
        pl2 = InStr(pl + 1, aStr, "]")
        While Mid(aStr, pl2 + 1, 1) = "]" And pl2 <> 0
            pl2 = InStr(pl2 + 2, aStr, "]")
        Wend
        'If Mid(aStr, pl2 + 1, 1) <> "]" Then
            If pl2 <> 0 Then
                If Mid(aStr, pl + 1, 1) = "/" Then
                    'END format
                    stackFormatNo.push Mid(aStr, pl + 2, pl2 - pl - 2)
                Else
                    'START format
                    stackFormatNo.push Mid(aStr, pl + 1, pl2 - pl - 1)
                End If
                stackFromNo.push pl2 + 1
                If lastStart <> 0 Then stackLenNo.push pl - lastStart
                lastStart = pl2 + 1
            End If
        'End If
    End If
    pl = InStr(pl + 1, aStr, "[")
    While Mid(aStr, pl + 1, 1) = "[" And pl <> 0
        pl = InStr(pl + 2, aStr, "[")
    Wend
Loop
If pl2 <> Len(aStr) Then ' More text
    stackFormatNo.push ""
    stackFromNo.push pl2 + 1
    stackLenNo.push Len(aStr) - pl2
End If

For pl = 1 To stackFormatNo.stackLevel
    stackFormat.push stackFormatNo.pop
Next pl
For pl = 1 To stackFromNo.stackLevel
    stackFrom.push stackFromNo.pop
Next pl
For pl = 1 To stackLenNo.stackLevel
    stackLen.push stackLenNo.pop
Next pl

End Sub

' Use of /
Private Sub SplitFormat2(ByRef aStr As String)
Dim pl As Long, pl2 As Long, pl3 As Long, lastStart As Long
Dim MustAdd1 As Long
Dim stackFormatNo As New clsStack
Dim stackFromNo As New clsStack
Dim stackLenNo As New clsStack

stackFormat.Clear
stackFrom.Clear
stackLen.Clear
pl = InStr(aStr, "/")
If pl > 1 Then
    While Mid(aStr, pl + 1, 1) = "/" And pl <> 0
        pl = InStr(pl + 2, aStr, "/")
    Wend
    stackFormatNo.push ""
    stackFromNo.push 1
    stackLenNo.push pl - 1
    lastStart = 0
End If
Do While pl <> 0
    pl2 = InStr(pl + 1, aStr, " ")
    MustAdd1 = 1
    pl3 = InStr(pl + 1, aStr, "/")
    If pl3 < pl2 And pl3 <> 0 Then
        pl2 = pl3: MustAdd1 = 0
    End If
    pl3 = InStr(pl + 1, aStr, vbNewLine)
    If pl3 < pl2 And pl3 <> 0 Then pl2 = pl3: MustAdd1 = 0
    
    If pl2 <> 0 Then
        stackFormatNo.push Mid(aStr, pl + 1, pl2 - pl - 1)
        stackFromNo.push pl2 + 1
        If lastStart <> 0 Then stackLenNo.push pl - lastStart
        lastStart = pl2 + MustAdd1
        
        pl = InStr(pl2, aStr, "/")
    Else
        pl = InStr(pl + 1, aStr, "/")
    End If
    
    While Mid(aStr, pl + 1, 1) = "/" And pl <> 0
        pl = InStr(pl + 2, aStr, "/")
    Wend
Loop
If pl2 <> Len(aStr) Then ' More text
    stackFormatNo.push ""
    stackFromNo.push pl2 + 1
    stackLenNo.push Len(aStr) - pl2
End If

For pl = 1 To stackFormatNo.stackLevel
    stackFormat.push stackFormatNo.pop
Next pl
For pl = 1 To stackFromNo.stackLevel
    stackFrom.push stackFromNo.pop
Next pl
For pl = 1 To stackLenNo.stackLevel
    stackLen.push stackLenNo.pop
Next pl

End Sub

Private Sub DrawTheTextOLD()
Dim aStr As String, printThis As String, i As Long, Lines As Long
Dim HeightOf1Line As Long, cHeight As Long, R As RECT, pl As Long
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long, NextMargin As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMargin As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

Picture1.FontBold = False
Picture1.FontItalic = False
Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
If Not UserControl.Enabled Then
    Picture1.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
Picture1.Cls
For i = 1 To stackLen.stackLevel
    TheFrom = stackFrom.pop
    TheLen = stackLen.pop
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.pop
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            Picture1.Print ""
            Picture1.CurrentX = LeftMargin
            
            NextMargin = 0
            HasPrintSomething = False
            MaxHeightOf1Line = HeightOf1Line
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                Picture1.CurrentX = LeftMarginNext
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
        Case "e": ' print a line
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(Picture1.Width, Picture1.CurrentY)
            Else
                Picture1.Line -(Picture1.Width, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
            End If
            Picture1.CurrentY = Picture1.CurrentY + 2
            Picture1.CurrentX = LeftMargin
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            Picture1.CurrentY = Picture1.CurrentY + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            aSingle = Picture1.CurrentY
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                Picture1.CurrentX = Picture1.CurrentX + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                Picture1.CurrentX = Picture1.CurrentX + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            Picture1.DrawWidth = 1
            Picture1.CurrentX = Picture1.CurrentX + 7
            Picture1.CurrentY = aSingle
    End Select
    If printThis = "" Then GoTo Nexti
    
    R.Left = Picture1.CurrentX
    R.Top = Picture1.CurrentY
    R.Right = Picture1.Width
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line) Then
        Picture1.Print printThis;
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            Picture1.Print LastprintThis;
            HasPrintSomething = False
            Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            Picture1.CurrentX = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo here
        Else
here:
            If HasPrintSomething Then
                Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                Picture1.CurrentX = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = Picture1.Width
            R.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                Picture1.Print printThis;
            Else
                R.Bottom = R.Bottom - HeightOf1Line
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                Picture1.CurrentX = LeftMargin + NextMargin
                Picture1.CurrentY = R.Bottom
                Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
            End If
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext
Next i

If Picture1.CurrentY + HeightOf1Line > R.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = R.Bottom
End If

If Picture1.Height > UserControl.ScaleHeight Then
    VScroll1.Enabled = True
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    If VScroll1.Value >= VScroll1.Min And VScroll1.Value <= VScroll1.Max Then
        VScroll1_Change
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    R.Left = Picture1.CurrentX
    R.Top = Picture1.CurrentY
    R.Right = Picture1.Width
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).R(LinkAreaNumber) = R
        'End If
        Picture1.Print printThis;
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                R.Left = Picture1.CurrentX
                R.Top = Picture1.CurrentY
                R.Right = Picture1.Width
                R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).R(LinkAreaNumber) = R
            'End If
            Picture1.Print LastprintThis;
            HasPrintSomething = False
            Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            Picture1.CurrentX = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo hereWithLink
        Else
hereWithLink:
            If HasPrintSomething Then
                Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                Picture1.CurrentX = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                Picture1.Print printThis;
            Else
                R.Bottom = R.Bottom - HeightOf1Line
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                Picture1.CurrentX = LeftMargin + NextMargin
                Picture1.CurrentY = R.Bottom
                'If WasLink Then
                    'Calculate rect
                    R.Left = Picture1.CurrentX
                    R.Top = Picture1.CurrentY
                    R.Right = Picture1.Width
                    R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                    Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
            End If
        End If
    End If

Return

End Sub



Private Sub DrawTheTextByLine(R As RECT, i As Long, LeftMargin As Long, LeftMarginNext As Long, NextMargin As Long, metrCell As Long, TableCellLeft() As Long, TableLineTop As Long, WasTable As Boolean, TableLineMaxHeight As Long, _
        TableCellHasBorder() As Boolean, TableLineFixedPixels As Long, CellPercentEat As Single, BorderColor, BorderColorBefore As Long, _
        CellMargin As Long, PixelsAfterCell As Long, PixelsAtTopBottomCell As Long, TableBorderWidth As Long, TableBorderWidth2 As Long, _
        MaxX As Long, HeightOf1Line As Long, MaxHeightOf1Line As Long, LastMaxHeightOf1Line As Long, MaxHeightOfBaseLine As Long, _
        MustMoveActiveLineToBaseLine As Long, TM As TEXTMETRIC, ParLine As Single, ParBefore As Long, ParAfter As Long, _
        DrawTextAtMaxY As Long, DrawTextAtMaxYStart As Long, DrawAroundPicture As Boolean, ExtraMarginLeft As Long, _
        ExtraMarginRight As Long, Align As AlignmentConstants, StoredY As Long, StoredY2 As Long)
Dim aStr As String, printThis As String, someThing As Long
Dim cHeight As Long, pl As Long, LastWhatIs As String
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single
Dim OldFontName As String, OldDrawStyle As Long, aTempVal As Single, aTempLong As Long, aTempLong2 As Long, aTempLong3 As Long
'Dim TableCellLeft() As Long, TableLineTop As Long, WasTable As Boolean, TableLineMaxHeight As Long
'Dim TableCellHasBorder() As Boolean, TableLineFixedPixels As Long, CellPercentEat As Single, BorderColor, BorderColorBefore As Long
'Dim CellMargin As Long, PixelsAfterCell As Long, PixelsAtTopBottomCell As Long, TableBorderWidth As Long, TableBorderWidth2 As Long
'Dim MaxX As Long
Dim StartLeft As Long, StartTop As Long, rRight As Long
Dim stackFromLevel As Long, stackLenLevel As Long, stackFormatLevel As Long
'Dim TM As TEXTMETRIC
Dim Alignment As AlignmentConstants, LastLineLink As Long
Dim PicName As String, PicHandle As Long, PicWidth As Long, PicHeight As Long

If Not mAutoRedraw Then Exit Sub

'PixelsAtTopBottomCell = 2
'PixelsAfterCell = 2
'TableBorderWidth = 1

On Error GoTo ErrHandle

'ReDim Links(0) As LinkAreaType

'Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

'Picture1.FontBold = False
'Picture1.FontItalic = False
'Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
OldFontName = Picture1.FontName
OldDrawStyle = Picture1.DrawStyle
'If Not PicTemp.Enabled Then
'    PicTemp.ForeColor = &H80000011
'End If

'HeightOf1Line = Picture1.TextHeight("astr")
'MaxHeightOf1Line = HeightOf1Line
'MaxX = Picture1.Width
'Picture1.Cls

Set PicTemp.Font = Picture1.Font
StartLeft = R.Left
StartTop = R.Top
stackFromLevel = stackFrom.stackLevel
stackLenLevel = stackLen.stackLevel
stackFormatLevel = stackFormat.stackLevel
For i = i To stackLenLevel
    TheFrom = stackFrom.popNo(stackFromLevel - i + 1)
    TheLen = stackLen.popNo(stackLenLevel - i + 1)
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.popNo(stackFormatLevel - i + 1)
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b":
            Picture1.FontBold = Not Picture1.FontBold
            PicTemp.FontBold = Picture1.FontBold
        Case "i":
            Picture1.FontItalic = Not Picture1.FontItalic
            PicTemp.FontItalic = Picture1.FontItalic
        Case "u":
            Picture1.FontUnderline = Not Picture1.FontUnderline
            PicTemp.FontUnderline = Picture1.FontUnderline
        Case "l":
            If Mid(WhatIs, 2, 1) = "g" Then ' align
                Select Case Mid(WhatIs, 3, 1)
                    Case "l": Alignment = vbLeftJustify
                    Case "c": Alignment = vbCenter
                    Case "r": Alignment = vbRightJustify
                End Select
                If Alignment = vbLeftJustify Then
                    i = i - 1
                    Exit Sub
                End If
            Else ' new line
                GoSub MoveTheLine
                LastLineLink = False
                R.Top = R.Top + HeightOf1Line * ParLine + ParAfter + ParBefore
                
                If DrawTextAtMaxY <> 0 Then
                    If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                        'R.Left = R.Left - ExtraMarginLeft
                        MaxX = MaxX + ExtraMarginRight
                        LeftMargin = LeftMargin - ExtraMarginLeft
                        If Align = vbCenter And ExtraMarginRight > 0 Then
                            ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                            LeftMargin = LeftMargin + ExtraMarginLeft
                            R.Right = LeftMargin
                            R.Top = DrawTextAtMaxYStart
                            ExtraMarginRight = 0
                        Else
                            R.Top = DrawTextAtMaxY + 3
                            ExtraMarginLeft = 0: ExtraMarginRight = 0
                            DrawTextAtMaxY = 0
                        End If
                    End If
                End If
                
                R.Left = LeftMargin
                R.Right = R.Left
                
                NextMargin = 0
                HasPrintSomething = False
                MaxHeightOf1Line = HeightOf1Line
                StartLeft = R.Left: StartTop = R.Top
                
                'GetTextMetrics Picture1.hdc, TM
                MaxHeightOfBaseLine = TM.tmAscent
            End If
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin - ExtraMarginLeft + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin - ExtraMarginLeft + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                R.Left = LeftMarginNext
                R.Right = LeftMarginNext
            ElseIf metrCell <> 0 And LastWhatIs = "a" Then
                R.Right = CellMargin + LeftMarginNext
                R.Left = R.Right
                LeftMargin = LeftMarginNext + CellMargin
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            PicTemp.FontSize = Val(Mid(WhatIs, 2))
SameAsFontSize:
            LastMaxHeightOf1Line = MaxHeightOf1Line
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
            
            aTempLong = TM.tmAscent
            GetTextMetrics Picture1.hdc, TM
                        
            If MaxHeightOfBaseLine < TM.tmAscent Then
                If HasPrintSomething Then
                    MustMoveActiveLineToBaseLine = MaxHeightOfBaseLine
                End If
                'If HasPrintSomething Then R.Top = R.Top + aTempLong - MaxHeightOfBaseLine
                R.Top = R.Top + aTempLong - MaxHeightOfBaseLine
                MaxHeightOfBaseLine = TM.tmAscent
            Else
                'If HasPrintSomething Then R.Top = R.Top + aTempLong - TM.tmAscent
                R.Top = R.Top + aTempLong - TM.tmAscent
                MustMoveActiveLineToBaseLine = 0
            End If
        
        Case "e": ' print a line
            Picture1.CurrentY = R.Top
            Picture1.CurrentX = R.Right
            aTempVal = InStr(WhatIs, "x")
            If aTempVal <> 0 Then 'user set a percent for the line
                aTempVal = Val(Mid(WhatIs, aTempVal + 1))
            End If
            If aTempVal <= 0 Then aTempVal = 100
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(aTempVal / 100 * MaxX, Picture1.CurrentY)
                R.Top = R.Top + 2
            Else
                Picture1.Line -(aTempVal / 100 * MaxX - 1, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
                R.Top = R.Top + Val(Mid(WhatIs, 2)) + 1
            End If
            R.Bottom = R.Top
            R.Left = LeftMargin
            StartLeft = R.Left: StartTop = R.Top
        Case "f": 'font
            On Error Resume Next
            If Val(Mid(WhatIs, 2)) < 1 Then
                Picture1.FontName = OldFontName
            Else
                Picture1.FontName = FontCol.Item(Val(Mid(WhatIs, 2)))
            End If
            PicTemp.FontName = Picture1.FontName
            On Error GoTo ErrHandle
            GoTo SameAsFontSize
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If PicTemp.Enabled Then
                    Picture1.ForeColor = vbBlue
                    PicTemp.ForeColor = vbBlue
                End If
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If PicTemp.Enabled Then
                    Picture1.ForeColor = OldForeColor
                    PicTemp.ForeColor = OldForeColor
                End If
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            R.Top = R.Top + Val(Mid(WhatIs, 2))
            StartTop = R.Top
        Case "c": ' Color
            If PicTemp.Enabled Then
                Picture1.ForeColor = Val(Mid(WhatIs, 2))
                PicTemp.ForeColor = Val(Mid(WhatIs, 2))
            End If
        Case "t": 'Bullet
            aSingle = R.Top
            If Mid(WhatIs, 2, 1) = "2" Then
                PicTemp.CurrentY = aSingle + HeightOf1Line / 2.5
                PicTemp.CurrentX = R.Left + 5
                PicTemp.DrawWidth = 3
                PicTemp.Line -(PicTemp.CurrentX + HeightOf1Line / 5, PicTemp.CurrentY + HeightOf1Line / 5), , BF
            Else
                PicTemp.CurrentY = aSingle + HeightOf1Line / 2
                PicTemp.CurrentX = R.Left + 6
                PicTemp.DrawWidth = 5
                PicTemp.Circle (PicTemp.CurrentX, PicTemp.CurrentY), HeightOf1Line \ 13
            End If
            HasPrintSomething = True
            PicTemp.DrawWidth = 1
            R.Left = PicTemp.CurrentX + 7
            R.Right = R.Left
            R.Top = aSingle
        Case "d":
            'user want to set the border style for drawing
            Picture1.DrawStyle = Val(Mid(WhatIs, 2))
            PicTemp.DrawStyle = Val(Mid(WhatIs, 2))
        Case "g": 'user want to paint a picture
          If WhatIs = "ga" Then
            DrawAroundPicture = True
          ElseIf WhatIs = "gn" Then
            DrawAroundPicture = False
          ElseIf WhatIs = "gr" Then
            If DrawTextAtMaxY <> 0 And R.Top < DrawTextAtMaxY + 3 Then
                R.Top = DrawTextAtMaxY + 3
                MaxX = MaxX + ExtraMarginRight
                R.Right = LeftMargin - ExtraMarginLeft
                R.Left = R.Right
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
            End If
          Else
            
            GoSub MoveTheLine
            aTempVal = 0
            PicWidth = 0
            PicHeight = InStr(3, WhatIs, "|")
            If PicHeight = 0 Then 'picture will paint using current dimensions
                If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
                    PicHandle = Val(Mid(WhatIs, 2))
                Else 'picture  by path
                    PicName = Mid(WhatIs, 2)
                End If
                Align = vbLeftJustify
            Else
                If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
                    PicHandle = Val(Mid(WhatIs, 3, PicHeight - 3))
                Else 'picture  by path
                    PicName = Mid(WhatIs, 2, PicHeight - 2)
                End If
                PicWidth = Val(Mid(WhatIs, PicHeight + 1))
                PicHeight = InStr(PicHeight + 1, WhatIs, "x")
                If PicHeight <> 0 Then
                    PicHeight = Val(Mid(WhatIs, PicHeight + 1))
                End If
                If InStr(Right(WhatIs, 2), "f") = 0 Then 'use percent for dimensions
                    'must be PicWidth <> 0 Or PicHeight <> 0
                    If PicWidth <> 0 Or PicHeight <> 0 Then
                        aTempVal = MaxX - R.Right
                    End If
                End If
                If InStr(Right(WhatIs, 2), "r") <> 0 Then
                    Align = vbRightJustify
                ElseIf InStr(Right(WhatIs, 2), "c") <> 0 Then
                    Align = vbCenter
                Else
                    Align = vbLeftJustify
                End If
            End If
            
            If DrawTextAtMaxY <> 0 Then
                R.Top = DrawTextAtMaxY + 3
                MaxX = MaxX + ExtraMarginRight
                R.Right = LeftMargin - ExtraMarginLeft
                R.Left = R.Right
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
            End If
            
            R.Top = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 2)
            If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
LoadByHandle:
                PaintPictureByHandleGgiPlus PicHandle, Picture1.hdc, R.Right, R.Top, PicWidth, PicHeight, aTempVal <> 0, MaxX - R.Right, 0, Align, True
                Set UserControl.Picture = Nothing
            Else 'picture  by path
                If UCase(Left(PicName, 10)) = "<APP.PATH>" Then
                    PicName = App.Path & Mid(PicName, 11)
                ElseIf UCase(Left(PicName, 3)) = "<R>" Then
                    If UCase(Right(PicName, 1)) = "I" Then 'Icon
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResIcon)
                    ElseIf UCase(Right(PicName, 1)) = "C" Then 'Cursor
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResCursor)
                    Else 'Bitmap
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResBitmap)
                    End If
                    PicHandle = UserControl.Picture.Handle
                    GoTo LoadByHandle
                End If
                PaintPictureGgiPlus PicName, Picture1.hdc, R.Right, R.Top, PicWidth, PicHeight, aTempVal <> 0, MaxX - R.Right, Align, True
            End If
            
            If DrawAroundPicture Then '/draw around picture
                If DrawTextAtMaxY < R.Top + PicHeight Then DrawTextAtMaxY = R.Top + PicHeight
                MaxX = MaxX + ExtraMarginRight
                If Align = vbRightJustify Then
                    'ExtraMarginLeft = 0
                    ExtraMarginRight = ExtraMarginRight + PicWidth + 3
                ElseIf Align = vbCenter Then
                    'ExtraMarginRight = ExtraMarginRight + PicWidth + 3
                    ExtraMarginRight = MaxX - ExtraMarginRight - R.Right + 3
                    'DrawTextAtMaxYStart = R.Top - IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
                    DrawTextAtMaxYStart = R.Top - IIf(HasPrintSomething, -1, -0)
                Else
                    ExtraMarginLeft = ExtraMarginLeft + LeftMargin - CellMargin + PicWidth + 3
                    'ExtraMarginRight = 0
                End If
                R.Right = LeftMargin + ExtraMarginLeft
                MaxX = MaxX - ExtraMarginRight
                
                If metrCell <> 0 Then
                    If TableLineMaxHeight < R.Top + PicHeight - TableLineTop Then TableLineMaxHeight = R.Top + PicHeight - TableLineTop
                End If
                StartLeft = R.Left: StartTop = R.Top
                'R.Top = R.Top - HeightOf1Line * ParLine + ParAfter + ParBefore - IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
                R.Top = R.Top - HeightOf1Line * ParLine + ParAfter + ParBefore '- IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
            Else
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
                
                R.Top = R.Top + PicHeight - MaxHeightOf1Line
                R.Bottom = R.Top + MaxHeightOf1Line
                R.Right = R.Right + PicWidth
                HasPrintSomething = True
                
                If metrCell <> 0 Then
                    If TableLineMaxHeight < R.Top + MaxHeightOf1Line - TableLineTop Then TableLineMaxHeight = R.Top + MaxHeightOf1Line - TableLineTop
                End If
                StartLeft = R.Left: StartTop = R.Bottom
            End If
                      
          End If
        Case "a": ' table cell
            If Val(Mid(WhatIs, 2)) = 0 Then 'END of cells
              If Mid(WhatIs, 2, 2) = "bc" Then 'user want to set the border color for the next table
                BorderColor = Val(Mid(WhatIs, 4))
              ElseIf Mid(WhatIs, 2, 1) = "b" Then 'user want to set the border width for the next table
                PixelsAfterCell = PixelsAfterCell - TableBorderWidth - TableBorderWidth2
                PixelsAtTopBottomCell = PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2
                
                TableBorderWidth = Val(Mid(WhatIs, 3))
                aTempVal = InStr(WhatIs, "x")
                If aTempVal = 0 Then
                    TableBorderWidth2 = 0
                Else
                    TableBorderWidth2 = Val(Mid(WhatIs, aTempVal + 1)) + 1
                    If TableBorderWidth2 = 1 Then TableBorderWidth2 = 0
                End If
                PixelsAfterCell = PixelsAfterCell + TableBorderWidth + TableBorderWidth2
                PixelsAtTopBottomCell = PixelsAtTopBottomCell + TableBorderWidth + TableBorderWidth2
              ElseIf Mid(WhatIs, 2, 2) = "mt" Then 'user want to set the margin at top and bottom for a cell
                PixelsAtTopBottomCell = Val(Mid(WhatIs, 4)) + TableBorderWidth + TableBorderWidth2
              ElseIf Mid(WhatIs, 2, 1) = "m" Then 'user want to set the margin at left and right for a cell
                PixelsAfterCell = Val(Mid(WhatIs, 3)) + TableBorderWidth + TableBorderWidth2
              Else 'user said that a table line is end
                WasTable = True
                If metrCell <> 0 Then 'we have cells, so lets print the borders
                    'fix round problem at right margin
                    If Abs(TableCellLeft(metrCell) + TableBorderWidth + TableBorderWidth2 - Picture1.Width) < TableBorderWidth + TableBorderWidth2 + 1 Then
                        TableCellLeft(metrCell) = Picture1.Width - TableBorderWidth - TableBorderWidth2
                    End If
                                        
                    If TableLineMaxHeight < R.Bottom - TableLineTop Then TableLineMaxHeight = R.Bottom - TableLineTop
                    If Not IsEmpty(BorderColor) Then
                        BorderColorBefore = Picture1.ForeColor
                        Picture1.ForeColor = BorderColor
                    End If
                    ReDim Preserve TableCellHasBorder(metrCell + 1)
                    For someThing = 0 To metrCell - 1
                        If TableCellHasBorder(someThing + 1) And TableBorderWidth > 0 Then 'has this cell border?
                            'cell's left border
                            Picture1.CurrentY = TableLineTop
                            Picture1.CurrentX = TableCellLeft(someThing)
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                            Else
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                            End If
                        
                            'cell's top border
                            Picture1.CurrentY = TableLineTop
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1), , BF
                            Else
                                Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1)
                            End If
                            
                            'cell's right border
                            Picture1.CurrentY = TableLineTop
                            Picture1.CurrentX = TableCellLeft(someThing + 1)
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                            Else
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                            End If
                            
                            'cell's bottom border
                            Picture1.CurrentX = TableCellLeft(someThing)
                            Picture1.CurrentY = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth - 1, Picture1.CurrentY + TableBorderWidth - 1), , BF
                            Else
                                Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth, Picture1.CurrentY + TableBorderWidth - 1)
                            End If
                            
                            'cell's second border
                            If TableBorderWidth2 > 0 Then
                                'cell's left border
                                If Not TableCellHasBorder(someThing) Then
                                    Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                    Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                    If TableBorderWidth2 > 2 Then
                                        Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                                    Else
                                        Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                                    End If
                                End If
                                
                                'cell's top border
                                Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                If TableBorderWidth2 > 2 Then
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2), , BF
                                Else
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2)
                                End If
                                
                                'cell's right border
                                Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                Picture1.CurrentX = TableCellLeft(someThing + 1) + TableBorderWidth + 1
                                If TableBorderWidth2 > 2 Then
                                    Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - IIf((TableCellHasBorder(someThing + 2) = False) Or (someThing = metrCell - 1), 0, 2) - TableBorderWidth - TableBorderWidth2), , BF
                                Else
                                    Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - IIf((TableCellHasBorder(someThing + 2) = False) Or (someThing = metrCell - 1), 0, 2) - TableBorderWidth - TableBorderWidth2)
                                End If
                                
                                'cell's bottom border
                                Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                Picture1.CurrentY = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth2 + 1
                                If TableBorderWidth2 > 2 Then
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2), , BF
                                Else
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2, Picture1.CurrentY + TableBorderWidth2 - 2)
                                End If
                            End If
                        End If
                    Next someThing
                    If Not IsEmpty(BorderColor) Then
                        Picture1.ForeColor = BorderColorBefore
                    End If
                    
                    LeftMarginNext = TableCellLeft(0)
                    R.Top = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell
                    R.Left = TableCellLeft(0)
                    R.Right = R.Left
                
                    'If we are around a picture, reset it
                    If DrawTextAtMaxY <> 0 Then
                        'R.Top = DrawTextAtMaxY + 3
                        MaxX = MaxX + ExtraMarginRight
                        'R.Right = LeftMargin - ExtraMarginLeft
                        'R.Left = R.Right
                        DrawTextAtMaxY = 0
                        ExtraMarginLeft = 0
                        ExtraMarginRight = 0
                    End If
                End If
                R.Bottom = R.Top
                
                metrCell = 0
                TableLineMaxHeight = 0
                LeftMargin = TableCellLeft(0)
                CellMargin = 0
                MaxX = Picture1.Width
                TableLineFixedPixels = 0
                CellPercentEat = 0
                StartLeft = R.Left: StartTop = R.Top
              End If
            Else ' we have a cell
                If WasTable Then
                    R.Top = R.Top - TableBorderWidth - TableBorderWidth2
                    WasTable = False
                End If
                If metrCell <> 0 Then
                    metrCell = metrCell + 1
                    If TableLineMaxHeight < R.Bottom - TableLineTop Then TableLineMaxHeight = R.Bottom - TableLineTop
                    R.Right = TableCellLeft(metrCell - 1) + PixelsAfterCell
                    R.Top = TableLineTop
                Else
                    'If we are around a picture, reset it
                    If DrawTextAtMaxY <> 0 Then
                        R.Top = DrawTextAtMaxY + 3
                        MaxX = MaxX + ExtraMarginRight
                        R.Right = LeftMargin - ExtraMarginLeft
                        R.Left = R.Right
                        DrawTextAtMaxY = 0
                        ExtraMarginLeft = 0
                        ExtraMarginRight = 0
                    End If
                    
                    TableLineTop = R.Top
                    metrCell = 1
                    ReDim TableCellLeft(0) As Long
                    TableCellLeft(0) = LeftMargin
                    R.Right = TableCellLeft(0) + PixelsAfterCell
                End If
                ReDim Preserve TableCellLeft(metrCell) As Long
                If InStr(Right(WhatIs, 2), "f") <> 0 Then 'this cell has fixed width
                    TableLineFixedPixels = TableLineFixedPixels + Val(Mid(WhatIs, 2))
                    TableCellLeft(metrCell) = TableCellLeft(metrCell - 1) + Val(Mid(WhatIs, 2)) + PixelsAfterCell
                Else 'this cell has NOT fixed width
                    'some maths here
                    aTempVal = 100 - CellPercentEat ' persent left
                    If aTempVal > 0 Then
                        aTempVal = Val(Mid(WhatIs, 2)) / aTempVal ' new persent (relative to pixels that have left)
                    Else
                        aTempVal = 0
                    End If
                    'compute pixels for this cell (in truth we set the max X position of the cell)
                    TableCellLeft(metrCell) = TableCellLeft(metrCell - 1) - PixelsAfterCell + Round((Picture1.Width - 1 - TableCellLeft(metrCell - 1)) * aTempVal)
                    'lets know how much percent we have eat
                    CellPercentEat = CellPercentEat + Val(Mid(WhatIs, 2))
                End If
                'fix round problem at right margin
                If Abs(TableCellLeft(metrCell) + TableBorderWidth + TableBorderWidth2 - Picture1.Width) < TableBorderWidth + TableBorderWidth2 + 1 Then
                    TableCellLeft(metrCell) = Picture1.Width - TableBorderWidth - TableBorderWidth2
                End If
                
                ReDim Preserve TableCellHasBorder(metrCell) As Boolean
                TableCellHasBorder(metrCell) = InStr(Right(WhatIs, 2), "n") = 0
                LeftMargin = TableCellLeft(metrCell - 1) + PixelsAfterCell '+ LeftMarginNext
                CellMargin = TableCellLeft(metrCell - 1) + PixelsAfterCell
                R.Left = CellMargin
                R.Right = CellMargin
                'LeftMarginNext = TableCellLeft(metrCell - 1) + PixelsAfterCell
                LeftMarginNext = 0
                MaxX = TableCellLeft(metrCell) - PixelsAfterCell + TableBorderWidth + TableBorderWidth2
                R.Top = R.Top + PixelsAtTopBottomCell
                StartLeft = R.Left: StartTop = R.Top
            End If
        Case "z":
            If Mid(WhatIs, 2, 1) = "2" Then
                If StoredY2 = 0 Then
                    StoredY2 = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 0)
                Else
                    R.Top = StoredY2
                    If Right(WhatIs, 1) = "." Then
                        If StoredY > StoredY2 Then R.Top = StoredY
                    End If
                    StoredY2 = 0
                End If
            Else
                If StoredY = 0 Then
                    StoredY = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 0)
                    If Right(WhatIs, 1) = "." Then
                        If StoredY2 > StoredY Then StoredY = StoredY2
                    End If
                Else
                    R.Top = StoredY
                    If Right(WhatIs, 1) = "." Then
                        If StoredY2 > StoredY Then R.Top = StoredY2
                    End If
                    StoredY = 0
                End If
            End If
        Case "p":
            If Mid(WhatIs, 2, 1) = "l" Then
                ParLine = Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "a" Then
                ParAfter = Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "b" Then
                ParBefore = Val(Mid(WhatIs, 3))
            End If
    End Select
    LastWhatIs = Left(WhatIs, 1)
pali:
    If printThis = "" Then GoTo Nexti Else LastWhatIs = ""
    
    WasTable = False
    
    R.Left = R.Right
    R.Right = MaxX
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2) Then
        If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
        Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
        aTempLong = Len(printThis) - Len(RTrim(printThis))
        If aTempLong > 0 Then
            aTempLong = PicTemp.TextWidth(RTrim(printThis))
            rRight = R.Left + aTempLong
        ElseIf rRight > 0 Then
            rRight = 0
        End If
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = MaxX
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(PicTemp.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2)) Then
                Exit Do
            End If
            rRight = 0
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            'calculate rect
            Call DrawText(PicTemp.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            Call DrawText(PicTemp.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            GoSub MoveTheLine
            HasPrintSomething = False
            R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + MaxHeightOf1Line * ParLine

            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            
            If DrawTextAtMaxY <> 0 Then
                If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                    R.Left = R.Left - ExtraMarginLeft
                    MaxX = MaxX + ExtraMarginRight
                    LeftMargin = LeftMargin - ExtraMarginLeft
                    If Align = vbCenter And ExtraMarginRight > 0 Then
                        'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                        ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                        LeftMargin = LeftMargin + ExtraMarginLeft
                        R.Right = LeftMargin
                        R.Top = DrawTextAtMaxYStart
                        ExtraMarginRight = 0
                    Else
                        R.Right = LeftMargin + NextMargin
                        R.Top = DrawTextAtMaxY + 3
                        ExtraMarginLeft = 0: ExtraMarginRight = 0
                        DrawTextAtMaxY = 0
                    End If
                    StartLeft = R.Right
                    StartTop = R.Top
                    GoTo pali
                End If
            End If
            
            R.Left = LeftMargin + NextMargin
            StartLeft = R.Left: StartTop = R.Top
            
            If printThis <> "" Then GoTo here
        Else
            If Left(WhatIs, 1) <> "s" Then LastMaxHeightOf1Line = MaxHeightOf1Line
here:
            If HasPrintSomething Then
                GoSub MoveTheLine
                R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + LastMaxHeightOf1Line * ParLine
                R.Left = LeftMargin + NextMargin
                StartLeft = R.Left: StartTop = R.Top
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = MaxX
            Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            If Abs(R.Top - (R.Bottom - HeightOf1Line)) <= 2 Then
                Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
                aTempLong = Len(printThis) - Len(RTrim(printThis))
                If aTempLong > 0 Then
                    aTempLong = PicTemp.TextWidth(RTrim(printThis))
                    rRight = R.Left + aTempLong
                ElseIf rRight > 0 Then
                    R.Right = 0
                End If
            Else
                'R.Bottom = R.Bottom - HeightOf1Line
                'aTempLong = R.Bottom '- HeightOf1Line
                While printThis <> "" 'R.Top < aTempLong
                    
                    If DrawTextAtMaxY <> 0 Then
                        If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                            R.Left = R.Left - ExtraMarginLeft
                            MaxX = MaxX + ExtraMarginRight
                            LeftMargin = LeftMargin - ExtraMarginLeft
                            If Align = vbCenter And ExtraMarginRight > 0 Then
                                'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                                ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                                LeftMargin = LeftMargin + ExtraMarginLeft
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxYStart
                                ExtraMarginRight = 0
                            Else
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxY + 3
                                ExtraMarginLeft = 0: ExtraMarginRight = 0
                                DrawTextAtMaxY = 0
                            End If
                            StartLeft = R.Right
                            StartTop = R.Top
                            GoTo pali
                        End If
                    End If
                    'R.Left = LeftMargin + NextMargin
                                        
                    R.Bottom = R.Top + HeightOf1Line
                    'print last text
                    Call DrawTextEx(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                    'calculate last rect
                    Call DrawText(PicTemp.hdc, Left(printThis, textParams.uiLengthDrawn), textParams.uiLengthDrawn, R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    aTempLong = textParams.uiLengthDrawn - Len(RTrim(Left(printThis, textParams.uiLengthDrawn)))
                    If aTempLong > 0 Then
                        aTempLong = PicTemp.TextWidth(RTrim(Left(printThis, textParams.uiLengthDrawn)))
                        rRight = R.Left + aTempLong
                    ElseIf rRight > 0 Then
                        rRight = 0
                    End If
                    
                    'StartLeft = R.Left
                    StartTop = R.Top
                    GoSub MoveTheLine
                    R.Top = R.Top + HeightOf1Line * ParLine
                    printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                Wend
                R.Top = R.Top - HeightOf1Line * ParLine
                
                'R.Top = R.Bottom
                'R.Bottom = R.Top + HeightOf1Line * ParLine
                'R.Top = R.Bottom - HeightOf1Line
                                
                                
                                
                'Call DrawTextEx(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL Or IIf(Alignment = vbCenter, DT_CENTER, DT_RIGHT), textParams)
                'aTempLong = Len(Left(printThis, textParams.uiLengthDrawn)) - Len(RTrim(Left(printThis, textParams.uiLengthDrawn)))
                'GoSub MoveTheLine
                'R.Left = LeftMargin + NextMargin
                'R.Top = R.Bottom
                'R.Bottom = R.Bottom + HeightOf1Line * ParLine
                'R.Top = R.Bottom - HeightOf1Line
                
                'StartLeft = R.Left: StartTop = R.Top
                'printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                'Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                'Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
                                
                'HERE
                'aTempLong = Len(printThis) - Len(RTrim(printThis))
                'If aTempLong > 0 Then
                '    aTempLong = PicTemp.TextWidth(RTrim(printThis))
                '    rRight = R.Left + aTempLong
                'End If
            End If
            MaxHeightOf1Line = HeightOf1Line: MaxHeightOfBaseLine = TM.tmAscent
            LastMaxHeightOf1Line = MaxHeightOf1Line
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext + CellMargin + ExtraMarginLeft
Next i

If Picture1.CurrentY + HeightOf1Line > R.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = R.Bottom
End If

If Picture1.Height > PicTemp.ScaleHeight Then
    VScroll1.Enabled = True
    aTempVal = VScroll1.Value / VScroll1.Max
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = PicTemp.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    aTempVal = CLng(aTempVal * VScroll1.Max)
    If aTempVal >= VScroll1.Min And aTempVal <= VScroll1.Max Then
        If VScroll1.Value = aTempVal Then
            VScroll1_Change
        Else
            VScroll1.Value = aTempVal
        End If
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

'Picture1.FontSize = OldFontSize
'Picture1.ForeColor = OldForeColor
'Picture1.FontName = OldFontName
'Picture1.DrawStyle = OldDrawStyle
If Not PicTemp.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    WasTable = False
    
    R.Left = R.Right
    R.Right = MaxX
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < MaxX) And (R.Bottom <= R.Top + MaxHeightOf1Line + 2) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).R(LinkAreaNumber) = R
            LastLineLink = LinkAreaNumber
        'End If
        If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
        Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
        aTempLong = Len(printThis) - Len(RTrim(printThis))
        If aTempLong > 0 Then
            aTempLong = PicTemp.TextWidth(RTrim(printThis))
            rRight = R.Left + aTempLong
        End If
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = MaxX
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(PicTemp.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2)) Then
                Exit Do
            End If
            rRight = 0
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                R.Right = MaxX
                Call DrawText(PicTemp.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).R(LinkAreaNumber) = R
                LastLineLink = LinkAreaNumber
            'End If
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            Call DrawText(PicTemp.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            GoSub MoveTheLine
            LastLineLink = 0
            HasPrintSomething = False
            R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + MaxHeightOf1Line * ParLine
            
            StartLeft = R.Left: StartTop = R.Top
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            
            If DrawTextAtMaxY <> 0 Then
                If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                    R.Left = R.Left - ExtraMarginLeft
                    MaxX = MaxX + ExtraMarginRight
                    LeftMargin = LeftMargin - ExtraMarginLeft
                    If Align = vbCenter And ExtraMarginRight > 0 Then
                        'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                        ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                        LeftMargin = LeftMargin + ExtraMarginLeft
                        R.Right = LeftMargin
                        R.Top = DrawTextAtMaxYStart
                        ExtraMarginRight = 0
                    Else
                        R.Right = LeftMargin + NextMargin
                        R.Top = DrawTextAtMaxY + 3
                        ExtraMarginLeft = 0: ExtraMarginRight = 0
                        DrawTextAtMaxY = 0
                    End If
                    StartLeft = R.Left: StartTop = R.Top
                    GoTo DoJodWithLink
                End If
            End If
            
            R.Left = LeftMargin + NextMargin
            If printThis <> "" Then GoTo hereWithLink
        Else
            If Left(WhatIs, 1) <> "s" Then LastMaxHeightOf1Line = MaxHeightOf1Line
hereWithLink:
            If HasPrintSomething Then
                GoSub MoveTheLine
                R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + LastMaxHeightOf1Line * ParLine
                R.Left = LeftMargin + NextMargin
                StartLeft = R.Left: StartTop = R.Top
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = MaxX
            Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            If Abs(R.Top - (R.Bottom - HeightOf1Line)) <= 2 Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                    LastLineLink = LinkAreaNumber
                'End If
                Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
                
                aTempLong = Len(printThis) - Len(RTrim(printThis))
                If aTempLong > 0 Then
                    aTempLong = PicTemp.TextWidth(RTrim(printThis))
                    rRight = R.Left + aTempLong
                End If
            Else
                ''
                aTempLong3 = R.Top
                While printThis <> "" 'R.Top < aTempLong
                    
                    If DrawTextAtMaxY <> 0 Then
                        If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                            'calculate 1 big rect
                            R.Top = aTempLong
                            LinkAreaNumber = LinkAreaNumber + 1
                            Links(UBound(Links)).R(LinkAreaNumber) = R
                            
                            R.Left = R.Left - ExtraMarginLeft
                            MaxX = MaxX + ExtraMarginRight
                            LeftMargin = LeftMargin - ExtraMarginLeft
                            If Align = vbCenter And ExtraMarginRight > 0 Then
                                'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                                ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                                LeftMargin = LeftMargin + ExtraMarginLeft
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxYStart
                                ExtraMarginRight = 0
                            Else
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxY + 3
                                ExtraMarginLeft = 0: ExtraMarginRight = 0
                                DrawTextAtMaxY = 0
                            End If
                            StartLeft = R.Left: StartTop = R.Top
                            GoTo DoJodWithLink
                        End If
                    End If
                    ''
                    
                    R.Bottom = R.Top + HeightOf1Line
                    'print last text
                    Call DrawTextEx(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                    'calculate last rect
                    Call DrawText(PicTemp.hdc, Left(printThis, textParams.uiLengthDrawn), textParams.uiLengthDrawn, R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    aTempLong = textParams.uiLengthDrawn - Len(RTrim(Left(printThis, textParams.uiLengthDrawn)))
                    If aTempLong > 0 Then
                        aTempLong = PicTemp.TextWidth(RTrim(Left(printThis, textParams.uiLengthDrawn)))
                        rRight = R.Left + aTempLong
                    ElseIf rRight > 0 Then
                        rRight = 0
                    End If
                    
                    'StartLeft = R.Left
                    StartTop = R.Top
                    GoSub MoveTheLine
                    R.Top = R.Top + HeightOf1Line * ParLine
                    printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                Wend
                R.Top = R.Top - HeightOf1Line * ParLine
                
                                
                'If WasLink Then
                    'calculate 1 big rect
                    aTempLong = R.Right
                    R.Right = MaxX
                    R.Top = aTempLong3
                    R.Bottom = R.Bottom - HeightOf1Line * ParLine
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                    R.Bottom = R.Bottom + HeightOf1Line * ParLine
                    R.Top = R.Bottom - HeightOf1Line

                    'calculate last rect
                    Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    R.Right = aTempLong
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                    LastLineLink = LinkAreaNumber
                'End If
                'Call DrawText(PicTemp.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)

                aTempLong = Len(printThis) - Len(RTrim(printThis))
                If aTempLong > 0 Then
                    aTempLong = PicTemp.TextWidth(RTrim(printThis))
                    rRight = R.Left + aTempLong
                End If
            End If
            MaxHeightOf1Line = HeightOf1Line: MaxHeightOfBaseLine = TM.tmAscent
            LastMaxHeightOf1Line = MaxHeightOf1Line
        End If
    End If

Return

MoveTheLine:

If rRight = 0 Then rRight = R.Right
If R.Bottom - R.Top < MaxHeightOf1Line Then
    aTempLong = MaxHeightOf1Line
Else
    aTempLong = R.Bottom - R.Top
End If
If Alignment = vbCenter Then
    aTempLong2 = CLng(StartLeft + ((MaxX - StartLeft) - (rRight - StartLeft)) \ 2) - 0
    ' from Picture1 to Picture1
    'BitBlt Picture1.hdc, aTempLong2, _
        StartTop, rRight - StartLeft + 1, aTempLong, Picture1.hdc, _
        StartLeft - 1, StartTop, vbSrcCopy
    
    ' from PicTemp to Picture1
    BitBlt Picture1.hdc, aTempLong2, _
        StartTop, rRight - StartLeft + 1, aTempLong, PicTemp.hdc, _
        StartLeft - 0, StartTop, vbSrcCopy
Else
    aTempLong2 = CLng(StartLeft + (MaxX - StartLeft) - (rRight - StartLeft)) - 0
    ' from Picture1 to Picture1
    'BitBlt Picture1.hdc, aTempLong2, _
        StartTop, rRight - StartLeft + 1, aTempLong, Picture1.hdc, _
        StartLeft - 1, StartTop, vbSrcCopy
    
    ' from PicTemp to Picture1
    BitBlt Picture1.hdc, aTempLong2, _
        StartTop, rRight - StartLeft + 1, aTempLong, PicTemp.hdc, _
        StartLeft - 0, StartTop, vbSrcCopy

End If
rRight = 0

If LastLineLink <> 0 Then
    Links(UBound(Links)).R(LastLineLink).Right = Links(UBound(Links)).R(LastLineLink).Right + aTempLong2 - StartLeft
    Links(UBound(Links)).R(LastLineLink).Left = Links(UBound(Links)).R(LastLineLink).Left + aTempLong2 - StartLeft
End If

' from Picture1 to Picture1
'If aTempLong2 > StartLeft Then
'    aTempLong = Picture1.ForeColor
'    Picture1.CurrentX = StartLeft - 1
'    Picture1.CurrentY = StartTop
'    Picture1.ForeColor = Picture1.BackColor
'
'    Picture1.Line -(aTempLong2 - IIf(metrCell = 0, 0, 0), StartTop + R.Bottom - R.Top), , BF
'    Picture1.ForeColor = aTempLong
'End If

Return

MoveActiveLineToBaseLine:

    'BitBlt PicTemp.hdc, LeftMargin + NextMargin, R.Top + MustMoveActiveLineToBaseLine - MaxHeightOfBaseLine, R.Right - LeftMargin - NextMargin, R.Bottom - R.Top, _
            PicTemp.hdc, LeftMargin + NextMargin, R.Top, vbSrcCopy
    BitBlt PicTemp.hdc, StartLeft, R.Top - MustMoveActiveLineToBaseLine + MaxHeightOfBaseLine, R.Left - StartLeft, LastMaxHeightOf1Line, _
            PicTemp.hdc, StartLeft, R.Top + TM.tmAscent - MaxHeightOfBaseLine, vbSrcCopy
            
    PicTemp.CurrentX = StartLeft
    PicTemp.CurrentY = R.Top + TM.tmAscent - MaxHeightOfBaseLine + 1
    aTempLong2 = PicTemp.ForeColor
    PicTemp.Tag = PicTemp.DrawStyle
    PicTemp.DrawStyle = 0
    PicTemp.ForeColor = PicTemp.BackColor
    'PicTemp.Line -(R.Right, R.Top + MustMoveActiveLineToBaseLine - MaxHeightOfBaseLine - 1), , BF
    PicTemp.Line -(R.Left, R.Top - MustMoveActiveLineToBaseLine + MaxHeightOfBaseLine - 1), , BF
    PicTemp.ForeColor = aTempLong2
    PicTemp.DrawStyle = Val(PicTemp.Tag)
    
    MustMoveActiveLineToBaseLine = 0
Return


End Sub


Private Sub DrawTheText()
Dim aStr As String, printThis As String, i As Long, someThing As Long
Dim HeightOf1Line As Long, cHeight As Long, R As RECT, pl As Long, LastWhatIs As String
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long, NextMargin As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMargin As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single
Dim OldFontName As String, OldDrawStyle As Long, aTempVal As Single
Dim TableCellLeft() As Long, TableLineTop As Long, metrCell As Long, WasTable As Boolean, TableLineMaxHeight As Long
Dim TableCellHasBorder() As Boolean, TableLineFixedPixels As Long, CellPercentEat As Single, BorderColor, BorderColorBefore As Long
Dim CellMargin As Long, PixelsAfterCell As Long, PixelsAtTopBottomCell As Long, TableBorderWidth As Long, TableBorderWidth2 As Long
Dim MaxX As Long, MaxHeightOfBaseLine As Long, StoredY As Long, StoredY2 As Long
Dim stackFromLevel As Long, stackLenLevel As Long, stackFormatLevel As Long
Dim aTempLong As Long, aTempLong2 As Long, TM As TEXTMETRIC, LastMaxHeightOf1Line As Long
'Dim ActiveLineHeight As Long
Dim MustMoveActiveLineToBaseLine As Long
Dim DrawTextAtMaxY As Long, DrawTextAtMaxYStart As Long, DrawAroundPicture As Boolean, ExtraMarginLeft As Long, ExtraMarginRight As Long

Dim Alignment As AlignmentConstants, ParBefore As Long, ParAfter As Long, ParLine As Single

Dim PicName As String, PicHandle As Long, PicWidth As Long, PicHeight As Long
Dim Align As AlignmentConstants


If Not mAutoRedraw Then Exit Sub

'SetTextAlign Picture1.hdc, TA_BASELINE
'SetTextAlign PicTemp.hdc, TA_BASELINE

PixelsAtTopBottomCell = 2
PixelsAfterCell = 2
TableBorderWidth = 1

ParLine = 1

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

Picture1.Height = MaxHeight
PicTemp.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

Picture1.FontBold = False
Picture1.FontItalic = False
Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
OldFontName = Picture1.FontName
OldDrawStyle = Picture1.DrawStyle
If Not UserControl.Enabled Then
    Picture1.ForeColor = &H80000011
    PicTemp.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
LastMaxHeightOf1Line = MaxHeightOf1Line
GetTextMetrics Picture1.hdc, TM
MaxHeightOfBaseLine = TM.tmAscent

MaxX = Picture1.Width
Picture1.Cls
PicTemp.Cls
stackFromLevel = stackFrom.stackLevel
stackLenLevel = stackLen.stackLevel
stackFormatLevel = stackFormat.stackLevel
For i = 1 To stackLenLevel
    TheFrom = stackFrom.popNo(stackFromLevel - i + 1)
    TheLen = stackLen.popNo(stackLenLevel - i + 1)
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
If InStr(printThis, "can see") <> 0 Then
    'Beep
End If
    
    WhatIs = stackFormat.popNo(stackFormatLevel - i + 1)
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            If Mid(WhatIs, 2, 1) = "g" Then ' align
                Select Case Mid(WhatIs, 3, 1)
                    Case "l": Alignment = vbLeftJustify
                    Case "c": Alignment = vbCenter
                    Case "r": Alignment = vbRightJustify
                End Select
            
                If Alignment <> vbLeftJustify Then
                    DrawTheTextByLine R, i, LeftMargin, LeftMarginNext, NextMargin, metrCell, TableCellLeft(), TableLineTop, _
                    WasTable, TableLineMaxHeight, TableCellHasBorder(), TableLineFixedPixels, CellPercentEat, _
                    BorderColor, BorderColorBefore, CellMargin, PixelsAfterCell, PixelsAtTopBottomCell, _
                    TableBorderWidth, TableBorderWidth2, MaxX, HeightOf1Line, MaxHeightOf1Line, LastMaxHeightOf1Line, MaxHeightOfBaseLine, _
                    MustMoveActiveLineToBaseLine, TM, ParLine, ParBefore, ParAfter, DrawTextAtMaxY, DrawTextAtMaxYStart, DrawAroundPicture, _
                    ExtraMarginLeft, ExtraMarginRight, Align, StoredY, StoredY2
                    GoTo Nexti
                End If
         
            Else ' new line
                R.Top = R.Top + HeightOf1Line * ParLine + ParAfter + ParBefore
                
                If DrawTextAtMaxY <> 0 Then
                    If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                        'R.Left = R.Left - ExtraMarginLeft
                        MaxX = MaxX + ExtraMarginRight
                        LeftMargin = LeftMargin - ExtraMarginLeft
                        If Align = vbCenter And ExtraMarginRight > 0 Then
                            ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                            LeftMargin = LeftMargin + ExtraMarginLeft
                            R.Right = LeftMargin
                            R.Top = DrawTextAtMaxYStart
                            ExtraMarginRight = 0
                        Else
                            R.Top = DrawTextAtMaxY + 3
                            ExtraMarginLeft = 0: ExtraMarginRight = 0
                            DrawTextAtMaxY = 0
                        End If
                    End If
                End If
                
                R.Left = LeftMargin
                R.Right = R.Left
                NextMargin = 0
                HasPrintSomething = False
                MaxHeightOf1Line = HeightOf1Line
                
                MaxHeightOfBaseLine = TM.tmAscent
            End If
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin - ExtraMarginLeft + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin - ExtraMarginLeft + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                R.Left = LeftMarginNext
                R.Right = LeftMarginNext
            ElseIf metrCell <> 0 And LastWhatIs = "a" Then
                R.Right = CellMargin + LeftMarginNext
                R.Left = R.Right
                LeftMargin = LeftMarginNext + CellMargin
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
SameAsFontSize:
            LastMaxHeightOf1Line = MaxHeightOf1Line
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
            
            aTempLong = TM.tmAscent
            GetTextMetrics Picture1.hdc, TM
                        
            If MaxHeightOfBaseLine < TM.tmAscent Then
                If HasPrintSomething Then
                    MustMoveActiveLineToBaseLine = MaxHeightOfBaseLine
                End If
                R.Top = R.Top + aTempLong - MaxHeightOfBaseLine
                MaxHeightOfBaseLine = TM.tmAscent
            Else
                'If HasPrintSomething Then R.Top = R.Top + aTempLong - TM.tmAscent
                R.Top = R.Top + aTempLong - TM.tmAscent
                MustMoveActiveLineToBaseLine = 0
            End If
                                    
        Case "e": ' print a line
            Picture1.CurrentY = R.Top
            Picture1.CurrentX = R.Right
            aTempVal = InStr(WhatIs, "x")
            If aTempVal <> 0 Then 'user set a percent for the line
                aTempVal = Val(Mid(WhatIs, aTempVal + 1))
            End If
            If aTempVal <= 0 Then aTempVal = 100
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(aTempVal / 100 * MaxX, Picture1.CurrentY)
                R.Top = R.Top + 2
            Else
                Picture1.Line -(aTempVal / 100 * MaxX - 1, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
                R.Top = R.Top + Val(Mid(WhatIs, 2)) + 1
            End If
            R.Bottom = R.Top
            R.Left = LeftMargin
        Case "f": 'font
            On Error Resume Next
            If Val(Mid(WhatIs, 2)) < 1 Then
                Picture1.FontName = OldFontName
            Else
                Picture1.FontName = FontCol.Item(Val(Mid(WhatIs, 2)))
            End If
            On Error GoTo ErrHandle
            GoTo SameAsFontSize
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            R.Top = R.Top + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            aSingle = R.Top
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                Picture1.CurrentX = R.Left + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                Picture1.CurrentX = R.Left + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            HasPrintSomething = True
            Picture1.DrawWidth = 1
            R.Left = Picture1.CurrentX + 7
            R.Right = R.Left
            R.Top = aSingle
        Case "d":
            'user want to set the border style for drawing
            Picture1.DrawStyle = Val(Mid(WhatIs, 2))
        Case "g": 'user want to paint a picture
          If WhatIs = "ga" Then
            DrawAroundPicture = True
          ElseIf WhatIs = "gn" Then
            DrawAroundPicture = False
          ElseIf WhatIs = "gr" Then
            If DrawTextAtMaxY <> 0 And R.Top < DrawTextAtMaxY + 3 Then
                R.Top = DrawTextAtMaxY + 3
                MaxX = MaxX + ExtraMarginRight
                R.Right = LeftMargin - ExtraMarginLeft
                R.Left = R.Right
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
            End If
          Else
            
            aTempVal = 0
            PicWidth = 0
            PicHeight = InStr(3, WhatIs, "|")
            If PicHeight = 0 Then 'picture will paint using current dimensions
                If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
                    PicHandle = Val(Mid(WhatIs, 2))
                Else 'picture  by path
                    PicName = Mid(WhatIs, 2)
                End If
                Align = vbLeftJustify
            Else
                If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
                    PicHandle = Val(Mid(WhatIs, 3, PicHeight - 3))
                Else 'picture  by path
                    PicName = Mid(WhatIs, 2, PicHeight - 2)
                End If
                PicWidth = Val(Mid(WhatIs, PicHeight + 1))
                PicHeight = InStr(PicHeight + 1, WhatIs, "x")
                If PicHeight <> 0 Then
                    PicHeight = Val(Mid(WhatIs, PicHeight + 1))
                End If
                If InStr(Right(WhatIs, 2), "f") = 0 Then 'use percent for dimensions
                    'must be PicWidth <> 0 Or PicHeight <> 0
                    If PicWidth <> 0 Or PicHeight <> 0 Then
                        aTempVal = MaxX - R.Right
                    End If
                End If
                If InStr(Right(WhatIs, 2), "r") <> 0 Then
                    Align = vbRightJustify
                ElseIf InStr(Right(WhatIs, 2), "c") <> 0 Then
                    Align = vbCenter
                Else
                    Align = vbLeftJustify
                End If
            End If
            
            If DrawTextAtMaxY <> 0 Then
                R.Top = DrawTextAtMaxY + 3
                MaxX = MaxX + ExtraMarginRight
                R.Right = LeftMargin - ExtraMarginLeft
                R.Left = R.Right
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
            End If
            
            R.Top = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 2)
            If Mid(WhatIs, 2, 1) = "|" Then 'picture by bitmap handle
LoadByHandle:
                PaintPictureByHandleGgiPlus PicHandle, Picture1.hdc, R.Right, R.Top, PicWidth, PicHeight, aTempVal <> 0, MaxX - R.Right, 0, Align, True
                Set UserControl.Picture = Nothing
            Else 'picture  by path
                If UCase(Left(PicName, 10)) = "<APP.PATH>" Then
                    PicName = App.Path & Mid(PicName, 11)
                ElseIf UCase(Left(PicName, 3)) = "<R>" Then
                    If UCase(Right(PicName, 1)) = "I" Then 'Icon
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResIcon)
                    ElseIf UCase(Right(PicName, 1)) = "C" Then 'Cursor
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResCursor)
                    Else 'Bitmap
                        Set UserControl.Picture = LoadResPicture(Val(Mid(PicName, 4)), vbResBitmap)
                    End If
                    PicHandle = UserControl.Picture.Handle
                    GoTo LoadByHandle
                End If
                PaintPictureGgiPlus PicName, Picture1.hdc, R.Right, R.Top, PicWidth, PicHeight, aTempVal <> 0, MaxX - R.Right, Align, True
            End If
            
            If DrawAroundPicture Then '/draw around picture
                If DrawTextAtMaxY < R.Top + PicHeight Then DrawTextAtMaxY = R.Top + PicHeight
                MaxX = MaxX + ExtraMarginRight
                If Align = vbRightJustify Then
                    'ExtraMarginLeft = 0
                    ExtraMarginRight = ExtraMarginRight + PicWidth + 3
                ElseIf Align = vbCenter Then
                    'ExtraMarginRight = ExtraMarginRight + PicWidth + 3
                    ExtraMarginRight = MaxX - ExtraMarginRight - R.Right + 3
                    'DrawTextAtMaxYStart = R.Top - IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
                    DrawTextAtMaxYStart = R.Top - IIf(HasPrintSomething, -1, -0)
                Else
                    ExtraMarginLeft = ExtraMarginLeft + LeftMargin - CellMargin + PicWidth + 3
                    'ExtraMarginRight = 0
                End If
                R.Right = LeftMargin + ExtraMarginLeft
                MaxX = MaxX - ExtraMarginRight
                
                If metrCell <> 0 Then
                    If TableLineMaxHeight < R.Top + PicHeight - TableLineTop Then TableLineMaxHeight = R.Top + PicHeight - TableLineTop
                End If
                'R.Top = R.Top - HeightOf1Line * ParLine + ParAfter + ParBefore - IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
                R.Top = R.Top - HeightOf1Line * ParLine + ParAfter + ParBefore '- IIf(HasPrintSomething, HeightOf1Line * ParLine + ParAfter + ParBefore, 2)
            Else
                DrawTextAtMaxY = 0
                ExtraMarginLeft = 0
                ExtraMarginRight = 0
                
                R.Top = R.Top + PicHeight - MaxHeightOf1Line
                R.Right = R.Right + PicWidth
                HasPrintSomething = True
                
                If metrCell <> 0 Then
                    If TableLineMaxHeight < R.Top + MaxHeightOf1Line - TableLineTop Then TableLineMaxHeight = R.Top + MaxHeightOf1Line - TableLineTop
                End If
            End If
          End If
        Case "a": ' table cell
            If Val(Mid(WhatIs, 2)) = 0 Then 'END of cells
              If Mid(WhatIs, 2, 2) = "bc" Then 'user want to set the border color for the next table
                BorderColor = Val(Mid(WhatIs, 4))
              ElseIf Mid(WhatIs, 2, 1) = "b" Then 'user want to set the border width for the next table
                PixelsAfterCell = PixelsAfterCell - TableBorderWidth - TableBorderWidth2
                PixelsAtTopBottomCell = PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2
                
                TableBorderWidth = Val(Mid(WhatIs, 3))
                aTempVal = InStr(WhatIs, "x")
                If aTempVal = 0 Then
                    TableBorderWidth2 = 0
                Else
                    TableBorderWidth2 = Val(Mid(WhatIs, aTempVal + 1)) + 1
                    If TableBorderWidth2 = 1 Then TableBorderWidth2 = 0
                End If
                PixelsAfterCell = PixelsAfterCell + TableBorderWidth + TableBorderWidth2
                PixelsAtTopBottomCell = PixelsAtTopBottomCell + TableBorderWidth + TableBorderWidth2
              ElseIf Mid(WhatIs, 2, 2) = "mt" Then 'user want to set the margin at top and bottom for a cell
                PixelsAtTopBottomCell = Val(Mid(WhatIs, 4)) + TableBorderWidth + TableBorderWidth2
              ElseIf Mid(WhatIs, 2, 1) = "m" Then 'user want to set the margin at left and right for a cell
                PixelsAfterCell = Val(Mid(WhatIs, 3)) + TableBorderWidth + TableBorderWidth2
              Else 'user said that a table line is end
                WasTable = True
                If metrCell <> 0 Then 'we have cells, so lets print the borders
                    'fix round problem at right margin
                    If Abs(TableCellLeft(metrCell) + TableBorderWidth + TableBorderWidth2 - Picture1.Width) < TableBorderWidth + TableBorderWidth2 + 1 Then
                        TableCellLeft(metrCell) = Picture1.Width - TableBorderWidth - TableBorderWidth2
                    End If
                                        
                    If TableLineMaxHeight < R.Bottom - TableLineTop Then TableLineMaxHeight = R.Bottom - TableLineTop
                    If Not IsEmpty(BorderColor) Then
                        BorderColorBefore = Picture1.ForeColor
                        Picture1.ForeColor = BorderColor
                    End If
                    ReDim Preserve TableCellHasBorder(metrCell + 1)
                    For someThing = 0 To metrCell - 1
                        If TableCellHasBorder(someThing + 1) And TableBorderWidth > 0 Then 'has this cell border?
                            'cell's left border
                            Picture1.CurrentY = TableLineTop
                            Picture1.CurrentX = TableCellLeft(someThing)
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                            Else
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                            End If
                        
                            'cell's top border
                            Picture1.CurrentY = TableLineTop
                            If TableBorderWidth > 1 Then
                                'if Picture1.DrawStyle <> 0 then
                                'delete any line, if was a line there fron any previus row
                                'in order to have a better line
                                If Picture1.DrawStyle <> 0 Then
                                    aTempLong = Picture1.ForeColor
                                    Picture1.ForeColor = Picture1.BackColor
                                    aTempLong2 = Picture1.DrawStyle
                                    Picture1.DrawStyle = 0
                                    Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1), , BF
                                    
                                    Picture1.CurrentX = TableCellLeft(someThing)
                                    Picture1.CurrentY = TableLineTop
                                    Picture1.ForeColor = aTempLong
                                    Picture1.DrawStyle = aTempLong2
                                End If
                                Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1), , BF
                            Else
                                'if Picture1.DrawStyle <> 0 then
                                'delete any line, if was a line there fron any previus row
                                'in order to have a better line
                                If Picture1.DrawStyle <> 0 Then
                                    aTempLong = Picture1.ForeColor
                                    Picture1.ForeColor = Picture1.BackColor
                                    aTempLong2 = Picture1.DrawStyle
                                    Picture1.DrawStyle = 0
                                    Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1)
                                    
                                    Picture1.CurrentX = TableCellLeft(someThing)
                                    Picture1.CurrentY = TableLineTop
                                    Picture1.ForeColor = aTempLong
                                    Picture1.DrawStyle = aTempLong2
                                End If
                                Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1)
                            End If
                            
                            
                            'cell's right border
                            Picture1.CurrentY = TableLineTop
                            Picture1.CurrentX = TableCellLeft(someThing + 1)
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                            Else
                                Picture1.Line -(Picture1.CurrentX + TableBorderWidth - 1, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                            End If
                            
                            'cell's bottom border
                            Picture1.CurrentX = TableCellLeft(someThing)
                            Picture1.CurrentY = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2
                            If TableBorderWidth > 1 Then
                                Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth - 1, Picture1.CurrentY + TableBorderWidth - 1), , BF
                            Else
                                Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth, Picture1.CurrentY + TableBorderWidth - 1)
                            End If
                            
                            'cell's second border
                            If TableBorderWidth2 > 0 Then
                                'cell's left border
                                If Not TableCellHasBorder(someThing) Then
                                    Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                    Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                    If TableBorderWidth2 > 2 Then
                                        Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2), , BF
                                    Else
                                        Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth - TableBorderWidth2)
                                    End If
                                End If
                                
                                'cell's top border
                                Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                If TableBorderWidth2 > 2 Then
                                    'if Picture1.DrawStyle <> 0 then
                                    'delete any line, if was a line there fron any previus row
                                    'in order to have a better line
                                    If Picture1.DrawStyle <> 0 Then
                                        aTempLong = Picture1.ForeColor
                                        Picture1.ForeColor = Picture1.BackColor
                                        aTempLong2 = Picture1.DrawStyle
                                        Picture1.DrawStyle = 0
                                        Picture1.Line -(TableCellLeft(someThing + 1), Picture1.CurrentY + TableBorderWidth - 1), , BF
                                        
                                        Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                        Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                        Picture1.ForeColor = aTempLong
                                        Picture1.DrawStyle = aTempLong2
                                    End If
                                    
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2), , BF
                                Else
                                    'if Picture1.DrawStyle <> 0 then
                                    'delete any line, if was a line there fron any previus row
                                    'in order to have a better line
                                    If Picture1.DrawStyle <> 0 Then
                                        aTempLong = Picture1.ForeColor
                                        Picture1.ForeColor = Picture1.BackColor
                                        aTempLong2 = Picture1.DrawStyle
                                        Picture1.DrawStyle = 0
                                        Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2)
                                        
                                        Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                        Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                        Picture1.ForeColor = aTempLong
                                        Picture1.DrawStyle = aTempLong2
                                    End If
                                    
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2)
                                End If
                                
                                'cell's right border
                                Picture1.CurrentY = TableLineTop + TableBorderWidth + 1
                                Picture1.CurrentX = TableCellLeft(someThing + 1) + TableBorderWidth + 1
                                If TableBorderWidth2 > 2 Then
                                    Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - IIf((TableCellHasBorder(someThing + 2) = False) Or (someThing = metrCell - 1), 0, 2) - TableBorderWidth - TableBorderWidth2), , BF
                                Else
                                    Picture1.Line -(Picture1.CurrentX + TableBorderWidth2 - 2, Picture1.CurrentY + TableLineMaxHeight + PixelsAtTopBottomCell - IIf((TableCellHasBorder(someThing + 2) = False) Or (someThing = metrCell - 1), 0, 2) - TableBorderWidth - TableBorderWidth2)
                                End If
                                
                                'cell's bottom border
                                Picture1.CurrentX = TableCellLeft(someThing) + TableBorderWidth + 1
                                Picture1.CurrentY = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell - TableBorderWidth2 + 1
                                If TableBorderWidth2 > 2 Then
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2 - 1, Picture1.CurrentY + TableBorderWidth2 - 2), , BF
                                Else
                                    Picture1.Line -(TableCellLeft(someThing + 1) + TableBorderWidth + TableBorderWidth2, Picture1.CurrentY + TableBorderWidth2 - 2)
                                End If
                            End If
                        End If
                    Next someThing
                    If Not IsEmpty(BorderColor) Then
                        Picture1.ForeColor = BorderColorBefore
                    End If
                    
                    LeftMarginNext = TableCellLeft(0)
                    R.Top = TableLineTop + TableLineMaxHeight + PixelsAtTopBottomCell
                    R.Left = TableCellLeft(0)
                    R.Right = R.Left
                    
                    'If we are around a picture, reset it
                    If DrawTextAtMaxY <> 0 Then
                        'R.Top = DrawTextAtMaxY + 3
                        MaxX = MaxX + ExtraMarginRight
                        'R.Right = LeftMargin - ExtraMarginLeft
                        'R.Left = R.Right
                        DrawTextAtMaxY = 0
                        ExtraMarginLeft = 0
                        ExtraMarginRight = 0
                    End If
                End If
                R.Bottom = R.Top
                
                metrCell = 0
                TableLineMaxHeight = 0
                LeftMargin = TableCellLeft(0)
                CellMargin = 0
                MaxX = Picture1.Width
                TableLineFixedPixels = 0
                CellPercentEat = 0
              End If
            Else ' we have a cell
                If WasTable Then
                    R.Top = R.Top - TableBorderWidth - TableBorderWidth2
                    WasTable = False
                End If
                If metrCell <> 0 Then
                    metrCell = metrCell + 1
                    If TableLineMaxHeight < R.Bottom - TableLineTop Then TableLineMaxHeight = R.Bottom - TableLineTop
                    R.Right = TableCellLeft(metrCell - 1) + PixelsAfterCell
                    R.Top = TableLineTop
                Else
                    'If we are around a picture, reset it
                    If DrawTextAtMaxY <> 0 Then
                        R.Top = DrawTextAtMaxY + 3
                        MaxX = MaxX + ExtraMarginRight
                        LeftMargin = LeftMargin - ExtraMarginLeft
                        R.Right = LeftMargin
                        R.Left = R.Right
                        DrawTextAtMaxY = 0
                        ExtraMarginLeft = 0
                        ExtraMarginRight = 0
                    End If
                
                    TableLineTop = R.Top
                    metrCell = 1
                    ReDim TableCellLeft(0) As Long
                    TableCellLeft(0) = LeftMargin
                    R.Right = TableCellLeft(0) + PixelsAfterCell
                End If
                ReDim Preserve TableCellLeft(metrCell) As Long
                If InStr(Right(WhatIs, 2), "f") <> 0 Then 'this cell has fixed width
                    TableLineFixedPixels = TableLineFixedPixels + Val(Mid(WhatIs, 2))
                    TableCellLeft(metrCell) = TableCellLeft(metrCell - 1) + Val(Mid(WhatIs, 2)) + PixelsAfterCell
                Else 'this cell has NOT fixed width
                    'some maths here
                    aTempVal = 100 - CellPercentEat ' persent left
                    If aTempVal > 0 Then
                        aTempVal = Val(Mid(WhatIs, 2)) / aTempVal ' new persent (relative to pixels that have left)
                    Else
                        aTempVal = 0
                    End If
                    'compute pixels for this cell (in truth we set the max X position of the cell)
                    TableCellLeft(metrCell) = TableCellLeft(metrCell - 1) - PixelsAfterCell + Round((Picture1.Width - 1 - TableCellLeft(metrCell - 1)) * aTempVal)
                    'lets know how much percent we have eat
                    CellPercentEat = CellPercentEat + Val(Mid(WhatIs, 2))
                End If
                'fix round problem at right margin
                If Abs(TableCellLeft(metrCell) + TableBorderWidth + TableBorderWidth2 - Picture1.Width) < TableBorderWidth + TableBorderWidth2 + 1 Then
                    TableCellLeft(metrCell) = Picture1.Width - TableBorderWidth - TableBorderWidth2
                End If
                
                ReDim Preserve TableCellHasBorder(metrCell) As Boolean
                TableCellHasBorder(metrCell) = InStr(Right(WhatIs, 2), "n") = 0
                LeftMargin = TableCellLeft(metrCell - 1) + PixelsAfterCell '+ LeftMarginNext
                CellMargin = TableCellLeft(metrCell - 1) + PixelsAfterCell
                R.Left = CellMargin
                R.Right = CellMargin
                'LeftMarginNext = TableCellLeft(metrCell - 1) + PixelsAfterCell
                LeftMarginNext = 0
                MaxX = TableCellLeft(metrCell) - PixelsAfterCell + TableBorderWidth + TableBorderWidth2
                R.Top = R.Top + PixelsAtTopBottomCell
            End If
        Case "z":
            If Mid(WhatIs, 2, 1) = "2" Then
                If StoredY2 = 0 Then
                    StoredY2 = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 0)
                Else
                    R.Top = StoredY2
                    If Right(WhatIs, 1) = "." Then
                        If StoredY > StoredY2 Then R.Top = StoredY
                    End If
                    StoredY2 = 0
                End If
            Else
                If StoredY = 0 Then
                    StoredY = R.Top + IIf(HasPrintSomething, MaxHeightOf1Line, 0)
                    If Right(WhatIs, 1) = "." Then
                        If StoredY2 > StoredY Then StoredY = StoredY2
                    End If
                Else
                    R.Top = StoredY
                    If Right(WhatIs, 1) = "." Then
                        If StoredY2 > StoredY Then R.Top = StoredY2
                    End If
                    StoredY = 0
                End If
            End If
        Case "p":
            If Mid(WhatIs, 2, 1) = "l" Then
                ParLine = Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "a" Then
                ParAfter = Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "b" Then
                ParBefore = Val(Mid(WhatIs, 3))
            End If
    End Select
    LastWhatIs = Left(WhatIs, 1)
    
pali:
    If printThis = "" Then GoTo Nexti Else LastWhatIs = ""
   
    
    WasTable = False
    
    R.Left = R.Right
    R.Right = MaxX
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2) Then
        If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
        Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = MaxX
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + MaxHeightOf1Line * ParLine
            
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            
            If DrawTextAtMaxY <> 0 Then
                If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                    R.Left = R.Left - ExtraMarginLeft
                    MaxX = MaxX + ExtraMarginRight
                    LeftMargin = LeftMargin - ExtraMarginLeft
                    If Align = vbCenter And ExtraMarginRight > 0 Then
                        'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                        ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                        LeftMargin = LeftMargin + ExtraMarginLeft
                        R.Right = LeftMargin
                        R.Top = DrawTextAtMaxYStart
                        ExtraMarginRight = 0
                    Else
                        R.Right = LeftMargin + NextMargin
                        R.Top = DrawTextAtMaxY + 3
                        ExtraMarginLeft = 0: ExtraMarginRight = 0
                        DrawTextAtMaxY = 0
                    End If
                    GoTo pali
                End If
            End If
            
            R.Left = LeftMargin + NextMargin
            If printThis <> "" Then GoTo here
        Else
            If Left(WhatIs, 1) <> "s" Then LastMaxHeightOf1Line = MaxHeightOf1Line
here:
            If HasPrintSomething Then
                R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + LastMaxHeightOf1Line * ParLine
                R.Left = LeftMargin + NextMargin
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = MaxX
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            If Abs(R.Top - (R.Bottom - HeightOf1Line)) <= 2 Then
                Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                While printThis <> "" 'R.Top < aTempLong
                    
                    If DrawTextAtMaxY <> 0 Then
                        If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                            R.Left = R.Left - ExtraMarginLeft
                            MaxX = MaxX + ExtraMarginRight
                            LeftMargin = LeftMargin - ExtraMarginLeft
                            If Align = vbCenter And ExtraMarginRight > 0 Then
                                'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                                ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                                LeftMargin = LeftMargin + ExtraMarginLeft
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxYStart
                                ExtraMarginRight = 0
                            Else
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxY + 3
                                ExtraMarginLeft = 0: ExtraMarginRight = 0
                                DrawTextAtMaxY = 0
                            End If
                            GoTo pali
                        End If
                    End If
                                        
                    R.Bottom = R.Top + HeightOf1Line
                    'print last text
                    Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                    'calculate last rect
                    Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    R.Top = R.Top + HeightOf1Line * ParLine
                    printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                Wend
                R.Top = R.Top - HeightOf1Line * ParLine
                'R.Top = R.Bottom
                'R.Bottom = R.Top + HeightOf1Line * ParLine
                'R.Top = R.Bottom - HeightOf1Line
                
                'printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                
                                
                'calculate last rect
                'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                'print last text
                'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            End If
            MaxHeightOf1Line = HeightOf1Line: MaxHeightOfBaseLine = TM.tmAscent
            LastMaxHeightOf1Line = MaxHeightOf1Line
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext + CellMargin + ExtraMarginLeft
Next i

If Picture1.CurrentY + HeightOf1Line > R.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = R.Bottom
End If

If Picture1.Height > UserControl.ScaleHeight Then
    VScroll1.Enabled = True
    aTempVal = VScroll1.Value / VScroll1.Max
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    aTempVal = CLng(aTempVal * VScroll1.Max)
    If aTempVal >= VScroll1.Min And aTempVal <= VScroll1.Max Then
        If VScroll1.Value = aTempVal Then
            VScroll1_Change
        Else
            VScroll1.Value = aTempVal
        End If
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
PicTemp.ForeColor = OldForeColor
Picture1.FontName = OldFontName
Picture1.DrawStyle = OldDrawStyle
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

stackFrom.Clear
stackLen.Clear
stackFormat.Clear

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    

    R.Left = R.Right
    R.Right = MaxX
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).R(LinkAreaNumber) = R
        'End If
        If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
        Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = MaxX
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < MaxX) And (R.Bottom <= R.Top + HeightOf1Line + 2)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                R.Right = MaxX
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).R(LinkAreaNumber) = R
            'End If
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + MaxHeightOf1Line * ParLine
            
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            
            If DrawTextAtMaxY <> 0 Then
                If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                    R.Left = R.Left - ExtraMarginLeft
                    MaxX = MaxX + ExtraMarginRight
                    LeftMargin = LeftMargin - ExtraMarginLeft
                    If Align = vbCenter And ExtraMarginRight > 0 Then
                        'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                        ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                        LeftMargin = LeftMargin + ExtraMarginLeft
                        R.Right = LeftMargin
                        R.Top = DrawTextAtMaxYStart
                        ExtraMarginRight = 0
                    Else
                        R.Right = LeftMargin + NextMargin
                        R.Top = DrawTextAtMaxY + 3
                        ExtraMarginLeft = 0: ExtraMarginRight = 0
                        DrawTextAtMaxY = 0
                    End If
                    GoTo DoJodWithLink
                End If
            End If
            
            R.Left = LeftMargin + NextMargin
            If printThis <> "" Then GoTo hereWithLink
        Else
            If Left(WhatIs, 1) <> "s" Then LastMaxHeightOf1Line = MaxHeightOf1Line
hereWithLink:
            If HasPrintSomething Then
                R.Top = R.Top + TM.tmAscent - MaxHeightOfBaseLine + LastMaxHeightOf1Line * ParLine
                R.Left = LeftMargin + NextMargin
                printThis = LTrim(printThis)
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = MaxX
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            
            If MustMoveActiveLineToBaseLine <> 0 Then GoSub MoveActiveLineToBaseLine
            If Abs(R.Top - (R.Bottom - HeightOf1Line)) <= 2 Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                aTempLong = R.Top
                While printThis <> "" 'R.Top < aTempLong
                    
                    If DrawTextAtMaxY <> 0 Then
                        If R.Top + HeightOf1Line >= DrawTextAtMaxY Then
                            'calculate 1 big rect
                            R.Top = aTempLong
                            LinkAreaNumber = LinkAreaNumber + 1
                            Links(UBound(Links)).R(LinkAreaNumber) = R
                            
                            R.Left = R.Left - ExtraMarginLeft
                            MaxX = MaxX + ExtraMarginRight
                            LeftMargin = LeftMargin - ExtraMarginLeft
                            If Align = vbCenter And ExtraMarginRight > 0 Then
                                'ExtraMarginLeft = MaxX - ExtraMarginRight - LeftMargin + PicWidth + 2 * 3
                                ExtraMarginLeft = MaxX - ExtraMarginRight + PicWidth + 2 * 3 - CellMargin
                                LeftMargin = LeftMargin + ExtraMarginLeft
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxYStart
                                ExtraMarginRight = 0
                            Else
                                R.Right = LeftMargin + NextMargin
                                R.Top = DrawTextAtMaxY + 3
                                ExtraMarginLeft = 0: ExtraMarginRight = 0
                                DrawTextAtMaxY = 0
                            End If
                            GoTo DoJodWithLink
                        End If
                    End If
                                        
                    R.Bottom = R.Top + HeightOf1Line
                    'print last text
                    Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                    R.Top = R.Top + HeightOf1Line * ParLine
                    LastprintThis = printThis
                    printThis = Mid(printThis, textParams.uiLengthDrawn + 1)
                Wend
                R.Top = R.Top - HeightOf1Line * ParLine
                
                'If WasLink Then
                    'calculate 1 big rect
                    R.Top = aTempLong
                    R.Bottom = R.Bottom - HeightOf1Line * ParLine
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                    R.Bottom = R.Bottom + HeightOf1Line * ParLine
                    R.Top = R.Bottom - HeightOf1Line
                    
                    'calculate last rect
                    R.Right = MaxX
                    Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
            End If
            MaxHeightOf1Line = HeightOf1Line: MaxHeightOfBaseLine = TM.tmAscent
            LastMaxHeightOf1Line = MaxHeightOf1Line
        End If
    End If

Return

MoveActiveLineToBaseLine:

    'BitBlt Picture1.hdc, LeftMargin + NextMargin, R.Top + MustMoveActiveLineToBaseLine - MaxHeightOfBaseLine, R.Right - LeftMargin - NextMargin, R.Bottom - R.Top, _
            Picture1.hdc, LeftMargin + NextMargin, R.Top, vbSrcCopy
    BitBlt Picture1.hdc, LeftMargin + NextMargin, R.Top - MustMoveActiveLineToBaseLine + MaxHeightOfBaseLine, R.Right - LeftMargin - NextMargin, LastMaxHeightOf1Line, _
            Picture1.hdc, LeftMargin + NextMargin, R.Top + TM.tmAscent - MaxHeightOfBaseLine, vbSrcCopy
            
    Picture1.CurrentX = LeftMargin + NextMargin
    Picture1.CurrentY = R.Top + TM.tmAscent - MaxHeightOfBaseLine + 1
    aTempLong2 = Picture1.ForeColor
    PicTemp.Tag = PicTemp.DrawStyle
    PicTemp.DrawStyle = 0
    Picture1.ForeColor = Picture1.BackColor
    'Picture1.Line -(R.Right, R.Top + MustMoveActiveLineToBaseLine - MaxHeightOfBaseLine - 1), , BF
    Picture1.Line -(R.Left, R.Top - MustMoveActiveLineToBaseLine + MaxHeightOfBaseLine - 1), , BF
    Picture1.ForeColor = aTempLong2
    PicTemp.DrawStyle = Val(PicTemp.Tag)
    
    MustMoveActiveLineToBaseLine = 0
Return

End Sub

Public Sub LocateTextAtXY(TextPosition As Long, x As Long, y As Long, TheActiveFont As StdFont, LeftMargin As Long, NextMargin As Long)
Dim aStr As String, printThis As String, i As Long, Lines As Long
Dim HeightOf1Line As Long, cHeight As Long, R As RECT, pl As Long
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single, aValue As Long
Dim LocateY As Boolean, TextFound As String, metr As Long
Dim CheckThis As String, ExtraI As Long

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

'Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

SplitFormat aStr

Picture1.FontBold = False
Picture1.FontItalic = False
Picture1.FontUnderline = False
OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
If Not UserControl.Enabled Then
    Picture1.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
'Picture1.Cls
For i = 1 To stackLen.stackLevel
    TheFrom = stackFrom.pop
    TheLen = stackLen.pop
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.pop
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin
            R.Right = R.Left
            If y >= R.Top And y <= R.Top + MaxHeightOf1Line Then
                If x <= R.Left Then
                    LocateY = True
                    TextFound = printThis
                    GoSub FindExactText
                    Exit Sub
                ElseIf x >= R.Left And x <= R.Right Then
                    LocateY = False
                    TextFound = printThis
                    GoSub FindExactText
                    Exit Sub
                Else
                    LocateY = True
                End If
            ElseIf LocateY And y <= R.Top Then
                TextFound = printThis
                GoSub FindExactText
                Exit Sub
            End If
            
            NextMargin = 0
            HasPrintSomething = False
            MaxHeightOf1Line = HeightOf1Line
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                R.Left = LeftMarginNext
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
        Case "e": ' print a line
            Picture1.CurrentY = R.Top
            Picture1.CurrentX = R.Left
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(Picture1.Width, Picture1.CurrentY)
                R.Top = R.Top + 2
            Else
                Picture1.Line -(Picture1.Width, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
                R.Top = R.Top + Val(Mid(WhatIs, 2)) + 1
            End If
            R.Left = LeftMargin
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            R.Top = R.Top + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            aSingle = R.Top
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                Picture1.CurrentX = R.Left + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                Picture1.CurrentX = R.Left + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            Picture1.DrawWidth = 1
            R.Left = Picture1.CurrentX + 7
            R.Right = R.Left
            R.Top = aSingle
    End Select
    If printThis = "" Then GoTo Nexti
    
    ExtraI = 0
    R.Left = R.Right
    R.Right = Picture1.Width
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line) Then
        CheckThis = printThis
        GoSub CheckIt
        'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + HeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            CheckThis = LastprintThis
            GoSub CheckIt
            ExtraI = Len(LastprintThis)
            'Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo here
        Else
here:
            If HasPrintSomething Then
                R.Top = R.Top + MaxHeightOf1Line
                R.Left = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = Picture1.Width
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                CheckThis = printThis
                GoSub CheckIt
                'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                aValue = R.Bottom - HeightOf1Line
                R.Bottom = R.Top
                NewprintThis = printThis
                Do While R.Bottom < aValue
                    R.Bottom = R.Bottom + HeightOf1Line
                    Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    CheckThis = printThis
                    GoSub CheckIt
                    ExtraI = ExtraI + Len(printThis)
                    printThis = Mid(NewprintThis, textParams.uiLengthDrawn + 1)
                Loop
                
                
                
                R.Left = LeftMargin + NextMargin
                R.Top = R.Bottom
                R.Bottom = R.Bottom + HeightOf1Line
                CheckThis = printThis
                GoSub CheckIt
                
            End If
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext
Next i

If Picture1.CurrentY + HeightOf1Line > R.Bottom Then
    Picture1.Height = Picture1.CurrentY + HeightOf1Line
Else
    Picture1.Height = R.Bottom
End If

If Picture1.Height > UserControl.ScaleHeight Then
    VScroll1.Enabled = True
    VScroll1.Max = Picture1.Height - Picture2.Height + IIf(Picture2.BorderStyle = 0, 0, 4)
    VScroll1.SmallChange = HeightOf1Line
    VScroll1.LargeChange = UserControl.ScaleHeight - IIf(Picture2.BorderStyle = 0, 5, 7)
    If VScroll1.Value >= VScroll1.Min And VScroll1.Value <= VScroll1.Max Then
        VScroll1_Change
    Else
        VScroll1.Value = 0
        Picture1.Top = 0
    End If
Else
    VScroll1.Enabled = False
    Picture1.Top = 0
End If

Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    R.Left = R.Right
    R.Right = Picture1.Width
    R.Bottom = R.Top + MaxHeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + MaxHeightOf1Line) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).R(LinkAreaNumber) = R
        'End If
        'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + MaxHeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + MaxHeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                'R.Left = Picture1.CurrentX
                'R.Top = Picture1.CurrentY
                R.Right = Picture1.Width
                'R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).R(LinkAreaNumber) = R
            'End If
            'Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo hereWithLink
        Else
hereWithLink:
            If HasPrintSomething Then
                R.Top = R.Top + MaxHeightOf1Line
                R.Left = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = Picture1.Width
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                'Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                R.Bottom = R.Bottom - HeightOf1Line
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                'Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                R.Left = LeftMargin + NextMargin
                R.Top = R.Bottom
                R.Bottom = R.Top + MaxHeightOf1Line
                'Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_EDITCONTROL)
                'If WasLink Then
                    'Calculate rect
                    'R.Left = Picture1.CurrentX
                    'R.Top = Picture1.CurrentY
                    R.Right = Picture1.Width
                    'R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                    Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
            End If
        End If
    End If

Return

CheckIt:
        If y >= R.Top And y <= R.Bottom Then
            If x <= R.Left Then
                TextFound = CheckThis
                GoSub FindExactText
                Exit Sub
            ElseIf x >= R.Left And x <= R.Right Then
                LocateY = False
                Debug.Print "ElseIf x >= r.Left And x <= r.Right Then"
                Debug.Print y, R.Top, R.Bottom
                TextFound = CheckThis
                GoSub FindExactText
                Exit Sub
            Else
                LocateY = True
            End If
        ElseIf LocateY And y <= R.Top Then
            TextFound = CheckThis
            GoSub FindExactText
            Exit Sub
        End If
Return

FindExactText:
If LocateY Then
    'nothing
    TextPosition = TheFrom + Len(TextFound)
    'MsgBox Mid(aStr, TextPosition, 20)
    x = LeftMargin
    y = R.Top '+ MaxHeightOf1Line
    'GoTo there
ElseIf x <= R.Left Then
    'nothing
    x = LeftMargin
    y = R.Top
    TextPosition = TheFrom + Len(printThis)
    Debug.Print "at left - "; Mid(aStr, TheFrom, 15); " - "; TextFound
    'GoTo there
ElseIf x >= R.Left And x <= R.Right Then
there:
    Do While TextFound = ""
        TheFrom = stackFrom.pop
        TheLen = stackLen.pop
        printThis = Mid(aStr, TheFrom, TheLen)
        If UseFormat = 0 Then
            printThis = Replace(printThis, "[[", "[")
            printThis = Replace(printThis, "]]", "]")
        Else
            printThis = Replace(printThis, "//", "/")
        End If
        TextFound = printThis
    Loop
    
    i = 1
    NewprintThis = Left(TextFound, i)
    Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    Do While x >= R.Right
        i = i + 1
        NewprintThis = Left(TextFound, i)
        If i = Len(TextFound) Then Exit Do
        Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_CALCRECT)
    Loop
    NewprintThis = Left(TextFound, i - 1)
    Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_CALCRECT)
    
    x = R.Right
    y = R.Top
    TextFound = Mid(TextFound, i)
    
    'count the [ and add this number
    pl = InStr(NewprintThis, "[")
    While pl <> 0
        metr = metr + 1
        pl = InStr(pl + 1, NewprintThis, "[")
    Wend
    'count the ] and add this number
    pl = InStr(NewprintThis, "]")
    While pl <> 0
        metr = metr + 1
        pl = InStr(pl + 1, NewprintThis, "]")
    Wend
    
    TextPosition = TheFrom + i + ExtraI - 1 + metr
    
    Set TheActiveFont = Picture1.Font
    
    'DrawTheTextFromPoint x, y, TextPosition, TheActiveFont, LeftMargin, NextMargin
End If

    Set TheActiveFont = Picture1.Font

Return

End Sub


Private Sub DrawTheTextFromToPoint(x As Long, y As Long, FromPosition As Long, ToPosition As Long, TheFont As StdFont, LeftMargin As Long, NextMargin As Long)
Dim aStr As String, printThis As String, i As Long, Lines As Long
Dim HeightOf1Line As Long, cHeight As Long, R As RECT, pl As Long
Dim NewprintThis As String, LastprintThis As String, WhatIs As String
Dim OldFontSize As Single, MaxHeightOf1Line As Long
Dim HasPrintSomething As Boolean, textParams As DRAWTEXTPARAMS, LinkAreaNumber As Long
Dim OldForeColor As Long, LeftMarginNext As Long, WasLink As Boolean
Dim TheFrom As Long, TheLen As Long, aSingle As Single

On Error GoTo ErrHandle

ReDim Links(0) As LinkAreaType

Picture1.Height = MaxHeight

textParams.cbSize = Len(textParams)

If UseFormat = 0 Then
    aStr = Replace(TheText, vbNewLine, "[l]")
Else
    aStr = Replace(TheText, vbNewLine, "/l ")
End If

If ToPosition = 0 Then
    aStr = Mid(aStr, FromPosition)
Else
    If Mid(aStr, ToPosition - 1, 2) = "[[" Then
        aStr = Mid(aStr, FromPosition, ToPosition - FromPosition + 1)
    Else
        aStr = Mid(aStr, FromPosition, ToPosition - FromPosition + 0)
    End If
End If
Debug.Print aStr
SplitFormat aStr

'Picture1.FontBold = False
'Picture1.FontItalic = False
'Picture1.FontUnderline = False
'OldFontSize = Picture1.FontSize
OldForeColor = Picture1.ForeColor
If Not UserControl.Enabled Then
    'Picture1.ForeColor = &H80000011
End If

HeightOf1Line = Picture1.TextHeight("astr")
MaxHeightOf1Line = HeightOf1Line
'Picture1.Cls
R.Left = x: R.Right = x
R.Top = y

SetBkColor Picture1.hdc, vbBlack
SetBkMode Picture1.hdc, OPAQUE
Picture1.ForeColor = vbWhite

LeftMarginNext = LeftMargin

Set Picture1.Font = TheFont
Picture1.Font.Bold = FontWasBold
Picture1.Font.Italic = FontWasItalic
Picture1.Font.Underline = FontWasUnderline
Picture1.Font.Size = FontWasSize

For i = 1 To stackLen.stackLevel
    TheFrom = stackFrom.pop
    TheLen = stackLen.pop
    printThis = Mid(aStr, TheFrom, TheLen)
    If UseFormat = 0 Then
        printThis = Replace(printThis, "[[", "[")
        printThis = Replace(printThis, "]]", "]")
    Else
        printThis = Replace(printThis, "//", "/")
    End If
    
    WhatIs = stackFormat.pop
    Select Case Left(WhatIs, 1)
        Case "":
        Case "b": Picture1.FontBold = Not Picture1.FontBold
        Case "i": Picture1.FontItalic = Not Picture1.FontItalic
        Case "u": Picture1.FontUnderline = Not Picture1.FontUnderline
        Case "l":
            'Picture1.Print ""
            'Picture1.CurrentX = LeftMargin
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin
            R.Right = R.Left
            
            NextMargin = 0
            HasPrintSomething = False
            MaxHeightOf1Line = HeightOf1Line
        Case "m":
            If Mid(WhatIs, 2, 1) = "+" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 3))
            ElseIf Mid(WhatIs, 2, 1) = "-" Then
                LeftMarginNext = LeftMargin + Val(Mid(WhatIs, 2))
            Else
                LeftMarginNext = Val(Mid(WhatIs, 2))
            End If
            NextMargin = 0
            If i = 1 Then
                'Picture1.CurrentX = LeftMarginNext
                R.Left = LeftMarginNext
                R.Right = LeftMarginNext
            End If
        Case "n":
            NextMargin = Val(Mid(WhatIs, 2))
        Case "s":
            Picture1.FontSize = Val(Mid(WhatIs, 2))
            HeightOf1Line = Picture1.TextHeight("astr")
            If MaxHeightOf1Line < HeightOf1Line Then MaxHeightOf1Line = HeightOf1Line
        Case "e": ' print a line
            Picture1.CurrentY = R.Top
            Picture1.CurrentX = R.Left
            If Val(Mid(WhatIs, 2)) < 2 Then
                Picture1.Line -(Picture1.Width, Picture1.CurrentY)
                R.Top = R.Top + 2
            Else
                Picture1.Line -(Picture1.Width, Picture1.CurrentY + Val(Mid(WhatIs, 2)) - 1), , BF
                R.Top = R.Top + Val(Mid(WhatIs, 2)) + 1
            End If
            'Picture1.CurrentY = Picture1.CurrentY + 2
            'Picture1.CurrentX = LeftMargin
            R.Left = LeftMargin
        Case "w": 'web link
            WasLink = Not WasLink
            If WasLink Then
                ReDim Preserve Links(UBound(Links) + 1) As LinkAreaType
                Links(UBound(Links)).Link = printThis
                'If UserControl.Enabled Then Picture1.ForeColor = vbBlue
                GoSub DoJodWithLink
                GoTo Nexti
            Else
                'If UserControl.Enabled Then Picture1.ForeColor = OldForeColor
                LinkAreaNumber = 0
            End If
        Case "y": ' Change the CurrentY
            'Picture1.CurrentY = Picture1.CurrentY + Val(Mid(WhatIs, 2))
            R.Top = R.Top + Val(Mid(WhatIs, 2))
        Case "c": ' Color
            'If UserControl.Enabled Then Picture1.ForeColor = Val(Mid(WhatIs, 2))
        Case "t": 'Bullet
            'aSingle = Picture1.CurrentY
            aSingle = R.Top
            If Mid(WhatIs, 2, 1) = "2" Then
                Picture1.CurrentY = aSingle + HeightOf1Line / 2.5
                'Picture1.CurrentX = Picture1.CurrentX + 5
                Picture1.CurrentX = R.Left + 5
                Picture1.DrawWidth = 3
                Picture1.Line -(Picture1.CurrentX + HeightOf1Line / 5, Picture1.CurrentY + HeightOf1Line / 5), , BF
            Else
                Picture1.CurrentY = aSingle + HeightOf1Line / 2
                'Picture1.CurrentX = Picture1.CurrentX + 6
                Picture1.CurrentX = R.Left + 6
                Picture1.DrawWidth = 5
                Picture1.Circle (Picture1.CurrentX, Picture1.CurrentY), HeightOf1Line \ 13
            End If
            Picture1.DrawWidth = 1
            'Picture1.CurrentX = Picture1.CurrentX + 7
            'Picture1.CurrentY = aSingle
            R.Left = Picture1.CurrentX + 7
            R.Right = R.Left
            R.Top = aSingle
    End Select
    If printThis = "" Then GoTo Nexti
    
    'R.Left = Picture1.CurrentX
    'R.Top = Picture1.CurrentY
    R.Left = R.Right
    R.Right = Picture1.Width
    R.Bottom = R.Top + HeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + HeightOf1Line) Then
        'Picture1.Print printThis;
        Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + MaxHeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + MaxHeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'Picture1.Print LastprintThis;
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            'Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            'Picture1.CurrentX = LeftMargin + NextMargin
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo here
        Else
here:
            If HasPrintSomething Then
                'Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                'Picture1.CurrentX = LeftMargin + NextMargin
                R.Top = R.Top + MaxHeightOf1Line
                R.Left = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = Picture1.Width
            'R.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                'Picture1.Print printThis;
                Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                R.Bottom = R.Bottom - HeightOf1Line
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                'Picture1.CurrentX = LeftMargin + NextMargin
                'Picture1.CurrentY = R.Bottom
                R.Left = LeftMargin + NextMargin
                R.Top = R.Bottom
                R.Bottom = R.Bottom + HeightOf1Line
                'Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_EDITCONTROL)
            End If
        End If
    End If
Nexti:
    LeftMargin = LeftMarginNext
Next i


'Picture1.FontSize = OldFontSize
Picture1.ForeColor = OldForeColor
If Not UserControl.Enabled Then
    VScroll1.Enabled = False
End If

Exit Sub

ErrHandle:
Beep
MsgBox Error
Resume Next
Exit Sub

DoJodWithLink:
    If printThis = "" Then Return
    
    'R.Left = Picture1.CurrentX
    'R.Top = Picture1.CurrentY
    R.Left = R.Right
    R.Right = Picture1.Width
    R.Bottom = R.Top + MaxHeightOf1Line
    cHeight = DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
    If (R.Right < Picture1.Width) And (R.Bottom = R.Top + MaxHeightOf1Line) Then
        'If WasLink Then
            LinkAreaNumber = LinkAreaNumber + 1
            Links(UBound(Links)).R(LinkAreaNumber) = R
        'End If
        'Picture1.Print printThis;
        Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
        HasPrintSomething = True
    Else
        LastprintThis = ""
        pl = 1
        While Mid(printThis, pl, 1) = " "
            pl = pl + 1
        Wend
        pl = InStr(pl, printThis, " ")
        Do While pl <> 0
            NewprintThis = Left(printThis, pl - 1)
            R.Right = Picture1.Width
            R.Bottom = R.Top + MaxHeightOf1Line
            Call DrawText(Picture1.hdc, NewprintThis, Len(NewprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
            If Not ((R.Right < Picture1.Width) And (R.Bottom = R.Top + MaxHeightOf1Line)) Then
                Exit Do
            End If
            LastprintThis = NewprintThis
            pl = InStr(pl + 1, printThis, " ")
        Loop
        If LastprintThis <> "" Then
            'If WasLink Then
                'Calculate rect
                'R.Left = Picture1.CurrentX
                'R.Top = Picture1.CurrentY
                R.Right = Picture1.Width
                'R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                
                LinkAreaNumber = LinkAreaNumber + 1
                Links(UBound(Links)).R(LinkAreaNumber) = R
            'End If
            'Picture1.Print LastprintThis;
            Call DrawText(Picture1.hdc, LastprintThis, Len(LastprintThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            HasPrintSomething = False
            'Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
            'Picture1.CurrentX = LeftMargin + NextMargin
            R.Top = R.Top + MaxHeightOf1Line
            R.Left = LeftMargin + NextMargin
            printThis = LTrim(Mid(printThis, Len(LastprintThis) + 1))
            If printThis <> "" Then GoTo hereWithLink
        Else
hereWithLink:
            If HasPrintSomething Then
                'Picture1.CurrentY = Picture1.CurrentY + MaxHeightOf1Line
                'Picture1.CurrentX = LeftMargin + NextMargin
                R.Top = R.Top + MaxHeightOf1Line
                R.Left = LeftMargin + NextMargin
            Else
                HasPrintSomething = True
            End If
            R.Left = LeftMargin + NextMargin
            R.Right = Picture1.Width
            'R.Top = Picture1.CurrentY
            Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                        
            If R.Top = R.Bottom - HeightOf1Line Then
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                'Picture1.Print printThis;
                Call DrawText(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL)
            Else
                R.Bottom = R.Bottom - HeightOf1Line
                'If WasLink Then
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
                Call DrawTextEx(Picture1.hdc, printThis, Len(printThis), R, DT_WORDBREAK Or DT_EDITCONTROL, textParams)
                'Picture1.CurrentX = LeftMargin + NextMargin
                'Picture1.CurrentY = R.Bottom
                R.Left = LeftMargin + NextMargin
                R.Top = R.Bottom
                R.Bottom = R.Top + MaxHeightOf1Line
                'Picture1.Print Mid(printThis, textParams.uiLengthDrawn + 1);
                Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_EDITCONTROL)
                'If WasLink Then
                    'Calculate rect
                    'R.Left = Picture1.CurrentX
                    'R.Top = Picture1.CurrentY
                    R.Right = Picture1.Width
                    'R.Bottom = Picture1.CurrentY + MaxHeightOf1Line
                    Call DrawText(Picture1.hdc, Mid(printThis, textParams.uiLengthDrawn + 1), Len(Mid(printThis, textParams.uiLengthDrawn + 1)), R, DT_WORDBREAK Or DT_CALCRECT Or DT_EDITCONTROL)
                    
                    LinkAreaNumber = LinkAreaNumber + 1
                    Links(UBound(Links)).R(LinkAreaNumber) = R
                'End If
            End If
        End If
    End If

Return

End Sub


