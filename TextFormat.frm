VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Draw formatted text (v. 2.0)"
   ClientHeight    =   6825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   455
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   502
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5280
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1635
   End
   Begin Project1.TextFormat TextFormat1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5895
      _ExtentX        =   9869
      _ExtentY        =   4366
      FormatMethod    =   0
      PrintAreaMaxHeight=   7000
      PointerForLink  =   "TextFormat.frx":0000
      RightMargin     =   5
      AutoRedraw      =   0   'False
      BackColor       =   15836555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "TextFormat.frx":001C
      Top             =   3720
      Width           =   5940
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   60
      TabIndex        =   3
      Top             =   5280
      Width           =   1620
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Has border"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   3180
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4260
      TabIndex        =   6
      Top             =   3120
      Width           =   1635
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim aPic As IPictureDisp

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
Private Declare Function GetDC Lib "user32" ( _
     ByVal hwnd As Long) As Long
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

Dim TheX As Single, TheY As Single

Private Sub Check1_Click()
TextFormat1.BorderStyle = IIf(Check1.Value = vbChecked, FixedSingle, None)
End Sub

Private Sub Check2_Click()
TextFormat1.Enabled = Check2.Value = vbChecked
End Sub


Private Sub Command2_Click()
TextFormat1.Text = Text1.Text
End Sub


Private Sub Form_Load()

Dim aTxt As String
aTxt = "Nothing [b] [s12]A bold [s8]" & "string [i]bold italic[/i] bold only[/b] [u] underline [/u] ok format "
aTxt = aTxt & vbNewLine & vbNewLine & "1  " & aTxt & "2 " & aTxt & aTxt & aTxt & aTxt & aTxt & " END" & vbNewLine & vbNewLine & "Some things more to print here, just to check this control. Very good!"

'UserControl11.Text = aTxt


TextFormat1.Fonts.Add "Arial"
TextFormat1.Fonts.Add "Times New Roman"

Dim AppPath As String

AppPath = App.Path
If Right(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

Set aPic = LoadPicture(AppPath & "FormatText.jpg")

'I am using the "[g|   " because if I want to copy the text from Text1
'first I replace the "[g|   " with "[g|ihxxx"
'If you want check the Sub Text1_KeyDown
Text1.Text = Replace(Text1.Text, "[g|ihxxx", "[g|   " & Trim(aPic.Handle))
TextFormat1.Text = Text1.Text

TextFormat1.AutoRedraw = True


End Sub

Private Sub Form_Resize()

    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Height < 2900 Then Me.Height = 2900
    If Me.Width < 3200 Then Me.Width = 3200
    
    TextFormat1.Width = Me.ScaleWidth - 2 * TextFormat1.Left
    TextFormat1.Height = (Me.ScaleHeight - 2 * TextFormat1.Top) / 2
    Check2.Top = 2 * TextFormat1.Top + TextFormat1.Height
    Check1.Top = 2 * TextFormat1.Top + TextFormat1.Height
    
    Label1.Top = Check2.Top
    
    Label1.Left = TextFormat1.Left + TextFormat1.Width - Label1.Width
        
    Text1.Top = Check1.Top + Check1.Height + TextFormat1.Top
    Text1.Width = TextFormat1.Width
    Text1.Height = Me.ScaleHeight - TextFormat1.Height - 5 * TextFormat1.Top - Command2.Height - Check2.Height
    
    Command2.Top = Me.ScaleHeight - Command2.Height - TextFormat1.Top
        
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 65 Then
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
ElseIf Shift = 2 And KeyCode = 45 Then ' Ctrl - Ins
    'ignore this code
    'it is used only to this program
    'just when I want to copy some text
    'replaces the real handle of the loaded bitmap (at ryn time)
    'with the code "[g|ihxxx"
    Dim pl As Long, pl2 As Long, aTxt As String
    aTxt = Text1.SelText
    If aTxt = "" Then Exit Sub
    
    KeyCode = 0
    pl = InStr(aTxt, "[g|   ")
    pl2 = InStr(pl + 3, aTxt, "|")
    If pl2 = 0 Then pl2 = InStr(pl + 1, aTxt, "]")
    If pl <> 0 And pl2 <> 0 Then
        aTxt = Left(aTxt, pl + 2) & "ihxxx" & Mid(aTxt, pl2)
    End If
    Clipboard.Clear
    Clipboard.SetText aTxt
End If
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 45 Then ' Ctrl - Ins
    KeyCode = 0
End If
End Sub

Private Sub TextFormat1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TheX = x
TheY = y
End Sub

Private Sub TextFormat1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Label1 = x & ", " & y
End Sub


Private Sub DrawFormatText()
   Dim graphics As Long, brush As Long, pen As Long
   Dim fontFam As Long, curFont As Long, strFormat As Long
   Dim rcLayout As RECTF   ' Designates the string drawing bounds
   Dim str As String
   
    Dim hGDI As Long
    Dim uInput As GdiplusStartupInput
    On Error Resume Next
    uInput.GdiplusVersion = 1
    Call GdiplusStartup(hGDI, uInput)
   
   ' Initializations
   Call GdipCreateFromHDC(Me.hdc, graphics) ' Initialize the graphics class - required for all drawing
   Call GdipCreateSolidFill(vbBlue, brush)    ' Create a brush to draw the text with
   ' Create a font family object to allow use to create a font
   ' We have no font collection here, so pass a NULL for that parameter
   Call GdipCreateFontFamilyFromName(StrConv("Arial", vbUnicode), 0, fontFam)
   ' Create the font from the specified font family name
   ' >> Note that we have changed the drawing Unit from pixels to points!!
   Call GdipCreateFont(fontFam, 12, FontStyleUnderline + FontStyleBoldItalic, UnitPoint, curFont)
   ' Create the StringFormat object
   ' We can pass NULL for the flags and language id if we want
   Call GdipCreateStringFormat(0, 0, strFormat)
   
   ' Set up the drawing area boundary
   rcLayout.Left = 1
   rcLayout.Top = 1
   rcLayout.Right = 120
   rcLayout.Bottom = 140
   
   ' Center-justify each line of text
   Call GdipSetStringFormatAlign(strFormat, StringAlignmentCenter)
   
   ' Center the block of text (top to bottom) in the rectangle.
   Call GdipSetStringFormatLineAlign(strFormat, StringAlignmentCenter)
   
   ' Draw the string within the boundary
   str = StrConv("Use StringFormat and RectF objects to center text in a rectangle.", vbUnicode)
   If GdipDrawString(graphics, str, -1, curFont, rcLayout, strFormat, brush) <> 0 Then
    MsgBox "error"
   End If
   
   ' Create a pen and draw the boundary around the text
   Call GdipCreatePen1(vbBlack, 1, UnitPixel, pen)
   Call GdipDrawRectangles(graphics, pen, rcLayout, 1)
      
   
   ' Cleanup
   Call GdipDeletePen(pen)
   Call GdipDeleteStringFormat(strFormat)
   Call GdipDeleteFont(curFont)     ' Delete the font object
   Call GdipDeleteFontFamily(fontFam)  ' Delete the font family object
   Call GdipDeleteBrush(brush)
   Call GdipDeleteGraphics(graphics)
     
    UnloadGDIplus hGDI
    
End Sub


